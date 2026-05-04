'use strict';

/**
 * wordConverter.js
 * Drives Microsoft Word (COM automation) via a PowerShell child process.
 * All conversion is delegated to scripts/word_convert.ps1; this module
 * handles spawning, manifest I/O, and stdout parsing.
 */

const { spawn }  = require('child_process');
const path       = require('path');
const fs         = require('fs');
const os         = require('os');
const crypto     = require('crypto');

// Resolve the PS script path for both development and electron-builder packaged modes.
// electron-builder moves files listed in asarUnpack to:
//   <app>/resources/app.asar.unpacked/<relative-path>
// process.resourcesPath exists in both modes; we probe the unpacked path first.
function resolveScriptPath() {
  if (process.resourcesPath) {
    const unpacked = path.join(
      process.resourcesPath, 'app.asar.unpacked', 'scripts', 'word_convert.ps1',
    );
    if (fs.existsSync(unpacked)) return unpacked;
  }
  return path.join(__dirname, 'scripts', 'word_convert.ps1');
}

// ---------------------------------------------------------------------------
// Dependency check
// Queries the registry — no Word process is launched, stays fast.
// ---------------------------------------------------------------------------

function checkWord() {
  return new Promise(resolve => {
    const cmd = `
      try {
        $null = Get-ItemProperty 'HKLM:\\SOFTWARE\\Classes\\Word.Application\\CLSID' -ErrorAction Stop
        $ver  = (Get-ChildItem 'HKLM:\\SOFTWARE\\Microsoft\\Office' -ErrorAction SilentlyContinue) |
                Where-Object { $_.PSChildName -match '^\\d+\\.\\d+$' } |
                Sort-Object   { [double]$_.PSChildName } -Descending   |
                Select-Object -First 1 -ExpandProperty PSChildName
        Write-Output "OK:$(if ($ver) { 'Office ' + $ver } else { 'Installed' })"
      } catch {
        Write-Output 'FAIL'
      }
    `;

    const ps = spawn('powershell.exe', [
      '-NoProfile', '-NonInteractive', '-Command', cmd,
    ]);

    let out = '';
    ps.stdout.on('data', d => { out += d.toString(); });
    ps.on('close', () => {
      const line = out.trim();
      resolve(line.startsWith('OK:')
        ? { ok: true,  version: line.slice(3) }
        : { ok: false, version: '' });
    });
    ps.on('error', () => resolve({ ok: false, version: '' }));
  });
}

// ---------------------------------------------------------------------------
// Batch conversion
//
// files      — Array<{ name: string, path: string }>
// outputDir  — absolute path to destination folder
// direction  — 'docx-to-pdf' | 'pdf-to-docx'
// onProgress — (index, fileName, 'converting'|'done'|'error', errMsg?) => void
//
// Returns Promise<Array<{ name, success, outPath, error }>>
//
// A single Word instance is opened by the PS script, processes all files
// sequentially, then is closed — no concurrent COM collisions possible.
// ---------------------------------------------------------------------------

function convertBatch(files, outputDir, direction, onProgress) {
  return new Promise((resolve, reject) => {
    // Build jobs manifest: each entry maps an input path to its output path
    const jobs = files.map(f => {
      const base = path.basename(f.path, path.extname(f.path));
      const ext  = direction === 'docx-to-pdf' ? '.pdf' : '.docx';
      return { name: f.name, input: f.path, output: path.join(outputDir, base + ext) };
    });

    // Write manifest to a temp file so paths with special chars are safe
    const manifestPath = path.join(
      os.tmpdir(),
      `ectd_${crypto.randomBytes(6).toString('hex')}.json`,
    );
    try {
      fs.writeFileSync(manifestPath, JSON.stringify(jobs, null, 0), 'utf8');
    } catch (err) {
      return reject(new Error(`Cannot write manifest: ${err.message}`));
    }

    // Pre-populate results — PS fills in success/error as it runs
    const results = jobs.map(j => ({ name: j.name, success: false, outPath: j.output, error: null }));

    const proc = spawn('powershell.exe', [
      '-NoProfile',
      '-NonInteractive',
      '-ExecutionPolicy', 'Bypass',
      '-File', resolveScriptPath(),
      '-ManifestPath', manifestPath,
      '-Direction',    direction,
    ]);

    // PS prints lines; chunks may split across line boundaries → use a buffer
    let lineBuffer = '';
    let stderr     = '';

    proc.stdout.on('data', chunk => {
      lineBuffer += chunk.toString();
      const lines = lineBuffer.split('\n');
      lineBuffer  = lines.pop(); // keep the incomplete trailing fragment
      lines.forEach(ln => parseLine(ln.trim(), jobs, results, onProgress));
    });

    proc.stderr.on('data', d => { stderr += d.toString(); });

    proc.on('close', code => {
      if (lineBuffer.trim()) parseLine(lineBuffer.trim(), jobs, results, onProgress);
      try { fs.unlinkSync(manifestPath); } catch (_) { /* best-effort */ }

      // If the process itself crashed and nothing was recorded, surface the error
      if (code !== 0 && results.every(r => !r.success && r.error === null)) {
        return reject(new Error(
          `Word automation process exited ${code}: ${stderr.trim().slice(0, 400)}`,
        ));
      }
      resolve(results);
    });

    proc.on('error', err => {
      try { fs.unlinkSync(manifestPath); } catch (_) { /* best-effort */ }
      reject(new Error(`Cannot spawn PowerShell: ${err.message}`));
    });
  });
}

// ---------------------------------------------------------------------------
// Stdout line protocol
//
//   STARTING:<index>
//   PROGRESS:<index>:done
//   PROGRESS:<index>:error:<message>
// ---------------------------------------------------------------------------

function parseLine(line, jobs, results, onProgress) {
  if (!line) return;

  let m;

  // Word started processing file at <index>
  m = line.match(/^STARTING:(\d+)$/);
  if (m) {
    const idx = parseInt(m[1], 10);
    if (jobs[idx] && onProgress) onProgress(idx, jobs[idx].name, 'converting');
    return;
  }

  // File at <index> converted successfully
  m = line.match(/^PROGRESS:(\d+):done$/);
  if (m) {
    const idx = parseInt(m[1], 10);
    results[idx].success = true;
    if (jobs[idx] && onProgress) onProgress(idx, jobs[idx].name, 'done');
    return;
  }

  // File at <index> failed — optional message follows the second colon
  m = line.match(/^PROGRESS:(\d+):error(?::(.*))?$/);
  if (m) {
    const idx = parseInt(m[1], 10);
    const msg = (m[2] || 'Conversion failed').trim();
    results[idx].success = false;
    results[idx].error   = msg;
    if (jobs[idx] && onProgress) onProgress(idx, jobs[idx].name, 'error', msg);
  }
}

module.exports = { checkWord, convertBatch };
