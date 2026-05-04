'use strict';

// ---------------------------------------------------------------------------
// State
// ---------------------------------------------------------------------------

const state = {
  direction: 'docx-to-pdf',   // 'docx-to-pdf' | 'pdf-to-docx'
  files: [],                   // Array<{ id, name, path, size, status }>
  outputDir: null,
  converting: false,
  lastResults: null,
};

let removeProgressListener = null;

// ---------------------------------------------------------------------------
// DOM refs
// ---------------------------------------------------------------------------

const depsBanner    = document.getElementById('depsBanner');
const depsMessage   = document.getElementById('depsMessage');
const wordBadge     = document.getElementById('wordBadge');

const btnDocxPdf    = document.getElementById('btnDocxPdf');
const btnPdfDocx    = document.getElementById('btnPdfDocx');

const dropZone      = document.getElementById('dropZone');
const dropHint      = document.getElementById('dropHint');
const fileInput     = document.getElementById('fileInput');

const fileList      = document.getElementById('fileList');
const fileCountEl   = document.getElementById('fileCount');
const clearAllBtn   = document.getElementById('clearAllBtn');

const outputPathText  = document.getElementById('outputPathText');
const selectOutputBtn = document.getElementById('selectOutputBtn');

const convertBtn    = document.getElementById('convertBtn');

const progressSection = document.getElementById('progressSection');
const progressLabel   = document.getElementById('progressLabel');
const progressPct     = document.getElementById('progressPct');
const progressFill    = document.getElementById('progressFill');

const resultsSection  = document.getElementById('resultsSection');
const resultSummary   = document.getElementById('resultSummary');
const resultText      = document.getElementById('resultText');
const openFolderBtn   = document.getElementById('openFolderBtn');
const resultErrors    = document.getElementById('resultErrors');

// ---------------------------------------------------------------------------
// Initialise
// ---------------------------------------------------------------------------

(async function init() {
  const deps = await window.api.checkDeps();
  renderDepsBanner(deps);
})();

// ---------------------------------------------------------------------------
// Dep banner
// ---------------------------------------------------------------------------

function renderDepsBanner(deps) {
  const ok = deps.word.ok;

  wordBadge.textContent = ok ? `✓ ${deps.word.version}` : '✗ Microsoft Word not found';
  wordBadge.className   = `dep-badge ${ok ? 'ok' : 'bad'}`;

  if (!ok) {
    depsMessage.textContent = 'Microsoft Word is required for conversion. Please install Microsoft Office.';
    depsBanner.className    = 'visible error';
  }
  // If Word is found the banner stays hidden
}

// ---------------------------------------------------------------------------
// Direction toggle
// ---------------------------------------------------------------------------

[btnDocxPdf, btnPdfDocx].forEach(btn => {
  btn.addEventListener('click', () => {
    state.direction = btn.dataset.dir;
    btnDocxPdf.classList.toggle('active', state.direction === 'docx-to-pdf');
    btnPdfDocx.classList.toggle('active', state.direction === 'pdf-to-docx');
    // Update drop zone hint and accepted types
    const isDocx = state.direction === 'docx-to-pdf';
    dropHint.textContent = isDocx ? 'Accepts .docx files' : 'Accepts .pdf files';
    fileInput.accept = isDocx ? '.docx' : '.pdf';
    clearFiles();
  });
});

// ---------------------------------------------------------------------------
// Drag & drop
// ---------------------------------------------------------------------------

dropZone.addEventListener('click', () => fileInput.click());

dropZone.addEventListener('dragover', e => {
  e.preventDefault();
  dropZone.classList.add('drag-over');
});

dropZone.addEventListener('dragleave', () => {
  dropZone.classList.remove('drag-over');
});

dropZone.addEventListener('drop', e => {
  e.preventDefault();
  dropZone.classList.remove('drag-over');
  const dropped = Array.from(e.dataTransfer.files);
  addFiles(dropped);
});

fileInput.addEventListener('change', () => {
  addFiles(Array.from(fileInput.files));
  fileInput.value = ''; // reset so same file can be re-added after clearing
});

// ---------------------------------------------------------------------------
// File management
// ---------------------------------------------------------------------------

function addFiles(rawFiles) {
  const ext = state.direction === 'docx-to-pdf' ? '.docx' : '.pdf';
  const valid = rawFiles.filter(f => f.name.toLowerCase().endsWith(ext));

  if (valid.length !== rawFiles.length) {
    const skipped = rawFiles.length - valid.length;
    // Non-blocking visual feedback via brief border flash — no alert
    dropZone.style.borderColor = '#f59e0b';
    setTimeout(() => { dropZone.style.borderColor = ''; }, 800);
    if (valid.length === 0) return;
  }

  valid.forEach(f => {
    // Deduplicate by path
    if (state.files.some(s => s.path === f.path)) return;
    state.files.push({
      id: crypto.randomUUID(),
      name: f.name,
      path: f.path,       // Electron File objects expose .path
      size: f.size,
      status: 'pending',
    });
  });

  renderFileList();
  updateConvertBtn();
}

function removeFile(id) {
  state.files = state.files.filter(f => f.id !== id);
  renderFileList();
  updateConvertBtn();
}

function clearFiles() {
  state.files = [];
  renderFileList();
  updateConvertBtn();
  hideResults();
}

clearAllBtn.addEventListener('click', clearFiles);

// ---------------------------------------------------------------------------
// Render file list
// ---------------------------------------------------------------------------

const FILE_ICONS = { docx: '📄', pdf: '📕' };

function renderFileList() {
  fileCountEl.textContent = state.files.length;
  fileList.innerHTML = '';
  state.files.forEach(f => {
    const ext  = f.name.split('.').pop().toLowerCase();
    const icon = FILE_ICONS[ext] || '📎';
    const size = formatBytes(f.size);

    const item = document.createElement('div');
    item.className = `file-item ${f.status}`;
    item.dataset.id = f.id;

    item.innerHTML = `
      <span class="file-icon">${icon}</span>
      <div class="file-meta">
        <div class="file-name" title="${escHtml(f.name)}">${escHtml(f.name)}</div>
        <div class="file-size">${size}</div>
      </div>
      <span class="file-status ${f.status}">${statusLabel(f.status)}</span>
      <button class="file-remove" data-id="${f.id}" title="Remove">×</button>
    `;

    item.querySelector('.file-remove').addEventListener('click', e => {
      e.stopPropagation();
      if (!state.converting) removeFile(f.id);
    });

    fileList.appendChild(item);
  });
}

function statusLabel(s) {
  return { pending: '—', converting: '⟳ Converting', done: '✓ Done', error: '✗ Failed' }[s] || s;
}

// ---------------------------------------------------------------------------
// Output directory
// ---------------------------------------------------------------------------

selectOutputBtn.addEventListener('click', async () => {
  const chosen = await window.api.openDirDialog();
  if (chosen) {
    state.outputDir = chosen;
    outputPathText.textContent = chosen;
    outputPathText.title = chosen;
    updateConvertBtn();
  }
});

// ---------------------------------------------------------------------------
// Convert button state
// ---------------------------------------------------------------------------

function updateConvertBtn() {
  const canConvert = state.files.length > 0 && state.outputDir && !state.converting;
  convertBtn.disabled = !canConvert;
  convertBtn.textContent = state.files.length > 0
    ? `Convert (${state.files.length} file${state.files.length !== 1 ? 's' : ''})`
    : 'Convert';
}

// ---------------------------------------------------------------------------
// Conversion
// ---------------------------------------------------------------------------

convertBtn.addEventListener('click', async () => {
  if (state.converting || !state.files.length || !state.outputDir) return;

  state.converting = true;
  updateConvertBtn();
  hideResults();
  resetFileStatuses();

  // Show progress bar
  progressSection.classList.add('visible');
  setProgress(0, state.files.length, 'Starting…');

  // Register progress listener
  if (removeProgressListener) removeProgressListener();
  removeProgressListener = window.api.onProgress(handleProgress);

  const payload = state.files.map(f => ({ name: f.name, path: f.path }));
  const results = await window.api.convert(payload, state.outputDir, state.direction);

  // Cleanup
  if (removeProgressListener) { removeProgressListener(); removeProgressListener = null; }

  state.converting = false;
  state.lastResults = results;

  updateConvertBtn();
  showResults(results);
});

function handleProgress({ current, total, fileName, status, percent }) {
  setProgress(current, total, `${status === 'converting' ? 'Converting' : status === 'done' ? 'Done' : 'Error'}: ${fileName}`);

  // Update individual file row
  const targetFile = state.files.find(f => f.name === fileName);
  if (targetFile && status !== 'converting') {
    targetFile.status = status;
    const row = fileList.querySelector(`[data-id="${targetFile.id}"]`);
    if (row) {
      row.className = `file-item ${status}`;
      const statusEl = row.querySelector('.file-status');
      statusEl.className = `file-status ${status}`;
      statusEl.textContent = statusLabel(status);
    }
  }
}

function setProgress(current, total, label) {
  const pct = total > 0 ? Math.round((current / total) * 100) : 0;
  progressFill.style.width = pct + '%';
  progressPct.textContent  = pct + '%';
  progressLabel.textContent = `${label} (${current}/${total})`;

  progressFill.classList.remove('done', 'errors');
  if (current === total) {
    const hasErrors = state.files.some(f => f.status === 'error');
    progressFill.classList.add(hasErrors ? 'errors' : 'done');
  }
}

function resetFileStatuses() {
  state.files.forEach(f => { f.status = 'pending'; });
  renderFileList();
}

// ---------------------------------------------------------------------------
// Results
// ---------------------------------------------------------------------------

function showResults(results) {
  const succeeded = results.filter(r => r.success).length;
  const failed    = results.filter(r => !r.success);

  resultsSection.classList.add('visible');
  resultSummary.className = `result-summary ${failed.length === 0 ? 'all-ok' : 'has-errors'}`;
  resultText.textContent = failed.length === 0
    ? `✓ All ${succeeded} file${succeeded !== 1 ? 's' : ''} converted successfully.`
    : `✓ ${succeeded} succeeded  ✗ ${failed.length} failed`;

  resultErrors.innerHTML = '';
  failed.forEach(r => {
    const el = document.createElement('div');
    el.className = 'result-error-item';
    el.innerHTML = `<span class="err-name">${escHtml(r.name)}</span>: ${escHtml(r.error || 'Unknown error')}`;
    resultErrors.appendChild(el);
  });
}

function hideResults() {
  resultsSection.classList.remove('visible');
  resultErrors.innerHTML = '';
  progressSection.classList.remove('visible');
}

openFolderBtn.addEventListener('click', () => {
  if (state.outputDir) window.api.openFolder(state.outputDir);
});

// ---------------------------------------------------------------------------
// Utilities
// ---------------------------------------------------------------------------

function formatBytes(bytes) {
  if (!bytes) return '—';
  if (bytes < 1024) return bytes + ' B';
  if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(1) + ' KB';
  return (bytes / (1024 * 1024)).toFixed(1) + ' MB';
}

function escHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
