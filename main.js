'use strict';

const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron');
const path = require('path');
const fs   = require('fs');
const wordConverter = require('./wordConverter');

let mainWindow = null;

// ---------------------------------------------------------------------------
// Window
// ---------------------------------------------------------------------------

function createWindow() {
  mainWindow = new BrowserWindow({
    width: 860,
    height: 680,
    minWidth: 680,
    minHeight: 540,
    title: 'Convertisseur PDF/Word',
    backgroundColor: '#f1f5f9',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false,
    },
  });

  mainWindow.loadFile(path.join(__dirname, 'renderer', 'index.html'));
  mainWindow.on('closed', () => { mainWindow = null; });
}

app.whenReady().then(async () => {
  // Run the Word check before the window appears so the user gets a clear
  // native dialog instead of a silent failure inside the app.
  const wordCheck = await wordConverter.checkWord();
  if (!wordCheck.ok) {
    dialog.showErrorBox(
      'Microsoft Word Not Found',
      'eCTD Precision Converter requires Microsoft Word (Office 2013 or later) ' +
      'to be installed on this machine.\n\n' +
      'Please install Microsoft Office and restart the application.\n\n' +
      'The application will open, but file conversions will not work until ' +
      'Word is available.',
    );
  }
  createWindow();
});

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit();
});

app.on('activate', () => {
  if (!mainWindow) createWindow();
});

// ---------------------------------------------------------------------------
// IPC: dependency check
// ---------------------------------------------------------------------------

ipcMain.handle('deps:check', async () => {
  const word = await wordConverter.checkWord();
  return { word };
});

// ---------------------------------------------------------------------------
// IPC: folder dialog
// ---------------------------------------------------------------------------

ipcMain.handle('dialog:openDir', async () => {
  const result = await dialog.showOpenDialog(mainWindow, {
    properties: ['openDirectory', 'createDirectory'],
    title: 'Select Output Folder',
  });
  return result.canceled ? null : result.filePaths[0];
});

// ---------------------------------------------------------------------------
// IPC: open folder in OS file explorer
// ---------------------------------------------------------------------------

ipcMain.handle('shell:openFolder', async (_event, folderPath) => {
  if (fs.existsSync(folderPath)) shell.openPath(folderPath);
});

// ---------------------------------------------------------------------------
// IPC: conversion
// Delegates entirely to wordConverter; bridges its per-file progress
// callbacks into the existing 'convert:progress' IPC event shape.
// ---------------------------------------------------------------------------

ipcMain.handle('convert:start', async (event, { files, outputDir, direction }) => {
  const total = files.length;

  const onProgress = (index, fileName, status) => {
    // 'converting' fires before the file starts  → current = index (not yet done)
    // 'done'/'error' fires after it finishes     → current = index + 1
    const current = status === 'converting' ? index : index + 1;
    event.sender.send('convert:progress', {
      current,
      total,
      fileName,
      status,
      percent: Math.round((current / total) * 100),
    });
  };

  try {
    const results = await wordConverter.convertBatch(files, outputDir, direction, onProgress);
    results.forEach(r => { if (!r.success && r.error) appendErrorLog(r.name, new Error(r.error)); });
    return results;
  } catch (err) {
    appendErrorLog('batch', err);
    // Surface as a single failed result per file so the UI can display it
    return files.map(f => ({ name: f.name, success: false, error: err.message }));
  }
});

// ---------------------------------------------------------------------------
// Error logging
// ---------------------------------------------------------------------------

function appendErrorLog(fileName, err) {
  try {
    const logDir = app.getPath('userData');
    const logPath = path.join(logDir, 'conversion-errors.log');
    const line = `[${new Date().toISOString()}] ${fileName}: ${err.message}\n`;
    fs.appendFileSync(logPath, line, 'utf8');
  } catch (_) { /* best-effort */ }
}
