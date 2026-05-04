'use strict';

const { contextBridge, ipcRenderer } = require('electron');

contextBridge.exposeInMainWorld('api', {

  // Dependency check
  checkDeps: () => ipcRenderer.invoke('deps:check'),

  // Open OS folder-picker dialog → returns path string or null
  openDirDialog: () => ipcRenderer.invoke('dialog:openDir'),

  // Open a folder in the OS file explorer
  openFolder: (folderPath) => ipcRenderer.invoke('shell:openFolder', folderPath),

  // Start conversion batch
  // files: Array<{ name: string, path: string }>
  // outputDir: string
  // direction: 'docx-to-pdf' | 'pdf-to-docx'
  convert: (files, outputDir, direction) =>
    ipcRenderer.invoke('convert:start', { files, outputDir, direction }),

  // Listen for per-file progress updates
  onProgress: (cb) => {
    const handler = (_event, data) => cb(data);
    ipcRenderer.on('convert:progress', handler);
    return () => ipcRenderer.removeListener('convert:progress', handler);
  },
});
