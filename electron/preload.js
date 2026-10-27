// Ponte segura entre a interface (renderer) e o processo principal.
const { contextBridge, ipcRenderer, webUtils } = require('electron');

contextBridge.exposeInMainWorld('lumina', {
  isElectron: true,
  platform: process.platform,
  openFile: (filters) => ipcRenderer.invoke('open-file', filters),
  readPath: (filePath) => ipcRenderer.invoke('read-path', filePath),
  saveFile: (opts) => ipcRenderer.invoke('save-file', opts),
  writeFile: (filePath, data) => ipcRenderer.invoke('write-file', { filePath, data }),
  exportPdf: (opts) => ipcRenderer.invoke('export-pdf', opts),
  copyImage: (dataUrl) => ipcRenderer.invoke('clipboard-write-image', dataUrl),
  readImage: () => ipcRenderer.invoke('clipboard-read-image'),
  getLaunchFiles: () => ipcRenderer.invoke('launch-files'),
  pathForFile: (file) => {
    try { return webUtils.getPathForFile(file); } catch { return null; }
  },
  onCloseRequest: (cb) => ipcRenderer.on('close-request', () => cb()),
  onOpenFiles: (cb) => ipcRenderer.on('open-files', (_e, files) => cb(files)),
  confirmClose: () => ipcRenderer.send('confirm-close'),
  toggleFullscreen: () => ipcRenderer.send('toggle-fullscreen'),
  setFullscreen: (on) => ipcRenderer.invoke('set-fullscreen', on),
});
