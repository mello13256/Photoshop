// Processo principal do Electron: janela, diálogos nativos e sistema de arquivos.
const { app, BrowserWindow, ipcMain, dialog, Menu, clipboard, ClipboardItem, nativeImage, shell } = require('electron');
const path = require('path');
const fs = require('fs/promises');

const SUPPORTED = /\.(lumina|png|jpe?g|webp|gif|bmp|svg|ico|avif)$/i;
let win = null;
let allowClose = false;
let launchPaths = process.argv.slice(app.isPackaged ? 1 : 2).filter((a) => SUPPORTED.test(a));

async function readFiles(paths) {
  const out = [];
  for (const p of paths) {
    try {
      out.push({ name: path.basename(p), path: p, data: await fs.readFile(p) });
    } catch (err) {
      console.error('Falha ao ler', p, err);
    }
  }
  return out;
}

function createWindow() {
  win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    backgroundColor: '#1c1d20',
    title: 'Lumina Studio',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: false,
    },
  });

  if (process.platform === 'darwin') {
    Menu.setApplicationMenu(Menu.buildFromTemplate([
      { role: 'appMenu' },
      { role: 'editMenu' },
      { role: 'windowMenu' },
    ]));
  } else {
    Menu.setApplicationMenu(null);
  }

  win.once('ready-to-show', () => {
    win.maximize();
    win.show();
  });
  win.loadFile(path.join(__dirname, '..', 'src', 'index.html'));

  // Links externos abrem no navegador padrão.
  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^https?:/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e) => e.preventDefault());

  win.on('close', (e) => {
    if (allowClose) return;
    e.preventDefault();
    win.webContents.send('close-request');
  });
  win.on('closed', () => { win = null; });

  // F12 abre as ferramentas de desenvolvedor (somente em desenvolvimento).
  win.webContents.on('before-input-event', (e, input) => {
    if (input.type !== 'keyDown') return;
    if (input.key === 'F12' && !app.isPackaged) {
      win.webContents.toggleDevTools();
      e.preventDefault();
    }
  });
}

// ---- IPC -----------------------------------------------------------------

ipcMain.handle('open-file', async (_e, filters) => {
  const r = await dialog.showOpenDialog(win, { properties: ['openFile', 'multiSelections'], filters });
  if (r.canceled) return [];
  return readFiles(r.filePaths);
});

ipcMain.handle('save-file', async (_e, { defaultPath, filters, data }) => {
  const r = await dialog.showSaveDialog(win, { defaultPath, filters });
  if (r.canceled || !r.filePath) return null;
  await fs.writeFile(r.filePath, Buffer.from(data));
  return r.filePath;
});

ipcMain.handle('write-file', async (_e, { filePath, data }) => {
  if (typeof filePath !== 'string' || !SUPPORTED.test(filePath)) throw new Error('Caminho inválido');
  await fs.writeFile(filePath, Buffer.from(data));
  return filePath;
});

// Electron 44+ usa a API assíncrona (estilo W3C); versões antigas usam writeImage/readImage.
ipcMain.handle('clipboard-write-image', async (_e, dataUrl) => {
  if (typeof clipboard.writeImage === 'function') {
    clipboard.writeImage(nativeImage.createFromDataURL(dataUrl));
    return true;
  }
  const png = nativeImage.createFromDataURL(dataUrl).toPNG();
  await clipboard.write([new ClipboardItem({ 'image/png': new Blob([png], { type: 'image/png' }) })]);
  return true;
});

ipcMain.handle('clipboard-read-image', async () => {
  if (typeof clipboard.readImage === 'function') {
    const img = clipboard.readImage();
    return img.isEmpty() ? null : img.toDataURL();
  }
  for (const item of await clipboard.read()) {
    const type = item.types.find((t) => t.startsWith('image/'));
    if (!type) continue;
    const blob = await item.getType(type);
    const buf = Buffer.from(await blob.arrayBuffer());
    return `data:${type};base64,${buf.toString('base64')}`;
  }
  return null;
});

ipcMain.handle('launch-files', async () => {
  const files = await readFiles(launchPaths);
  launchPaths = [];
  return files;
});

ipcMain.on('confirm-close', () => {
  allowClose = true;
  win?.close();
});

ipcMain.on('toggle-fullscreen', () => {
  if (win) win.setFullScreen(!win.isFullScreen());
});

// ---- Ciclo de vida -------------------------------------------------------

if (!app.requestSingleInstanceLock()) {
  app.quit();
} else {
  app.on('second-instance', async (_e, argv) => {
    if (!win) return;
    if (win.isMinimized()) win.restore();
    win.focus();
    const files = await readFiles(argv.filter((a) => SUPPORTED.test(a)));
    if (files.length) win.webContents.send('open-files', files);
  });

  app.on('open-file', async (e, p) => {
    e.preventDefault();
    if (win) win.webContents.send('open-files', await readFiles([p]));
    else launchPaths.push(p);
  });

  app.whenReady().then(() => {
    createWindow();
    app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
  });
}
