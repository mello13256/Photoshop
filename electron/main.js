// Processo principal do Electron: janela, protocolo app://, diálogos nativos, PDF e sistema de arquivos.
const {
  app, BrowserWindow, ipcMain, dialog, Menu, clipboard, ClipboardItem, nativeImage, shell, protocol, net, session,
} = require('electron');
const path = require('path');
const os = require('os');
const fs = require('fs/promises');
const { pathToFileURL } = require('url');

const ROOT = path.join(__dirname, '..', 'src');
const SUPPORTED = /\.(lumina|lumdoc|lumsheet|lumslides|png|jpe?g|webp|gif|bmp|svg|ico|avif|docx|xlsx|xlsm|csv|tsv|pptx|pdf|md|markdown|txt|html?)$/i;

let win = null;
let allowClose = false;
let launchPaths = process.argv.slice(app.isPackaged ? 1 : 2).filter((a) => SUPPORTED.test(a));

// O protocolo app:// dá a todas as páginas a mesma origem (necessário para os apps em iframes).
protocol.registerSchemesAsPrivileged([
  { scheme: 'app', privileges: { standard: true, secure: true, supportFetchAPI: true, corsEnabled: true } },
]);

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

function registerProtocol() {
  protocol.handle('app', (req) => {
    const { pathname } = new URL(req.url);
    const file = path.normalize(path.join(ROOT, decodeURIComponent(pathname)));
    if (!file.startsWith(ROOT + path.sep)) return new Response('Proibido', { status: 403 });
    return net.fetch(pathToFileURL(file).toString());
  });
}

function buildContextMenu(params) {
  const items = [];
  for (const s of params.dictionarySuggestions.slice(0, 6)) {
    items.push({ label: s, click: () => win.webContents.replaceMisspelling(s) });
  }
  if (params.misspelledWord) {
    if (!params.dictionarySuggestions.length) items.push({ label: 'Sem sugestões', enabled: false });
    items.push({ label: 'Adicionar ao dicionário', click: () => session.defaultSession.addWordToSpellCheckerDictionary(params.misspelledWord) });
    items.push({ type: 'separator' });
  }
  if (params.isEditable) {
    items.push(
      { label: 'Recortar', role: 'cut', enabled: params.editFlags.canCut },
      { label: 'Copiar', role: 'copy', enabled: params.editFlags.canCopy },
      { label: 'Colar', role: 'paste', enabled: params.editFlags.canPaste },
      { type: 'separator' },
      { label: 'Selecionar tudo', role: 'selectAll' },
    );
  } else if (params.selectionText) {
    items.push({ label: 'Copiar', role: 'copy' });
  }
  return items;
}

function createWindow() {
  win = new BrowserWindow({
    width: 1440,
    height: 900,
    minWidth: 960,
    minHeight: 600,
    backgroundColor: '#14151a',
    title: 'Lumina Studio',
    icon: path.join(__dirname, '..', 'build', 'icon.png'),
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: true,
      spellcheck: true,
    },
  });

  if (process.platform === 'darwin') {
    Menu.setApplicationMenu(Menu.buildFromTemplate([{ role: 'appMenu' }, { role: 'editMenu' }, { role: 'windowMenu' }]));
  } else {
    Menu.setApplicationMenu(null);
  }

  win.once('ready-to-show', () => {
    win.maximize();
    win.show();
  });
  win.loadURL('app://lumina/index.html');

  win.webContents.setWindowOpenHandler(({ url }) => {
    if (/^(https?|mailto):/.test(url)) shell.openExternal(url);
    return { action: 'deny' };
  });
  win.webContents.on('will-navigate', (e, url) => {
    if (!url.startsWith('app://lumina/')) e.preventDefault();
  });
  win.webContents.on('context-menu', (_e, params) => {
    const items = buildContextMenu(params);
    if (items.length) Menu.buildFromTemplate(items).popup({ window: win });
  });

  win.on('close', (e) => {
    if (allowClose) return;
    e.preventDefault();
    win.webContents.send('close-request');
  });
  win.on('closed', () => { win = null; });

  win.webContents.on('before-input-event', (e, input) => {
    if (input.type === 'keyDown' && input.key === 'F12' && !app.isPackaged) {
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

ipcMain.handle('read-path', async (_e, filePath) => {
  if (typeof filePath !== 'string' || !SUPPORTED.test(filePath)) throw new Error('Caminho inválido');
  const [f] = await readFiles([filePath]);
  if (!f) throw new Error('Arquivo não encontrado');
  return f;
});

ipcMain.handle('save-file', async (_e, { defaultPath, filters, data }) => {
  const r = await dialog.showSaveDialog(win, { defaultPath, filters });
  if (r.canceled || !r.filePath) return null;
  await fs.writeFile(r.filePath, typeof data === 'string' ? data : Buffer.from(data));
  return r.filePath;
});

ipcMain.handle('write-file', async (_e, { filePath, data }) => {
  if (typeof filePath !== 'string' || !SUPPORTED.test(filePath)) throw new Error('Caminho inválido');
  await fs.writeFile(filePath, typeof data === 'string' ? data : Buffer.from(data));
  return filePath;
});

// Gera um PDF a partir de um documento HTML completo, em uma janela oculta e isolada.
ipcMain.handle('export-pdf', async (_e, { html, defaultPath, options = {} }) => {
  const target = await dialog.showSaveDialog(win, { defaultPath, filters: [{ name: 'PDF', extensions: ['pdf'] }] });
  if (target.canceled || !target.filePath) return null;
  const tmp = path.join(os.tmpdir(), `lumina-print-${process.pid}-${Date.now()}.html`);
  await fs.writeFile(tmp, html, 'utf8');
  const printer = new BrowserWindow({
    show: false,
    webPreferences: { sandbox: true, contextIsolation: true, javascript: false },
  });
  try {
    await printer.loadFile(tmp);
    const pdf = await printer.webContents.printToPDF({ printBackground: true, preferCSSPageSize: true, ...options });
    await fs.writeFile(target.filePath, pdf);
    return target.filePath;
  } finally {
    printer.destroy();
    fs.unlink(tmp).catch(() => {});
  }
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

ipcMain.handle('set-fullscreen', (_e, on) => {
  if (win) win.setFullScreen(!!on);
  return true;
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
    registerProtocol();
    try { session.defaultSession.setSpellCheckerLanguages(['pt-BR', 'en-US']); } catch { /* macOS usa o corretor do sistema */ }
    createWindow();
    app.on('activate', () => { if (!BrowserWindow.getAllWindows().length) createWindow(); });
  });

  app.on('window-all-closed', () => {
    if (process.platform !== 'darwin') app.quit();
  });
}
