// Ponte com o aplicativo desktop (Electron) e com a "casca" da suíte, com alternativas para o navegador.

function findNative() {
  try {
    if (window.lumina) return window.lumina;
    if (window.parent !== window && window.parent.lumina) return window.parent.lumina;
  } catch { /* origem diferente */ }
  return null;
}

export const native = findNative();
export const isNative = !!native?.isElectron;

export function shell() {
  try {
    return window.parent !== window ? window.parent.luminaShell || null : null;
  } catch {
    return null;
  }
}

export const extOf = (name) => (String(name).split('.').pop() || '').toLowerCase();
export const baseName = (name) => String(name).replace(/\.[^.\\/]+$/, '');
export const fileName = (p) => String(p).split(/[\\/]/).pop();

// ---- Abrir ---------------------------------------------------------------

// Retorna [{ name, path, data: Uint8Array }]
const localize = (f) => ({ ...f, data: new Uint8Array(f.data) });

export async function pickFiles(filters, { multiple = true } = {}) {
  if (isNative) return (await native.openFile(filters)).map(localize);
  const exts = filters.flatMap((f) => f.extensions).filter((e) => e !== '*');
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = multiple;
    if (exts.length) input.accept = exts.map((e) => `.${e}`).join(',');
    input.onchange = async () => {
      const out = [];
      for (const f of input.files) out.push({ name: f.name, path: null, data: new Uint8Array(await f.arrayBuffer()) });
      resolve(out);
    };
    input.click();
  });
}

export async function readPath(path) {
  if (!isNative) throw new Error('Disponível apenas no aplicativo desktop.');
  return localize(await native.readPath(path));
}

export async function pickImageDataUrl() {
  const [f] = await pickFiles([{ name: 'Imagens', extensions: ['png', 'jpg', 'jpeg', 'gif', 'webp', 'svg', 'bmp'] }], { multiple: false });
  if (!f) return null;
  return bytesToDataUrl(f.data, mimeOf(f.name));
}

export function mimeOf(name) {
  const e = extOf(name);
  return {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', webp: 'image/webp', svg: 'image/svg+xml', bmp: 'image/bmp',
    pdf: 'application/pdf', html: 'text/html', htm: 'text/html', md: 'text/markdown', txt: 'text/plain', csv: 'text/csv', json: 'application/json',
    docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    pptx: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
  }[e] || 'application/octet-stream';
}

export function bytesToDataUrl(bytes, mime) {
  let bin = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) bin += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));
  return `data:${mime};base64,${btoa(bin)}`;
}

export const decodeText = (bytes) => new TextDecoder().decode(bytes);

// ---- Salvar --------------------------------------------------------------

// data: string | Uint8Array | ArrayBuffer | Blob. Retorna o caminho (ou nome no navegador) ou null.
export async function saveFile(data, { name, filters, path = null }) {
  if (data instanceof Blob) data = new Uint8Array(await data.arrayBuffer());
  else if (data instanceof ArrayBuffer) data = new Uint8Array(data);
  if (isNative) {
    if (path) return native.writeFile(path, data);
    return native.saveFile({ defaultPath: name, filters, data });
  }
  const blob = new Blob([data], { type: mimeOf(name) });
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = name;
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  return name;
}

// Exporta um documento HTML completo como PDF.
export async function exportPdf(html, { name, options = {} } = {}) {
  if (isNative) return native.exportPdf({ html, defaultPath: name, options });
  const w = window.open('', '_blank');
  if (!w) throw new Error('O navegador bloqueou a janela de impressão.');
  w.document.open();
  w.document.write(html);
  w.document.close();
  setTimeout(() => { w.focus(); w.print(); }, 400);
  return name;
}

// ---- Bibliotecas de terceiros (carregadas sob demanda) -------------------

const loaded = new Map();
export function loadVendor(file, globalName) {
  if (globalName && window[globalName]) return Promise.resolve(window[globalName]);
  if (!loaded.has(file)) {
    loaded.set(file, new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = new URL(`../vendor/${file}`, import.meta.url).href;
      s.onload = () => resolve(globalName ? window[globalName] : true);
      s.onerror = () => { loaded.delete(file); reject(new Error(`Não foi possível carregar ${file}. Rode "npm install".`)); };
      document.head.append(s);
    }));
  }
  return loaded.get(file);
}

// ---- Integração com a casca ---------------------------------------------

export function setTitle(t) {
  document.title = t;
  shell()?.setTitle?.(t, window);
}

export function addRecent(entry) {
  shell()?.addRecent?.(entry);
}

// ---- Tema ----------------------------------------------------------------

const THEME_KEY = 'lumina.theme';
export function getThemePref() {
  try { return localStorage.getItem(THEME_KEY) || 'system'; } catch { return 'system'; }
}
export function resolvedTheme(pref = getThemePref()) {
  if (pref === 'system') return matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light';
  return pref;
}
export function applyTheme() {
  document.documentElement.dataset.theme = resolvedTheme();
}
export function setThemePref(pref) {
  try { localStorage.setItem(THEME_KEY, pref); } catch { /* indisponível */ }
  applyTheme();
  for (const f of document.querySelectorAll('iframe')) {
    try { f.contentWindow.luminaApp?.themeChanged?.(); f.contentDocument.documentElement.dataset.theme = resolvedTheme(); } catch { /* ignorar */ }
  }
}
export function initTheme() {
  applyTheme();
  matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
  window.addEventListener('storage', (e) => { if (e.key === THEME_KEY) applyTheme(); });
}

// ---- Registro do app na casca -------------------------------------------

// api: { openFiles(files), newDocument(template), beforeClose(): Promise<boolean>, unsaved(): string[] }
export function registerApp(api) {
  // Arquivos vindos da casca pertencem a outro contexto (iframe): copia os bytes para este.
  if (api.openFiles) {
    const open = api.openFiles;
    api.openFiles = (files) => open(files.map((f) => ({ ...f, data: new Uint8Array(f.data) })));
  }
  window.luminaApp = api;
  window.addEventListener('beforeunload', (e) => {
    if (!isNative && window.parent === window && api.unsaved?.().length) { e.preventDefault(); e.returnValue = ''; }
  });
}

// Lê arquivos abertos diretamente quando o app roda sozinho (fora da casca).
export function standaloneBoot(api) {
  if (window.parent !== window) return;
  if (isNative) native.getLaunchFiles?.().then((f) => f?.length && api.openFiles(f));
}

export function fmtDate(d = new Date()) {
  return d.toLocaleDateString('pt-BR', { day: '2-digit', month: 'short', year: 'numeric' }) + ' ' +
    d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' });
}
