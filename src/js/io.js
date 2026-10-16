// Abrir/salvar arquivos. Usa a API nativa do Electron quando disponível e recorre ao navegador caso contrário.
import { Doc, Layer, documentFromImage } from './document.js';
import { createCanvas, canvasToBlob, loadImage } from './utils.js';
import { renderText } from './tools.js';

const native = window.lumina;
export const isNative = !!native?.isElectron;

export const PROJECT_EXT = 'lumina';
const IMAGE_EXTS = ['png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg', 'ico', 'avif'];

export const OPEN_FILTERS = [
  { name: 'Todos os suportados', extensions: [PROJECT_EXT, ...IMAGE_EXTS] },
  { name: 'Projeto Lumina', extensions: [PROJECT_EXT] },
  { name: 'Imagens', extensions: IMAGE_EXTS },
];

const extOf = (name) => (name.split('.').pop() || '').toLowerCase();
const baseName = (name) => name.replace(/\.[^.]+$/, '');

// Retorna [{ name, path, data: Uint8Array }]
export async function pickFiles() {
  if (isNative) return native.openFile(OPEN_FILTERS);
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.multiple = true;
    input.accept = [PROJECT_EXT, ...IMAGE_EXTS].map((e) => `.${e}`).join(',');
    input.onchange = async () => {
      const out = [];
      for (const f of input.files) out.push({ name: f.name, path: null, data: new Uint8Array(await f.arrayBuffer()) });
      resolve(out);
    };
    input.click();
  });
}

export async function fileToDocument({ name, path, data }) {
  const ext = extOf(name);
  if (ext === PROJECT_EXT) {
    const doc = await loadProject(new TextDecoder().decode(data), baseName(name));
    doc.filePath = path;
    doc.name = name;
    return doc;
  }
  const type = ext === 'svg' ? 'image/svg+xml' : ext === 'jpg' ? 'image/jpeg' : `image/${ext}`;
  const url = URL.createObjectURL(new Blob([data], { type }));
  try {
    const img = await loadImage(url);
    const doc = documentFromImage(img, name);
    doc.exportPath = path;
    return doc;
  } finally {
    URL.revokeObjectURL(url);
  }
}

export async function blobToDocument(blob, name = 'Colado') {
  const url = URL.createObjectURL(blob);
  try {
    return documentFromImage(await loadImage(url), name);
  } finally {
    URL.revokeObjectURL(url);
  }
}

// ---- Projeto (.lumina) ---------------------------------------------------

export function serializeProject(doc) {
  return JSON.stringify({
    app: 'Lumina Studio',
    format: 1,
    width: doc.width,
    height: doc.height,
    active: doc.active,
    layers: doc.layers.map((l) => ({
      name: l.name, visible: l.visible, opacity: l.opacity, blend: l.blend, locked: l.locked,
      text: l.text ? { ...l.text, bounds: undefined } : null,
      png: l.canvas.toDataURL('image/png'),
    })),
  });
}

export async function loadProject(json, name) {
  const p = JSON.parse(json);
  if (!p || p.format !== 1 || !Array.isArray(p.layers)) throw new Error('Arquivo de projeto inválido.');
  const doc = new Doc(p.width, p.height, name);
  for (const d of p.layers) {
    const img = await loadImage(d.png);
    const c = createCanvas(p.width, p.height);
    c.getContext('2d').drawImage(img, 0, 0);
    const l = new Layer(d.name, p.width, p.height, c);
    Object.assign(l, { visible: d.visible !== false, opacity: d.opacity ?? 1, blend: d.blend || 'source-over', locked: !!d.locked, text: d.text || null });
    if (l.text) renderText(l.text, p.width, p.height); // recalcula limites
    doc.layers.push(l);
  }
  doc.active = Math.min(p.active || 0, doc.layers.length - 1);
  doc.commit('Abrir');
  doc.savedIndex = doc.hIndex;
  return doc;
}

// ---- Salvar --------------------------------------------------------------

async function saveBytes(bytes, defaultName, filters, path) {
  if (isNative) {
    if (path) {
      await native.writeFile(path, bytes);
      return path;
    }
    return native.saveFile({ defaultPath: defaultName, filters, data: bytes });
  }
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([bytes]));
  a.download = defaultName;
  a.click();
  setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  return defaultName;
}

export async function saveProject(doc, saveAs = false) {
  const bytes = new TextEncoder().encode(serializeProject(doc));
  const name = `${baseName(doc.name)}.${PROJECT_EXT}`;
  const path = await saveBytes(bytes, name, [{ name: 'Projeto Lumina', extensions: [PROJECT_EXT] }], saveAs ? null : doc.filePath);
  if (!path) return false;
  if (isNative) {
    doc.filePath = path;
    doc.name = path.split(/[\\/]/).pop();
  } else doc.name = name;
  doc.markSaved();
  return true;
}

export const EXPORT_FORMATS = {
  png: { mime: 'image/png', name: 'PNG', ext: 'png' },
  jpg: { mime: 'image/jpeg', name: 'JPEG', ext: 'jpg' },
  webp: { mime: 'image/webp', name: 'WebP', ext: 'webp' },
};

export async function exportImage(doc, { format = 'png', quality = 0.92, scale = 1, background = '#ffffff' } = {}) {
  const f = EXPORT_FORMATS[format];
  const flat = doc.flattenToCanvas();
  const w = Math.max(1, Math.round(doc.width * scale)), h = Math.max(1, Math.round(doc.height * scale));
  const out = createCanvas(w, h);
  const ctx = out.getContext('2d');
  if (format === 'jpg') {
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, w, h);
  }
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(flat, 0, 0, w, h);
  const blob = await canvasToBlob(out, f.mime, quality);
  const bytes = new Uint8Array(await blob.arrayBuffer());
  return saveBytes(bytes, `${baseName(doc.name)}.${f.ext}`, [{ name: f.name, extensions: [f.ext] }], null);
}

// ---- Área de transferência ----------------------------------------------

export async function writeClipboardImage(canvas) {
  try {
    if (isNative) {
      await native.copyImage(canvas.toDataURL('image/png'));
      return true;
    }
    const blob = await canvasToBlob(canvas, 'image/png');
    await navigator.clipboard.write([new ClipboardItem({ 'image/png': blob })]);
    return true;
  } catch {
    return false;
  }
}

export async function readClipboardImage() {
  try {
    if (isNative) {
      const url = await native.readImage();
      return url ? loadImage(url) : null;
    }
    const items = await navigator.clipboard.read();
    for (const it of items) {
      const type = it.types.find((t) => t.startsWith('image/'));
      if (type) {
        const url = URL.createObjectURL(await it.getType(type));
        const img = await loadImage(url);
        URL.revokeObjectURL(url);
        return img;
      }
    }
  } catch { /* sem permissão ou vazio */ }
  return null;
}
