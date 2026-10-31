// Leitura de pacotes Office Open XML (.docx, .pptx, .xlsx): zip, XML, relações e cores do tema.
import { loadVendor, bytesToDataUrl } from './native.js';

export const EMU_PER_PX = 9525; // 96 dpi
export const EMU_PER_PT = 12700;

export async function openPackage(data) {
  const JSZip = await loadVendor('jszip.min.js', 'JSZip');
  const zip = await JSZip.loadAsync(data instanceof Uint8Array ? data : new Uint8Array(data));
  return new Package(zip);
}

export class Package {
  constructor(zip) {
    this.zip = zip;
    this.xmlCache = new Map();
    this.relsCache = new Map();
  }

  has(path) {
    return !!this.zip.file(path);
  }

  async text(path) {
    const f = this.zip.file(path);
    return f ? f.async('string') : null;
  }

  async bytes(path) {
    const f = this.zip.file(path);
    return f ? f.async('uint8array') : null;
  }

  async xml(path) {
    if (!path) return null;
    if (!this.xmlCache.has(path)) {
      this.xmlCache.set(path, this.text(path).then((t) => {
        if (t == null) return null;
        const doc = new DOMParser().parseFromString(t, 'application/xml');
        return doc.getElementsByTagName('parsererror').length ? null : doc;
      }));
    }
    return this.xmlCache.get(path);
  }

  // Relações da parte: Map id -> { target (caminho absoluto no pacote), type, external }
  async rels(partPath) {
    if (this.relsCache.has(partPath)) return this.relsCache.get(partPath);
    const dir = partPath.includes('/') ? partPath.slice(0, partPath.lastIndexOf('/')) : '';
    const name = partPath.slice(partPath.lastIndexOf('/') + 1);
    const doc = await this.xml(`${dir ? `${dir}/` : ''}_rels/${name}.rels`);
    const map = new Map();
    if (doc) {
      for (const r of doc.getElementsByTagName('Relationship')) {
        const external = r.getAttribute('TargetMode') === 'External';
        const target = r.getAttribute('Target') || '';
        map.set(r.getAttribute('Id'), {
          target: external ? target : resolvePath(dir, target),
          type: (r.getAttribute('Type') || '').split('/').pop(),
          external,
        });
      }
    }
    this.relsCache.set(partPath, map);
    return map;
  }

  async relByType(partPath, type) {
    for (const r of (await this.rels(partPath)).values()) if (r.type === type) return r;
    return null;
  }

  async dataUrl(path) {
    const b = await this.bytes(path);
    if (!b) return null;
    return bytesToDataUrl(b, mimeOfPath(path));
  }
}

export function resolvePath(dir, target) {
  if (target.startsWith('/')) return target.slice(1);
  const parts = dir ? dir.split('/') : [];
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.' && seg !== '') parts.push(seg);
  }
  return parts.join('/');
}

export function mimeOfPath(p) {
  const e = p.split('.').pop().toLowerCase();
  return {
    png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', webp: 'image/webp',
    svg: 'image/svg+xml', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf',
  }[e] || 'application/octet-stream';
}

// Formatos que o navegador consegue exibir.
export const isWebImage = (p) => /\.(png|jpe?g|gif|bmp|webp|svg)$/i.test(p);

// ---- Navegação por nome local (independe do prefixo de namespace) ----------------

export function kids(node, name) {
  const out = [];
  if (!node) return out;
  for (const c of node.children) if (!name || c.localName === name) out.push(c);
  return out;
}

export function kid(node, name) {
  if (!node) return null;
  for (const c of node.children) if (c.localName === name) return c;
  return null;
}

// Caminho de filhos: path(node, 'spPr', 'xfrm', 'off')
export function path(node, ...names) {
  let n = node;
  for (const name of names) {
    n = kid(n, name);
    if (!n) return null;
  }
  return n;
}

export function all(node, name) {
  return node ? [...node.getElementsByTagNameNS('*', name)] : [];
}

export function first(node, name) {
  return node ? node.getElementsByTagNameNS('*', name)[0] || null : null;
}

export function attr(node, name) {
  if (!node) return null;
  if (node.hasAttribute(name)) return node.getAttribute(name);
  for (const a of node.attributes) if (a.localName === name) return a.value;
  return null;
}

// r:id (relação) — distinto de um atributo "id" comum no mesmo elemento
export function rid(node, name = 'id') {
  return node?.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships', name) || null;
}

export function num(node, name, def = null) {
  const v = attr(node, name);
  if (v === null || v === '') return def;
  const n = Number(v);
  return Number.isFinite(n) ? n : def;
}

// Booleano no estilo OOXML: <w:b/>, <w:b w:val="0"/>, val="false"/"off"
export function flag(node) {
  if (!node) return null;
  const v = attr(node, 'val');
  return v === null ? true : !['0', 'false', 'off', 'none'].includes(v);
}

// Conteúdo de mc:AlternateContent: prefere Choice quando reconhecível, senão Fallback.
export function unwrapAlternate(node) {
  if (node?.localName !== 'AlternateContent') return node;
  const choice = kid(node, 'Choice');
  const fallback = kid(node, 'Fallback');
  return (choice && choice.firstElementChild ? choice : fallback) || choice || fallback;
}

// ---- Cores (DrawingML) -----------------------------------------------------------

export async function readTheme(pkg, path) {
  const doc = await pkg.xml(path);
  const theme = { colors: {}, major: 'Calibri Light', minor: 'Calibri' };
  if (!doc) return theme;
  const scheme = first(doc, 'clrScheme');
  for (const c of kids(scheme)) {
    const v = kid(c, 'srgbClr') ? attr(kid(c, 'srgbClr'), 'val') : kid(c, 'sysClr') ? attr(kid(c, 'sysClr'), 'lastClr') || (attr(kid(c, 'sysClr'), 'val') === 'window' ? 'FFFFFF' : '000000') : null;
    if (v) theme.colors[c.localName] = `#${v.toLowerCase()}`;
  }
  const major = first(doc, 'majorFont');
  const minor = first(doc, 'minorFont');
  if (major) theme.major = attr(kid(major, 'latin'), 'typeface') || theme.major;
  if (minor) theme.minor = attr(kid(minor, 'latin'), 'typeface') || theme.minor;
  return theme;
}

const PRESET = {
  black: '000000', white: 'ffffff', red: 'ff0000', green: '008000', blue: '0000ff', yellow: 'ffff00', cyan: '00ffff', magenta: 'ff00ff',
  gray: '808080', grey: '808080', darkGray: 'a9a9a9', lightGray: 'd3d3d3', orange: 'ffa500', navy: '000080', purple: '800080', silver: 'c0c0c0',
};

const SCHEME_ALIAS = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2', text1: 'dk1', text2: 'dk2', background1: 'lt1', background2: 'lt2', dark1: 'dk1', dark2: 'dk2', light1: 'lt1', light2: 'lt2', hyperlink: 'hlink', followedHyperlink: 'folHlink' };

export function schemeColor(name, theme, map = null) {
  let n = name;
  if (map && map[n]) n = map[n];
  n = SCHEME_ALIAS[n] || n;
  return theme?.colors?.[n] || null;
}

// Lê um elemento de cor DrawingML (srgbClr, schemeClr…) com seus modificadores.
// Retorna { hex, alpha } ou null.
export function readColor(node, theme, map = null, phClr = null) {
  if (!node) return null;
  let c = null;
  for (const el of node.localName.endsWith('Clr') ? [node] : kids(node)) {
    let hex = null;
    switch (el.localName) {
      case 'srgbClr': hex = `#${(attr(el, 'val') || '000000').toLowerCase()}`; break;
      case 'schemeClr': hex = attr(el, 'val') === 'phClr' ? phClr : schemeColor(attr(el, 'val'), theme, map); break;
      case 'sysClr': hex = `#${(attr(el, 'lastClr') || (attr(el, 'val') === 'window' ? 'ffffff' : '000000')).toLowerCase()}`; break;
      case 'prstClr': hex = `#${PRESET[attr(el, 'val')] || '000000'}`; break;
      case 'scrgbClr': hex = rgbHex(num(el, 'r', 0) / 1000 * 2.55, num(el, 'g', 0) / 1000 * 2.55, num(el, 'b', 0) / 1000 * 2.55); break;
      case 'hslClr': hex = hslHex(num(el, 'hue', 0) / 60000, num(el, 'sat', 0) / 100000, num(el, 'lum', 0) / 100000); break;
      default: continue;
    }
    if (!hex) return null;
    c = applyMods(hex, el);
    break;
  }
  return c;
}

function applyMods(hex, el) {
  let [r, g, b] = hexRgb(hex);
  let alpha = 1;
  for (const m of kids(el)) {
    const v = num(m, 'val', 0) / 100000;
    switch (m.localName) {
      case 'alpha': alpha = v; break;
      case 'lumMod': case 'lumOff': {
        let [h, s, l] = rgbHsl(r, g, b);
        l = m.localName === 'lumMod' ? l * v : l + v;
        [r, g, b] = hslRgb(h, s, Math.max(0, Math.min(1, l)));
        break;
      }
      case 'tint': r += (255 - r) * (1 - v); g += (255 - g) * (1 - v); b += (255 - b) * (1 - v); break;
      case 'shade': r *= v; g *= v; b *= v; break;
      case 'satMod': {
        let [h, s, l] = rgbHsl(r, g, b);
        s = Math.min(1, s * v);
        [r, g, b] = hslRgb(h, s, l);
        break;
      }
      default: break;
    }
  }
  return { hex: rgbHex(r, g, b), alpha };
}

// Cor do Word (w:color, w:shd, bordas): valor hexadecimal ou themeColor com themeTint/themeShade.
// valAttr: 'val' (w:color), 'fill' (w:shd) ou 'color' (bordas).
export function wordColor(node, theme, valAttr = 'val') {
  if (!node) return null;
  const tc = attr(node, valAttr === 'fill' ? 'themeFill' : 'themeColor');
  let hex = null;
  if (tc) {
    hex = schemeColor(tc, theme);
    const tint = attr(node, valAttr === 'fill' ? 'themeFillTint' : 'themeTint');
    const shade = attr(node, valAttr === 'fill' ? 'themeFillShade' : 'themeShade');
    if (hex && (tint || shade)) {
      let [r, g, b] = hexRgb(hex);
      if (tint) { const t = parseInt(tint, 16) / 255; r = 255 - (255 - r) * t; g = 255 - (255 - g) * t; b = 255 - (255 - b) * t; }
      if (shade) { const s = parseInt(shade, 16) / 255; r *= s; g *= s; b *= s; }
      hex = rgbHex(r, g, b);
    }
  }
  if (!hex) {
    const v = attr(node, valAttr);
    if (v && /^[0-9a-f]{6}$/i.test(v)) hex = `#${v.toLowerCase()}`;
  }
  return hex;
}

export function cssColor(c) {
  if (!c) return null;
  return c.alpha < 1 ? `rgba(${hexRgb(c.hex).join(',')},${+c.alpha.toFixed(3)})` : c.hex;
}

export function hexRgb(hex) {
  const n = parseInt(hex.replace('#', '').slice(0, 6), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
}

export function rgbHex(r, g, b) {
  return `#${[r, g, b].map((x) => Math.round(Math.max(0, Math.min(255, x))).toString(16).padStart(2, '0')).join('')}`;
}

function rgbHsl(r, g, b) {
  r /= 255; g /= 255; b /= 255;
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  const l = (max + min) / 2;
  if (max === min) return [0, 0, l];
  const d = max - min;
  const s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
  let h;
  if (max === r) h = (g - b) / d + (g < b ? 6 : 0);
  else if (max === g) h = (b - r) / d + 2;
  else h = (r - g) / d + 4;
  return [h / 6, s, l];
}

function hslRgb(h, s, l) {
  if (s === 0) return [l * 255, l * 255, l * 255];
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s;
  const p = 2 * l - q;
  const f = (t) => {
    if (t < 0) t += 1;
    if (t > 1) t -= 1;
    if (t < 1 / 6) return p + (q - p) * 6 * t;
    if (t < 1 / 2) return q;
    if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6;
    return p;
  };
  return [f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255];
}

function hslHex(hDeg, s, l) {
  return rgbHex(...hslRgb(hDeg / 360, s, l));
}

// Fonte CSS com alternativas métricas para as fontes padrão do Office.
const FALLBACK = {
  calibri: 'Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  'calibri light': '"Calibri Light", Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  cambria: 'Cambria, Caladea, Georgia, serif',
  'times new roman': '"Times New Roman", Tinos, "Liberation Serif", serif',
  arial: 'Arial, Arimo, "Liberation Sans", Helvetica, sans-serif',
  'courier new': '"Courier New", Cousine, "Liberation Mono", monospace',
  aptos: 'Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  'aptos display': '"Aptos Display", Aptos, Calibri, Carlito, "Segoe UI", Arial, sans-serif',
  georgia: 'Georgia, serif',
  verdana: 'Verdana, "DejaVu Sans", sans-serif',
  tahoma: 'Tahoma, Verdana, sans-serif',
  'segoe ui': '"Segoe UI", Inter, system-ui, sans-serif',
  consolas: 'Consolas, "Cascadia Mono", Menlo, monospace',
};

export function cssFont(name) {
  if (!name) return null;
  const k = name.toLowerCase().trim();
  if (FALLBACK[k]) return FALLBACK[k];
  const serif = /serif|times|garamond|georgia|cambria|book|roman/i.test(name) && !/sans/i.test(name);
  return `"${name.replace(/"/g, '')}", ${serif ? 'Georgia, serif' : 'Calibri, Carlito, Arial, sans-serif'}`;
}
