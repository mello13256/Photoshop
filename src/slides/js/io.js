// Arquivos do Lumina Apresentações: .lumslides (nativo), PPTX, PDF e PNG.
import { loadVendor, decodeText } from '../../shared/native.js';
import { escapeHtml } from '../../shared/ui.js';
import { renderSlide, SLIDE_CSS, themeOf, resolveColor, bgCss, isDark, solidOf } from './render.js';
import { W, H } from './themes.js';

export function serialize(pres) {
  // Mantém apenas as imagens usadas
  const used = new Set();
  for (const s of pres.slides) {
    if (s.bg?.type === 'image') used.add(s.bg.asset);
    for (const e of s.elements) if (e.type === 'image' && e.image.asset) used.add(e.image.asset);
  }
  const assets = Object.fromEntries(Object.entries(pres.assets || {}).filter(([k]) => used.has(k)));
  return JSON.stringify({ app: 'Lumina Apresentações', format: 1, ...pres, assets });
}

export function parse(bytes) {
  const p = JSON.parse(decodeText(bytes));
  if (p.format !== 1 || !Array.isArray(p.slides)) throw new Error('Arquivo .lumslides inválido.');
  p.assets ||= {};
  return p;
}

// ---- HTML / PDF ----------------------------------------------------------------

export function presentationHtml(pres, chartImages = {}) {
  const pages = pres.slides.map((s) => {
    const node = renderSlide(s, pres, { mode: 'export', chartImages });
    node.style.transform = `scale(${1280 / W})`;
    node.style.transformOrigin = '0 0';
    return `<div class="page">${node.outerHTML}</div>`;
  }).join('');
  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${escapeHtml(pres.title)}</title><style>
@page { size: 1280px 720px; margin: 0; }
html, body { margin: 0; background: #fff; }
.page { width: 1280px; height: 720px; overflow: hidden; break-after: page; page-break-after: always; position: relative; }
.page:last-child { break-after: auto; page-break-after: auto; }
${SLIDE_CSS}
</style></head><body>${pages}</body></html>`;
}

// ---- Imagem (PNG) via SVG foreignObject -------------------------------------------

export async function slideToPng(slide, pres, { width = 1920, chartImages = {} } = {}) {
  const node = renderSlide(slide, pres, { mode: 'export', chartImages });
  node.setAttribute('xmlns', 'http://www.w3.org/1999/xhtml');
  const xml = new XMLSerializer().serializeToString(node);
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="${W}" height="${H}"><foreignObject width="100%" height="100%"><style xmlns="http://www.w3.org/1999/xhtml">${SLIDE_CSS}</style>${xml}</foreignObject></svg>`;
  const img = new Image();
  img.src = `data:image/svg+xml;charset=utf-8,${encodeURIComponent(svg)}`;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = width;
  c.height = Math.round((width * H) / W);
  const ctx = c.getContext('2d');
  ctx.drawImage(img, 0, 0, c.width, c.height);
  return c.toDataURL('image/png');
}

// ---- PPTX ------------------------------------------------------------------------

const IN = (px) => px / 144; // 1920 px = 13,333 pol
const PT = (px) => Math.max(6, Math.round(px * 0.5 * 10) / 10);
const hex = (c) => {
  if (!c || c === 'transparent') return undefined;
  if (c.startsWith('linear-gradient')) c = solidOf(c);
  if (!c) return undefined;
  const m = /^#?([0-9a-f]{6})$/i.exec(c.trim());
  if (m) return m[1].toUpperCase();
  const r = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c);
  if (r) return [r[1], r[2], r[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
  return undefined;
};

function gradientPng(bg) {
  const c = document.createElement('canvas');
  c.width = 1920;
  c.height = 1080;
  const ctx = c.getContext('2d');
  const a = ((bg.angle ?? 135) - 90) * (Math.PI / 180);
  const cx = 960, cy = 540, r = Math.abs(960 * Math.cos(a)) + Math.abs(540 * Math.sin(a));
  const g = ctx.createLinearGradient(cx - Math.cos(a) * r, cy - Math.sin(a) * r, cx + Math.cos(a) * r, cy + Math.sin(a) * r);
  g.addColorStop(0, bg.from);
  g.addColorStop(1, bg.to);
  ctx.fillStyle = g;
  ctx.fillRect(0, 0, 1920, 1080);
  return c.toDataURL('image/jpeg', 0.9);
}

// Converte HTML do texto em parágrafos/trechos do PptxGenJS.
function runsFromHtml(html, base) {
  const root = document.createElement('div');
  root.innerHTML = html || '';
  const runs = [];
  const walk = (node, fmt, para) => {
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        const t = n.textContent.replace(/​/g, '');
        if (t) runs.push({ text: t, options: { ...base, ...fmt, ...para } });
        continue;
      }
      if (n.nodeType !== 1) continue;
      const tag = n.tagName;
      if (tag === 'BR') { runs.push({ text: '', options: { ...base, ...para, breakLine: true } }); continue; }
      const f = { ...fmt };
      if (tag === 'B' || tag === 'STRONG') f.bold = true;
      if (tag === 'I' || tag === 'EM') f.italic = true;
      if (tag === 'U') f.underline = { style: 'sng' };
      if (tag === 'S' || tag === 'STRIKE') f.strike = 'sngStrike';
      const st = n.style;
      if (st?.color) f.color = hex(st.color);
      if (st?.fontWeight === 'bold' || +st?.fontWeight >= 600) f.bold = true;
      if (st?.fontStyle === 'italic') f.italic = true;
      const fsz = /([\d.]+)px/.exec(st?.fontSize || '');
      if (fsz) f.fontSize = PT(+fsz[1]);
      if (st?.fontFamily) f.fontFace = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
      if (st?.backgroundColor && hex(st.backgroundColor)) f.highlight = hex(st.backgroundColor);
      if (st?.textTransform === 'uppercase') f.caps = true;
      if (tag === 'SUP') f.superscript = true;
      if (tag === 'SUB') f.subscript = true;
      if (tag === 'A') f.hyperlink = { url: n.getAttribute('href') || '' };
      if (tag === 'P' || tag === 'DIV' || tag === 'LI' || /^H\d$/.test(tag)) {
        const p = { ...para };
        if (tag === 'LI') p.bullet = n.parentElement?.tagName === 'OL' ? { type: 'number' } : true;
        if (n.style?.textAlign) p.align = n.style.textAlign === 'justify' ? 'justify' : n.style.textAlign;
        const startLen = runs.length;
        walk(n, f, p);
        if (runs.length === startLen) runs.push({ text: '', options: { ...base, ...p } });
        runs[runs.length - 1].options.breakLine = true;
        continue;
      }
      walk(n, f, para);
    }
  };
  walk(root, {}, {});
  if (runs.length) delete runs[runs.length - 1].options.breakLine;
  return runs.length ? runs : [{ text: '', options: base }];
}

const SHAPE_MAP = {
  rect: 'rect', round: 'roundRect', ellipse: 'ellipse', triangle: 'triangle', diamond: 'diamond', hexagon: 'hexagon',
  star: 'star5', arrow: 'rightArrow', chevron: 'chevron', callout: 'wedgeRoundRectCallout', line: 'line',
};

export async function exportPptx(pres, chartImages = {}) {
  const PptxGenJS = await loadVendor('pptxgen.bundle.js', 'PptxGenJS');
  const pptx = new PptxGenJS();
  pptx.layout = 'LAYOUT_WIDE';
  pptx.title = pres.title;
  pptx.author = 'Lumina Apresentações';
  const theme = themeOf(pres);
  if (pres.customTheme?.imported) pptx.theme = { headFontFace: theme.head.split(',')[0].replace(/["']/g, ''), bodyFontFace: theme.body.split(',')[0].replace(/["']/g, '') };
  for (const s of pres.slides) {
    const slide = pptx.addSlide();
    const bg = s.bg || theme.bg;
    if (bg.type === 'gradient') slide.background = { data: gradientPng(bg) };
    else if (bg.type === 'image' && pres.assets[bg.asset]) slide.background = { data: pres.assets[bg.asset] };
    else slide.background = { color: hex(bg.color) || 'FFFFFF' };
    for (const e of s.elements) {
      const box = { x: IN(e.x), y: IN(e.y), w: IN(e.w), h: IN(e.h), rotate: e.rot || 0 };
      if (e.type === 'text' || (e.type === 'shape' && e.text?.html)) {
        const t = e.text;
        const role = e.role || 'body';
        const color = resolveColor(t.color, theme, role === 'title' || role === 'big' || role === 'quote' ? theme.title : theme.text);
        const font = (t.font === 'head' || (!t.font && ['title', 'big', 'quote'].includes(role)) ? theme.head : !t.font || t.font === 'body' ? theme.body : t.font).split(',')[0].replace(/["']/g, '').trim();
        const base = { fontFace: font, fontSize: PT(t.size), color: hex(color), bold: !!t.bold, italic: !!t.italic };
        const margin = t.pad ? Math.round((t.pad.reduce((a, b) => a + b, 0) / 4) * 0.5) : e.type === 'shape' ? 12 : 3;
        const opts = { ...box, align: t.align || 'left', valign: t.valign === 'middle' ? 'middle' : t.valign === 'bottom' ? 'bottom' : 'top', margin, lineSpacingMultiple: Math.round(((t.line || 1.2) / 1.2) * 100) / 100, fit: 'shrink' };
        if (e.type === 'shape') {
          Object.assign(opts, shapeOpts(pptx, e, theme));
        }
        if (e.opacity < 1 && e.type === 'shape') opts.fill = { ...opts.fill, transparency: Math.round((1 - e.opacity) * 100) };
        slide.addText(runsFromHtml(t.html, base), opts);
      } else if (e.type === 'shape') {
        slide.addShape(pptx.ShapeType[SHAPE_MAP[e.shape.kind]] || pptx.ShapeType.rect, { ...box, ...shapeOpts(pptx, e, theme) });
      } else if (e.type === 'image' && pres.assets[e.image.asset]) {
        const src = pres.assets[e.image.asset];
        const dims = await imageSize(src);
        slide.addImage({
          data: src, ...box,
          sizing: e.image.fit === 'fill' || !dims ? undefined : e.image.fit === 'contain' ? { type: 'contain', w: box.w, h: box.h } : { type: 'cover', w: box.w, h: box.h },
          rounding: false, transparency: e.opacity < 1 ? Math.round((1 - e.opacity) * 100) : undefined,
          flipH: !!e.image.flipH, flipV: !!e.image.flipV,
        });
      } else if (e.type === 'table') {
        const t = e.table;
        const rows = t.rows.map((row, r) => row.map((cell, c) => {
          const head = t.header && r === 0;
          const own = t.cells?.[r]?.[c];
          const tmp = document.createElement('div');
          tmp.innerHTML = (cell || '').replace(/<br\s*\/?>/gi, '\n');
          const bg = own?.bg || (t.imported ? null : head ? theme.accent : r % 2 ? theme.accent2 : null);
          const color = own?.color || (head ? (isDark(own?.bg || theme.accent) ? '#ffffff' : '#111111') : theme.text);
          return { text: tmp.textContent, options: { bold: head || /<b>|<strong>/i.test(cell || ''), color: hex(color), fill: hex(bg) ? { color: hex(bg) } : undefined } };
        }));
        slide.addTable(rows, {
          ...box, fontSize: PT(t.size || 30), fontFace: theme.body.split(',')[0].replace(/["']/g, ''), border: { type: 'solid', pt: 1, color: hex(theme.accent2) || 'DDDDDD' }, autoPage: false,
          colW: t.cols?.length === t.rows[0]?.length ? t.cols.map((f) => f * box.w) : undefined,
        });
      } else if (e.type === 'chart') {
        const ch = e.chart;
        const type = { bar: pptx.ChartType.bar, column: pptx.ChartType.bar, line: pptx.ChartType.line, area: pptx.ChartType.area, pie: pptx.ChartType.pie, doughnut: pptx.ChartType.doughnut }[ch.kind] || pptx.ChartType.bar;
        const data = ch.series.map((sr) => ({ name: sr.name, labels: ch.labels, values: sr.values.map((v) => +v || 0) }));
        slide.addChart(type, data, {
          ...box, barDir: ch.kind === 'bar' ? 'bar' : 'col', showTitle: !!ch.title, title: ch.title, showLegend: ch.series.length > 1 || ['pie', 'doughnut'].includes(ch.kind),
          legendPos: 'b', chartColors: (ch.colors?.length ? ch.colors : chartPalette(theme)).map((c) => hex(c) || '888888'), catAxisLabelColor: hex(theme.text), valAxisLabelColor: hex(theme.text), titleColor: hex(theme.title),
          legendColor: hex(theme.text),
        });
        void chartImages;
      }
    }
    if (s.notes) slide.addNotes(s.notes);
  }
  return pptx.write({ outputType: 'uint8array' });
}

function shapeOpts(pptx, e, theme) {
  const s = e.shape;
  const fill = resolveColor(s.fill, theme, theme.accent);
  const stroke = s.stroke ? resolveColor(s.stroke, theme, '#000') : null;
  const o = { shape: pptx.ShapeType[SHAPE_MAP[s.kind]] || pptx.ShapeType.rect };
  if (s.kind === 'line') o.line = { color: hex(stroke || fill), width: Math.max(1, (s.strokeW || 6) * 0.5) };
  else {
    o.fill = hex(fill) ? { color: hex(fill) } : { type: 'none' };
    if (stroke) o.line = { color: hex(stroke), width: (s.strokeW || 4) * 0.5 };
  }
  if (s.kind === 'round') o.rectRadius = Math.min(0.5, (s.radius || 24) / Math.min(e.w, e.h));
  return o;
}

export function chartPalette(theme) {
  return [theme.accent, '#f0603f', '#2f9e6e', '#f5b400', '#9b5de5', '#00b4d8', '#ef476f', '#6c757d'];
}

function imageSize(src) {
  return new Promise((resolve) => {
    const i = new Image();
    i.onload = () => resolve({ w: i.naturalWidth, h: i.naturalHeight });
    i.onerror = () => resolve(null);
    i.src = src;
  });
}

export { bgCss, W, H };
