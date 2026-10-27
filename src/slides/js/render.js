// Renderização de slides em DOM (edição, miniaturas, apresentação e exportação).
import { THEMES, W, H } from './themes.js';

export const SHAPES = {
  rect: 'Retângulo', round: 'Retângulo arredondado', ellipse: 'Elipse', triangle: 'Triângulo', diamond: 'Losango', hexagon: 'Hexágono',
  star: 'Estrela', arrow: 'Seta', chevron: 'Divisa', callout: 'Balão', line: 'Linha',
};

function starPoints(n = 5, inner = 0.45) {
  const pts = [];
  for (let i = 0; i < n * 2; i++) {
    const r = i % 2 ? 50 * inner : 50;
    const a = (i * Math.PI) / n - Math.PI / 2;
    pts.push(`${(50 + Math.cos(a) * r).toFixed(2)},${(50 + Math.sin(a) * r * 1.05 + 2).toFixed(2)}`);
  }
  return pts.join(' ');
}

const PATHS = {
  triangle: '<polygon points="50,0 100,100 0,100"/>',
  diamond: '<polygon points="50,0 100,50 50,100 0,50"/>',
  hexagon: '<polygon points="25,0 75,0 100,50 75,100 25,100 0,50"/>',
  arrow: '<polygon points="0,30 62,30 62,6 100,50 62,94 62,70 0,70"/>',
  chevron: '<polygon points="0,0 72,0 100,50 72,100 0,100 28,50"/>',
  star: `<polygon points="${starPoints()}"/>`,
  callout: '<path d="M8,0 H92 Q100,0 100,8 V64 Q100,72 92,72 H40 L18,100 L24,72 H8 Q0,72 0,64 V8 Q0,0 8,0 Z"/>',
};

export function themeOf(pres) {
  return THEMES[pres.theme] || THEMES.claro;
}

export function resolveColor(c, theme, fallback) {
  if (!c) return fallback;
  if (c === 'accent') return theme.accent;
  if (c === 'accent2') return theme.accent2;
  if (c === 'title') return theme.title;
  if (c === 'text') return theme.text;
  return c;
}

export function bgCss(bg, theme, assets = {}) {
  const b = bg || theme.bg;
  if (b.type === 'gradient') return `linear-gradient(${b.angle ?? 135}deg, ${b.from}, ${b.to})`;
  if (b.type === 'image' && assets[b.asset]) return `center / cover no-repeat url("${assets[b.asset]}")`;
  return b.color || '#ffffff';
}

const esc = (s) => String(s ?? '').replace(/[&<>"]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' })[c]);

// opts: { mode: 'edit'|'thumb'|'present'|'export', chartImages, editingId }
export function renderSlide(slide, pres, opts = {}) {
  const theme = themeOf(pres);
  const mode = opts.mode || 'edit';
  const root = document.createElement('div');
  root.className = `slide mode-${mode}`;
  root.style.cssText = `width:${W}px;height:${H}px;position:relative;overflow:hidden;background:${bgCss(slide.bg, theme, pres.assets)};`;
  slide.elements.forEach((e, i) => {
    const node = renderElement(e, theme, pres, opts);
    node.style.zIndex = String(i + 1);
    root.append(node);
  });
  return root;
}

export function renderElement(e, theme, pres, opts = {}) {
  const mode = opts.mode || 'edit';
  const n = document.createElement('div');
  n.className = `sl-el sl-${e.type}`;
  n.dataset.id = e.id;
  const style = [
    'position:absolute', `left:${e.x}px`, `top:${e.y}px`, `width:${e.w}px`, `height:${e.h}px`,
    `transform:rotate(${e.rot || 0}deg)`, `opacity:${e.opacity ?? 1}`,
  ];
  if (e.shadow) style.push('filter:drop-shadow(0 18px 32px rgba(0,0,0,.28))');
  n.style.cssText = style.join(';');

  if (e.type === 'shape') {
    const s = e.shape;
    const fill = resolveColor(s.fill, theme, theme.accent);
    const stroke = s.stroke ? resolveColor(s.stroke, theme, '#000') : null;
    if (s.kind === 'line') {
      n.innerHTML = `<svg viewBox="0 0 100 100" preserveAspectRatio="none" style="position:absolute;inset:0;width:100%;height:100%;overflow:visible"><line x1="0" y1="50" x2="100" y2="50" stroke="${stroke || fill}" stroke-width="${Math.max(2, s.strokeW || 6)}" vector-effect="non-scaling-stroke" stroke-linecap="round"/></svg>`;
    } else if (['rect', 'round', 'ellipse'].includes(s.kind)) {
      const bgDiv = document.createElement('div');
      bgDiv.style.cssText = `position:absolute;inset:0;background:${fill || 'transparent'};border-radius:${s.kind === 'ellipse' ? '50%' : s.kind === 'round' ? `${s.radius ?? 24}px` : `${s.kind === 'rect' && s.radius && s.radius < 12 ? s.radius : 0}px`};${stroke ? `border:${s.strokeW || 4}px solid ${stroke};` : ''}`;
      n.append(bgDiv);
    } else {
      n.insertAdjacentHTML('beforeend', `<svg viewBox="0 0 100 100" preserveAspectRatio="none" style="position:absolute;inset:0;width:100%;height:100%;overflow:visible"><g fill="${fill || 'none'}" stroke="${stroke || 'none'}" stroke-width="${s.strokeW || 0}" vector-effect="non-scaling-stroke" stroke-linejoin="round">${PATHS[s.kind] || ''}</g></svg>`);
      n.querySelectorAll('polygon,path').forEach((p) => p.setAttribute('vector-effect', 'non-scaling-stroke'));
    }
    if (e.text && (e.text.html || mode === 'edit')) n.append(textBox(e, theme, mode, opts));
  } else if (e.type === 'text') {
    n.append(textBox(e, theme, mode, opts));
  } else if (e.type === 'image') {
    const src = pres.assets?.[e.image.asset];
    if (src) {
      const img = document.createElement('img');
      img.src = src;
      img.alt = e.alt || '';
      img.draggable = false;
      img.style.cssText = `position:absolute;inset:0;width:100%;height:100%;object-fit:${e.image.fit || 'cover'};border-radius:${e.image.radius || 0}px;${e.image.filter ? `filter:${e.image.filter};` : ''}`;
      n.append(img);
    } else if (mode === 'edit' || mode === 'thumb') {
      n.innerHTML = `<div class="img-ph" style="position:absolute;inset:0;display:flex;flex-direction:column;align-items:center;justify-content:center;gap:16px;background:${theme.accent2};color:${theme.text};font:32px ${theme.body};border-radius:${e.image.radius || 0}px"><svg viewBox="0 0 24 24" width="96" height="96" fill="none" stroke="currentColor" stroke-width="1.5"><rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="9" cy="9" r="2"/><path d="m21 15-5-5L5 21"/></svg>${mode === 'edit' ? `<span>${esc(e.ph || 'Imagem')}</span>` : ''}</div>`;
    }
  } else if (e.type === 'table') {
    n.append(tableBox(e, theme, mode));
  } else if (e.type === 'chart') {
    const img = opts.chartImages?.[e.id];
    if (mode === 'edit') {
      const c = document.createElement('canvas');
      c.className = 'sl-chart-canvas';
      c.style.cssText = 'position:absolute;inset:0;width:100%;height:100%';
      n.append(c);
    } else if (img) {
      n.innerHTML = `<img src="${img}" style="position:absolute;inset:0;width:100%;height:100%;object-fit:contain">`;
    } else {
      n.innerHTML = `<div style="position:absolute;inset:0;display:grid;place-items:center;background:${theme.accent2};color:${theme.text};font:40px ${theme.body}">📊</div>`;
    }
  }
  if (mode === 'present' && e.anim && e.anim.type !== 'none') n.dataset.anim = e.anim.type;
  return n;
}

function textBox(e, theme, mode, opts) {
  const t = e.text;
  const box = document.createElement('div');
  box.className = 'sl-textbox';
  const pad = e.type === 'shape' ? 24 : 6;
  box.style.cssText = `position:absolute;inset:0;padding:${pad}px;display:flex;flex-direction:column;justify-content:${{ top: 'flex-start', middle: 'center', bottom: 'flex-end' }[t.valign || 'top']};`;
  const inner = document.createElement('div');
  inner.className = 'sl-text';
  const role = e.role || 'body';
  const color = resolveColor(t.color, theme, role === 'title' || role === 'big' || role === 'quote' ? theme.title : theme.text);
  const font = t.font === 'head' || (!t.font && (role === 'title' || role === 'big' || role === 'quote')) ? theme.head : t.font === 'body' || !t.font ? theme.body : t.font;
  inner.style.cssText = `font-family:${font};font-size:${t.size}px;line-height:${t.line || 1.25};color:${color};text-align:${t.align || 'left'};font-weight:${t.bold ? 700 : 400};${t.italic ? 'font-style:italic;' : ''}outline:none;word-wrap:break-word;--accent:${theme.accent};${role === 'title' ? 'letter-spacing:-0.01em;' : ''}`;
  inner.innerHTML = t.html || '';
  if (!t.html && e.ph && (mode === 'edit' || mode === 'thumb') && opts.editingId !== e.id) {
    inner.dataset.ph = mode === 'edit' ? e.ph : '';
    inner.classList.add('is-empty');
  }
  box.append(inner);
  return box;
}

function tableBox(e, theme, mode) {
  const t = e.table;
  const wrap = document.createElement('div');
  wrap.style.cssText = 'position:absolute;inset:0;overflow:hidden';
  const fs = t.size || 30;
  let html = `<table class="sl-table" style="width:100%;height:100%;border-collapse:collapse;table-layout:fixed;font:${fs}px ${theme.body};color:${theme.text}">`;
  t.rows.forEach((row, r) => {
    html += '<tr>';
    row.forEach((cell, c) => {
      const head = t.header && r === 0;
      const bg = head ? theme.accent : r % 2 ? theme.accent2 : 'transparent';
      const color = head ? (isDark(theme.accent) ? '#ffffff' : '#111111') : theme.text;
      html += `<td data-r="${r}" data-c="${c}" style="padding:12px 18px;border-bottom:2px solid ${theme.accent2};background:${bg};color:${color};font-weight:${head ? 700 : 400};vertical-align:middle">${cell || ''}</td>`;
    });
    html += '</tr>';
  });
  html += '</table>';
  wrap.innerHTML = html;
  void mode;
  return wrap;
}

export function isDark(hex) {
  const m = /^#?([0-9a-f]{6})$/i.exec(String(hex).trim());
  if (!m) return false;
  const n = parseInt(m[1], 16);
  const r = (n >> 16) & 255, g = (n >> 8) & 255, b = n & 255;
  return 0.299 * r + 0.587 * g + 0.114 * b < 150;
}

// CSS compartilhado pelo editor e pela exportação.
export const SLIDE_CSS = `
.slide { font-kerning: normal; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
.sl-text p { margin: 0 0 0.35em; }
.sl-text p:last-child { margin-bottom: 0; }
.sl-text ul, .sl-text ol { margin: 0; padding-left: 1.2em; }
.sl-text li { margin: 0 0 0.45em; padding-left: 0.2em; }
.sl-text li::marker { color: var(--accent); }
.sl-text a { color: var(--accent); }
.sl-text b, .sl-text strong { font-weight: 700; }
`;

export { W, H };
