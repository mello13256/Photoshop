// Formatos de arquivo do Lumina Texto: .lumdoc (nativo), DOCX, PDF, HTML, Markdown e TXT.
import { loadVendor, extOf, baseName, decodeText, bytesToDataUrl } from '../../shared/native.js';
import { escapeHtml } from '../../shared/ui.js';

export const PAGE_SIZES = {
  A4: ['A4 (21 × 29,7 cm)', 21, 29.7],
  Letter: ['Carta (21,6 × 27,9 cm)', 21.59, 27.94],
  Legal: ['Ofício (21,6 × 35,6 cm)', 21.59, 35.56],
  A5: ['A5 (14,8 × 21 cm)', 14.8, 21],
  A3: ['A3 (29,7 × 42 cm)', 29.7, 42],
};

export const STYLE_PRESETS = {
  padrao: { label: 'Padrão', font: 'Calibri, Carlito, "Segoe UI", Arial, sans-serif', headFont: '', size: 11, line: 1.5, space: 8, indent: 0, align: 'left', headColor: '#1f2a44', accent: '#3b6cf6', headCase: 'none', headScale: 1 },
  abnt: { label: 'ABNT (trabalhos acadêmicos)', font: 'Arial, Helvetica, sans-serif', headFont: '', size: 12, line: 1.5, space: 0, indent: 1.25, align: 'justify', headColor: '#000000', accent: '#000000', headCase: 'uppercase', headScale: 0 },
  moderno: { label: 'Moderno', font: '"Segoe UI Variable", "Segoe UI", Inter, system-ui, sans-serif', headFont: '', size: 11, line: 1.65, space: 10, indent: 0, align: 'left', headColor: '#3b6cf6', accent: '#3b6cf6', headCase: 'none', headScale: 1 },
  classico: { label: 'Clássico', font: 'Georgia, "Times New Roman", serif', headFont: 'Georgia, serif', size: 12, line: 1.5, space: 6, indent: 0.75, align: 'justify', headColor: '#222222', accent: '#7a4b1e', headCase: 'none', headScale: 1 },
  compacto: { label: 'Compacto', font: 'Arial, Helvetica, sans-serif', headFont: '', size: 10, line: 1.2, space: 4, indent: 0, align: 'left', headColor: '#111827', accent: '#2563eb', headCase: 'none', headScale: 0.85 },
};

export function defaultPage() {
  return { size: 'A4', orientation: 'portrait', margins: { top: 2.5, right: 2.5, bottom: 2.5, left: 2.5 }, pageNumbers: true, header: '', footer: '' };
}

export function defaultStyle(preset = 'padrao') {
  return { preset, ...STYLE_PRESETS[preset] };
}

export function pageDims(page) {
  const [, w, h] = PAGE_SIZES[page.size] || PAGE_SIZES.A4;
  return page.orientation === 'landscape' ? [h, w] : [w, h];
}

// Variáveis CSS do estilo do documento.
export function styleVars(st) {
  const scale = st.headScale;
  const v = {
    '--doc-font': st.font,
    '--doc-head-font': st.headFont || st.font,
    '--doc-size': `${st.size}pt`,
    '--doc-line': st.line,
    '--doc-space': `${st.space}pt`,
    '--doc-indent': `${st.indent}cm`,
    '--doc-align': st.align,
    '--doc-head-color': st.headColor,
    '--doc-accent': st.accent,
    '--doc-head-case': st.headCase,
  };
  return Object.entries(v).map(([k, x]) => `${k}:${x}`).join(';') + (scale === 0 ? ';--abnt:1' : '');
}

export function applyStyleVars(node, st) {
  node.setAttribute('style', styleVars(st));
  node.classList.toggle('abnt-heads', st.headScale === 0);
  node.style.setProperty('--hs', st.headScale || 1);
}

// ---- Nativo ----------------------------------------------------------------

export function serialize(doc, html) {
  return JSON.stringify({
    app: 'Lumina Texto', format: 1, title: doc.title, html, page: doc.page, style: doc.style,
    versions: doc.versions.slice(-25), savedAt: new Date().toISOString(),
  });
}

export function parseNative(text) {
  const d = JSON.parse(text);
  if (d.format !== 1 || typeof d.html !== 'string') throw new Error('Arquivo .lumdoc inválido.');
  return d;
}

// ---- Importação --------------------------------------------------------------

let purify = null;
export async function ensurePurify() {
  purify = await loadVendor('purify.min.js', 'DOMPurify');
  return purify;
}

export function sanitize(html) {
  if (!purify) return escapeHtml(html.replace(/<[^>]+>/g, ' '));
  return purify.sanitize(html, {
    FORBID_TAGS: ['style', 'script', 'iframe', 'object', 'embed', 'form', 'input', 'button', 'meta', 'link'],
    FORBID_ATTR: ['contenteditable'],
    ALLOW_DATA_ATTR: true,
  });
}

// Converte arquivo aberto em { title, html, page?, style?, versions? }.
export async function importFile({ name, data }) {
  const ext = extOf(name);
  const title = baseName(name);
  if (ext === 'lumdoc') return { ...parseNative(decodeText(data)), native: true };
  if (ext === 'docx' || ext === 'docm' || ext === 'dotx') {
    try {
      const { importDocx } = await import('./docx-import.js');
      const res = await importDocx(data, defaultStyle('padrao'));
      return { title: res.title || title, html: sanitize(res.html) || '<p><br></p>', page: { ...defaultPage(), ...res.page }, style: res.style };
    } catch (err) {
      console.warn('Importador próprio falhou, usando o modo simplificado:', err);
      return importDocxSimple(data, title);
    }
  }
  const text = decodeText(data);
  if (ext === 'md' || ext === 'markdown') return { title, html: mdToHtml(text) };
  if (ext === 'html' || ext === 'htm') {
    const body = /<body[^>]*>([\s\S]*)<\/body>/i.exec(text)?.[1] ?? text;
    return { title, html: sanitize(body) };
  }
  return { title, html: text.split(/\r?\n/).map((l) => `<p>${escapeHtml(l) || '<br>'}</p>`).join('') };
}

async function importDocxSimple(data, title) {
  const mammoth = await loadVendor('mammoth.browser.min.js', 'mammoth');
  const res = await mammoth.convertToHtml(
    { arrayBuffer: data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength) },
    {
      styleMap: ["p[style-name='Title'] => p.title:fresh", "p[style-name='Subtitle'] => p.subtitle:fresh", "p[style-name='Quote'] => blockquote:fresh"],
      convertImage: mammoth.images.imgElement((image) => image.read('base64').then((b64) => ({ src: `data:${image.contentType};base64,${b64}` }))),
    });
  return { title, html: sanitize(res.value) || '<p><br></p>' };
}

// ---- Markdown ------------------------------------------------------------------

function inlineMd(s) {
  let t = escapeHtml(s);
  t = t.replace(/!\[([^\]]*)\]\(([^)\s]+)\)/g, (_, alt, src) => `<img src="${src}" alt="${alt}">`);
  t = t.replace(/\[([^\]]+)\]\(([^)\s]+)\)/g, '<a href="$2">$1</a>');
  t = t.replace(/`([^`]+)`/g, '<code>$1</code>');
  t = t.replace(/\*\*([^*]+)\*\*|__([^_]+)__/g, (_, a, b) => `<strong>${a || b}</strong>`);
  t = t.replace(/(^|[^*])\*([^*]+)\*/g, '$1<em>$2</em>').replace(/(^|[^\w_])_([^_]+)_/g, '$1<em>$2</em>');
  t = t.replace(/~~([^~]+)~~/g, '<s>$1</s>');
  return t;
}

export function mdToHtml(md) {
  const lines = md.replace(/\r\n?/g, '\n').split('\n');
  const out = [];
  let i = 0;
  const listStack = [];
  const closeLists = (depth = 0) => {
    while (listStack.length > depth) out.push(`</${listStack.pop().tag}>`);
  };
  while (i < lines.length) {
    const line = lines[i];
    if (/^```/.test(line)) {
      closeLists();
      const buf = [];
      i++;
      while (i < lines.length && !/^```/.test(lines[i])) buf.push(lines[i++]);
      i++;
      out.push(`<pre><code>${escapeHtml(buf.join('\n'))}</code></pre>`);
      continue;
    }
    let m;
    if ((m = /^(#{1,6})\s+(.*)$/.exec(line))) {
      closeLists();
      const lvl = Math.min(4, m[1].length);
      out.push(`<h${lvl}>${inlineMd(m[2])}</h${lvl}>`);
      i++;
      continue;
    }
    if (/^\s*([-*_])\s*\1\s*\1[\s\1]*$/.test(line)) { closeLists(); out.push('<hr>'); i++; continue; }
    if (/^>\s?/.test(line)) {
      closeLists();
      const buf = [];
      while (i < lines.length && /^>\s?/.test(lines[i])) buf.push(lines[i++].replace(/^>\s?/, ''));
      out.push(`<blockquote>${inlineMd(buf.join(' '))}</blockquote>`);
      continue;
    }
    if (/^\|.*\|\s*$/.test(line) && /^\|?\s*:?-+/.test(lines[i + 1] || '')) {
      closeLists();
      const row = (l) => l.trim().replace(/^\||\|$/g, '').split('|').map((c) => inlineMd(c.trim()));
      const head = row(line);
      i += 2;
      const body = [];
      while (i < lines.length && /^\|.*\|\s*$/.test(lines[i])) body.push(row(lines[i++]));
      out.push(`<table><tbody><tr>${head.map((c) => `<th>${c}</th>`).join('')}</tr>${body.map((r) => `<tr>${r.map((c) => `<td>${c || '<br>'}</td>`).join('')}</tr>`).join('')}</tbody></table>`);
      continue;
    }
    if ((m = /^(\s*)([-*+]|\d+[.)])\s+(.*)$/.exec(line))) {
      const depth = Math.floor(m[1].replace(/\t/g, '  ').length / 2) + 1;
      const ordered = /\d/.test(m[2]);
      let content = m[3];
      let check = null;
      const cm = /^\[( |x|X)\]\s+(.*)$/.exec(content);
      if (cm) { check = cm[1].toLowerCase() === 'x'; content = cm[2]; }
      while (listStack.length > depth) out.push(`</${listStack.pop().tag}>`);
      if (listStack.length === depth && listStack[depth - 1].tag !== (ordered ? 'ol' : 'ul')) out.push(`</${listStack.pop().tag}>`);
      while (listStack.length < depth) {
        const tag = ordered ? 'ol' : 'ul';
        listStack.push({ tag });
        out.push(check !== null ? '<ul class="checklist">' : `<${tag}>`);
      }
      out.push(`<li${check !== null ? ` data-checked="${check}"` : ''}>${inlineMd(content)}</li>`);
      i++;
      continue;
    }
    if (!line.trim()) { closeLists(); i++; continue; }
    closeLists();
    const buf = [];
    while (i < lines.length && lines[i].trim() && !/^(#{1,6}\s|```|>|\s*[-*+]\s|\s*\d+[.)]\s|\|)/.test(lines[i])) buf.push(lines[i++]);
    if (!buf.length) buf.push(lines[i++]);
    out.push(`<p>${inlineMd(buf.join(' '))}</p>`);
  }
  closeLists();
  return out.join('') || '<p><br></p>';
}

export function htmlToMd(html) {
  const root = document.createElement('div');
  root.innerHTML = html;
  const inline = (n) => {
    if (n.nodeType === 3) return n.textContent.replace(/​/g, '');
    if (n.nodeType !== 1) return '';
    const inner = [...n.childNodes].map(inline).join('');
    switch (n.tagName) {
      case 'STRONG': case 'B': return inner.trim() ? `**${inner}**` : inner;
      case 'EM': case 'I': return inner.trim() ? `*${inner}*` : inner;
      case 'S': case 'STRIKE': case 'DEL': return `~~${inner}~~`;
      case 'CODE': return `\`${inner}\``;
      case 'A': return `[${inner}](${n.getAttribute('href') || ''})`;
      case 'IMG': return `![${n.alt || ''}](${n.getAttribute('src')})`;
      case 'BR': return '  \n';
      default: return inner;
    }
  };
  const block = (n, depth = 0) => {
    if (n.nodeType === 3) return n.textContent.trim() ? n.textContent : '';
    if (n.nodeType !== 1) return '';
    const t = n.tagName;
    if (/^H[1-6]$/.test(t)) return `${'#'.repeat(+t[1])} ${inline(n).trim()}\n\n`;
    if (t === 'P' && n.classList.contains('title')) return `# ${inline(n).trim()}\n\n`;
    if (t === 'P' || t === 'DIV' && !n.classList.length) return `${inline(n).trim()}\n\n`;
    if (t === 'BLOCKQUOTE') return `> ${inline(n).trim()}\n\n`;
    if (t === 'PRE') return `\`\`\`\n${n.textContent}\n\`\`\`\n\n`;
    if (t === 'HR') return '---\n\n';
    if (t === 'UL' || t === 'OL') {
      const check = n.classList.contains('checklist');
      let i = 1;
      const s = [...n.children].map((li) => {
        const nested = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
        const own = [...li.childNodes].filter((c) => !nested.includes(c)).map(inline).join('').trim();
        const marker = check ? `- [${li.dataset.checked === 'true' ? 'x' : ' '}]` : t === 'OL' ? `${i++}.` : '-';
        return `${'  '.repeat(depth)}${marker} ${own}\n${nested.map((x) => block(x, depth + 1)).join('')}`;
      }).join('');
      return depth ? s : `${s}\n`;
    }
    if (t === 'TABLE') {
      const rows = [...n.rows].map((r) => [...r.cells].map((c) => inline(c).trim().replace(/\|/g, '\\|')));
      if (!rows.length) return '';
      const w = Math.max(...rows.map((r) => r.length));
      const line = (r) => `| ${Array.from({ length: w }, (_, i) => r[i] || '').join(' | ')} |`;
      return `${line(rows[0])}\n| ${Array(w).fill('---').join(' | ')} |\n${rows.slice(1).map(line).join('\n')}\n\n`;
    }
    if (n.classList.contains('page-break') || n.classList.contains('toc')) return '';
    return [...n.childNodes].map((c) => block(c, depth)).join('');
  };
  return [...root.childNodes].map((n) => block(n)).join('').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

// ---- HTML e PDF --------------------------------------------------------------

let contentCss = null;
export async function getContentCss() {
  if (!contentCss) contentCss = await (await fetch(new URL('../content.css', import.meta.url))).text();
  return contentCss;
}

export async function standaloneHtml(doc, html, { forPrint = false } = {}) {
  const css = await getContentCss();
  const [w, h] = pageDims(doc.page);
  const m = doc.page.margins;
  const vars = styleVars(doc.style);
  return `<!doctype html>
<html lang="pt-BR"><head><meta charset="utf-8"><title>${escapeHtml(doc.title)}</title>
<style>
${css}
@page { size: ${w}cm ${h}cm; margin: ${m.top}cm ${m.right}cm ${m.bottom}cm ${m.left}cm; }
html, body { margin: 0; background: #fff; }
${forPrint ? '' : `.doc-content { max-width: ${w - m.left - m.right}cm; margin: 2cm auto; }`}
.abnt-heads h1, .abnt-heads h2, .abnt-heads h3, .abnt-heads h4 { font-size: 1em; }
.abnt-heads h2 { text-transform: none; }
</style></head>
<body><div class="doc-content${doc.style.headScale === 0 ? ' abnt-heads' : ''}" style="${vars}">${html}</div></body></html>`;
}

// ---- DOCX ----------------------------------------------------------------------

const cmToTwip = (cm) => Math.round(cm * 567);

function cssColorToHex(c) {
  if (!c) return null;
  const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?\)/.exec(c);
  if (m) {
    if (m[4] !== undefined && +m[4] === 0) return null;
    return [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
  }
  const h = /^#([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(c.trim());
  if (h) return (h[1].length === 3 ? h[1].split('').map((x) => x + x).join('') : h[1]).toUpperCase();
  return null;
}

async function imageToPng(src) {
  const img = new Image();
  img.src = src;
  await img.decode();
  const c = document.createElement('canvas');
  c.width = img.naturalWidth;
  c.height = img.naturalHeight;
  c.getContext('2d').drawImage(img, 0, 0);
  const blob = await new Promise((r) => c.toBlob(r, 'image/png'));
  return { data: new Uint8Array(await blob.arrayBuffer()), w: img.naturalWidth, h: img.naturalHeight };
}

export async function exportDocx(doc, html) {
  const D = await loadVendor('docx.iife.js', 'docx');
  const st = doc.style;
  const [pw] = pageDims(doc.page);
  const contentWidthPx = (pw - doc.page.margins.left - doc.page.margins.right) * 37.8;
  const root = document.createElement('div');
  root.innerHTML = html;
  const baseHalfPts = Math.round(st.size * 2);
  const lineVal = Math.round(st.line * 240);
  const headSizes = st.headScale === 0 ? [1, 1, 1, 1] : [1.85, 1.4, 1.17, 1].map((x) => x * (st.headScale || 1));
  const headColor = cssColorToHex(st.headColor) || '1F2A44';
  const firstFont = (f) => (f || '').split(',')[0].replace(/["']/g, '').trim() || 'Calibri';
  let listInstance = 0;
  const NUM_FORMATS = { 1: 'DECIMAL', a: 'LOWER_LETTER', A: 'UPPER_LETTER', i: 'LOWER_ROMAN', I: 'UPPER_ROMAN' };
  const numRef = (list) => (list.getAttribute('type') && NUM_FORMATS[list.getAttribute('type')] ? `lumina-num-${list.getAttribute('type')}` : 'lumina-num');

  const runsOf = async (node, fmt = {}) => {
    const out = [];
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        const text = n.textContent.replace(/​/g, '');
        if (!text) continue;
        const parts = text.split('\t');
        parts.forEach((p, i) => {
          if (i) out.push(new D.TextRun({ ...fmt, children: [new D.Tab()] }));
          if (p) out.push(new D.TextRun({ ...fmt, text: p }));
        });
        continue;
      }
      if (n.nodeType !== 1) continue;
      const t = n.tagName;
      if (t === 'BR') { out.push(new D.TextRun({ ...fmt, break: 1 })); continue; }
      if (t === 'SUP' && footnoteIds.has(n.querySelector('a')?.getAttribute('href'))) {
        out.push(new D.FootnoteReferenceRun(footnoteIds.get(n.querySelector('a').getAttribute('href'))));
        continue;
      }
      if (t === 'IMG') { const ir = await imageRun(n); if (ir) out.push(ir); continue; }
      const f = { ...fmt };
      if (t === 'STRONG' || t === 'B') f.bold = true;
      if (t === 'EM' || t === 'I') f.italics = true;
      if (t === 'U') f.underline = {};
      if (t === 'S' || t === 'STRIKE' || t === 'DEL') f.strike = true;
      if (t === 'SUP') f.superScript = true;
      if (t === 'SUB') f.subScript = true;
      if (t === 'CODE') f.font = 'Consolas';
      if (t === 'MARK') {
        const fill = cssColorToHex(n.style.backgroundColor) || 'FFFF00';
        const named = { FFFF00: 'yellow', '00FF00': 'green', '00FFFF': 'cyan', FF00FF: 'magenta', '0000FF': 'blue', FF0000: 'red', '000080': 'darkBlue', '008080': 'darkCyan', '008000': 'darkGreen', '800080': 'darkMagenta', '800000': 'darkRed', '808000': 'darkYellow', '808080': 'darkGray', C0C0C0: 'lightGray' }[fill];
        if (named) f.highlight = named;
        else f.shading = { type: D.ShadingType.CLEAR, fill, color: 'auto' };
      }
      const s = n.style;
      if (s) {
        if (s.fontWeight === 'bold' || +s.fontWeight >= 600) f.bold = true;
        if (s.fontWeight === 'normal' || (+s.fontWeight && +s.fontWeight < 600)) f.bold = false;
        if (s.fontStyle === 'normal') f.italics = false;
        if (s.textTransform === 'uppercase') f.allCaps = true;
        if (s.fontVariant === 'small-caps') f.smallCaps = true;
        const ls = /(-?[\d.]+)pt/.exec(s.letterSpacing || '');
        if (ls) f.characterSpacing = Math.round(+ls[1] * 20);
        if (s.fontStyle === 'italic') f.italics = true;
        if (s.textDecoration?.includes('underline') || s.textDecorationLine?.includes('underline')) f.underline = {};
        if (s.textDecoration?.includes('line-through') || s.textDecorationLine?.includes('line-through')) f.strike = true;
        const col = cssColorToHex(s.color);
        if (col) f.color = col;
        const bg = cssColorToHex(s.backgroundColor);
        if (bg && t !== 'MARK') f.shading = { type: D.ShadingType.CLEAR, fill: bg, color: 'auto' };
        if (s.fontFamily) f.font = firstFont(s.fontFamily);
        const fs = /([\d.]+)(pt|px)/.exec(s.fontSize || '');
        if (fs) f.size = Math.round((fs[2] === 'px' ? +fs[1] * 0.75 : +fs[1]) * 2);
      }
      if (t === 'A' && n.getAttribute('href') && !n.getAttribute('href').startsWith('#')) {
        out.push(new D.ExternalHyperlink({ link: n.href, children: await runsOf(n, { ...f, color: '1155CC', underline: {} }) }));
        continue;
      }
      out.push(...(await runsOf(n, f)));
    }
    return out;
  };

  const imageRun = async (img) => {
    try {
      const { data, w, h } = await imageToPng(img.src);
      const shownW = Math.min(contentWidthPx, img.width || parseFloat(img.style.width) || w);
      const scale = shownW / w;
      return new D.ImageRun({ type: 'png', data, transformation: { width: Math.round(w * scale), height: Math.round(h * scale) } });
    } catch {
      return null;
    }
  };

  // Espaçamento/recuo definidos no próprio bloco (vindos do Word ou do editor)
  const lenPt = (v) => {
    const m = /(-?[\d.]+)(pt|px|cm|mm|em)?/.exec(v || '');
    if (!m) return null;
    const x = +m[1];
    return { pt: x, px: x * 0.75, cm: x * 28.3465, mm: x * 2.83465, em: x * st.size }[m[2] || 'px'];
  };
  const blockProps = (n, ctx = {}) => {
    const s = n.style || {};
    const o = {};
    const spacing = {};
    const mt = lenPt(s.marginTop), mb = lenPt(s.marginBottom);
    if (mt !== null) spacing.before = Math.round(mt * 20);
    if (mb !== null) spacing.after = Math.round(mb * 20);
    if (s.lineHeight) {
      if (/pt|px/.test(s.lineHeight)) { spacing.line = Math.round(lenPt(s.lineHeight) * 20); spacing.lineRule = D.LineRuleType ? D.LineRuleType.EXACT : 'exact'; }
      else if (+s.lineHeight) spacing.line = Math.round(+s.lineHeight * 240);
    }
    if (Object.keys(spacing).length) o.spacing = spacing;
    const indent = {};
    const ml = lenPt(s.marginLeft), mr = lenPt(s.marginRight), ti = lenPt(s.textIndent);
    if (ml) indent.left = Math.round(ml * 20);
    if (mr) indent.right = Math.round(mr * 20);
    if (ti !== null) { if (ti < 0) indent.hanging = Math.round(-ti * 20); else indent.firstLine = Math.round(ti * 20); }
    else if (!ctx.noIndent && st.indent && n.tagName === 'P') indent.firstLine = cmToTwip(st.indent);
    if (Object.keys(indent).length) o.indent = indent;
    const bg = cssColorToHex(s.backgroundColor);
    if (bg) o.shading = { type: D.ShadingType.CLEAR, fill: bg, color: 'auto' };
    return o;
  };

  const alignOf = (n) => {
    const a = n.style?.textAlign || (n.classList?.contains('align-center') ? 'center' : '');
    return { center: D.AlignmentType.CENTER, right: D.AlignmentType.RIGHT, justify: D.AlignmentType.JUSTIFIED, left: D.AlignmentType.LEFT }[a];
  };

  const blocks = async (node, ctx = {}) => {
    const out = [];
    for (const n of node.childNodes) {
      if (n.nodeType === 3) {
        if (n.textContent.trim()) out.push(new D.Paragraph({ children: [new D.TextRun(n.textContent)] }));
        continue;
      }
      if (n.nodeType !== 1) continue;
      const t = n.tagName;
      const common = { alignment: alignOf(n) };
      if (/^H[1-4]$/.test(t) || n.classList.contains('title') || n.classList.contains('subtitle')) {
        const lvl = n.classList.contains('title') ? 0 : n.classList.contains('subtitle') ? -1 : +t[1];
        const size = lvl === 0 ? 2.6 : lvl === -1 ? 1.3 : headSizes[lvl - 1];
        const text = st.headCase === 'uppercase' && lvl === 1 ? n.textContent.toUpperCase() : null;
        const bp = blockProps(n, { noIndent: true });
        out.push(new D.Paragraph({
          ...common,
          ...bp,
          heading: lvl === 0 ? D.HeadingLevel.TITLE : lvl > 0 ? D.HeadingLevel[`HEADING_${lvl}`] : undefined,
          spacing: { before: lvl > 0 ? 240 : 0, after: 120, ...(bp.spacing || {}) },
          keepNext: true,
          children: text ? [new D.TextRun({ text, bold: true, size: Math.round(baseHalfPts * size), color: headColor, font: firstFont(st.headFont || st.font) })]
            : await runsOf(n, { bold: lvl !== -1, size: Math.round(baseHalfPts * size), color: lvl === -1 ? '5F6B7A' : headColor, font: firstFont(st.headFont || st.font) }),
        }));
        continue;
      }
      if (t === 'P' || (t === 'DIV' && !n.classList.contains('page-break') && !n.classList.contains('toc') && !n.classList.contains('callout'))) {
        if (n.querySelector('p, h1, h2, h3, table, ul, ol')) { out.push(...(await blocks(n, ctx))); continue; }
        if (n.classList.contains('footnotes')) continue;
        out.push(new D.Paragraph({
          ...common,
          ...blockProps(n, ctx),
          children: await runsOf(n),
        }));
        continue;
      }
      if (t === 'BLOCKQUOTE') {
        out.push(new D.Paragraph({
          ...common, indent: { left: cmToTwip(1) },
          border: { left: { style: D.BorderStyle.SINGLE, size: 12, color: cssColorToHex(st.accent) || '3B6CF6', space: 8 } },
          children: await runsOf(n, { italics: true, color: '4A5361' }),
        }));
        continue;
      }
      if (t === 'PRE') {
        const lines = n.textContent.split('\n');
        out.push(new D.Paragraph({
          shading: { type: D.ShadingType.CLEAR, fill: 'F4F6F9', color: 'auto' },
          children: lines.flatMap((l, i) => [...(i ? [new D.TextRun({ break: 1 })] : []), new D.TextRun({ text: l, font: 'Consolas', size: Math.round(baseHalfPts * 0.88) })]),
        }));
        continue;
      }
      if (t === 'HR' && n.nextElementSibling?.classList.contains('footnotes')) continue;
      if (t === 'HR') {
        out.push(new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: 'D5DAE3', space: 1 } }, children: [] }));
        continue;
      }
      if (n.classList.contains('page-break')) {
        out.push(new D.Paragraph({ children: [new D.PageBreak()] }));
        continue;
      }
      if (n.classList.contains('toc')) {
        out.push(new D.Paragraph({ children: [new D.TextRun({ text: n.querySelector('.toc-title')?.textContent || 'Sumário', bold: true })] }));
        for (const a of n.querySelectorAll('a')) {
          const lvl = a.classList.contains('l3') ? 2 : a.classList.contains('l2') ? 1 : 0;
          out.push(new D.Paragraph({ indent: { left: lvl * 400 }, children: [new D.TextRun(a.textContent)] }));
        }
        continue;
      }
      if (n.classList.contains('callout')) {
        const body = n.querySelector('.callout-body') || n;
        out.push(new D.Paragraph({
          shading: { type: D.ShadingType.CLEAR, fill: 'EEF4FF', color: 'auto' },
          children: [new D.TextRun(`${n.querySelector('.callout-icon')?.textContent || ''}  `), ...(await runsOf(body))],
        }));
        continue;
      }
      if (t === 'UL' || t === 'OL') {
        out.push(...(await listItems(n, 0, t === 'OL' ? listInstance++ : null)));
        continue;
      }
      if (t === 'TABLE') {
        out.push(await table(n));
        out.push(new D.Paragraph({ children: [] }));
        continue;
      }
      if (t === 'FIGURE') { out.push(...(await blocks(n, ctx))); continue; }
      out.push(...(await blocks(n, ctx)));
    }
    return out;
  };

  const listItems = async (list, level, instance) => {
    const out = [];
    const check = list.classList.contains('checklist');
    for (const li of list.children) {
      if (li.tagName !== 'LI') continue;
      const nested = [...li.children].filter((c) => c.tagName === 'UL' || c.tagName === 'OL');
      const tmp = li.cloneNode(true);
      tmp.querySelectorAll(':scope > ul, :scope > ol').forEach((x) => x.remove());
      const runs = await runsOf(tmp);
      if (check) runs.unshift(new D.TextRun(li.dataset.checked === 'true' ? '☒ ' : '☐ '));
      out.push(new D.Paragraph({
        children: runs,
        ...(check ? { indent: { left: 360 * (level + 1) } }
          : instance !== null ? { numbering: { reference: numRef(list), level, instance } } : { bullet: { level } }),
      }));
      for (const nl of nested) out.push(...(await listItems(nl, level + 1, nl.tagName === 'OL' ? listInstance++ : null)));
    }
    return out;
  };

  const borderOf = (css) => {
    if (!css || /none/.test(css)) return { style: D.BorderStyle.NONE, size: 0, color: 'auto' };
    const w = parseFloat(css) || 1;
    const color = cssColorToHex(css.match(/rgba?\([^)]+\)|#[0-9a-f]{3,6}/i)?.[0] || '') || '000000';
    return { style: /double/.test(css) ? D.BorderStyle.DOUBLE : /dashed/.test(css) ? D.BorderStyle.DASHED : /dotted/.test(css) ? D.BorderStyle.DOTTED : D.BorderStyle.SINGLE, size: Math.max(2, Math.round(w * 6)), color };
  };

  const table = async (tbl) => {
    const rows = [];
    const cols = [...tbl.querySelectorAll(':scope > colgroup > col')].map((c) => parseFloat(c.style.width) || 0);
    const tableW = /px/.test(tbl.style.width) ? parseFloat(tbl.style.width) : null;
    for (const tr of tbl.rows) {
      const cells = [];
      for (const cell of tr.cells) {
        const isHead = cell.tagName === 'TH';
        const cs = cell.style;
        const inner = await blocks(cell, { noIndent: true });
        const headFmt = isHead && cs.fontWeight !== 'normal' ? { bold: true } : {};
        const kids = inner.length ? inner : [new D.Paragraph({ children: await runsOf(cell, headFmt) })];
        const bg = cssColorToHex(cs.backgroundColor);
        const fill = bg || (isHead && !cs.backgroundColor ? 'F1F4F9' : null);
        const borders = {};
        if (cs.borderTopStyle || cs.border) for (const [k, side] of [['top', 'Top'], ['bottom', 'Bottom'], ['left', 'Left'], ['right', 'Right']]) {
          const style = cs[`border${side}Style`];
          if (style) borders[k] = borderOf(style === 'none' ? 'none' : `${cs[`border${side}Width`]} ${style} ${cs[`border${side}Color`]}`);
        }
        cells.push(new D.TableCell({
          children: kids,
          columnSpan: cell.colSpan > 1 ? cell.colSpan : undefined,
          rowSpan: cell.rowSpan > 1 ? cell.rowSpan : undefined,
          shading: fill ? { type: D.ShadingType.CLEAR, fill, color: 'auto' } : undefined,
          verticalAlign: cs.verticalAlign === 'middle' ? D.VerticalAlign.CENTER : cs.verticalAlign === 'bottom' ? D.VerticalAlign.BOTTOM : undefined,
          borders: Object.keys(borders).length ? borders : undefined,
          margins: { top: 60, bottom: 60, left: 100, right: 100 },
        }));
      }
      rows.push(new D.TableRow({ children: cells, tableHeader: tr.querySelector('th') ? true : undefined }));
    }
    const widthTw = tableW ? Math.round(tableW * 15) : Math.round(contentWidthPx * 15);
    return new D.Table({
      rows,
      width: tableW ? { size: widthTw, type: D.WidthType.DXA } : { size: 100, type: D.WidthType.PERCENTAGE },
      columnWidths: cols.length && cols.every((x) => x > 0) ? cols.map((pct) => Math.round((pct / 100) * widthTw)) : undefined,
      alignment: tbl.style.marginLeft === 'auto' ? D.AlignmentType.CENTER : undefined,
    });
  };

  // Notas de rodapé (geradas pelo importador ou pelo usuário no mesmo formato)
  const footnoteIds = new Map();
  const footnotes = {};
  let fnSeq = 1;
  for (const p of root.querySelectorAll('.footnotes > p[id^="nota-"]')) {
    const id = fnSeq++;
    footnoteIds.set(`#${p.id}`, id);
    const clone = p.cloneNode(true);
    clone.querySelector('sup')?.remove();
    footnotes[id] = { children: [new D.Paragraph({ children: await runsOf(clone) })] };
  }

  const children = await blocks(root);
  const [w, h] = pageDims({ ...doc.page, orientation: 'portrait' });
  const m = doc.page.margins;
  const document_ = new D.Document({
    title: doc.title,
    creator: 'Lumina Texto',
    styles: {
      default: {
        document: {
          run: { font: firstFont(st.font), size: baseHalfPts },
          paragraph: { spacing: { line: lineVal, after: Math.round(st.space * 20) }, alignment: st.align === 'justify' ? D.AlignmentType.JUSTIFIED : undefined },
        },
      },
    },
    numbering: {
      config: [null, ...Object.keys(NUM_FORMATS)].map((type) => ({
        reference: type ? `lumina-num-${type}` : 'lumina-num',
        levels: Array.from({ length: 9 }, (_, l) => ({
          level: l,
          format: type && l === 0 ? D.LevelFormat[NUM_FORMATS[type]] : [D.LevelFormat.DECIMAL, D.LevelFormat.LOWER_LETTER, D.LevelFormat.LOWER_ROMAN][l % 3],
          text: `%${l + 1}${type === 'a' || type === 'A' ? ')' : '.'}`,
          alignment: D.AlignmentType.START,
          style: { paragraph: { indent: { left: 720 * (l + 1), hanging: 360 } } },
        })),
      })),
    },
    footnotes,
    sections: [{
      properties: {
        page: {
          size: { width: cmToTwip(w), height: cmToTwip(h), orientation: doc.page.orientation === 'landscape' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT },
          margin: { top: cmToTwip(m.top), right: cmToTwip(m.right), bottom: cmToTwip(m.bottom), left: cmToTwip(m.left) },
        },
      },
      headers: doc.page.header ? {
        default: new D.Header({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.TextRun({ text: doc.page.header, size: 18, color: '666666' })] })] }),
      } : undefined,
      footers: doc.page.pageNumbers || doc.page.footer ? {
        default: new D.Footer({
          children: [new D.Paragraph({
            alignment: D.AlignmentType.CENTER,
            children: [
              ...(doc.page.footer ? [new D.TextRun({ text: doc.page.footer, size: 18, color: '666666' })] : []),
              ...(doc.page.footer && doc.page.pageNumbers ? [new D.TextRun({ text: '   ·   ', size: 18, color: '999999' })] : []),
              ...(doc.page.pageNumbers ? [new D.TextRun({ children: [D.PageNumber.CURRENT], size: 18, color: '777777' })] : []),
            ],
          })],
        }),
      } : undefined,
      children: children.length ? children : [new D.Paragraph({ children: [] })],
    }],
  });
  return D.Packer.toBlob(document_);
}

export function plainText(html) {
  const d = document.createElement('div');
  d.innerHTML = html.replace(/<\/(p|h[1-6]|li|blockquote|pre|tr|div)>/gi, '</$1>\n').replace(/<br\s*\/?>/gi, '\n');
  return d.textContent.replace(/​/g, '').replace(/\n{3,}/g, '\n\n').trim() + '\n';
}

export { bytesToDataUrl };
