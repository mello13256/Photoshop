// Importador de .docx próprio. Preserva fontes, tamanhos, cores, realces, espaçamentos, recuos,
// listas, tabelas (mesclas, sombreamento, bordas), imagens, links, notas de rodapé, cabeçalho/rodapé
// e configuração de página.
import {
  openPackage, kids, kid, path, attr, num, flag, first, all, wordColor, readTheme, cssFont, isWebImage, unwrapAlternate, EMU_PER_PX,
} from '../../shared/ooxml.js';
import { escapeHtml } from '../../shared/ui.js';
import { PAGE_SIZES } from './io.js';

const TWIP_PER_CM = 567;
const twipCm = (t) => Math.round((t / TWIP_PER_CM) * 100) / 100;
const twipPt = (t) => Math.round((t / 20) * 10) / 10;

const HIGHLIGHT = {
  yellow: '#ffff00', green: '#00ff00', cyan: '#00ffff', magenta: '#ff00ff', blue: '#0000ff', red: '#ff0000',
  darkBlue: '#000080', darkCyan: '#008080', darkGreen: '#008000', darkMagenta: '#800080', darkRed: '#800000',
  darkYellow: '#808000', darkGray: '#808080', lightGray: '#c0c0c0', black: '#000000', white: '#ffffff',
};

const OL_TYPES = { decimal: '1', lowerLetter: 'a', upperLetter: 'A', lowerRoman: 'i', upperRoman: 'I', decimalZero: '1' };

// ---------------------------------------------------------------------------
// Propriedades

function readRPr(n, theme) {
  const o = {};
  if (!n) return o;
  const fonts = kid(n, 'rFonts');
  if (fonts) {
    const f = attr(fonts, 'ascii') || attr(fonts, 'hAnsi') || attr(fonts, 'cs');
    const tf = attr(fonts, 'asciiTheme') || attr(fonts, 'hAnsiTheme');
    if (f) o.font = f;
    else if (tf) o.font = /^major/.test(tf) ? theme.major : theme.minor;
  }
  const sz = kid(n, 'sz');
  if (sz) o.size = num(sz, 'val', 22) / 2;
  const color = kid(n, 'color');
  if (color) {
    const v = attr(color, 'val');
    o.color = v === 'auto' && !attr(color, 'themeColor') ? 'auto' : wordColor(color, theme) || 'auto';
  }
  for (const [tag, key] of [['b', 'b'], ['i', 'i'], ['strike', 'strike'], ['dstrike', 'strike'], ['caps', 'caps'], ['smallCaps', 'smallCaps'], ['vanish', 'hidden']]) {
    const x = kid(n, tag);
    if (x) o[key] = flag(x);
  }
  const u = kid(n, 'u');
  if (u) o.u = attr(u, 'val') !== 'none';
  const va = kid(n, 'vertAlign');
  if (va) o.va = attr(va, 'val');
  const hl = kid(n, 'highlight');
  if (hl && HIGHLIGHT[attr(hl, 'val')]) o.hl = HIGHLIGHT[attr(hl, 'val')];
  const shd = kid(n, 'shd');
  if (shd) { const f = wordColor(shd, theme, 'fill'); if (f) o.shd = f; }
  const sp = kid(n, 'spacing');
  if (sp) o.spacing = num(sp, 'val', 0) / 20;
  const rs = kid(n, 'rStyle');
  if (rs) o.rStyle = attr(rs, 'val');
  return o;
}

function readPPr(n, theme) {
  const o = {};
  if (!n) return o;
  const jc = kid(n, 'jc');
  if (jc) o.align = { both: 'justify', distribute: 'justify', center: 'center', right: 'right', end: 'right', left: 'left', start: 'left' }[attr(jc, 'val')] || 'left';
  const sp = kid(n, 'spacing');
  if (sp) {
    if (attr(sp, 'before') !== null) o.before = twipPt(num(sp, 'before', 0));
    if (attr(sp, 'after') !== null) o.after = twipPt(num(sp, 'after', 0));
    if (attr(sp, 'beforeAutospacing') === '1' || attr(sp, 'beforeAutospacing') === 'true') o.before = 14;
    if (attr(sp, 'afterAutospacing') === '1' || attr(sp, 'afterAutospacing') === 'true') o.after = 14;
    if (attr(sp, 'line') !== null) {
      const rule = attr(sp, 'lineRule') || 'auto';
      const v = num(sp, 'line', 240);
      o.line = rule === 'auto' ? { mult: Math.round((v / 240) * 100) / 100 } : { pt: twipPt(v) };
    }
  }
  const ind = kid(n, 'ind');
  if (ind) {
    const l = attr(ind, 'left') ?? attr(ind, 'start');
    const r = attr(ind, 'right') ?? attr(ind, 'end');
    if (l !== null) o.left = twipCm(+l);
    if (r !== null) o.right = twipCm(+r);
    if (attr(ind, 'hanging') !== null) o.first = -twipCm(num(ind, 'hanging', 0));
    else if (attr(ind, 'firstLine') !== null) o.first = twipCm(num(ind, 'firstLine', 0));
  }
  const shd = kid(n, 'shd');
  if (shd) { const f = wordColor(shd, theme, 'fill'); if (f) o.shd = f; }
  const numPr = kid(n, 'numPr');
  if (numPr) {
    o.numId = attr(kid(numPr, 'numId'), 'val');
    o.ilvl = num(kid(numPr, 'ilvl'), 'val', 0);
  }
  const ol = kid(n, 'outlineLvl');
  if (ol) o.outline = num(ol, 'val', 9);
  if (kid(n, 'pageBreakBefore')) o.pageBreak = flag(kid(n, 'pageBreakBefore'));
  if (kid(n, 'contextualSpacing')) o.contextual = flag(kid(n, 'contextualSpacing'));
  const bdr = kid(n, 'pBdr');
  if (bdr) {
    o.borders = {};
    for (const side of ['top', 'bottom', 'left', 'right']) {
      const b = kid(bdr, side);
      if (b && !['nil', 'none'].includes(attr(b, 'val'))) o.borders[side] = borderCss(b, theme);
    }
  }
  const ps = kid(n, 'pStyle');
  if (ps) o.pStyle = attr(ps, 'val');
  return o;
}

function borderCss(b, theme) {
  const w = Math.max(1, Math.round((num(b, 'sz', 4) / 8) * 1.333));
  const c = wordColor(b, theme, 'color') || '#000000';
  const style = { double: 'double', dotted: 'dotted', dashed: 'dashed', dashSmallGap: 'dashed' }[attr(b, 'val')] || 'solid';
  return `${attr(b, 'val') === 'double' ? Math.max(3, w) : w}px ${style} ${c}`;
}

function readBorders(n, theme) {
  const o = {};
  if (!n) return o;
  for (const side of ['top', 'left', 'bottom', 'right', 'start', 'end', 'insideH', 'insideV']) {
    const b = kid(n, side);
    if (!b) continue;
    const key = side === 'start' ? 'left' : side === 'end' ? 'right' : side;
    o[key] = ['nil', 'none'].includes(attr(b, 'val')) ? 'none' : borderCss(b, theme);
  }
  return o;
}

// ---------------------------------------------------------------------------
// Estilos

class Styles {
  constructor(doc, theme) {
    this.theme = theme;
    this.byId = new Map();
    this.cache = new Map();
    this.defaults = { r: {}, p: {} };
    this.defaultPara = null;
    this.defaultTable = null;
    if (!doc) return;
    const dd = first(doc, 'docDefaults');
    if (dd) {
      this.defaults.r = readRPr(path(dd, 'rPrDefault', 'rPr'), theme);
      this.defaults.p = readPPr(path(dd, 'pPrDefault', 'pPr'), theme);
    }
    for (const s of all(doc, 'style')) {
      if (s.parentNode?.localName !== 'styles') continue;
      const id = attr(s, 'styleId');
      const type = attr(s, 'type');
      const st = {
        id, type,
        name: (attr(kid(s, 'name'), 'val') || id || '').toLowerCase(),
        basedOn: attr(kid(s, 'basedOn'), 'val'),
        r: readRPr(kid(s, 'rPr'), theme),
        p: readPPr(kid(s, 'pPr'), theme),
        tbl: kid(s, 'tblPr'),
        tcPr: kid(s, 'tcPr'),
        cond: {},
      };
      for (const c of kids(s, 'tblStylePr')) {
        st.cond[attr(c, 'type')] = { r: readRPr(kid(c, 'rPr'), theme), tcPr: kid(c, 'tcPr'), p: readPPr(kid(c, 'pPr'), theme) };
      }
      this.byId.set(id, st);
      if (attr(s, 'default') === '1' || attr(s, 'default') === 'true') {
        if (type === 'paragraph') this.defaultPara = id;
        if (type === 'table') this.defaultTable = id;
      }
    }
  }

  // Cadeia basedOn resolvida (do mais genérico para o mais específico).
  chain(id) {
    const out = [];
    const seen = new Set();
    let cur = this.byId.get(id);
    while (cur && !seen.has(cur.id)) {
      seen.add(cur.id);
      out.unshift(cur);
      cur = cur.basedOn ? this.byId.get(cur.basedOn) : null;
    }
    return out;
  }

  resolved(id) {
    if (!id) id = this.defaultPara;
    const key = `p:${id}`;
    if (!this.cache.has(key)) {
      const chain = this.chain(id);
      const r = {}, p = {};
      for (const s of chain) { Object.assign(r, s.r); Object.assign(p, s.p); }
      delete p.pStyle;
      this.cache.set(key, { r, p, names: chain.map((s) => s.name), ids: chain.map((s) => s.id) });
    }
    return this.cache.get(key);
  }

  charStyle(id) {
    const key = `c:${id}`;
    if (!this.cache.has(key)) {
      const r = {};
      for (const s of this.chain(id)) Object.assign(r, s.r);
      delete r.rStyle;
      this.cache.set(key, r);
    }
    return this.cache.get(key);
  }

  tableStyle(id) {
    const chain = this.chain(id || this.defaultTable);
    const out = { borders: {}, cond: {}, r: {}, p: {}, cellShd: null };
    for (const s of chain) {
      Object.assign(out.borders, readBorders(kid(s.tbl, 'tblBorders'), this.theme));
      Object.assign(out.r, s.r);
      Object.assign(out.p, s.p);
      const shd = kid(s.tcPr, 'shd');
      if (shd) out.cellShd = wordColor(shd, this.theme, 'fill') || out.cellShd;
      for (const [k, v] of Object.entries(s.cond)) {
        const prev = out.cond[k] || { r: {}, p: {}, shd: null, borders: {} };
        const shd2 = kid(v.tcPr, 'shd');
        out.cond[k] = {
          r: { ...prev.r, ...v.r },
          p: { ...prev.p, ...v.p },
          shd: (shd2 && wordColor(shd2, this.theme, 'fill')) || prev.shd,
          borders: { ...prev.borders, ...readBorders(kid(v.tcPr, 'tcBorders'), this.theme) },
        };
      }
    }
    return out;
  }
}

// ---------------------------------------------------------------------------
// Numeração

class Numbering {
  constructor(doc) {
    this.abstract = new Map();
    this.nums = new Map();
    if (!doc) return;
    for (const a of all(doc, 'abstractNum')) {
      const levels = new Map();
      for (const l of kids(a, 'lvl')) {
        levels.set(num(l, 'ilvl', 0), {
          fmt: attr(kid(l, 'numFmt'), 'val') || 'decimal',
          text: attr(kid(l, 'lvlText'), 'val') || '',
          start: num(kid(l, 'start'), 'val', 1),
          font: attr(path(l, 'rPr', 'rFonts'), 'ascii') || '',
          left: twipCm(num(path(l, 'pPr', 'ind'), 'left', null) ?? num(path(l, 'pPr', 'ind'), 'start', 0)),
        });
      }
      this.abstract.set(attr(a, 'abstractNumId'), levels);
    }
    for (const n of all(doc, 'num')) {
      const overrides = new Map();
      for (const o of kids(n, 'lvlOverride')) {
        const so = kid(o, 'startOverride');
        if (so) overrides.set(num(o, 'ilvl', 0), num(so, 'val', 1));
      }
      this.nums.set(attr(n, 'numId'), { abs: attr(kid(n, 'abstractNumId'), 'val'), overrides });
    }
  }

  level(numId, ilvl) {
    const n = this.nums.get(numId);
    if (!n) return null;
    const lv = this.abstract.get(n.abs)?.get(ilvl);
    if (!lv) return { fmt: 'bullet', text: '•', start: 1, abs: n.abs };
    return { ...lv, start: n.overrides.get(ilvl) ?? lv.start, abs: n.abs };
  }
}

// ---------------------------------------------------------------------------
// Conversor

export async function importDocx(data, docStyleBase) {
  const pkg = await openPackage(data);
  const ct = await pkg.xml('[Content_Types].xml');
  let mainPath = 'word/document.xml';
  if (ct) for (const o of all(ct, 'Override')) if (/document\.main\+xml|document\.macroEnabled\.main\+xml/.test(attr(o, 'ContentType') || '')) mainPath = attr(o, 'PartName').replace(/^\//, '');
  const docXml = await pkg.xml(mainPath);
  if (!docXml) throw new Error('Documento do Word inválido.');
  const rels = await pkg.rels(mainPath);
  const relPath = (type) => [...rels.values()].find((r) => r.type === type)?.target;
  const theme = await readTheme(pkg, relPath('theme'));
  const styles = new Styles(await pkg.xml(relPath('styles')), theme);
  const numbering = new Numbering(await pkg.xml(relPath('numbering')));
  const conv = new Converter(pkg, styles, numbering, theme, mainPath, rels);
  const notesXml = await pkg.xml(relPath('footnotes'));
  const endXml = await pkg.xml(relPath('endnotes'));
  conv.noteSources = { footnote: notesXml, endnote: endXml };
  conv.noteRels = {
    footnote: relPath('footnotes') ? await pkg.rels(relPath('footnotes')) : null,
    endnote: relPath('endnotes') ? await pkg.rels(relPath('endnotes')) : null,
  };

  const body = first(docXml, 'body');
  const style = conv.documentStyle(docStyleBase);
  conv.docStyle = style;
  let html = await conv.blocks(body);
  html += await conv.notesHtml();

  // Configuração de página (última seção)
  const sect = kid(body, 'sectPr') || all(body, 'sectPr').pop();
  const page = await conv.pageSetup(sect);
  const core = await pkg.xml('docProps/core.xml');
  const title = core ? first(core, 'title')?.textContent?.trim() : '';
  return { html: html || '<p><br></p>', page, style, title };
}

class Converter {
  constructor(pkg, styles, numbering, theme, mainPath, rels) {
    this.pkg = pkg;
    this.styles = styles;
    this.numbering = numbering;
    this.theme = theme;
    this.partPath = mainPath;
    this.rels = rels;
    this.notes = [];
    this.noteNumber = 0;
    this.counters = new Map();
    this.imageCache = new Map();
    this.fields = []; // campos complexos podem atravessar parágrafos (ex.: TOC)
  }

  // ---- Estilo do documento (vira o "estilo do documento" do Lumina) ---------

  documentStyle(base) {
    const n = this.styles.resolved(null);
    const r = { ...this.styles.defaults.r, ...n.r };
    const p = { ...this.styles.defaults.p, ...n.p };
    const h1Id = this.headingStyleId(1);
    const h1 = h1Id ? this.styles.resolved(h1Id) : null;
    const h1r = h1 ? { ...this.styles.defaults.r, ...h1.r } : null;
    const size = r.size || 10;
    const linkStyle = [...this.styles.byId.values()].find((s) => s.name === 'hyperlink');
    const st = {
      ...base,
      preset: 'custom',
      font: cssFont(r.font || this.theme.minor),
      headFont: h1r?.font && h1r.font !== r.font ? cssFont(h1r.font) : '',
      size,
      line: p.line?.mult ?? (p.line?.pt ? Math.round((p.line.pt / size) * 100) / 100 : 1),
      space: p.after ?? 0,
      indent: p.first > 0 ? p.first : 0,
      align: p.align === 'justify' || p.align === 'center' || p.align === 'right' ? p.align : 'left',
      headColor: h1r?.color && h1r.color !== 'auto' ? h1r.color : '#000000',
      accent: (linkStyle?.r.color && linkStyle.r.color !== 'auto' ? linkStyle.r.color : null) || this.theme.colors.hlink || this.theme.colors.accent1 || '#3b6cf6',
      headCase: 'none',
      headScale: h1r?.size ? Math.max(0.3, Math.round((h1r.size / (1.85 * size)) * 100) / 100) : 1,
    };
    this.base = { r, p };
    return st;
  }

  headingStyleId(level) {
    for (const s of this.styles.byId.values()) if (s.type === 'paragraph' && s.name === `heading ${level}`) return s.id;
    return null;
  }

  // ---- Blocos -----------------------------------------------------------------

  async blocks(container) {
    const items = [];
    await this.collect(container, items);
    return this.assemble(items);
  }

  // Achata o corpo numa sequência de itens { kind: 'p'|'table'|'raw', ... }.
  async collect(container, items) {
    for (let node of kids(container)) {
      node = unwrapAlternate(node);
      if (!node) continue;
      switch (node.localName) {
        case 'p': items.push(...(await this.paragraph(node))); break;
        case 'tbl': items.push({ kind: 'raw', html: await this.table(node) }); break;
        case 'sdt': {
          const gallery = attr(path(node, 'sdtPr', 'docPartObj', 'docPartGallery'), 'val') || '';
          if (/table of contents/i.test(gallery)) {
            const firstP = first(kid(node, 'sdtContent'), 'p');
            const t = firstP && this.isTocHeading(firstP) ? firstP.textContent.trim() : 'Sumário';
            items.push({ kind: 'raw', html: `<div class="toc" contenteditable="false"><div class="toc-title">${escapeHtml(t || 'Sumário')}</div></div>` });
          } else await this.collect(kid(node, 'sdtContent'), items);
          break;
        }
        case 'customXml': case 'smartTag': case 'ins': case 'moveTo': await this.collect(node, items); break;
        default: break;
      }
    }
  }

  isTocHeading(p) {
    const id = attr(path(p, 'pPr', 'pStyle'), 'val');
    return !!id && this.styles.resolved(id).names.some((n) => /toc heading|título do sumário/.test(n));
  }

  // Monta as listas (ul/ol aninhadas) e aplica espaçamento contextual.
  assemble(items) {
    let out = '';
    const stack = []; // { tag, numId, level, html: [] }
    const closeTo = (depth) => {
      while (stack.length > depth) {
        const l = stack.pop();
        const html = `<${l.tag}${l.attrs}>${l.items.join('')}</${l.tag}>`;
        if (stack.length) {
          const parent = stack[stack.length - 1];
          parent.items[parent.items.length - 1] = parent.items[parent.items.length - 1].replace(/<\/li>$/, `${html}</li>`);
        } else out += html;
      }
    };
    for (let i = 0; i < items.length; i++) {
      const it = items[i];
      if (it.kind === 'p' && it.contextual && items[i + 1]?.kind === 'p' && items[i + 1].styleId === it.styleId) it.css.push('margin-bottom:0');
      if (it.kind === 'p' && it.list) {
        const { level, numId, lv } = it.list;
        const tag = lv.fmt === 'bullet' || lv.fmt === 'none' ? 'ul' : 'ol';
        while (stack.length && (stack[stack.length - 1].level > level || (stack[stack.length - 1].level === level && (stack[stack.length - 1].tag !== tag || stack[stack.length - 1].abs !== lv.abs)))) closeTo(stack.length - 1);
        if (!stack.length || stack[stack.length - 1].level < level) {
          const key = `${lv.abs}:${level}`;
          const count = this.counters.get(key) || 0;
          let attrs = '';
          if (tag === 'ol') {
            const type = OL_TYPES[lv.fmt] || '1';
            if (type !== '1') attrs += ` type="${type}"`;
            const start = lv.start + count;
            if (start !== 1) attrs += ` start="${start}"`;
          } else {
            const bullet = /o/.test(lv.text) && /courier/i.test(lv.font) ? 'circle' : /[§▪■]/.test(lv.text) || /wingdings/i.test(lv.font) ? 'square' : lv.fmt === 'none' ? 'none' : '';
            if (bullet) attrs += ` style="list-style-type:${bullet}"`;
          }
          // Lista recuada além do padrão do Word (ex.: estilo "Marcador 2")
          const extra = Math.round(((it.list.left || lv.left || 0) - 0.635 * (level + 1)) * 100) / 100;
          if (!stack.length && extra > 0.3) attrs = attrs.includes('style="') ? attrs.replace('style="', `style="margin-left:${extra}cm;`) : `${attrs} style="margin-left:${extra}cm"`;
          stack.push({ tag, numId, level, abs: lv.abs, attrs, items: [] });
        }
        this.counters.set(`${lv.abs}:${level}`, (this.counters.get(`${lv.abs}:${level}`) || 0) + 1);
        // Reinicia os níveis abaixo
        for (const k of [...this.counters.keys()]) { const [a, l] = k.split(':'); if (a === lv.abs && +l > level) this.counters.delete(k); }
        const cssLi = it.css.filter((c) => !/^(margin-left|text-indent)/.test(c));
        stack[stack.length - 1].items.push(`<li${cssLi.length ? ` style="${cssLi.join(';')}"` : ''}>${it.inner || '<br>'}</li>`);
        continue;
      }
      closeTo(0);
      if (it.kind === 'raw') { out += it.html; continue; }
      if (it.kind === 'p') out += this.blockHtml(it);
    }
    closeTo(0);
    return out;
  }

  blockHtml(it) {
    const style = it.css.length ? ` style="${it.css.join(';')}"` : '';
    const cls = it.cls ? ` class="${it.cls}"` : '';
    return `<${it.tag}${cls}${style}>${it.inner || '<br>'}</${it.tag}>`;
  }

  // Converte um parágrafo em itens (podem vir quebras de página e caixas de texto junto).
  async paragraph(p) {
    const pPr = kid(p, 'pPr');
    const direct = readPPr(pPr, this.theme);
    const sid = direct.pStyle || this.styles.defaultPara;
    const st = this.styles.resolved(sid);
    const pp = { ...this.styles.defaults.p, ...st.p, ...direct };
    const out = [];
    if (pp.pageBreak) out.push({ kind: 'raw', html: '<div class="page-break"></div>' });

    // Tipo de bloco
    let tag = 'p', cls = '', headLevel = 0;
    const names = st.names;
    if (names.includes('title') || names.includes('título')) { cls = 'title'; }
    else if (names.includes('subtitle') || names.includes('subtítulo')) { cls = 'subtitle'; }
    else {
      const h = names.map((n) => /^(heading|título) ([1-9])$/.exec(n)).filter(Boolean).pop();
      const lvl = h ? +h[2] : pp.outline !== undefined && pp.outline < 9 ? pp.outline + 1 : 0;
      if (lvl) { headLevel = Math.min(4, lvl); tag = `h${headLevel}`; }
      else if (names.some((n) => /^(quote|intense quote|citação|citação intensa)$/.test(n))) tag = 'blockquote';
    }

    // Base com que o CSS do Lumina já renderiza este bloco
    const ds = this.docStyle;
    const base = { ...this.base.r, color: 'auto' };
    if (headLevel || cls) {
      const scale = cls === 'title' ? 2.6 : cls === 'subtitle' ? 1.3 : [1.85 * ds.headScale, 1.4 * ds.headScale, 1.17 * ds.headScale, 1][headLevel - 1];
      base.size = Math.round(ds.size * scale * 10) / 10;
      base.b = cls !== 'subtitle';
      base.headFont = ds.headFont || ds.font;
      base.color = cls === 'subtitle' ? '#5f6b7a' : ds.headColor;
      base.heading = true;
    }
    if (tag === 'blockquote') base.i = true;

    const runBase = { ...this.styles.defaults.r, ...st.r };
    const ctx = { base, runBase, deferred: [], fields: this.fields };
    const startedInToc = this.fields.some((f) => f.toc);
    let inner = await this.inline(p, ctx);
    if (inner.includes('<span data-toc="1"></span>')) {
      const before_ = inner.split('<span data-toc="1"></span>')[0];
      if (before_.replace(/<[^>]+>/g, '').trim()) out.push({ kind: 'p', tag: 'p', cls: '', css: [], inner: before_ });
      out.push({ kind: 'raw', html: '<div class="toc" contenteditable="false"><div class="toc-title">Sumário</div></div>' });
      return out;
    }
    if (startedInToc) {
      if (!inner.replace(/<[^>]+>/g, '').trim()) return out;
      inner = inner.replace(/^(<span class="tab">\t<\/span>)+/, '');
    }

    const css = [];
    const listInfo = pp.numId && pp.numId !== '0' ? this.numbering.level(pp.numId, pp.ilvl || 0) : null;
    if (!listInfo) {
      if (pp.align && pp.align !== (headLevel || cls || tag === 'blockquote' ? 'left' : ds.align)) css.push(`text-align:${pp.align}`);
      if (pp.left) css.push(`margin-left:${pp.left}cm`);
      if (pp.right) css.push(`margin-right:${pp.right}cm`);
      const first_ = pp.first ?? 0;
      if (first_ !== (tag === 'p' && !cls ? ds.indent : 0)) css.push(`text-indent:${first_}cm`);
    } else if (pp.align && pp.align !== 'left') css.push(`text-align:${pp.align}`);
    const before = pp.before ?? 0;
    const after = pp.after ?? 0;
    if (headLevel || cls || before) css.push(`margin-top:${before}pt`);
    if (headLevel || cls || after !== ds.space) css.push(`margin-bottom:${after}pt`);
    if (pp.line) {
      const lh = pp.line.mult ?? null;
      if (lh !== null && (Math.abs(lh - ds.line) > 0.01 || headLevel || cls)) css.push(`line-height:${lh}`);
      if (pp.line.pt) css.push(`line-height:${pp.line.pt}pt`);
    }
    if (pp.shd) css.push(`background-color:${pp.shd}`);
    if (pp.borders) for (const [side, v] of Object.entries(pp.borders)) css.push(`border-${side}:${v}`, `padding-${side}:2pt`);

    const item = { kind: 'p', tag, cls, css, inner, contextual: pp.contextual, styleId: sid };
    if (listInfo) item.list = { level: pp.ilvl || 0, numId: pp.numId, lv: listInfo, left: pp.left || 0 };
    // Parágrafo vazio que só contém uma quebra de página
    if (!(ctx.onlyBreak && !inner.replace(/<[^>]+>/g, '').trim())) out.push(item);
    for (const b of ctx.pageBreaksAfter || []) out.push(b);
    out.push(...ctx.deferred);
    return out;
  }

  // ---- Conteúdo em linha --------------------------------------------------------

  async inline(node, ctx, linkHref = null) {
    const segs = [];
    await this.walkInline(node, ctx, segs, linkHref);
    // Junta trechos vizinhos com a mesma formatação
    let html = '';
    let i = 0;
    while (i < segs.length) {
      const s = segs[i];
      if (s.raw !== undefined) { html += s.raw; i++; continue; }
      let text = s.text;
      let j = i + 1;
      while (j < segs.length && segs[j].raw === undefined && segs[j].key === s.key && segs[j].href === s.href) { text += segs[j].text; j++; }
      let h = escapeHtml(text).replace(/ {2}/g, ' \u00a0');
      h = wrapRun(h, s.props, ctx.base, this.docStyle);
      if (s.href) h = `<a href="${escapeHtml(s.href)}">${h}</a>`;
      html += h;
      i = j;
    }
    return html;
  }

  async walkInline(node, ctx, segs, href) {
    for (let n of kids(node)) {
      n = unwrapAlternate(n);
      if (!n) continue;
      switch (n.localName) {
        case 'r': await this.run(n, ctx, segs, href); break;
        case 'hyperlink': {
          const id = attr(n, 'id');
          const anchor = attr(n, 'anchor');
          const target = id ? this.rels.get(id)?.target : anchor ? `#${anchor}` : null;
          await this.walkInline(n, ctx, segs, target || href);
          break;
        }
        case 'fldSimple': {
          const instr = (attr(n, 'instr') || '').trim();
          const m = /^HYPERLINK\s+"([^"]+)"/i.exec(instr);
          await this.walkInline(n, ctx, segs, m ? m[1] : href);
          break;
        }
        case 'ins': case 'moveTo': case 'smartTag': case 'customXml': case 'bdo': case 'dir':
          await this.walkInline(n, ctx, segs, href);
          break;
        case 'sdt': await this.walkInline(kid(n, 'sdtContent'), ctx, segs, href); break;
        case 'bookmarkStart': {
          const name = attr(n, 'name');
          if (name && !name.startsWith('_')) segs.push({ raw: `<a id="${escapeHtml(name)}"></a>` });
          break;
        }
        case 'oMath': case 'oMathPara': segs.push({ raw: `<span class="math">${escapeHtml(n.textContent)}</span>` }); break;
        default: break;
      }
    }
  }

  async run(r, ctx, segs, href) {
    const direct = readRPr(kid(r, 'rPr'), this.theme);
    const charSt = direct.rStyle ? this.styles.charStyle(direct.rStyle) : {};
    const props = { ...ctx.runBase, ...charSt, ...direct };
    delete props.rStyle;
    if (props.hidden) return;
    // Campos complexos (HYPERLINK, TOC, PAGE…)
    const fieldHref = () => ctx.fields.map((f) => f.href).filter(Boolean).pop();
    const inToc = () => ctx.fields.some((f) => f.toc);
    const key = propsKey(props);
    const hidden = () => inToc() || ctx.fields.some((f) => f.mode === 'instr');
    const push = (text) => {
      if (hidden()) return;
      segs.push({ text, props, key, href: fieldHref() || href });
    };
    const pushRaw = (raw) => { if (!hidden()) segs.push({ raw }); };
    for (let c of kids(r)) {
      c = unwrapAlternate(c);
      if (!c) continue;
      switch (c.localName) {
        case 't': push(c.textContent); break;
        case 'tab': case 'ptab': pushRaw('<span class="tab">\t</span>'); break;
        case 'br': case 'cr': {
          const type = attr(c, 'type');
          if (type === 'page') {
            ctx.pageBreaksAfter ||= [];
            ctx.pageBreaksAfter.push({ kind: 'raw', html: '<div class="page-break"></div>' });
            ctx.onlyBreak = true;
          } else pushRaw('<br>');
          break;
        }
        case 'noBreakHyphen': push('‑'); break;
        case 'softHyphen': push('­'); break;
        case 'sym': {
          const code = parseInt(attr(c, 'char') || '0', 16);
          push(code >= 0xf000 ? ({ 0xf0fc: '✓', 0xf0fb: '✗', 0xf06e: '■', 0xf0a8: '☐', 0xf0fe: '☒', 0xf0b7: '•' }[code] || '•') : String.fromCharCode(code));
          break;
        }
        case 'fldChar': {
          const t = attr(c, 'fldCharType');
          if (t === 'begin') ctx.fields.push({ mode: 'instr', instr: '' });
          else if (t === 'separate') {
            const f = ctx.fields[ctx.fields.length - 1];
            if (f) {
              f.mode = 'result';
              const m = /^\s*HYPERLINK\s+(?:\\l\s+)?"([^"]+)"/i.exec(f.instr);
              if (m) f.href = /\\l/.test(f.instr) ? `#${m[1]}` : m[1];
            }
          } else if (t === 'end') ctx.fields.pop();
          break;
        }
        case 'instrText': {
          const f = ctx.fields[ctx.fields.length - 1];
          if (f) {
            f.instr += c.textContent;
            if (/^\s*TOC\b/.test(f.instr) && !f.toc) {
              f.toc = true;
              segs.push({ raw: '<span data-toc="1"></span>' });
            }
          }
          break;
        }
        case 'drawing': if (!hidden()) pushRaw(await this.drawing(c, ctx)); break;
        case 'pict': case 'object': if (!hidden()) pushRaw(await this.vml(c, ctx)); break;
        case 'footnoteReference': case 'endnoteReference': {
          const kind = c.localName === 'footnoteReference' ? 'footnote' : 'endnote';
          const n = ++this.noteNumber;
          this.notes.push({ kind, id: attr(c, 'id'), n });
          segs.push({ raw: `<sup><a href="#nota-${n}" id="ref-${n}">${n}</a></sup>` });
          break;
        }
        default: break;
      }
    }
  }

  // ---- Imagens ----------------------------------------------------------------

  async image(relId, rels = this.rels) {
    const rel = rels.get(relId);
    if (!rel || rel.external) return null;
    if (!isWebImage(rel.target)) return null;
    if (!this.imageCache.has(rel.target)) this.imageCache.set(rel.target, this.pkg.dataUrl(rel.target));
    return this.imageCache.get(rel.target);
  }

  async drawing(d, ctx) {
    const holder = kid(d, 'inline') || kid(d, 'anchor');
    if (!holder) return '';
    const ext = kid(holder, 'extent');
    const w = Math.round(num(ext, 'cx', 0) / EMU_PER_PX);
    const blip = first(holder, 'blip');
    const docPr = kid(holder, 'docPr');
    const alt = attr(docPr, 'descr') || attr(docPr, 'title') || '';
    // Caixa de texto (wps:txbx)
    const txbx = first(holder, 'txbxContent');
    if (txbx && !blip) {
      ctx.deferred.push({ kind: 'raw', html: `<div class="callout"><div class="callout-body">${await this.blocks(txbx)}</div></div>` });
      return '';
    }
    if (!blip) return '';
    const src = await this.image(attr(blip, 'embed') || attr(blip, 'link'), this.rels);
    if (!src) return `<span class="img-missing" title="Imagem em formato não suportado">[imagem]</span>`;
    let cls = '';
    if (holder.localName === 'anchor') {
      const ph = attr(path(holder, 'positionH', 'align'), 'val') || path(holder, 'positionH', 'align')?.textContent;
      const wrap = kids(holder).find((k) => /^wrap/.test(k.localName))?.localName;
      if (wrap === 'wrapSquare' || wrap === 'wrapTight' || wrap === 'wrapThrough') cls = ph === 'right' ? 'align-right' : ph === 'center' ? 'align-center' : 'align-left';
      else if (wrap === 'wrapTopAndBottom') cls = ph === 'right' ? '' : 'align-center';
    }
    return `<img src="${src}" alt="${escapeHtml(alt)}"${cls ? ` class="${cls}"` : ''} style="width:${w}px">`;
  }

  async vml(node, ctx) {
    const img = first(node, 'imagedata');
    if (!img) {
      const tb = first(node, 'txbxContent');
      if (tb) ctx.deferred.push({ kind: 'raw', html: `<div class="callout"><div class="callout-body">${await this.blocks(tb)}</div></div>` });
      return '';
    }
    const src = await this.image(attr(img, 'id'), this.rels);
    if (!src) return '';
    const shape = img.parentNode;
    const m = /width:\s*([\d.]+)(pt|px|in)/.exec(attr(shape, 'style') || '');
    const w = m ? Math.round(+m[1] * (m[2] === 'pt' ? 4 / 3 : m[2] === 'in' ? 96 : 1)) : null;
    return `<img src="${src}" alt=""${w ? ` style="width:${w}px"` : ''}>`;
  }

  // ---- Tabelas -----------------------------------------------------------------

  async table(tbl) {
    const tblPr = kid(tbl, 'tblPr');
    const tst = this.styles.tableStyle(attr(kid(tblPr, 'tblStyle'), 'val'));
    const borders = { ...tst.borders, ...readBorders(kid(tblPr, 'tblBorders'), this.theme) };
    const look = kid(tblPr, 'tblLook');
    const lookVal = parseInt(attr(look, 'val') || '0', 16);
    const firstRow = attr(look, 'firstRow') === '1' || (!attr(look, 'firstRow') && (lookVal & 0x20));
    const bandRows = !(attr(look, 'noHBand') === '1' || (!attr(look, 'noHBand') && (lookVal & 0x200)));
    const firstCol = attr(look, 'firstColumn') === '1' || (!attr(look, 'firstColumn') && (lookVal & 0x80));
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => num(g, 'w', 0));
    const totalW = grid.reduce((a, b) => a + b, 0);
    const tblW = kid(tblPr, 'tblW');
    const jc = attr(kid(tblPr, 'jc'), 'val');
    const hasBorders = Object.values(borders).some((b) => b && b !== 'none');

    // Primeiro passo: células com coluna inicial, para calcular mesclas verticais
    const rows = kids(tbl, 'tr').map((tr) => {
      let col = 0;
      const cells = [];
      for (const tc of kids(tr).flatMap((k) => (k.localName === 'sdt' ? kids(kid(k, 'sdtContent'), 'tc') : k.localName === 'tc' ? [k] : []))) {
        const tcPr = kid(tc, 'tcPr');
        const span = num(kid(tcPr, 'gridSpan'), 'val', 1);
        const vm = kid(tcPr, 'vMerge');
        cells.push({ tc, tcPr, col, span, vMerge: vm ? attr(vm, 'val') || 'continue' : null, rowspan: 1 });
        col += span;
      }
      const trPr = kid(tr, 'trPr');
      return { cells, header: !!kid(trPr, 'tblHeader'), height: num(kid(trPr, 'trHeight'), 'val', 0) };
    });
    for (let r = 0; r < rows.length; r++) {
      for (const c of rows[r].cells) {
        if (c.vMerge !== 'restart') continue;
        let rr = r + 1;
        while (rr < rows.length) {
          const below = rows[rr].cells.find((x) => x.col === c.col);
          if (!below || below.vMerge !== 'continue') break;
          below.skip = true;
          c.rowspan++;
          rr++;
        }
      }
    }

    const style = [];
    if (tblW && attr(tblW, 'type') === 'pct') style.push(`width:${Math.min(100, num(tblW, 'w', 5000) / 50)}%`);
    else if (tblW && attr(tblW, 'type') === 'dxa' && num(tblW, 'w', 0) > 0) style.push(`width:${Math.round(num(tblW, 'w', 0) / 15)}px`, 'max-width:100%');
    else if (totalW) style.push(`width:${Math.round(totalW / 15)}px`, 'max-width:100%');
    if (jc === 'center') style.push('margin-left:auto', 'margin-right:auto');
    let html = `<table${style.length ? ` style="${style.join(';')}"` : ''}>`;
    if (totalW) html += `<colgroup>${grid.map((w) => `<col style="width:${((w / totalW) * 100).toFixed(2)}%">`).join('')}</colgroup>`;
    let bodyRow = 0;
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      const isFirst = r === 0 && firstRow;
      const band = !isFirst && bandRows ? (bodyRow++ % 2 === 0 ? 'band1Horz' : 'band2Horz') : null;
      html += `<tr${row.height ? ` style="height:${Math.round(row.height / 15)}px"` : ''}>`;
      for (const c of row.cells) {
        if (c.skip) continue;
        const conds = [];
        if (band && tst.cond[band]) conds.push(tst.cond[band]);
        if (firstCol && c.col === 0 && tst.cond.firstCol) conds.push(tst.cond.firstCol);
        if (isFirst && tst.cond.firstRow) conds.push(tst.cond.firstRow);
        const css = [];
        let shd = tst.cellShd;
        for (const cd of conds) if (cd.shd) shd = cd.shd;
        const ownShd = kid(c.tcPr, 'shd');
        if (ownShd) shd = wordColor(ownShd, this.theme, 'fill') || shd;
        if (shd) css.push(`background-color:${shd}`);
        const cb = { ...Object.assign({}, ...conds.map((cd) => cd.borders)), ...readBorders(kid(c.tcPr, 'tcBorders'), this.theme) };
        if (!hasBorders && !Object.keys(cb).length) css.push('border:none');
        else {
          for (const side of ['top', 'right', 'bottom', 'left']) {
            const inner = side === 'top' || side === 'bottom' ? borders.insideH : borders.insideV;
            const edge = (side === 'top' && r === 0) || (side === 'bottom' && r + c.rowspan >= rows.length) || (side === 'left' && c.col === 0) || (side === 'right' && c.col + c.span >= grid.length);
            const b = cb[side] ?? (edge ? borders[side] : inner) ?? 'none';
            css.push(`border-${side}:${b}`);
          }
        }
        const va = attr(kid(c.tcPr, 'vAlign'), 'val');
        if (va === 'center') css.push('vertical-align:middle');
        else if (va === 'bottom') css.push('vertical-align:bottom');
        const condR = Object.assign({}, tst.r, ...conds.map((cd) => cd.r));
        const condP = Object.assign({}, tst.p, ...conds.map((cd) => cd.p));
        const inner = await this.cellContent(c.tc, condR, condP);
        const tag = row.header || (isFirst && conds.length && conds[conds.length - 1].r.b) ? 'th' : 'td';
        if (tag === 'th' && !shd) css.push('background-color:transparent');
        if (tag === 'th' && !condR.b) css.push('font-weight:normal');
        const span = `${c.span > 1 ? ` colspan="${c.span}"` : ''}${c.rowspan > 1 ? ` rowspan="${c.rowspan}"` : ''}`;
        html += `<${tag}${span} style="${css.join(';')}">${inner}</${tag}>`;
      }
      html += '</tr>';
    }
    return `${html}</table>`;
  }

  async cellContent(tc, condR, condP) {
    const items = [];
    const savedDefaults = this.styles.defaults;
    // Formatação condicional do estilo de tabela participa como base dos trechos
    this.styles.defaults = { r: { ...savedDefaults.r, ...condR }, p: { ...savedDefaults.p, ...condP } };
    try {
      await this.collect(tc, items);
    } finally {
      this.styles.defaults = savedDefaults;
    }
    // Parágrafo único sem formatação de bloco: conteúdo direto na célula
    if (items.length === 1 && items[0].kind === 'p' && !items[0].list && items[0].tag === 'p' && !items[0].cls) {
      const css = items[0].css.filter((c) => !/^(margin-(top|bottom)|text-indent|line-height)/.test(c));
      return css.length ? `<p style="${css.join(';')};margin:0">${items[0].inner || '<br>'}</p>` : items[0].inner || '<br>';
    }
    for (const it of items) if (it.kind === 'p') it.css = it.css.filter((c) => !/^text-indent/.test(c));
    return this.assemble(items);
  }

  // ---- Notas de rodapé e de fim -----------------------------------------------------

  async notesHtml() {
    if (!this.notes.length) return '';
    let html = '<hr><div class="footnotes">';
    for (const note of this.notes) {
      const src = this.noteSources[note.kind];
      const el = src && all(src, note.kind).find((x) => attr(x, 'id') === note.id);
      let inner = '';
      if (el) {
        const savedRels = this.rels;
        this.rels = this.noteRels[note.kind] || savedRels;
        const items = [];
        await this.collect(el, items);
        this.rels = savedRels;
        inner = items.filter((i) => i.kind === 'p').map((i) => i.inner).join(' ');
      }
      html += `<p id="nota-${note.n}" style="font-size:0.85em;margin-bottom:2pt"><sup><a href="#ref-${note.n}">${note.n}</a></sup> ${inner}</p>`;
    }
    return `${html}</div>`;
  }

  // ---- Página, cabeçalho e rodapé ----------------------------------------------

  async pageSetup(sect) {
    const page = { size: 'A4', orientation: 'portrait', margins: { top: 2.5, right: 3, bottom: 2.5, left: 3 }, pageNumbers: false, header: '', footer: '' };
    if (!sect) return page;
    const pg = kid(sect, 'pgSz');
    if (pg) {
      const w = num(pg, 'w', 11906) / TWIP_PER_CM, h = num(pg, 'h', 16838) / TWIP_PER_CM;
      const landscape = attr(pg, 'orient') === 'landscape' || w > h;
      const [pw, ph] = landscape ? [h, w] : [w, h];
      let best = 'A4', bestD = Infinity;
      for (const [k, [, sw, sh]] of Object.entries(PAGE_SIZES)) {
        const d = Math.abs(sw - pw) + Math.abs(sh - ph);
        if (d < bestD) { bestD = d; best = k; }
      }
      page.size = best;
      page.orientation = landscape ? 'landscape' : 'portrait';
    }
    const mar = kid(sect, 'pgMar');
    if (mar) {
      for (const k of ['top', 'right', 'bottom', 'left']) {
        const v = num(mar, k, null);
        if (v !== null) page.margins[k] = Math.max(0, twipCm(Math.abs(v)));
      }
    }
    // Cabeçalho e rodapé padrão
    for (const kind of ['header', 'footer']) {
      const refs = kids(sect, `${kind}Reference`);
      const ref = refs.find((x) => attr(x, 'type') === 'default') || refs[0];
      const rel = ref && this.rels.get(attr(ref, 'id'));
      if (!rel) continue;
      const xml = await this.pkg.xml(rel.target);
      if (!xml) continue;
      const instr = [...all(xml, 'instrText').map((x) => x.textContent), ...all(xml, 'fldSimple').map((x) => attr(x, 'instr'))].join(' ');
      if (/\bPAGE\b/.test(instr)) page.pageNumbers = true;
      // Texto (sem o número da página, que vem do campo)
      const lines = all(xml, 'p').map((p) => {
        let t = '';
        let inField = 0, result = false;
        for (const r of all(p, 'r')) {
          for (const c of kids(r)) {
            if (c.localName === 'fldChar') {
              const ft = attr(c, 'fldCharType');
              if (ft === 'begin') { inField++; result = false; } else if (ft === 'separate') result = true; else if (ft === 'end') { inField--; result = false; }
            } else if (c.localName === 't' && (!inField || (result && !/\bPAGE\b|NUMPAGES/.test(instr)))) t += c.textContent;
            else if (c.localName === 'tab') t += '  ';
          }
        }
        return t.trim();
      }).filter(Boolean);
      page[kind] = lines.join(' · ');
    }
    return page;
  }
}

function propsKey(p) {
  return JSON.stringify([p.font, p.size, p.color, p.b, p.i, p.u, p.strike, p.va, p.hl, p.shd, p.caps, p.smallCaps, p.spacing]);
}

// Aplica as diferenças entre a formatação do trecho e a base que o CSS já renderiza.
function wrapRun(html, p, base, ds) {
  const css = [];
  const fontCss = p.font ? cssFont(p.font) : null;
  const baseFont = base.heading ? base.headFont : ds.font;
  if (fontCss && fontCss !== baseFont) css.push(`font-family:${fontCss.replace(/"/g, "'")}`);
  const size = p.size ?? 10;
  const baseSize = base.heading ? base.size : ds.size;
  if (Math.abs(size - baseSize) > 0.05) css.push(`font-size:${size}pt`);
  const color = p.color && p.color !== 'auto' ? p.color : null;
  if (base.heading) {
    const eff = color || '#000000';
    if (eff.toLowerCase() !== String(base.color).toLowerCase()) css.push(`color:${eff}`);
  } else if (color && color !== '#000000') css.push(`color:${color}`);
  if (p.shd) css.push(`background-color:${p.shd}`);
  if (p.caps) css.push('text-transform:uppercase');
  if (p.smallCaps) css.push('font-variant:small-caps');
  if (p.spacing) css.push(`letter-spacing:${p.spacing}pt`);
  if (base.heading && !p.b && base.b) css.push('font-weight:normal');
  if (base.i && p.i === false) css.push('font-style:normal');
  let h = html;
  if (css.length) h = `<span style="${css.join(';')}">${h}</span>`;
  if (p.b && !(base.heading && base.b)) h = `<strong>${h}</strong>`;
  if (p.i && !base.i) h = `<em>${h}</em>`;
  if (p.u) h = `<u>${h}</u>`;
  if (p.strike) h = `<s>${h}</s>`;
  if (p.va === 'superscript') h = `<sup>${h}</sup>`;
  if (p.va === 'subscript') h = `<sub>${h}</sub>`;
  if (p.hl) h = `<mark style="background-color:${p.hl}">${h}</mark>`;
  return h;
}
