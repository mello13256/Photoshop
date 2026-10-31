// Leitor de .xlsx/.xlsm próprio (JSZip + XML). Lê valores, fórmulas (incl. compartilhadas), estilos com cores de tema,
// formatos numéricos, mesclas, congelamento, larguras, formatação condicional, validação, notas, hiperlinks,
// nomes definidos, filtros, cores de abas, gráficos e imagens.
import {
  openPackage, kids, kid, path, attr, num, first, all, rid, readTheme, isWebImage, hexRgb, rgbHex,
} from '../../shared/ooxml.js';
import { Workbook, Sheet } from './model.js';
import { shiftFormula, toPortugueseFormula } from './formula.js';
import { parseAddr, quoteSheet } from './values.js';

const addr = (a) => { const p = a ? parseAddr(a) : null; return p ? { row: p.r, col: p.c } : null; };

const BUILTIN_FMT = {
  0: 'General', 1: '0', 2: '0.00', 3: '#,##0', 4: '#,##0.00', 9: '0%', 10: '0.00%', 11: '0.00E+00', 12: '# ?/?', 13: '# ??/??',
  14: 'dd/mm/yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 17: 'mmm-yy', 18: 'h:mm AM/PM', 19: 'h:mm:ss AM/PM', 20: 'hh:mm', 21: 'hh:mm:ss',
  22: 'dd/mm/yyyy hh:mm', 37: '#,##0 ;(#,##0)', 38: '#,##0 ;[Red](#,##0)', 39: '#,##0.00;(#,##0.00)', 40: '#,##0.00;[Red](#,##0.00)',
  44: '_("R$"* #,##0.00_)', 45: 'mm:ss', 46: '[h]:mm:ss', 47: 'mm:ss.0', 48: '##0.0E+0', 49: '@',
};

// Paleta indexada padrão do Excel (índices 0–63)
const INDEXED = [
  '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF', '000000', 'FFFFFF', 'FF0000', '00FF00', '0000FF', 'FFFF00', 'FF00FF', '00FFFF',
  '800000', '008000', '000080', '808000', '800080', '008080', 'C0C0C0', '808080', '9999FF', '993366', 'FFFFCC', 'CCFFFF', '660066', 'FF8080', '0066CC', 'CCCCFF',
  '000080', 'FF00FF', 'FFFF00', '00FFFF', '800080', '800000', '008080', '0000FF', '00CCFF', 'CCFFFF', 'CCFFCC', 'FFFF99', '99CCFF', 'FF99CC', 'CC99FF', 'FFCC99',
  '3366FF', '33CCCC', '99CC00', 'FFCC00', 'FF9900', 'FF6600', '666699', '969696', '003366', '339966', '003300', '333300', '993300', '993366', '333399', '333333',
];

const BORDER_W = { thin: 'thin', hair: 'thin', dotted: 'dashed', dashed: 'dashed', dashDot: 'dashed', dashDotDot: 'dashed', medium: 'medium', mediumDashed: 'dashed', mediumDashDot: 'dashed', thick: 'thick', double: 'thick', slantDashDot: 'dashed' };
const OPS = { greaterThan: 'gt', lessThan: 'lt', between: 'between', equal: 'eq', notEqual: 'ne', greaterThanOrEqual: 'gte', lessThanOrEqual: 'lte', notBetween: 'notBetween' };
const OP_SYM = { gt: '>', lt: '<', eq: '=', ne: '<>', gte: '>=', lte: '<=' };

export async function importXlsxFile(data) {
  const pkg = await openPackage(data);
  const root = await pkg.rels('');
  const wbPath = [...root.values()].find((r) => r.type === 'officeDocument')?.target || 'xl/workbook.xml';
  const wbXml = await pkg.xml(wbPath);
  if (!wbXml) throw new Error('Pasta de trabalho do Excel inválida.');
  const rels = await pkg.rels(wbPath);
  const relOf = (type) => [...rels.values()].find((r) => r.type === type)?.target;
  const theme = await readTheme(pkg, relOf('theme'));
  const ctx = new Ctx(pkg, theme);
  ctx.date1904 = ['1', 'true'].includes(attr(first(wbXml, 'workbookPr'), 'date1904'));
  await ctx.loadStrings(relOf('sharedStrings'));
  await ctx.loadStyles(relOf('styles'));

  const wb = new Workbook();
  wb.sheets = [];
  const sheetEls = all(first(wbXml, 'sheets'), 'sheet');
  const sheetNames = sheetEls.map((s) => attr(s, 'name'));
  for (const sEl of sheetEls) {
    if (attr(sEl, 'state') === 'veryHidden') continue;
    const rel = rels.get(rid(sEl));
    if (!rel || rel.type !== 'worksheet') continue;
    const s = new Sheet(String(attr(sEl, 'name')).slice(0, 31));
    try {
      await ctx.sheet(rel.target, s);
    } catch (err) {
      console.warn('Falha ao ler a planilha', attr(sEl, 'name'), err);
    }
    if (attr(sEl, 'state') === 'hidden') s.hiddenSheet = true;
    wb.sheets.push(s);
  }
  if (!wb.sheets.length) wb.sheets.push(new Sheet('Planilha1'));
  // Gráficos usam o id interno da planilha
  for (const s of wb.sheets) {
    for (const ch of s.charts) {
      const src = ch.externalSheet ? wb.sheets.find((x) => x.name === ch.externalSheet) : null;
      ch.sheetId = (src || s).id;
      delete ch.externalSheet;
    }
  }
  // Nomes definidos (globais)
  for (const dn of all(first(wbXml, 'definedNames'), 'definedName')) {
    const name = attr(dn, 'name');
    if (!name || name.startsWith('_xlnm') || attr(dn, 'hidden') === '1') continue;
    const ref = dn.textContent.trim();
    if (!ref || /#REF!/.test(ref)) continue;
    const local = attr(dn, 'localSheetId');
    const scoped = local !== null && sheetNames[+local] && !ref.includes('!') ? `${quoteSheet(sheetNames[+local])}!${ref}` : ref;
    wb.names[name.toUpperCase()] = toPortugueseFormula(scoped.replace(/^=/, ''));
  }
  const active = num(first(wbXml, 'workbookView'), 'activeTab', 0);
  wb.active = Math.min(Math.max(0, active), wb.sheets.length - 1);
  if (pkg.has('xl/vbaProject.bin')) wb.vbaProject = await pkg.bytes('xl/vbaProject.bin');
  wb.recalc();
  return wb;
}


class Ctx {
  constructor(pkg, theme) {
    this.pkg = pkg;
    this.theme = theme;
    this.strings = [];
    this.xfs = [];
    this.dxfs = [];
    this.numFmts = { ...BUILTIN_FMT };
    this.indexed = INDEXED;
  }

  async loadStrings(p) {
    const xml = p && await this.pkg.xml(p);
    if (!xml) return;
    for (const si of kids(xml.documentElement, 'si')) {
      // Texto simples ou rico (ignora fonética rPh)
      let t = '';
      for (const k of kids(si)) {
        if (k.localName === 't') t += k.textContent;
        else if (k.localName === 'r') t += kid(k, 't')?.textContent || '';
      }
      this.strings.push(t.replace(/_x([0-9A-Fa-f]{4})_/g, (_, h) => String.fromCharCode(parseInt(h, 16))));
    }
  }

  color(node) {
    if (!node) return null;
    if (attr(node, 'auto') === '1') return null;
    let hex = null;
    const rgb = attr(node, 'rgb');
    if (rgb) hex = `#${rgb.slice(-6).toLowerCase()}`;
    else if (attr(node, 'theme') !== null) {
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      hex = this.theme.colors[order[num(node, 'theme', 0)]] || null;
    } else if (attr(node, 'indexed') !== null) {
      const i = num(node, 'indexed', 0);
      if (i === 64 || i === 65) return null;
      hex = this.indexed[i] ? `#${this.indexed[i].toLowerCase()}` : null;
    }
    const tint = num(node, 'tint', 0);
    if (hex && tint) hex = applyTint(hex, tint);
    return hex;
  }

  async loadStyles(p) {
    const xml = p && await this.pkg.xml(p);
    if (!xml) return;
    const root = xml.documentElement;
    const custom = kids(path(root, 'colors', 'indexedColors'), 'rgbColor').map((c) => (attr(c, 'rgb') || '').slice(-6));
    if (custom.length) this.indexed = custom;
    for (const nf of all(kid(root, 'numFmts'), 'numFmt')) this.numFmts[num(nf, 'numFmtId', 0)] = attr(nf, 'formatCode');
    const fonts = kids(kid(root, 'fonts'), 'font').map((f) => this.font(f));
    const fills = kids(kid(root, 'fills'), 'fill').map((f) => this.fill(f));
    const borders = kids(kid(root, 'borders'), 'border').map((b) => this.border(b));
    for (const xf of kids(kid(root, 'cellXfs'), 'xf')) {
      const st = {};
      Object.assign(st, fonts[num(xf, 'fontId', 0)] || {});
      const fill = fills[num(xf, 'fillId', 0)];
      if (fill) st.bg = fill;
      const bd = borders[num(xf, 'borderId', 0)];
      if (bd) st.bd = bd;
      const al = kid(xf, 'alignment');
      if (al) {
        const h = attr(al, 'horizontal');
        if (['left', 'center', 'right'].includes(h)) st.ha = h;
        else if (h === 'centerContinuous' || h === 'distributed' || h === 'justify') st.ha = 'center';
        const v = attr(al, 'vertical');
        if (v) st.va = v === 'center' ? 'middle' : v === 'top' ? 'top' : 'bottom';
        if (attr(al, 'wrapText') === '1' || attr(al, 'wrapText') === 'true') st.wrap = true;
        if (attr(al, 'indent')) st.indent = num(al, 'indent', 0);
      }
      const fmt = fmtFromCode(this.numFmts[num(xf, 'numFmtId', 0)]);
      if (fmt) Object.assign(st, fmt);
      this.xfs.push(st);
    }
    for (const dxf of kids(kid(root, 'dxfs'), 'dxf')) {
      const st = {};
      const f = kid(dxf, 'font');
      if (f) {
        const c = this.color(kid(f, 'color'));
        if (c) st.color = c;
        if (kid(f, 'b') && attr(kid(f, 'b'), 'val') !== '0') st.bold = true;
        if (kid(f, 'i') && attr(kid(f, 'i'), 'val') !== '0') st.italic = true;
      }
      const pf = path(dxf, 'fill', 'patternFill');
      if (pf) {
        const c = this.color(kid(pf, 'bgColor')) || this.color(kid(pf, 'fgColor'));
        if (c) st.bg = c;
      }
      this.dxfs.push(st);
    }
  }

  font(f) {
    const st = {};
    if (kid(f, 'b') && attr(kid(f, 'b'), 'val') !== '0') st.b = true;
    if (kid(f, 'i') && attr(kid(f, 'i'), 'val') !== '0') st.i = true;
    if (kid(f, 'u') && attr(kid(f, 'u'), 'val') !== 'none') st.u = true;
    if (kid(f, 'strike') && attr(kid(f, 'strike'), 'val') !== '0') st.st = true;
    const sz = num(kid(f, 'sz'), 'val', 11);
    if (sz !== 11) st.size = sz;
    const name = attr(kid(f, 'name'), 'val');
    if (name && !/^(Calibri|Aptos.*|Aptos Narrow)$/i.test(name)) st.font = name;
    const c = this.color(kid(f, 'color'));
    if (c && c !== '#000000') st.color = c;
    return st;
  }

  fill(f) {
    const pf = kid(f, 'patternFill');
    if (pf) {
      const type = attr(pf, 'patternType');
      if (!type || type === 'none') return null;
      const c = this.color(kid(pf, 'fgColor')) || this.color(kid(pf, 'bgColor'));
      return c && c !== '#ffffff' ? c : type === 'solid' ? c : null;
    }
    const gf = kid(f, 'gradientFill');
    if (gf) return this.color(first(gf, 'color'));
    return null;
  }

  border(b) {
    const side = (name) => {
      const n = kid(b, name) || (name === 'left' ? kid(b, 'start') : name === 'right' ? kid(b, 'end') : null);
      const style = attr(n, 'style');
      if (!style || style === 'none') return undefined;
      return { w: BORDER_W[style] || 'thin', color: this.color(kid(n, 'color')) || '#1d2129' };
    };
    const o = { t: side('top'), r: side('right'), b: side('bottom'), l: side('left') };
    return o.t || o.r || o.b || o.l ? o : null;
  }

  // ---- Planilha ----------------------------------------------------------------

  async sheet(p, s) {
    const xml = await this.pkg.xml(p);
    if (!xml) throw new Error('parte ausente');
    const root = xml.documentElement;
    const rels = await this.pkg.rels(p);

    const tab = path(root, 'sheetPr', 'tabColor');
    if (tab) s.tabColor = this.color(tab);
    const view = path(root, 'sheetViews', 'sheetView');
    if (view && attr(view, 'showGridLines') === '0') s.showGrid = false;
    const pane = kid(view, 'pane');
    if (pane && (attr(pane, 'state') === 'frozen' || attr(pane, 'state') === 'frozenSplit')) s.freeze = { r: num(pane, 'ySplit', 0), c: num(pane, 'xSplit', 0) };
    const fmtPr = kid(root, 'sheetFormatPr');
    const defH = num(fmtPr, 'defaultRowHeight', null);
    if (defH && Math.abs(defH - 15) > 0.5) s.rows.def = Math.round(defH * (4 / 3) + 4);
    const defW = num(fmtPr, 'defaultColWidth', null) ?? (num(fmtPr, 'baseColWidth', null) !== null ? num(fmtPr, 'baseColWidth', 8) + 0.71 : null);
    if (defW) s.cols.def = Math.round(defW * 7 + 5);
    for (const col of kids(kid(root, 'cols'), 'col')) {
      const min = num(col, 'min', 1) - 1, max = Math.min(num(col, 'max', 1) - 1, min + 300);
      const w = num(col, 'width', null);
      for (let i = min; i <= max; i++) {
        if (w !== null && (attr(col, 'customWidth') === '1' || Math.abs(w * 7 + 5 - s.cols.def) > 2)) s.cols.setSize(i, w * 7 + 5);
        if (attr(col, 'hidden') === '1' || attr(col, 'hidden') === 'true') s.cols.setHidden(i, true);
      }
    }

    // Células
    const shared = new Map();
    for (const row of kids(kid(root, 'sheetData'), 'row')) {
      const rn = num(row, 'r', 0) - 1;
      const ht = num(row, 'ht', null);
      if (ht !== null && (attr(row, 'customHeight') === '1' || Math.abs(ht - 15) > 0.6)) s.rows.setSize(rn, ht * (4 / 3) + 1);
      if (attr(row, 'hidden') === '1' || attr(row, 'hidden') === 'true') s.rows.setHidden(rn, true);
      const rowStyle = attr(row, 'customFormat') === '1' ? num(row, 's', 0) : null;
      let colCursor = 0;
      for (const c of kids(row, 'c')) {
        const ref = attr(c, 'r');
        const pos = ref ? addr(ref) : { row: rn, col: colCursor };
        const r = pos.row, cc = pos.col;
        colCursor = cc + 1;
        const out = { v: null };
        const t = attr(c, 't') || 'n';
        const v = kid(c, 'v')?.textContent;
        const f = kid(c, 'f');
        if (f) {
          const ft = attr(f, 't');
          let text = f.textContent;
          if (ft === 'shared') {
            const si = attr(f, 'si');
            if (text) shared.set(si, { text, r, c: cc });
            else if (shared.has(si)) { const m = shared.get(si); text = shiftFormula(m.text, r - m.r, cc - m.c); }
          }
          if (text) {
            try { out.f = toPortugueseFormula(text.replace(/_xlfn\._xlws\.|_xlfn\.|_xlws\./g, '').replace(/^=/, '')); } catch { out.f = text; }
          }
        }
        if (t === 's') out.v = this.strings[+v] ?? '';
        else if (t === 'inlineStr') out.v = all(kid(c, 'is'), 't').map((x) => x.textContent).join('');
        else if (t === 'str') out.v = v ?? '';
        else if (t === 'b') out.v = v === '1';
        else if (t === 'e') out.v = out.f ? null : v || null;
        else if (t === 'd') { const d = new Date(v); out.v = Number.isNaN(+d) ? v : d.getTime() / 86400000 + 25569; }
        else if (v !== undefined && v !== '') out.v = +v;
        if (typeof out.v === 'number' && this.date1904) {
          // datas do sistema 1904
          const st0 = this.xfs[num(c, 's', 0)] || {};
          if (['date', 'datetime', 'longdate'].includes(st0.fmt)) out.v += 1462;
        }
        if (out.f) out.v = null;
        const sIdx = attr(c, 's') !== null ? num(c, 's', 0) : rowStyle;
        const st = sIdx !== null ? this.xfs[sIdx] : null;
        if (st && Object.keys(st).length) out.s = { ...st };
        if (out.v !== null || out.f || out.s) s.set(r, cc, out);
      }
    }

    for (const m of all(kid(root, 'mergeCells'), 'mergeCell')) {
      const rg = rangeOf(attr(m, 'ref'));
      if (rg) s.merges.push(rg);
    }
    const af = kid(root, 'autoFilter');
    if (af) { const rg = rangeOf(attr(af, 'ref')); if (rg) s.filter = { ...rg, criteria: {} }; }

    // Hiperlinks
    for (const h of all(kid(root, 'hyperlinks'), 'hyperlink')) {
      const rg = rangeOf(attr(h, 'ref'));
      if (!rg) continue;
      const rel = rels.get(rid(h));
      const link = rel?.external ? rel.target : attr(h, 'location') ? `#${attr(h, 'location')}` : null;
      if (!link) continue;
      const cell = s.get(rg.r1, rg.c1) || { v: null };
      s.set(rg.r1, rg.c1, { ...cell, link });
    }

    this.conditional(root, s);
    this.validations(root, s);

    // Notas (comentários)
    for (const rel of rels.values()) {
      if (rel.type !== 'comments') continue;
      const cx = await this.pkg.xml(rel.target);
      for (const cm of all(cx, 'comment')) {
        const pos = addr(attr(cm, 'ref') || '');
        if (!pos) continue;
        const text = all(kid(cm, 'text'), 't').map((x) => x.textContent).join('').trim();
        const cell = s.get(pos.row, pos.col) || { v: null };
        s.set(pos.row, pos.col, { ...cell, note: text });
      }
    }

    // Desenhos: gráficos e imagens
    const drawing = kid(root, 'drawing');
    const dRel = drawing && rels.get(rid(drawing));
    if (dRel) await this.drawing(dRel.target, s);

    const u = s.used();
    s.rows.count = Math.max(1000, u.r + 200);
    s.cols.count = Math.max(26, u.c + 10);
  }

  conditional(root, s) {
    for (const cf of kids(root, 'conditionalFormatting')) {
      const ranges = (attr(cf, 'sqref') || first(cf, 'sqref')?.textContent || '').split(/\s+/).map(rangeOf).filter(Boolean);
      if (!ranges.length) continue;
      const base = ranges[0];
      for (const rule of kids(cf, 'cfRule')) {
        const type = attr(rule, 'type');
        const style = this.dxfs[num(rule, 'dxfId', -1)];
        const formulas = kids(rule, 'formula').map((x) => x.textContent);
        const toPt = (f) => { try { return toPortugueseFormula(f.replace(/_xlfn\./g, '')); } catch { return f; } };
        let out = null;
        if (type === 'colorScale') {
          const colors = kids(kid(rule, 'colorScale'), 'color').map((c) => this.color(c));
          out = colors.length >= 3 ? { type: 'scale3', params: { min: colors[0], mid: colors[1], max: colors[2] } } : { type: 'scale2', params: { min: colors[0], max: colors[1] } };
        } else if (type === 'dataBar') out = { type: 'bar', params: { color: this.color(path(rule, 'dataBar', 'color')) || '#638ec6' } };
        else if (type === 'cellIs') {
          const op = OPS[attr(rule, 'operator')] || 'eq';
          const a = cfValue(formulas[0]), b = cfValue(formulas[1]);
          if (['gt', 'lt', 'eq'].includes(op) && a.literal) out = { type: op, params: { a: a.v } };
          else if (op === 'between' && a.literal && b.literal) out = { type: 'between', params: { a: a.v, b: b.v } };
          else {
            const cell = `${colLetters(base.c1)}${base.r1 + 1}`;
            const expr = op === 'between' ? `E(${cell}>=${formulas[0]};${cell}<=${formulas[1]})` : op === 'notBetween' ? `OU(${cell}<${formulas[0]};${cell}>${formulas[1]})` : `${cell}${OP_SYM[op]}${formulas[0]}`;
            out = { type: 'formula', params: { f: toPt(expr) } };
          }
        } else if (type === 'containsText' && attr(rule, 'text') !== null) out = { type: 'text', params: { a: attr(rule, 'text') } };
        else if (type === 'duplicateValues') out = { type: 'dup', params: {} };
        else if (type === 'uniqueValues') out = { type: 'unique', params: {} };
        else if (type === 'top10') out = { type: attr(rule, 'bottom') === '1' ? 'bottom' : 'top', params: { n: num(rule, 'rank', 10), percent: attr(rule, 'percent') === '1' } };
        else if (type === 'aboveAverage') out = { type: attr(rule, 'aboveAverage') === '0' ? 'belowAvg' : 'aboveAvg', params: {} };
        else if (type === 'containsBlanks') out = { type: 'blank', params: {} };
        else if (type === 'containsErrors') out = { type: 'error', params: {} };
        else if (formulas[0]) out = { type: 'formula', params: { f: toPt(formulas[0]) } };
        if (!out) continue;
        if (!['scale2', 'scale3', 'bar'].includes(out.type)) out.style = style && Object.keys(style).length ? style : { bg: '#ffc9c9', color: '#9c0006' };
        for (const rg of ranges) {
          const rule2 = { ...out, params: { ...out.params }, id: `cf${Math.random().toString(36).slice(2, 9)}`, range: rg };
          if (rule2.type === 'formula' && rg !== base) rule2.params.f = shiftFormula(out.params.f, rg.r1 - base.r1, rg.c1 - base.c1);
          s.cf.push(rule2);
        }
      }
    }
  }

  validations(root, s) {
    for (const dv of all(kid(root, 'dataValidations'), 'dataValidation')) {
      const type = attr(dv, 'type') || 'any';
      if (type === 'any') continue;
      const ranges = (attr(dv, 'sqref') || '').split(/\s+/).map(rangeOf).filter(Boolean);
      const f1 = kid(dv, 'formula1')?.textContent ?? '';
      const f2 = kid(dv, 'formula2')?.textContent ?? '';
      const showError = attr(dv, 'showErrorMessage') === '1' || attr(dv, 'showErrorMessage') === 'true';
      const base = { msg: attr(dv, 'error') || '', strict: showError && !['warning', 'information'].includes(attr(dv, 'errorStyle')), showError };
      for (const rg of ranges) {
        let v;
        if (type === 'list') {
          if (/^".*"$/s.test(f1)) v = { type: 'list', items: f1.slice(1, -1).split(',').map((x) => x.trim()).filter(Boolean) };
          else v = { type: 'list', items: [], ref: toPortugueseFormula(f1.replace(/^=/, '')) };
        } else {
          v = { type, op: OPS[attr(dv, 'operator') || 'between'] || 'between', a: f1, b: f2 };
          if (type === 'custom') v = { type: 'custom', f: toPortugueseFormula(f1) };
        }
        s.validations.push({ range: rg, ...base, ...v });
      }
    }
  }

  async drawing(p, s) {
    const xml = await this.pkg.xml(p);
    if (!xml) return;
    const rels = await this.pkg.rels(p);
    const px = (anchor, which) => {
      const n = kid(anchor, which);
      if (!n) return null;
      const col = +(kid(n, 'col')?.textContent || 0), row = +(kid(n, 'row')?.textContent || 0);
      const colOff = +(kid(n, 'colOff')?.textContent || 0), rowOff = +(kid(n, 'rowOff')?.textContent || 0);
      return { x: s.cols.pos(col) + colOff / 9525, y: s.rows.pos(row) + rowOff / 9525 };
    };
    for (const anchor of kids(xml.documentElement)) {
      if (!/Anchor$/.test(anchor.localName)) continue;
      const from = px(anchor, 'from') || (kid(anchor, 'pos') ? { x: num(kid(anchor, 'pos'), 'x', 0) / 9525, y: num(kid(anchor, 'pos'), 'y', 0) / 9525 } : null);
      if (!from) continue;
      let to = px(anchor, 'to');
      if (!to) {
        const ext = kid(anchor, 'ext');
        to = { x: from.x + num(ext, 'cx', 4572000) / 9525, y: from.y + num(ext, 'cy', 2743200) / 9525 };
      }
      const box = { x: Math.round(from.x), y: Math.round(from.y), w: Math.max(80, Math.round(to.x - from.x)), h: Math.max(60, Math.round(to.y - from.y)) };
      const chartEl = first(anchor, 'chart');
      if (chartEl) {
        const rel = rels.get(rid(chartEl));
        const ch = rel && await this.chart(rel.target, s);
        if (ch) s.charts.push({ ...ch, ...box });
        continue;
      }
      const blip = first(anchor, 'blip');
      if (blip) {
        const rel = rels.get(rid(blip, 'embed'));
        if (rel && !rel.external && isWebImage(rel.target)) {
          const src = await this.pkg.dataUrl(rel.target);
          if (src) s.charts.push({ id: `img${Math.random().toString(36).slice(2, 9)}`, type: 'image', src, ...box });
        }
      }
    }
  }

  async chart(p, s) {
    const xml = await this.pkg.xml(p);
    if (!xml) return null;
    const plot = first(xml, 'plotArea');
    const typeNode = kids(plot).find((k) => /Chart$/.test(k.localName));
    if (!typeNode) return null;
    const tn = typeNode.localName;
    let type = 'column';
    if (tn.startsWith('bar')) type = attr(kid(typeNode, 'barDir'), 'val') === 'bar' ? 'bar' : 'column';
    else if (tn.startsWith('line') || tn.startsWith('radar') || tn.startsWith('stock')) type = 'line';
    else if (tn.startsWith('area')) type = 'area';
    else if (tn.startsWith('doughnut')) type = 'doughnut';
    else if (tn.startsWith('pie') || tn.startsWith('ofPie')) type = 'pie';
    else if (tn.startsWith('scatter') || tn.startsWith('bubble')) type = 'scatter';
    const grouping = attr(kid(typeNode, 'grouping'), 'val');
    // Junta todas as referências das séries num único intervalo
    const refs = [];
    const colors = [];
    let sheetName = null;
    for (const ser of kids(typeNode, 'ser')) {
      for (const k of ['tx', 'cat', 'val', 'xVal', 'yVal']) {
        const f = first(kid(ser, k), 'f')?.textContent;
        if (!f) continue;
        const m = /^(?:'?((?:[^']|'')+?)'?!)?\$?([A-Z]+)\$?(\d+)(?::\$?([A-Z]+)\$?(\d+))?$/.exec(f.trim());
        if (!m) continue;
        if (m[1]) sheetName ||= m[1].replace(/''/g, "'");
        const a = addr(`${m[2]}${m[3]}`), b = m[4] ? addr(`${m[4]}${m[5]}`) : a;
        refs.push({ r1: a.row, c1: a.col, r2: b.row, c2: b.col, kind: k });
      }
      const c = first(path(ser, 'spPr', 'solidFill'), 'srgbClr') || first(path(ser, 'spPr', 'solidFill'), 'schemeClr');
      if (c) colors.push(c.localName === 'srgbClr' ? `#${attr(c, 'val').toLowerCase()}` : this.theme.colors[attr(c, 'val')] || null);
    }
    if (!refs.length) return null;
    const range = {
      r1: Math.min(...refs.map((r) => r.r1)), c1: Math.min(...refs.map((r) => r.c1)),
      r2: Math.max(...refs.map((r) => r.r2)), c2: Math.max(...refs.map((r) => r.c2)),
    };
    const vals = refs.filter((r) => r.kind === 'val' || r.kind === 'yVal');
    const seriesIn = vals.length && vals[0].r1 === vals[0].r2 && vals[0].c1 !== vals[0].c2 ? 'rows' : 'cols';
    const titleNode = first(xml, 'title');
    const deleted = attr(first(xml, 'autoTitleDeleted'), 'val') === '1';
    let title = titleNode && !deleted ? all(titleNode, 't').map((t) => t.textContent).join('') : '';
    if (!title && titleNode && !deleted) title = first(titleNode, 'v')?.textContent || '';
    const legendPos = attr(path(xml.documentElement, 'chart', 'legend', 'legendPos'), 'val');
    const hasLegend = !!path(xml.documentElement, 'chart', 'legend');
    const out = {
      id: `ch${Math.random().toString(36).slice(2, 9)}`, type, title, range, seriesIn,
      legend: !hasLegend ? 'none' : legendPos === 't' ? 'top' : legendPos === 'r' ? 'right' : 'bottom',
      stacked: grouping === 'stacked' || grouping === 'percentStacked',
      externalSheet: sheetName && sheetName !== s.name ? sheetName : null,
    };
    // Sem cores explícitas, o Excel usa as cores de ênfase do tema
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((a) => this.theme.colors[a]).filter(Boolean);
    if (colors.filter(Boolean).length) out.colors = colors.map((c, i) => c || accents[i % accents.length]).filter(Boolean);
    else if (accents.length) out.colors = accents;
    return out;
  }
}

function cfValue(f) {
  if (f === undefined) return { literal: false };
  if (/^-?\d+(\.\d+)?$/.test(f)) return { literal: true, v: +f };
  if (/^"(.*)"$/.test(f)) return { literal: true, v: f.slice(1, -1) };
  return { literal: false };
}

function colLetters(c) {
  let s = '';
  for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

function rangeOf(ref) {
  if (!ref) return null;
  const [a, b] = ref.replace(/\$/g, '').split(':');
  const pa = addr(a), pb = b ? addr(b) : pa;
  if (!pa || !pb) return null;
  return { r1: Math.min(pa.row, pb.row), c1: Math.min(pa.col, pb.col), r2: Math.max(pa.row, pb.row), c2: Math.max(pa.col, pb.col) };
}

// Tonalidade do Excel (aplicada à luminosidade HSL)
function applyTint(hex, tint) {
  const [r, g, b] = hexRgb(hex).map((x) => x / 255);
  const max = Math.max(r, g, b), min = Math.min(r, g, b);
  let h = 0, s = 0, l = (max + min) / 2;
  if (max !== min) {
    const d = max - min;
    s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
    h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
    h /= 6;
  }
  l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
  const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
  const f = (t) => { if (t < 0) t += 1; if (t > 1) t -= 1; return t < 1 / 6 ? p + (q - p) * 6 * t : t < 1 / 2 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p; };
  return s === 0 ? rgbHex(l * 255, l * 255, l * 255) : rgbHex(f(h + 1 / 3) * 255, f(h) * 255, f(h - 1 / 3) * 255);
}

// Formato numérico do Excel -> formato do Lumina
export function fmtFromCode(full) {
  if (!full || full === 'General') return null;
  const nf = full.split(';')[0];
  const decs = (x) => ((x.split('.')[1] || '').match(/^[0#]+/)?.[0] || '').replace(/#/g, '').length;
  if (nf === '@') return { fmt: 'text' };
  if (/R\$|\[\$R\$|\$|€|£/.test(nf) && !/[dmyh]/i.test(nf.replace(/"[^"]*"/g, ''))) {
    const cur = /€/.test(nf) ? 'EUR' : /R\$/.test(nf) ? 'BRL' : 'USD';
    if (/^_|\*/.test(nf)) return { fmt: 'accounting', cur, dec: decs(nf) };
    return { fmt: 'currency', cur, dec: decs(nf) };
  }
  if (nf.includes('%')) return { fmt: 'percent', dec: decs(nf) };
  if (/E[+-]/i.test(nf)) return { fmt: 'scientific', dec: decs(nf) };
  const bare = nf.replace(/"[^"]*"|\[[^\]]*\]|\\./g, '');
  if (/[dy]/i.test(bare) && /h/i.test(bare)) return { fmt: 'datetime' };
  if (/dddd/i.test(bare) || /mmmm/i.test(bare)) return { fmt: 'custom', code: full };
  if (/[dy]/i.test(bare)) return { fmt: 'date' };
  if (/h|s/i.test(bare) && /[hms]/i.test(bare)) return /\[h\]/i.test(nf) ? { fmt: 'custom', code: full } : { fmt: 'time' };
  if (/^#,##0(\.0+)?$/.test(nf) || /^0\.0+$/.test(nf) || /^#,##0(\.0+)?_\)?$/.test(nf)) return { fmt: 'number', dec: decs(nf) };
  if (/^0$/.test(nf)) return { fmt: 'number', dec: 0 };
  return { fmt: 'custom', code: full };
}
