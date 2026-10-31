// Arquivos do Lumina Planilhas: .lumsheet (nativo), XLSX (ExcelJS), CSV e PDF.
import { loadVendor, decodeText, extOf, baseName } from '../../shared/native.js';
import { escapeHtml } from '../../shared/ui.js';
import { Workbook, Sheet, KEY } from './model.js';
import { parseInput, display } from './format.js';
import { isErr, isMatrix, serialFromParts, colName } from './values.js';
import { toExcelFormula, toPortugueseFormula, shiftFormula, tokenize } from './formula.js';

// ---- Nativo -------------------------------------------------------------------

export function serialize(wb) {
  return JSON.stringify(wb.toJSON());
}

export async function importFile({ name, data }) {
  const ext = extOf(name);
  if (ext === 'lumsheet') return Workbook.from(JSON.parse(decodeText(data)));
  if (ext === 'xlsx' || ext === 'xlsm') {
    try {
      const { importXlsxFile } = await import('./xlsx-import.js');
      const wb = await importXlsxFile(data);
      await readLuminaMacros(data, wb);
      return wb;
    } catch (err) {
      console.warn('Leitor próprio falhou; usando ExcelJS:', err);
      return importXlsx(data);
    }
  }
  if (ext === 'csv' || ext === 'tsv') return importCsv(decodeText(data), ext === 'tsv' ? '\t' : null, baseName(name));
  throw new Error('Formato não suportado');
}

// ---- CSV -----------------------------------------------------------------------

export function parseDelimited(text, delim = null) {
  text = text.replace(/^﻿/, '');
  if (!delim) {
    const sample = text.split(/\r?\n/).slice(0, 10).join('\n');
    const count = (d) => (sample.match(new RegExp(d === '\t' ? '\\t' : `\\${d}`, 'g')) || []).length;
    delim = [';', ',', '\t', '|'].sort((a, b) => count(b) - count(a))[0];
  }
  const rows = [];
  let row = [], cur = '', q = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (q) {
      if (ch === '"') { if (text[i + 1] === '"') { cur += '"'; i++; } else q = false; }
      else cur += ch;
    } else if (ch === '"' && cur === '') q = true;
    else if (ch === delim) { row.push(cur); cur = ''; }
    else if (ch === '\n' || ch === '\r') {
      if (ch === '\r' && text[i + 1] === '\n') i++;
      row.push(cur); rows.push(row); row = []; cur = '';
    } else cur += ch;
  }
  if (cur !== '' || row.length) { row.push(cur); rows.push(row); }
  return rows;
}

export function importCsv(text, delim, name) {
  const wb = new Workbook();
  const s = wb.sheet;
  s.name = (name || 'Planilha1').slice(0, 31);
  parseDelimited(text, delim).forEach((row, r) => row.forEach((val, c) => {
    if (val === '') return;
    const p = parseInput(val);
    s.set(r, c, { v: p.v, s: p.fmt ? { fmt: p.fmt, ...(p.dec !== undefined ? { dec: p.dec } : {}) } : undefined });
  }));
  wb.recalc();
  return wb;
}

export function exportCsv(wb, sheet, delim = ';') {
  const u = sheet.used();
  const lines = [];
  for (let r = 0; r <= u.r; r++) {
    const row = [];
    for (let c = 0; c <= u.c; c++) {
      const cell = sheet.get(r, c);
      const v = wb.value(sheet, r, c);
      let t = v === null ? '' : display(v, cell?.s || {});
      if (/[";\n\r,]/.test(t) || t.includes(delim)) t = `"${t.replace(/"/g, '""')}"`;
      row.push(t);
    }
    lines.push(row.join(delim));
  }
  return `﻿${lines.join('\r\n')}\r\n`;
}

// ---- XLSX ----------------------------------------------------------------------

const argbToHex = (c) => (c?.argb && c.argb.length >= 6 ? `#${c.argb.slice(-6).toLowerCase()}` : null);
const hexToArgb = (h) => `FF${h.replace('#', '').toUpperCase()}`;

function fmtFromNumFmt(full) {
  if (!full || full === 'General') return null;
  const nf = full.split(';')[0];
  if (/R\$|\$|€/.test(nf)) return { fmt: 'currency', cur: nf.includes('€') ? 'EUR' : /US\$|"\$"|^\$/.test(nf) && !nf.includes('R$') ? 'USD' : 'BRL', dec: (nf.split('.')[1] || '').replace(/[^0]/g, '').length };
  if (nf.includes('%')) return { fmt: 'percent', dec: (nf.split('.')[1] || '').replace(/[^0]/g, '').length };
  if (/[dy]/i.test(nf) && /h/i.test(nf)) return { fmt: 'datetime' };
  if (/[dy]/i.test(nf)) return { fmt: 'date' };
  if (/h/i.test(nf) && /m/i.test(nf)) return { fmt: 'time' };
  if (/^#,##0(\.0+)?$/.test(nf) || /^0\.0+$/.test(nf)) return { fmt: 'number', dec: (nf.split('.')[1] || '').length };
  if (/^0$/.test(nf)) return { fmt: 'number', dec: 0 };
  return { fmt: 'custom', code: full };
}

function numFmtFor(s) {
  const dec = s.dec ?? 2;
  const d = dec > 0 ? `.${'0'.repeat(dec)}` : '';
  switch (s.fmt) {
    case 'number': return `#,##0${d}`;
    case 'currency': case 'accounting': return `"${{ BRL: 'R$', USD: 'US$', EUR: '€' }[s.cur || 'BRL']}" #,##0${d};-"${{ BRL: 'R$', USD: 'US$', EUR: '€' }[s.cur || 'BRL']}" #,##0${d}`;
    case 'percent': return `0${(s.dec ?? 0) > 0 ? `.${'0'.repeat(s.dec)}` : ''}%`;
    case 'date': return 'dd/mm/yyyy';
    case 'datetime': return 'dd/mm/yyyy hh:mm';
    case 'time': return 'hh:mm:ss';
    case 'longdate': return 'dddd, d "de" mmmm "de" yyyy';
    case 'scientific': return `0${d}E+00`;
    case 'text': return '@';
    case 'custom': return s.code || 'General';
    default: return null;
  }
}

const BORDER_MAP_IN = { thin: 'thin', hair: 'thin', dotted: 'dashed', dashed: 'dashed', medium: 'medium', mediumDashed: 'dashed', thick: 'thick', double: 'thick' };

export async function importXlsx(data) {
  const ExcelJS = await loadVendor('exceljs.min.js', 'ExcelJS');
  const x = new ExcelJS.Workbook();
  await x.xlsx.load(data.buffer.slice(data.byteOffset, data.byteOffset + data.byteLength));
  const wb = new Workbook();
  wb.sheets = [];
  for (const ws of x.worksheets) {
    if (ws.state === 'veryHidden') continue;
    const s = new Sheet(ws.name.slice(0, 31));
    ws.eachRow({ includeEmpty: false }, (row, rn) => {
      if (row.height) s.rows.setSize(rn - 1, row.height * (4 / 3));
      if (row.hidden) s.rows.setHidden(rn - 1, true);
      row.eachCell({ includeEmpty: false }, (cell, cn) => {
        const r = rn - 1, c = cn - 1;
        const out = { v: null };
        let val = cell.value;
        if (val && typeof val === 'object' && !(val instanceof Date)) {
          if (val.formula || val.sharedFormula) {
            let f = val.formula;
            if (!f && val.sharedFormula) {
              const master = ws.getCell(val.sharedFormula);
              const mf = master.value?.formula;
              if (mf) f = shiftFormula(mf, r - (master.row - 1), c - (master.col - 1));
            }
            if (f) out.f = toPortugueseFormula(f.replace(/^=/, '').replace(/_xlfn\.|_xlws\./g, ''));
            val = val.result ?? null;
            if (val && typeof val === 'object' && val.error) val = null;
          } else if (val.richText) val = val.richText.map((t) => t.text).join('');
          else if (val.text !== undefined) val = typeof val.text === 'object' ? val.text.richText?.map((t) => t.text).join('') : val.text;
          else if (val.error) val = null;
        }
        if (val instanceof Date) {
          out.v = serialFromParts(val.getUTCFullYear(), val.getUTCMonth() + 1, val.getUTCDate(), val.getUTCHours(), val.getUTCMinutes(), val.getUTCSeconds());
        } else out.v = val ?? null;
        const st = {};
        const f = cell.font || {};
        if (f.bold) st.b = true;
        if (f.italic) st.i = true;
        if (f.underline) st.u = true;
        if (f.strike) st.st = true;
        if (f.size && f.size !== 11) st.size = f.size;
        if (f.name && !/^(Calibri|Aptos.*)$/i.test(f.name)) st.font = f.name;
        const fc = argbToHex(f.color);
        if (fc && fc !== '#000000') st.color = fc;
        const fill = cell.fill;
        if (fill?.type === 'pattern' && fill.pattern === 'solid') { const bg = argbToHex(fill.fgColor); if (bg && bg !== '#ffffff') st.bg = bg; }
        const al = cell.alignment || {};
        if (al.horizontal && ['left', 'center', 'right'].includes(al.horizontal)) st.ha = al.horizontal;
        if (al.vertical) st.va = al.vertical === 'middle' ? 'middle' : al.vertical === 'top' ? 'top' : 'bottom';
        if (al.wrapText) st.wrap = true;
        const nf = fmtFromNumFmt(cell.numFmt);
        if (nf) Object.assign(st, nf);
        else if (cell.value instanceof Date) st.fmt = 'date';
        const bd = cell.border;
        if (bd) {
          const side = (b) => (b?.style ? { w: BORDER_MAP_IN[b.style] || 'thin', color: argbToHex(b.color) || '#1d2129' } : undefined);
          const o = { t: side(bd.top), r: side(bd.right), b: side(bd.bottom), l: side(bd.left) };
          if (o.t || o.r || o.b || o.l) st.bd = o;
        }
        if (Object.keys(st).length) out.s = st;
        if (out.v !== null || out.f || out.s) s.set(r, c, out);
      });
    });
    ws.columns?.forEach((col, i) => {
      if (col.width) s.cols.setSize(i, col.width * 7 + 5);
      if (col.hidden) s.cols.setHidden(i, true);
    });
    for (const m of ws.model?.merges || []) {
      const [a, b] = m.split(':');
      const pa = cellOf(a), pb = cellOf(b);
      if (pa && pb) s.merges.push({ r1: pa.r, c1: pa.c, r2: pb.r, c2: pb.c });
    }
    const v = ws.views?.[0];
    if (v?.state === 'frozen') s.freeze = { r: v.ySplit || 0, c: v.xSplit || 0 };
    if (v && v.showGridLines === false) s.showGrid = false;
    const u = s.used();
    s.rows.count = Math.max(1000, u.r + 200);
    s.cols.count = Math.max(26, u.c + 10);
    wb.sheets.push(s);
  }
  if (!wb.sheets.length) wb.sheets.push(new Sheet('Planilha1'));
  wb.active = 0;
  wb.recalc();
  return wb;
}

function cellOf(a) {
  const m = /^([A-Z]+)(\d+)$/.exec(a);
  if (!m) return null;
  let c = 0;
  for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64;
  return { r: +m[2] - 1, c: c - 1 };
}

export async function exportXlsx(wb, chartImages = {}, { macroEnabled = false } = {}) {
  const ExcelJS = await loadVendor('exceljs.min.js', 'ExcelJS');
  const x = new ExcelJS.Workbook();
  x.creator = 'Lumina Planilhas';
  x.created = new Date();
  for (const s of wb.sheets) {
    const views = [];
    if (s.freeze.r || s.freeze.c) views.push({ state: 'frozen', xSplit: s.freeze.c, ySplit: s.freeze.r, showGridLines: s.showGrid });
    else views.push({ showGridLines: s.showGrid });
    const ws = x.addWorksheet(s.name, { views, properties: { tabColor: s.tabColor ? { argb: hexToArgb(s.tabColor) } : undefined } });
    const u = s.used();
    for (let c = 0; c <= Math.max(u.c, 0); c++) {
      const col = ws.getColumn(c + 1);
      if (s.cols.sizes.has(c)) col.width = Math.max(1, (s.cols.size(c) - 5) / 7);
      if (s.cols.hidden.has(c)) col.hidden = true;
    }
    for (const [i, size] of s.rows.sizes) ws.getRow(i + 1).height = size * 0.75;
    for (const i of s.rows.hidden) ws.getRow(i + 1).hidden = true;
    const spillMap = wb.spills.get(s.id) || new Map();
    const writeValue = (wc, v) => {
      if (isErr(v)) wc.value = { error: v.code === '#CIRC!' || v.code === '#SPILL!' || v.code === '#CALC!' ? '#VALUE!' : v.code };
      else wc.value = v;
    };
    for (const [k, cell] of s.cells) {
      const r = Math.floor(k / 16384), c = k % 16384;
      const wc = ws.getCell(r + 1, c + 1);
      if (cell.f) {
        let res = cell.val;
        if (isMatrix(res)) res = res[0]?.[0] ?? null;
        if (isErr(res)) res = { error: '#VALUE!' };
        wc.value = { formula: toExcelFormula(cell.f), result: res ?? undefined };
      } else if (cell.link && cell.v !== null && cell.v !== undefined) {
        const text = display(cell.v, cell.s || {});
        wc.value = cell.link.startsWith('#') ? { text, hyperlink: cell.link } : { text, hyperlink: cell.link, tooltip: cell.link };
      } else if (cell.v !== null && cell.v !== undefined) writeValue(wc, cell.v);
      if (cell.note) wc.note = cell.note;
      applyStyleOut(wc, cell.s);
    }
    // Valores despejados por matrizes dinâmicas
    for (const [k, sp] of spillMap) {
      const r = Math.floor(k / 16384), c = k % 16384;
      const wc = ws.getCell(r + 1, c + 1);
      if (!wc.value) writeValue(wc, wb.value(s, r, c));
      void sp;
    }
    for (const m of s.merges) ws.mergeCells(m.r1 + 1, m.c1 + 1, m.r2 + 1, m.c2 + 1);
    if (s.filter) ws.autoFilter = { from: { row: s.filter.r1 + 1, column: s.filter.c1 + 1 }, to: { row: s.filter.r2 + 1, column: s.filter.c2 + 1 } };
    // Imagens (gráficos viram gráficos nativos depois, no pós-processamento)
    for (const ch of s.charts) {
      if (ch.type !== 'image' || !ch.src) continue;
      const ext = /^data:image\/(png|jpe?g|gif)/.exec(ch.src)?.[1]?.replace('jpg', 'jpeg');
      if (!ext) continue;
      const id = x.addImage({ base64: ch.src, extension: ext === 'jpeg' ? 'jpeg' : ext });
      ws.addImage(id, { tl: anchorOf(s, ch.x, ch.y), ext: { width: ch.w, height: ch.h } });
    }
    writeConditional(ws, s);
    writeValidations(ws, s);
  }
  for (const [name, ref] of Object.entries(wb.names || {})) {
    try { x.definedNames.add(toExcelFormula(ref).replace(/^=/, ''), name); } catch { /* nome inválido no Excel */ }
  }
  void chartImages;
  const buf = new Uint8Array(await x.xlsx.writeBuffer());
  return postProcessXlsx(buf, wb, macroEnabled && wb.vbaProject ? wb.vbaProject : null);
}

function anchorOf(s, x, y) {
  const col = s.cols.indexAt(x), row = s.rows.indexAt(y);
  const cw = s.cols.size(col) || 1, rh = s.rows.size(row) || 1;
  return { col: col + (x - s.cols.pos(col)) / cw, row: row + (y - s.rows.pos(row)) / rh };
}

const A1 = (r, c) => `${colName(c)}${r + 1}`;
const rangeA1 = (rg) => (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? A1(rg.r1, rg.c1) : `${A1(rg.r1, rg.c1)}:${A1(rg.r2, rg.c2)}`);
const absA1 = (rg) => `$${colName(rg.c1)}$${rg.r1 + 1}:$${colName(rg.c2)}$${rg.r2 + 1}`;
const argb = (h) => ({ argb: hexToArgb(h) });

function dxfOf(style = {}) {
  const st = {};
  if (style.color || style.bold || style.italic) st.font = { ...(style.color ? { color: argb(style.color) } : {}), ...(style.bold ? { bold: true } : {}), ...(style.italic ? { italic: true } : {}) };
  if (style.bg) st.fill = { type: 'pattern', pattern: 'solid', bgColor: argb(style.bg), fgColor: argb(style.bg) };
  return st;
}

function writeConditional(ws, s) {
  for (const rule of s.cf) {
    const ref = rangeA1(rule.range);
    const tl = A1(rule.range.r1, rule.range.c1);
    const p = rule.params || {};
    const style = dxfOf(rule.style);
    const lit = (v) => (typeof v === 'number' ? String(v) : `"${String(v ?? '').replace(/"/g, '""')}"`);
    let r = null;
    switch (rule.type) {
      case 'scale2': r = { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'max' }], color: [argb(p.min || '#f8696b'), argb(p.max || '#63be7b')] }; break;
      case 'scale3': r = { type: 'colorScale', cfvo: [{ type: 'min' }, { type: 'percentile', value: 50 }, { type: 'max' }], color: [argb(p.min || '#f8696b'), argb(p.mid || '#ffeb84'), argb(p.max || '#63be7b')] }; break;
      case 'bar': r = { type: 'dataBar', minLength: 0, maxLength: 100, cfvo: [{ type: 'min' }, { type: 'max' }], color: argb(p.color || '#5b8cff') }; break;
      case 'gt': case 'lt': case 'eq': r = { type: 'cellIs', operator: { gt: 'greaterThan', lt: 'lessThan', eq: 'equal' }[rule.type], formulae: [lit(p.a)], style }; break;
      case 'between': r = { type: 'cellIs', operator: 'between', formulae: [lit(p.a), lit(p.b)], style }; break;
      case 'text': r = { type: 'containsText', operator: 'containsText', text: String(p.a ?? ''), style }; break;
      case 'dup': r = { type: 'expression', formulae: [`COUNTIF(${absA1(rule.range)},${tl})>1`], style }; break;
      case 'unique': r = { type: 'expression', formulae: [`COUNTIF(${absA1(rule.range)},${tl})=1`], style }; break;
      case 'top': case 'bottom': r = { type: 'top10', rank: p.n || 10, percent: !!p.percent, bottom: rule.type === 'bottom', style }; break;
      case 'aboveAvg': case 'belowAvg': r = { type: 'aboveAverage', aboveAverage: rule.type === 'aboveAvg', style }; break;
      case 'blank': r = { type: 'expression', formulae: [`LEN(TRIM(${tl}))=0`], style }; break;
      case 'error': r = { type: 'expression', formulae: [`ISERROR(${tl})`], style }; break;
      case 'formula': r = { type: 'expression', formulae: [toExcelFormula(String(p.f || '').replace(/^=/, ''))], style }; break;
      default: break;
    }
    if (r) ws.addConditionalFormatting({ ref, rules: [{ priority: 1, ...r }] });
  }
}

function writeValidations(ws, s) {
  const OP = { between: 'between', notBetween: 'notBetween', eq: 'equal', ne: 'notEqual', gt: 'greaterThan', lt: 'lessThan', gte: 'greaterThanOrEqual', lte: 'lessThanOrEqual' };
  for (const v of s.validations) {
    const common = { allowBlank: true, showErrorMessage: v.showError !== false, errorStyle: v.strict === false ? 'warning' : 'stop', error: v.msg || undefined };
    let dv;
    if (v.type === 'list') dv = { type: 'list', formulae: [v.ref ? toExcelFormula(v.ref) : `"${v.items.join(',').replace(/"/g, '""')}"`] };
    else if (v.type === 'custom') dv = { type: 'custom', formulae: [toExcelFormula(v.f || 'VERDADEIRO')] };
    else dv = { type: v.type, operator: OP[v.op] || 'between', formulae: [v.a, v.b].filter((x) => x !== undefined && x !== '').map(String) };
    try { ws.dataValidations.add(rangeA1(v.range), { ...common, ...dv }); } catch { /* ignora regra inválida */ }
  }
}

// ---- Pós-processamento do .xlsx: gráficos nativos e macros do Lumina ------------------

async function postProcessXlsx(buf, wb, vba = null) {
  const JSZip = await loadVendor('jszip.min.js', 'JSZip');
  const zip = await JSZip.loadAsync(buf);
  let ct = await zip.file('[Content_Types].xml').async('string');
  const addOverride = (part, type) => {
    if (!ct.includes(`PartName="/${part}"`)) ct = ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
  };
  const wbRelsPath = 'xl/_rels/workbook.xml.rels';
  let wbRels = await zip.file(wbRelsPath).async('string');
  // Planilhas na ordem do workbook.xml
  const wbXml = await zip.file('xl/workbook.xml').async('string');
  const sheetRids = [...wbXml.matchAll(/<sheet [^>]*r:id="([^"]+)"/g)].map((m) => m[1]);
  const target = (rId) => new RegExp(`Id="${rId}"[^>]*Target="([^"]+)"`).exec(wbRels)?.[1] || new RegExp(`Target="([^"]+)"[^>]*Id="${rId}"`).exec(wbRels)?.[1];
  let chartN = 0, drawingN = 100;
  for (let i = 0; i < wb.sheets.length; i++) {
    const s = wb.sheets[i];
    const t = target(sheetRids[i]);
    if (!t) continue;
    const sheetPath = `xl/${t.replace(/^\/?xl\//, '').replace(/^\//, '')}`;
    if (!zip.file(sheetPath)) continue;
    // O ExcelJS grava o <extLst> da planilha antes do <legacyDrawing>; o formato exige que ele seja o último
    {
      let x = await zip.file(sheetPath).async('string');
      const fixed = moveSheetExtLst(x);
      if (fixed !== x) { x = fixed; zip.file(sheetPath, x); }
    }
    const charts = s.charts.filter((c) => c.type !== 'image' && c.range);
    if (!charts.length) continue;
    const sheetName = sheetPath.split('/').pop();
    const relsPath = `xl/worksheets/_rels/${sheetName}.rels`;
    let sheetXml = await zip.file(sheetPath).async('string');
    let rels = zip.file(relsPath) ? await zip.file(relsPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    // Reaproveita o desenho criado pelo ExcelJS (imagens) ou cria um novo
    let drawingPath;
    const existing = /<drawing r:id="([^"]+)"\/>/.exec(sheetXml);
    if (existing) {
      const m = new RegExp(`Id="${existing[1]}"[^>]*Target="([^"]+)"`).exec(rels) || new RegExp(`Target="([^"]+)"[^>]*Id="${existing[1]}"`).exec(rels);
      drawingPath = m ? `xl/${m[1].replace(/^\.\.\//, '')}` : null;
    }
    if (!drawingPath) {
      drawingPath = `xl/drawings/drawing${drawingN++}.xml`;
      zip.file(drawingPath, '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"></xdr:wsDr>');
      rels = rels.replace('</Relationships>', `<Relationship Id="rIdLuminaDr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/${drawingPath.split('/').pop()}"/></Relationships>`);
      const tag = '<drawing r:id="rIdLuminaDr"/>';
      // Posição no nível da planilha: depois de pageMargins/pageSetup/headerFooter e antes de legacyDrawing/tableParts/extLst
      const from = Math.max(sheetXml.indexOf('<pageMargins'), sheetXml.lastIndexOf('</sheetData>'), 0);
      const before = ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst'].map((k) => sheetXml.indexOf(k, from)).filter((k) => k > 0);
      sheetXml = before.length ? sheetXml.slice(0, Math.min(...before)) + tag + sheetXml.slice(Math.min(...before)) : sheetXml.replace('</worksheet>', `${tag}</worksheet>`);
      if (!/xmlns:r=/.test(sheetXml.slice(0, 500))) sheetXml = sheetXml.replace('<worksheet ', '<worksheet xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
      zip.file(sheetPath, sheetXml);
      zip.file(relsPath, rels);
      addOverride(drawingPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
    }
    let dxml = await zip.file(drawingPath).async('string');
    if (!dxml.includes('xmlns:c=')) dxml = dxml.replace('<xdr:wsDr ', '<xdr:wsDr xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    else if (!dxml.includes('xmlns:r=')) dxml = dxml.replace('<xdr:wsDr ', '<xdr:wsDr xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ');
    const dRelsPath = drawingPath.replace('drawings/', 'drawings/_rels/') + '.rels';
    let dRels = zip.file(dRelsPath) ? await zip.file(dRelsPath).async('string') : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>';
    let anchors = '';
    for (const ch of charts) {
      const n = ++chartN;
      const chartPath = `xl/charts/chart${n}.xml`;
      zip.file(chartPath, chartXml(wb, s, ch));
      addOverride(chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const rid = `rIdLuminaCh${n}`;
      dRels = dRels.replace('</Relationships>', `<Relationship Id="${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart" Target="../charts/chart${n}.xml"/></Relationships>`);
      const from = anchorOf(s, ch.x, ch.y), to = anchorOf(s, ch.x + ch.w, ch.y + ch.h);
      const mk = (a) => `<xdr:col>${Math.floor(a.col)}</xdr:col><xdr:colOff>${Math.round((a.col % 1) * (s.cols.size(Math.floor(a.col)) || 1) * 9525)}</xdr:colOff><xdr:row>${Math.floor(a.row)}</xdr:row><xdr:rowOff>${Math.round((a.row % 1) * (s.rows.size(Math.floor(a.row)) || 1) * 9525)}</xdr:rowOff>`;
      anchors += `<xdr:twoCellAnchor editAs="oneCell"><xdr:from>${mk(from)}</xdr:from><xdr:to>${mk(to)}</xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${1000 + n}" name="Gráfico ${n}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`;
    }
    dxml = dxml.replace('</xdr:wsDr>', `${anchors}</xdr:wsDr>`);
    zip.file(drawingPath, dxml);
    zip.file(dRelsPath, dRels);
  }
  // Macros do Lumina: parte customXml (o Excel preserva, mas não executa)
  if ((wb.macros || []).length) {
    const json = JSON.stringify(wb.macros).replace(/]]>/g, ']]]]><![CDATA[>');
    zip.file('customXml/luminaMacros.xml', `<?xml version="1.0" encoding="UTF-8"?><luminaMacros xmlns="urn:lumina:macros"><![CDATA[${json}]]></luminaMacros>`);
    if (!/Extension="xml"/.test(ct)) ct = ct.replace('</Types>', '<Default Extension="xml" ContentType="application/xml"/></Types>');
    wbRels = wbRels.replace('</Relationships>', '<Relationship Id="rIdLuminaMacros" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml" Target="../customXml/luminaMacros.xml"/></Relationships>');
    zip.file(wbRelsPath, wbRels);
  }
  // Projeto VBA preservado (arquivo .xlsm)
  if (vba) {
    zip.file('xl/vbaProject.bin', vba);
    ct = ct.replace('application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml', 'application/vnd.ms-excel.sheet.macroEnabled.main+xml');
    if (!/Extension="bin"/.test(ct)) ct = ct.replace('</Types>', '<Default Extension="bin" ContentType="application/vnd.ms-office.vbaProject"/></Types>');
    wbRels = wbRels.replace('</Relationships>', '<Relationship Id="rIdLuminaVba" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/></Relationships>');
    zip.file(wbRelsPath, wbRels);
  }
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE' });
}

// Move o <extLst> de nível superior para o fim da planilha.
function moveSheetExtLst(xml) {
  const from = Math.max(xml.indexOf('<pageMargins'), xml.lastIndexOf('</sheetData>'));
  if (from < 0) return xml;
  const start = xml.indexOf('<extLst', from);
  if (start < 0) return xml;
  const end = xml.indexOf('</extLst>', start);
  if (end < 0) return xml;
  const block = xml.slice(start, end + 9);
  const rest = xml.slice(0, start) + xml.slice(end + 9);
  if (rest.indexOf('<', start) === rest.indexOf('</worksheet>', start)) return xml;
  return rest.replace('</worksheet>', `${block}</worksheet>`);
}

async function readLuminaMacros(data, wb) {
  try {
    const JSZip = await loadVendor('jszip.min.js', 'JSZip');
    const zip = await JSZip.loadAsync(data);
    const f = zip.file('customXml/luminaMacros.xml');
    if (!f) return;
    const xml = await f.async('string');
    const json = /<!\[CDATA\[([\s\S]*)\]\]>/.exec(xml)?.[1]?.replace(/]]]]><!\[CDATA\[>/g, ']]>');
    if (json) wb.macros = JSON.parse(json);
  } catch (err) {
    console.warn('Macros ignoradas:', err);
  }
}

const xesc = (s) => String(s ?? '').replace(/[<>&"]/g, (c) => ({ '<': '&lt;', '>': '&gt;', '&': '&amp;', '"': '&quot;' })[c]);

// XML de um gráfico nativo (DrawingML) com referências às células e valores em cache.
function chartXml(wb, s, ch) {
  const src = wb.byId(ch.sheetId) || s;
  const sheetRef = quoteSheetName(src.name);
  const { r1, c1, r2, c2 } = ch.range;
  const val = (r, c) => wb.value(src, r, c);
  const isText = (v) => typeof v === 'string' && v !== '';
  const byRows = ch.seriesIn === 'rows';
  // Monta as séries em termos de (linhas, colunas) das células
  const lines = []; // cada série: lista de coordenadas
  if (!byRows) for (let c = c1; c <= c2; c++) lines.push(Array.from({ length: r2 - r1 + 1 }, (_, i) => [r1 + i, c]));
  else for (let r = r1; r <= r2; r++) lines.push(Array.from({ length: c2 - c1 + 1 }, (_, i) => [r, c1 + i]));
  let picked = lines;
  if (ch.pick) picked = ch.pick.map((k) => lines[k]).filter(Boolean);
  const header = ch.header ?? picked.some((l) => isText(val(...l[0])));
  const hasLabels = picked.length > 1 && picked[0].slice(header ? 1 : 0).some(([r, c]) => isText(val(r, c)));
  const labelLine = hasLabels ? picked[0] : null;
  const series = hasLabels ? picked.slice(1) : picked;
  const refOf = (cells) => `${sheetRef}!$${colName(cells[0][1])}$${cells[0][0] + 1}:$${colName(cells[cells.length - 1][1])}$${cells[cells.length - 1][0] + 1}`;
  const numCache = (cells) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${cells.length}"/>${cells.map(([r, c], i) => { const v = val(r, c); return typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : ''; }).join('')}</c:numCache>`;
  const strCache = (cells) => `<c:strCache><c:ptCount val="${cells.length}"/>${cells.map(([r, c], i) => `<c:pt idx="${i}"><c:v>${xesc(val(r, c) === null ? '' : display(val(r, c), src.get(r, c)?.s || {}))}</c:v></c:pt>`).join('')}</c:strCache>`;
  const t = ch.type;
  const scatter = t === 'scatter';
  const pie = t === 'pie' || t === 'doughnut';
  const colors = ch.colors?.length ? ch.colors : ['#3b6cf6', '#f0603f', '#2f9e6e', '#f5b400', '#9b5de5', '#00b4d8', '#ef476f', '#6c757d'];
  const serXml = series.map((cells, i) => {
    const data = header ? cells.slice(1) : cells;
    const cats = labelLine ? (header ? labelLine.slice(1) : labelLine) : null;
    const color = colors[i % colors.length].replace('#', '').toUpperCase();
    const spPr = pie ? '' : t === 'line' || scatter ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></a:ln></c:spPr>` : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
    const dPts = pie ? data.map((_, k) => `<c:dPt><c:idx val="${k}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${colors[k % colors.length].replace('#', '').toUpperCase()}"/></a:solidFill></c:spPr></c:dPt>`).join('') : '';
    const tx = header ? `<c:tx><c:strRef><c:f>${refOf([cells[0]])}</c:f>${strCache([cells[0]])}</c:strRef></c:tx>` : '';
    const catXml = cats ? (scatter ? `<c:xVal><c:numRef><c:f>${refOf(cats)}</c:f>${numCache(cats)}</c:numRef></c:xVal>` : `<c:cat><c:strRef><c:f>${refOf(cats)}</c:f>${strCache(cats)}</c:strRef></c:cat>`) : '';
    const valXml = scatter ? `<c:yVal><c:numRef><c:f>${refOf(data)}</c:f>${numCache(data)}</c:numRef></c:yVal>` : `<c:val><c:numRef><c:f>${refOf(data)}</c:f>${numCache(data)}</c:numRef></c:val>`;
    const extra = t === 'line' ? '<c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker><c:smooth val="0"/>' : scatter ? '<c:marker><c:symbol val="circle"/><c:size val="6"/></c:marker><c:smooth val="0"/>' : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${spPr}${t === 'column' || t === 'bar' ? '<c:invertIfNegative val="0"/>' : ''}${dPts}${scatter ? '' : catXml}${scatter ? catXml : ''}${valXml}${extra}</c:ser>`;
  }).join('');
  const axes = '<c:axId val="500000001"/><c:axId val="500000002"/>';
  const grouping = ch.stacked ? 'stacked' : 'clustered';
  let plot;
  if (t === 'column' || t === 'bar') plot = `<c:barChart><c:barDir val="${t === 'bar' ? 'bar' : 'col'}"/><c:grouping val="${grouping}"/><c:varyColors val="0"/>${serXml}<c:gapWidth val="80"/>${ch.stacked ? '<c:overlap val="100"/>' : ''}${axes}</c:barChart>`;
  else if (t === 'line') plot = `<c:lineChart><c:grouping val="${ch.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${serXml}<c:marker val="1"/>${axes}</c:lineChart>`;
  else if (t === 'area') plot = `<c:areaChart><c:grouping val="${ch.stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${serXml}${axes}</c:areaChart>`;
  else if (t === 'pie') plot = `<c:pieChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/></c:pieChart>`;
  else if (t === 'doughnut') plot = `<c:doughnutChart><c:varyColors val="1"/>${serXml}<c:firstSliceAng val="0"/><c:holeSize val="55"/></c:doughnutChart>`;
  else plot = `<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${serXml}${axes}</c:scatterChart>`;
  const catAx = scatter
    ? '<c:valAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="out"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>'
    : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${t === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="${t === 'area' ? 'midCat' : 'between'}"/></c:valAx>`;
  const title = ch.title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="1"/></a:pPr><a:r><a:rPr lang="pt-BR" sz="1400" b="1"/><a:t>${xesc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
  const legend = ch.legend === 'none' || (!pie && series.length < 2) ? '' : `<c:legend><c:legendPos val="${ch.legend === 'top' ? 't' : ch.legend === 'right' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend>`;
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot}${pie ? '' : catAx + valAx}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:spPr><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill><a:ln><a:noFill/></a:ln></c:spPr><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="pt-BR"/></a:p></c:txPr></c:chartSpace>`;
}

const quoteSheetName = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

function applyStyleOut(wc, st) {
  if (!st) return;
  const font = {};
  if (st.b) font.bold = true;
  if (st.i) font.italic = true;
  if (st.u) font.underline = true;
  if (st.st) font.strike = true;
  if (st.size) font.size = st.size;
  if (st.font) font.name = st.font.split(',')[0].replace(/"/g, '');
  if (st.color) font.color = { argb: hexToArgb(st.color) };
  if (Object.keys(font).length) wc.font = font;
  if (st.bg) wc.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: hexToArgb(st.bg) } };
  const al = {};
  if (st.ha) al.horizontal = st.ha;
  if (st.va) al.vertical = st.va === 'middle' ? 'middle' : st.va;
  if (st.wrap) al.wrapText = true;
  if (Object.keys(al).length) wc.alignment = al;
  const nf = numFmtFor(st);
  if (nf) wc.numFmt = nf;
  if (st.bd) {
    const side = (b) => (b ? { style: b.w === 'dashed' ? 'dashed' : b.w || 'thin', color: { argb: hexToArgb(b.color || '#1d2129') } } : undefined);
    wc.border = { top: side(st.bd.t), right: side(st.bd.r), bottom: side(st.bd.b), left: side(st.bd.l) };
  }
}

// ---- HTML / PDF ----------------------------------------------------------------

export function sheetHtml(wb, sheet, rg = null, chartImages = {}) {
  const u = rg || { r1: 0, c1: 0, r2: sheet.used().r, c2: sheet.used().c };
  const merged = new Set();
  let html = '<table>';
  html += '<colgroup>';
  for (let c = u.c1; c <= u.c2; c++) html += `<col style="width:${sheet.cols.size(c)}px">`;
  html += '</colgroup>';
  for (let r = u.r1; r <= u.r2; r++) {
    if (sheet.rows.hidden.has(r)) continue;
    html += `<tr style="height:${sheet.rows.size(r)}px">`;
    for (let c = u.c1; c <= u.c2; c++) {
      if (merged.has(KEY(r, c)) || sheet.cols.hidden.has(c)) continue;
      const m = sheet.mergeAt(r, c);
      let span = '';
      if (m) {
        for (let rr = m.r1; rr <= m.r2; rr++) for (let cc = m.c1; cc <= m.c2; cc++) if (rr !== r || cc !== c) merged.add(KEY(rr, cc));
        span = ` rowspan="${m.r2 - m.r1 + 1}" colspan="${m.c2 - m.c1 + 1}"`;
      }
      const cell = sheet.get(r, c);
      const st = cell?.s || {};
      const v = wb.value(sheet, r, c);
      const fmt = cell?.autoFmt && !st.fmt ? { ...st, fmt: cell.autoFmt } : st;
      const text = v === null ? '' : display(v, fmt);
      const css = [];
      if (st.b) css.push('font-weight:700');
      if (st.i) css.push('font-style:italic');
      if (st.u) css.push('text-decoration:underline');
      if (st.color) css.push(`color:${st.color}`);
      if (st.bg) css.push(`background:${st.bg}`);
      if (st.size) css.push(`font-size:${st.size}pt`);
      css.push(`text-align:${st.ha || (typeof v === 'number' ? 'right' : 'left')}`);
      if (st.wrap) css.push('white-space:pre-wrap');
      if (st.bd) for (const [k, side] of [['t', 'top'], ['r', 'right'], ['b', 'bottom'], ['l', 'left']]) if (st.bd[k]) css.push(`border-${side}:${st.bd[k].w === 'thick' ? 3 : st.bd[k].w === 'medium' ? 2 : 1}px solid ${st.bd[k].color || '#000'}`);
      html += `<td${span} style="${css.join(';')}">${escapeHtml(text)}</td>`;
    }
    html += '</tr>';
  }
  html += '</table>';
  for (const ch of sheet.charts) if (chartImages[ch.id]) html += `<div class="chart"><img src="${chartImages[ch.id]}" style="width:${ch.w}px;height:${ch.h}px"></div>`;
  return html;
}

export function printableHtml(title, body, { landscape = true, gridlines = true } = {}) {
  return `<!doctype html><html lang="pt-BR"><head><meta charset="utf-8"><title>${escapeHtml(title)}</title><style>
@page { size: A4 ${landscape ? 'landscape' : 'portrait'}; margin: 1.2cm; }
body { font: 10pt Calibri, Carlito, Arial, sans-serif; color: #111; margin: 0; }
h1 { font-size: 13pt; margin: 0 0 8px; }
table { border-collapse: collapse; table-layout: fixed; }
td { padding: 2px 5px; overflow: hidden; white-space: nowrap; ${gridlines ? 'border: 1px solid #d0d4da;' : ''} vertical-align: bottom; }
.chart { margin-top: 16px; break-inside: avoid; }
tr { break-inside: avoid; }
</style></head><body>${body}</body></html>`;
}

export { colName, tokenize };
