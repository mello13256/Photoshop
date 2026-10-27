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
  if (ext === 'xlsx') return importXlsx(data);
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

export async function exportXlsx(wb, chartImages = {}) {
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
      } else if (cell.v !== null && cell.v !== undefined) writeValue(wc, cell.v);
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
    for (const ch of s.charts) {
      const img = chartImages[ch.id];
      if (!img) continue;
      const id = x.addImage({ base64: img, extension: 'png' });
      ws.addImage(id, { tl: { col: s.cols.indexAt(ch.x), row: s.rows.indexAt(ch.y) }, ext: { width: ch.w, height: ch.h } });
    }
  }
  return new Uint8Array(await x.xlsx.writeBuffer());
}

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
