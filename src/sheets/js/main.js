// Lumina Planilhas — aplicação principal.
import {
  el, $, toast, dialog, prompt, field, select, menubar, toolButton, sep, popover, closePopovers, colorPopover, dropdown,
  contextMenu, clamp, escapeHtml, comboOf, isDialogOpen, debounce,
} from '../../shared/ui.js';
import { icon } from '../../shared/icons.js';
import {
  pickFiles, saveFile, exportPdf, isNative, registerApp, standaloneBoot, setTitle, addRecent, extOf, fileName, initTheme,
} from '../../shared/native.js';
import { Workbook, KEY, RC, stripCell, clone, rangeText } from './model.js';
import { Grid, CellEditor, norm, inRange, RH_W, CH_H } from './grid.js';
import { parseInput, display, FORMATS, editText, fmtDate } from './format.js';
import {
  isErr, isMatrix, colName, parseAddr, compare, numberToText, colIndex,
} from './values.js';
import { shiftFormula, parse, tokenize } from './formula.js';
import { FUNCS, CATEGORIES } from './functions.js';
import { ChartLayer, CHART_TYPES, defaultChartPlacement } from './charts.js';
import * as io from './io.js';
import { SHEET_TEMPLATES } from './templates.js';

const S = { books: [], book: null, seq: 1, clipboard: null, painter: null };
const wb = () => S.book.wb;
const sheet = () => S.book.wb.sheet;

// ---------------------------------------------------------------------------
// Grade, editor e gráficos

const sink = $('#kbd');
const bar = $('#formula-input');
const nameBox = $('#name-box');

const appApi = {
  get wb() { return S.book.wb; },
  editor: null,
  onSelect: () => selectionChanged(),
  onScroll: () => { charts.position(); positionValidationButton(); if (editor.active) editor.position(); },
  onZoom: (z) => { $('#zoom-val').textContent = `${Math.round(z * 100)}%`; charts.position(); },
  onContextMenu: (e, h) => cellMenu(e, h),
  onCellClick: (r, c, e) => cellClicked(r, c, e),
  resizeAxis: (axis, i, size) => resizeAxis(axis, i, size),
  autofit: (axis, i) => autofit(axis, i),
  fill: (from, to) => fill(from, to),
  fillDownToNeighbors: () => fillDownToNeighbors(),
  startEdit: (o) => editor.begin(o),
  focusGrid: () => focusGrid(),
  setCellInput: (r, c, v, o) => setCellInput(r, c, v, o),
  moveAfterEdit: (d) => moveAfterEdit(d),
  returnToSheet: (id) => returnToSheet(id),
  afterRender: () => positionValidationButton(),
  updateChart: (id, ch, label) => updateChart(id, ch, label),
  editChart: (id) => chartDialog(id),
  chartMenu: (id, e) => chartMenu(id, e),
  onChartSelect: () => {},
};
const grid = new Grid($('#grid'), appApi);
const editor = new CellEditor(appApi, grid, bar);
appApi.editor = editor;
const charts = new ChartLayer(appApi, grid);

function focusGrid() {
  if (!isDialogOpen()) sink.focus({ preventScroll: true });
}

// ---------------------------------------------------------------------------
// Pastas de trabalho (abas de arquivo)

function makeBook({ wb: w, title, filePath = null, ext = 'lumsheet' } = {}) {
  const b = { id: S.seq++, wb: w || new Workbook(), title: title || 'Planilha sem título', filePath, ext, dirty: false, view: {} };
  b.wb.onChange(() => {
    if (!b.dirty) { b.dirty = true; renderTitle(); }
    if (b === S.book) refresh();
  });
  return b;
}

function isPristine(b) {
  return b && !b.dirty && !b.filePath && b.wb.sheets.length === 1 && !b.wb.sheets[0].cells.size && b.title === 'Planilha sem título';
}

function addBook(b) {
  saveView();
  if (S.books.length === 1 && isPristine(S.books[0])) S.books = [];
  S.books.push(b);
  activate(b);
}

function saveView() {
  const b = S.book;
  if (!b) return;
  b.view[b.wb.sheet.id] = { sel: { ...grid.sel }, sx: grid.scroller.scrollLeft, sy: grid.scroller.scrollTop };
}

function activate(b) {
  if (editor.active) editor.commit();
  saveView();
  S.book = b;
  restoreView();
  renderAll();
}

function restoreView() {
  const v = S.book.view[sheet().id];
  grid.copyRange = null;
  grid.resize();
  grid.sel = v?.sel ? { ...v.sel } : { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
  grid.scroller.scrollLeft = v?.sx || 0;
  grid.scroller.scrollTop = v?.sy || 0;
  grid.invalidate();
}

async function closeBook(b = S.book) {
  if (b.dirty && !(await askSave(b))) return false;
  const i = S.books.indexOf(b);
  S.books.splice(i, 1);
  if (!S.books.length) S.books.push(makeBook());
  S.book = null;
  activate(S.books[Math.min(i, S.books.length - 1)]);
  return true;
}

async function askSave(b) {
  if (b !== S.book) activate(b);
  let choice = 'cancel';
  const ok = await dialog({
    title: 'Salvar alterações?', okText: 'Salvar',
    body: el('p', {}, `A planilha "${b.title}" tem alterações não salvas.`),
    extraButtons: [{ label: 'Não salvar', onClick: (close) => { choice = 'discard'; close(false); } }],
  });
  if (ok) return save(b);
  return choice === 'discard';
}

function newBook(tpl = 'blank') {
  const t = SHEET_TEMPLATES[tpl] || SHEET_TEMPLATES.blank;
  const b = makeBook({ wb: t.build(), title: t.title });
  addBook(b);
}

async function openFiles(files) {
  for (const f of files) {
    try {
      if (extOf(f.name) === 'pptx') { toast('Arquivos .pptx abrem no Lumina Apresentações.'); continue; }
      const w = await io.importFile(f);
      const ext = extOf(f.name);
      addBook(makeBook({ wb: w, title: f.name.replace(/\.[^.]+$/, ''), filePath: f.path, ext }));
      addRecent({ app: 'sheets', name: f.name, path: f.path });
    } catch (err) {
      console.error(err);
      toast(`Não foi possível abrir ${f.name}: ${err.message}`);
    }
  }
}

async function openDialog() {
  const files = await pickFiles([
    { name: 'Planilhas', extensions: ['lumsheet', 'xlsx', 'csv', 'tsv'] },
    { name: 'Planilha Lumina', extensions: ['lumsheet'] }, { name: 'Excel', extensions: ['xlsx'] }, { name: 'CSV', extensions: ['csv', 'tsv'] },
  ]);
  if (files.length) openFiles(files);
}

function chartImages(w = wb()) {
  const out = {};
  for (const s of w.sheets) for (const ch of s.charts) {
    const img = charts.image(ch.id);
    if (img) out[ch.id] = img;
  }
  return out;
}

async function save(b = S.book, saveAs = false) {
  if (editor.active) editor.commit();
  const keep = !saveAs && b.filePath && ['xlsx', 'csv', 'tsv'].includes(b.ext);
  try {
    let path;
    if (keep) path = await writeFormat(b, b.ext, b.filePath);
    else {
      path = await saveFile(io.serialize(b.wb), {
        name: `${b.title}.lumsheet`, filters: [{ name: 'Planilha Lumina', extensions: ['lumsheet'] }],
        path: !saveAs && b.ext === 'lumsheet' ? b.filePath : null,
      });
      if (path) b.ext = 'lumsheet';
    }
    if (!path) return false;
    if (isNative) b.filePath = path;
    b.dirty = false;
    addRecent({ app: 'sheets', name: fileName(path), path: isNative ? path : null });
    toast(isNative ? `Salvo: ${fileName(path)}` : 'Planilha baixada.');
    renderTitle();
    return true;
  } catch (err) {
    console.error(err);
    toast(`Erro ao salvar: ${err.message}`);
    return false;
  }
}

async function writeFormat(b, fmt, path = null) {
  const name = `${b.title}.${fmt}`;
  if (fmt === 'xlsx') return saveFile(await io.exportXlsx(b.wb, chartImages(b.wb)), { name, filters: [{ name: 'Pasta de trabalho do Excel', extensions: ['xlsx'] }], path });
  if (fmt === 'csv' || fmt === 'tsv') {
    const s = b.wb.sheet;
    if (b.wb.sheets.length > 1) toast(`CSV guarda só a planilha atual ("${s.name}").`);
    return saveFile(io.exportCsv(b.wb, s, fmt === 'tsv' ? '\t' : ';'), { name, filters: [{ name: 'CSV', extensions: [fmt] }], path });
  }
  throw new Error('Formato desconhecido');
}

async function exportAs(fmt) {
  if (editor.active) editor.commit();
  try {
    if (fmt === 'pdf') {
      const b = S.book;
      const html = io.printableHtml(b.title, io.sheetHtml(b.wb, sheet(), null, chartImages()), { landscape: true, gridlines: sheet().showGrid });
      const out = await exportPdf(html, { name: `${b.title}.pdf` });
      if (out) toast(isNative ? `PDF exportado: ${fileName(out)}` : 'Use "Salvar como PDF" na janela de impressão.');
      return;
    }
    const out = await writeFormat(S.book, fmt);
    if (out) toast(`Exportado: ${fileName(out)}`);
  } catch (err) {
    console.error(err);
    toast(`Erro ao exportar: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// Edição de células

function isLocked() { return false; }
appApi.isLocked = isLocked;

function cellAt(s, r, c) {
  return s.get(r, c);
}

function setCellInput(r, c, text, opts = {}) {
  const s = sheet();
  const w = wb();
  const apply = (rr, cc, txt) => {
    const old = cellAt(s, rr, cc);
    const style = old?.s ? { ...old.s } : undefined;
    let cell;
    if (typeof txt === 'string' && txt.startsWith('=') && txt.length > 1) {
      let f = txt.slice(1);
      f = autoClose(f);
      cell = { v: null, f, s: style };
      try { parse(f); } catch (e) { toast(`Erro na fórmula: ${e.message}`); }
    } else {
      const p = parseInput(txt);
      cell = { v: p.v, s: style };
      if (p.forceText) cell.forceText = true;
      if (p.fmt && (!style?.fmt || style.fmt === 'general')) cell.s = { ...(style || {}), fmt: p.fmt, ...(p.dec !== undefined ? { dec: p.dec } : {}), ...(p.cur ? { cur: p.cur } : {}) };
    }
    if (old?.note) cell.note = old.note;
    w.setCell(s, rr, cc, cell);
  };
  w.transact('Editar célula', () => {
    if (opts.fillSelection) {
      const sel = grid.sel;
      for (let rr = sel.r1; rr <= sel.r2; rr++) for (let cc = sel.c1; cc <= sel.c2; cc++) {
        const t = text.startsWith('=') ? `=${shiftFormula(text.slice(1), rr - r, cc - c)}` : text;
        apply(rr, cc, t);
      }
    } else apply(r, c, text);
  });
  validateCell(r, c);
}

// Fecha parênteses esquecidos, como o Excel.
function autoClose(f) {
  let depth = 0, inStr = false;
  for (const ch of f) {
    if (ch === '"') inStr = !inStr;
    else if (!inStr && ch === '(') depth++;
    else if (!inStr && ch === ')') depth--;
  }
  return depth > 0 ? f + ')'.repeat(depth) : f;
}

function validateCell(r, c) {
  const v = sheet().validations.find((x) => inRange(x.range, r, c));
  if (!v || v.type !== 'list') return;
  const val = wb().value(sheet(), r, c);
  if (val === null || val === '') return;
  if (!v.items.some((i) => String(i).toLowerCase() === String(val).toLowerCase())) toast(`Valor fora da lista permitida: ${v.items.join(', ')}`);
}

function moveAfterEdit([dr, dc]) {
  const sel = grid.sel;
  const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
  if (multi) {
    // Move dentro da seleção, como no Excel
    let { r, c } = sel;
    if (dr) { r += dr; if (r > sel.r2) { r = sel.r1; c = c + 1 > sel.c2 ? sel.c1 : c + 1; } if (r < sel.r1) r = sel.r2; }
    if (dc) { c += dc; if (c > sel.c2) { c = sel.c1; r = r + 1 > sel.r2 ? sel.r1 : r + 1; } if (c < sel.c1) c = sel.c2; }
    grid.sel = { ...sel, r, c };
    grid.invalidate();
    selectionChanged();
  } else {
    grid.select(sel.r + dr, sel.c + dc);
  }
  grid.scrollIntoView(grid.sel.r, grid.sel.c);
}

function returnToSheet(id) {
  if (sheet().id === id) return;
  const i = wb().sheets.findIndex((s) => s.id === id);
  if (i >= 0) { wb().active = i; restoreView(); renderTabs(); }
}

// Itera células da seleção limitadas à área usada (para seleções de colunas inteiras).
function selectionCells(rg = grid.sel, { createWithin = true } = {}) {
  const s = sheet();
  const u = s.used();
  const r2 = rg.r2 >= s.rows.count - 1 ? Math.max(u.r, rg.r1) : rg.r2;
  const c2 = rg.c2 >= s.cols.count - 1 ? Math.max(u.c, rg.c1) : rg.c2;
  const out = [];
  for (let r = rg.r1; r <= r2; r++) for (let c = rg.c1; c <= c2; c++) out.push([r, c]);
  void createWithin;
  return out;
}

function applyStyle(patch, label = 'Formatar') {
  const s = sheet();
  wb().transact(label, () => {
    for (const [r, c] of selectionCells()) {
      const old = s.get(r, c);
      const st = { ...(old?.s || {}) };
      for (const [k, v] of Object.entries(patch)) {
        if (v === null || v === undefined) delete st[k];
        else st[k] = v;
      }
      wb().setCell(s, r, c, { ...(old ? stripCell(old) : { v: null }), s: st });
    }
  });
}

function toggleStyle(key) {
  const cur = sheet().get(grid.sel.r, grid.sel.c)?.s?.[key];
  applyStyle({ [key]: cur ? null : true }, 'Formatar');
}

function clearRange(what = 'contents') {
  const s = sheet();
  const label = { contents: 'Limpar conteúdo', formats: 'Limpar formatos', all: 'Limpar tudo' }[what];
  wb().transact(label, () => {
    for (const [r, c] of selectionCells()) {
      const old = s.get(r, c);
      if (!old) continue;
      if (what === 'all') wb().setCell(s, r, c, null);
      else if (what === 'formats') wb().setCell(s, r, c, { ...stripCell(old), s: undefined });
      else wb().setCell(s, r, c, old.s ? { v: null, s: old.s } : null);
    }
  });
}

function setBorders(kind, color = '#1d2129', w = 'thin') {
  const s = sheet();
  const sel = grid.sel;
  const rg = selectionCells();
  const b = { w, color };
  wb().transact('Bordas', () => {
    for (const [r, c] of rg) {
      const old = s.get(r, c);
      const st = { ...(old?.s || {}) };
      const bd = { ...(st.bd || {}) };
      const top = r === sel.r1, bottom = r === Math.min(sel.r2, rg[rg.length - 1][0]), left = c === sel.c1, right = c === Math.min(sel.c2, rg[rg.length - 1][1]);
      switch (kind) {
        case 'all': Object.assign(bd, { t: b, r: b, b, l: b }); break;
        case 'outer': case 'thick':
          if (top) bd.t = kind === 'thick' ? { ...b, w: 'medium' } : b;
          if (bottom) bd.b = kind === 'thick' ? { ...b, w: 'medium' } : b;
          if (left) bd.l = kind === 'thick' ? { ...b, w: 'medium' } : b;
          if (right) bd.r = kind === 'thick' ? { ...b, w: 'medium' } : b;
          break;
        case 'bottom': if (bottom) bd.b = b; break;
        case 'top': if (top) bd.t = b; break;
        case 'left': if (left) bd.l = b; break;
        case 'right': if (right) bd.r = b; break;
        case 'doubleBottom': if (bottom) bd.b = { ...b, w: 'thick' }; break;
        case 'none': delete bd.t; delete bd.b; delete bd.l; delete bd.r; break;
        default:
      }
      if (Object.keys(bd).length) st.bd = bd; else delete st.bd;
      wb().setCell(s, r, c, { ...(old ? stripCell(old) : { v: null }), s: st });
    }
  });
}

function mergeCells(center = true) {
  const s = sheet();
  const sel = { ...grid.sel };
  const existing = s.merges.filter((m) => !(m.r2 < sel.r1 || m.r1 > sel.r2 || m.c2 < sel.c1 || m.c1 > sel.c2));
  if (existing.length && existing.every((m) => m.r1 >= sel.r1 && m.r2 <= sel.r2 && m.c1 >= sel.c1 && m.c2 <= sel.c2) && sel.r1 === existing[0].r1 && sel.c1 === existing[0].c1 && existing.length === 1 && existing[0].r2 === sel.r2 && existing[0].c2 === sel.c2) {
    wb().transact('Desfazer mesclagem', () => { wb().snapSheet(s); s.merges = s.merges.filter((m) => m !== existing[0]); });
    return;
  }
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) return;
  let others = 0;
  for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) if ((r !== sel.r1 || c !== sel.c1) && wb().value(s, r, c) !== null) others++;
  const go = () => wb().transact('Mesclar células', () => {
    wb().snapSheet(s);
    s.merges = s.merges.filter((m) => !existing.includes(m));
    for (let r = sel.r1; r <= sel.r2; r++) for (let c = sel.c1; c <= sel.c2; c++) if (r !== sel.r1 || c !== sel.c1) s.set(r, c, null);
    s.merges.push({ r1: sel.r1, c1: sel.c1, r2: sel.r2, c2: sel.c2 });
    if (center) {
      const old = s.get(sel.r1, sel.c1);
      s.set(sel.r1, sel.c1, { ...(old ? stripCell(old) : { v: null }), s: { ...(old?.s || {}), ha: 'center', va: 'middle' } });
    }
  });
  if (others) {
    dialog({ title: 'Mesclar células', body: el('p', {}, 'Mesclar mantém apenas o valor da célula superior esquerda. Continuar?'), okText: 'Mesclar' }).then((ok) => ok && go());
  } else go();
  grid.sel = { ...sel, r: sel.r1, c: sel.c1 };
}

function setFormat(fmt, extra = {}) {
  applyStyle({ fmt, dec: extra.dec ?? null, cur: extra.cur ?? null, code: extra.code ?? null }, 'Formato de número');
}

function stepDecimals(d) {
  const st = sheet().get(grid.sel.r, grid.sel.c)?.s || {};
  let fmt = st.fmt;
  if (!fmt || fmt === 'general') {
    fmt = 'number';
    const v = wb().value(sheet(), grid.sel.r, grid.sel.c);
    const cur = typeof v === 'number' ? (numberToText(v).split(',')[1] || '').length : 0;
    applyStyle({ fmt, dec: clamp(cur + d, 0, 12) }, 'Casas decimais');
    return;
  }
  const def = { percent: 0 }[fmt] ?? 2;
  applyStyle({ dec: clamp((st.dec ?? def) + d, 0, 12) }, 'Casas decimais');
}

// ---------------------------------------------------------------------------
// Linhas e colunas

function resizeAxis(axis, i, size) {
  const s = sheet();
  const sel = grid.sel;
  const isCol = axis === 'col';
  const whole = isCol ? sel.r1 === 0 && sel.r2 >= s.rows.count - 1 : sel.c1 === 0 && sel.c2 >= s.cols.count - 1;
  const targets = whole && (isCol ? i >= sel.c1 && i <= sel.c2 : i >= sel.r1 && i <= sel.r2)
    ? Array.from({ length: (isCol ? sel.c2 - sel.c1 : sel.r2 - sel.r1) + 1 }, (_, k) => (isCol ? sel.c1 : sel.r1) + k) : [i];
  wb().transact(isCol ? 'Largura da coluna' : 'Altura da linha', () => {
    wb().snapSheet(s);
    for (const t of targets) (isCol ? s.cols : s.rows).setSize(t, size);
  });
  grid.updateSpacer();
}

function autofit(axis, i) {
  const s = sheet();
  const ctx = grid.ctx;
  if (axis === 'col') {
    let max = 30;
    const u = s.used();
    for (let r = 0; r <= u.r; r++) {
      const cell = s.get(r, i);
      const v = wb().value(s, r, i);
      if (v === null) continue;
      const st = cell?.s || {};
      ctx.font = `${st.b ? '700 ' : ''}${(st.size || 11) * (4 / 3)}px ${st.font || 'Calibri, Carlito, Arial'}`;
      max = Math.max(max, ctx.measureText(display(v, cell?.autoFmt && !st.fmt ? { ...st, fmt: cell.autoFmt } : st)).width + 14);
    }
    resizeAxis('col', i, Math.min(600, max));
  } else {
    let max = 20;
    for (const [k, cell] of s.cells) {
      const [r] = RC(k);
      if (r !== i) continue;
      max = Math.max(max, (cell.s?.size || 11) * (4 / 3) * 1.5);
    }
    resizeAxis('row', i, max);
  }
}

function insertCells(axis, where) {
  const sel = grid.sel;
  const s = sheet();
  if (axis === 'row') {
    const n = sel.r2 - sel.r1 + 1;
    wb().insert(s, 'row', where === 'before' ? sel.r1 : sel.r2 + 1, n);
  } else {
    const n = sel.c2 - sel.c1 + 1;
    wb().insert(s, 'col', where === 'before' ? sel.c1 : sel.c2 + 1, n);
  }
  grid.updateSpacer();
}

function deleteCells(axis) {
  const sel = grid.sel;
  const s = sheet();
  if (axis === 'row') wb().insert(s, 'row', sel.r1, -(sel.r2 - sel.r1 + 1));
  else wb().insert(s, 'col', sel.c1, -(sel.c2 - sel.c1 + 1));
  grid.select(sel.r1, sel.c1);
  grid.updateSpacer();
}

function hideAxis(axis, hide) {
  const s = sheet();
  const sel = grid.sel;
  wb().transact(hide ? 'Ocultar' : 'Reexibir', () => {
    wb().snapSheet(s);
    if (axis === 'row') for (let r = sel.r1; r <= Math.min(sel.r2, s.rows.count - 1); r++) s.rows.setHidden(r, hide);
    else for (let c = sel.c1; c <= Math.min(sel.c2, s.cols.count - 1); c++) s.cols.setHidden(c, hide);
    if (!hide && axis === 'row' && sel.r1 === sel.r2) { s.rows.setHidden(sel.r1 - 1, false); s.rows.setHidden(sel.r1 + 1, false); }
    if (!hide && axis === 'col' && sel.c1 === sel.c2) { s.cols.setHidden(sel.c1 - 1, false); s.cols.setHidden(sel.c1 + 1, false); }
  });
  grid.updateSpacer();
}

function freeze(kind) {
  const s = sheet();
  const v = grid.sel;
  wb().transact('Congelar painéis', () => {
    wb().snapSheet(s);
    s.freeze = kind === 'row' ? { r: 1, c: 0 } : kind === 'col' ? { r: 0, c: 1 } : kind === 'sel' ? { r: v.r, c: v.c } : { r: 0, c: 0 };
  });
  grid.scroller.scrollTop = 0;
  grid.scroller.scrollLeft = 0;
}

// ---------------------------------------------------------------------------
// Preenchimento (alça) e séries

const MONTHS = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const MONTHS3 = MONTHS.map((m) => m.slice(0, 3));
const DAYS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
const DAYS_SHORT = ['domingo', 'segunda', 'terça', 'quarta', 'quinta', 'sexta', 'sábado'];
const DAYS3 = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sáb'];
const LISTS = [MONTHS, MONTHS3, DAYS, DAYS_SHORT, DAYS3];

function matchCase(word, sample) {
  if (sample === sample.toUpperCase()) return word.toUpperCase();
  if (sample[0] === sample[0].toUpperCase()) return word[0].toUpperCase() + word.slice(1);
  return word;
}

// Gera o valor k (0-based, a partir do fim da origem) de uma série a partir das células de origem.
function seriesGen(cells) {
  const vals = cells.map((c) => (c && !c.f ? c.v : null));
  if (cells.some((c) => c?.f)) return null;
  if (vals.every((v) => typeof v === 'number') && vals.length) {
    if (vals.length === 1) {
      const isDate = ['date', 'datetime'].includes(cells[0]?.s?.fmt);
      return isDate ? (k) => vals[0] + k + 1 : null;
    }
    const n = vals.length;
    const xs = vals.map((_, i) => i);
    const mx = (n - 1) / 2, my = vals.reduce((a, b) => a + b, 0) / n;
    let sxy = 0, sxx = 0;
    xs.forEach((x, i) => { sxy += (x - mx) * (vals[i] - my); sxx += (x - mx) ** 2; });
    const slope = sxy / sxx, icpt = my - slope * mx;
    return (k) => +(icpt + slope * (n + k)).toPrecision(15);
  }
  if (vals.length === 1 && typeof vals[0] === 'string') {
    const v = vals[0];
    const low = v.toLowerCase();
    for (const list of LISTS) {
      const i = list.indexOf(low);
      if (i >= 0) return (k) => matchCase(list[(i + k + 1) % list.length], v);
    }
    const m = /^(.*?)(\d+)(\D*)$/.exec(v);
    if (m) return (k) => `${m[1]}${String(+m[2] + k + 1).padStart(m[2].startsWith('0') ? m[2].length : 0, '0')}${m[3]}`;
  }
  return null;
}

function fill(from, to) {
  const s = sheet();
  const down = to.r2 > from.r2, up = to.r1 < from.r1, right = to.c2 > from.c2;
  const vertical = down || up;
  wb().transact('Preencher', () => {
    if (vertical) {
      for (let c = from.c1; c <= from.c2; c++) {
        const src = [];
        for (let r = from.r1; r <= from.r2; r++) src.push(s.get(r, c));
        const gen = seriesGen(up ? [...src].reverse() : src);
        const targets = down ? range(from.r2 + 1, to.r2) : range(to.r1, from.r1 - 1).reverse();
        targets.forEach((r, k) => {
          const srcIdx = down ? k % src.length : src.length - 1 - (k % src.length);
          const srcR = from.r1 + srcIdx;
          writeFill(s, r, c, src[srcIdx], srcR, c, gen ? gen(k) : undefined);
        });
      }
    } else {
      for (let r = from.r1; r <= from.r2; r++) {
        const src = [];
        for (let c = from.c1; c <= from.c2; c++) src.push(s.get(r, c));
        const gen = seriesGen(right ? src : [...src].reverse());
        const targets = right ? range(from.c2 + 1, to.c2) : range(to.c1, from.c1 - 1).reverse();
        targets.forEach((c, k) => {
          const srcIdx = right ? k % src.length : src.length - 1 - (k % src.length);
          writeFill(s, r, c, src[srcIdx], r, from.c1 + srcIdx, gen ? gen(k) : undefined);
        });
      }
    }
  });
  grid.setRange(norm(Math.min(from.r1, to.r1), Math.min(from.c1, to.c1), Math.max(from.r2, to.r2), Math.max(from.c2, to.c2)), { r: from.r1, c: from.c1 });
}

const range = (a, b) => (b < a ? [] : Array.from({ length: b - a + 1 }, (_, i) => a + i));

function writeFill(s, r, c, src, sr, sc, genValue) {
  if (!src) { wb().setCell(s, r, c, null); return; }
  const n = clone(stripCell(src));
  if (n.f) n.f = shiftFormula(n.f, r - sr, c - sc);
  else if (genValue !== undefined) n.v = genValue;
  wb().setCell(s, r, c, n);
}

function fillDownToNeighbors() {
  const s = sheet();
  const sel = grid.sel;
  const probe = (c) => {
    if (c < 0) return -1;
    let r = sel.r2 + 1;
    if (wb().value(s, r, c) === null) return -1;
    while (wb().value(s, r + 1, c) !== null) r++;
    return r;
  };
  const end = Math.max(probe(sel.c1 - 1), probe(sel.c2 + 1));
  if (end > sel.r2) fill({ ...sel }, { ...sel, r2: end });
}

// ---------------------------------------------------------------------------
// Área de transferência

function rangeMatrix(rg) {
  const s = sheet();
  const out = [];
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    for (let c = rg.c1; c <= rg.c2; c++) row.push(clone(stripCell(s.get(r, c))));
    out.push(row);
  }
  return out;
}

function clipRange() {
  const s = sheet();
  const u = s.used();
  const sel = grid.sel;
  return { r1: sel.r1, c1: sel.c1, r2: Math.min(sel.r2, Math.max(u.r, sel.r1)), c2: Math.min(sel.c2, Math.max(u.c, sel.c1)) };
}

function onCopy(e, cut = false) {
  if (editor.active) return;
  e.preventDefault();
  const rg = clipRange();
  const s = sheet();
  const rows = [];
  let html = '<table>';
  for (let r = rg.r1; r <= rg.r2; r++) {
    const row = [];
    html += '<tr>';
    for (let c = rg.c1; c <= rg.c2; c++) {
      const cell = s.get(r, c);
      const v = wb().value(s, r, c);
      const t = v === null ? '' : display(v, cell?.autoFmt && !cell?.s?.fmt ? { ...(cell?.s || {}), fmt: cell.autoFmt } : cell?.s || {});
      row.push(/[\t\n"]/.test(t) ? `"${t.replace(/"/g, '""')}"` : t);
      const st = cell?.s || {};
      html += `<td style="${st.b ? 'font-weight:bold;' : ''}${st.bg ? `background:${st.bg};` : ''}${st.color ? `color:${st.color};` : ''}">${escapeHtml(t)}</td>`;
    }
    rows.push(row.join('\t'));
    html += '</tr>';
  }
  html += '</table>';
  const tsv = rows.join('\r\n');
  e.clipboardData.setData('text/plain', tsv);
  e.clipboardData.setData('text/html', html);
  S.clipboard = { tsv, sheetId: s.id, range: rg, cells: rangeMatrix(rg), cut };
  grid.copyRange = { ...rg, sheet: s.id };
  grid.invalidate();
  toast(cut ? 'Recortado. Cole com Ctrl+V.' : `Copiado ${rg.r2 - rg.r1 + 1} × ${rg.c2 - rg.c1 + 1}.`, 1400);
}

function onPaste(e, valuesOnly = false) {
  if (editor.active) return;
  e.preventDefault();
  const text = e.clipboardData?.getData('text/plain') ?? '';
  const html = e.clipboardData?.getData('text/html') ?? '';
  const internal = S.clipboard && (S.clipboard.tsv === text.replace(/\r?\n$/, '') || S.clipboard.tsv === text);
  if (internal) pasteInternal(S.clipboard, valuesOnly);
  else pasteExternal(text, html);
}

function pasteInternal(clip, valuesOnly = false) {
  const s = sheet();
  const sel = grid.sel;
  const h = clip.cells.length, w = clip.cells[0].length;
  // Repete o conteúdo se a seleção for múltipla do tamanho copiado
  const reps = (sel.r2 - sel.r1 + 1) % h === 0 && (sel.c2 - sel.c1 + 1) % w === 0 && (sel.r2 - sel.r1 + 1) * (sel.c2 - sel.c1 + 1) > 1
    ? [(sel.r2 - sel.r1 + 1) / h, (sel.c2 - sel.c1 + 1) / w] : [1, 1];
  const srcSheet = wb().byId(clip.sheetId);
  wb().transact(clip.cut ? 'Mover' : 'Colar', () => {
    const values = valuesOnly ? clip.cells.map((row, i) => row.map((_, j) => (srcSheet ? wb().value(srcSheet, clip.range.r1 + i, clip.range.c1 + j) : null))) : null;
    if (clip.cut && srcSheet) {
      for (let r = clip.range.r1; r <= clip.range.r2; r++) for (let c = clip.range.c1; c <= clip.range.c2; c++) wb().setCell(srcSheet, r, c, null);
    }
    for (let a = 0; a < reps[0]; a++) for (let b = 0; b < reps[1]; b++) {
      clip.cells.forEach((row, i) => row.forEach((cell, j) => {
        const r = sel.r1 + a * h + i, c = sel.c1 + b * w + j;
        if (valuesOnly) {
          const v = values[i][j];
          const old = s.get(r, c);
          wb().setCell(s, r, c, { v: isMatrix(v) ? v[0][0] : isErr(v) ? v.code : v, s: old?.s || cell?.s });
          return;
        }
        if (!cell) { wb().setCell(s, r, c, null); return; }
        const n = clone(cell);
        if (n.f && !clip.cut) n.f = shiftFormula(n.f, r - (clip.range.r1 + i), c - (clip.range.c1 + j));
        wb().setCell(s, r, c, n);
      }));
    }
  });
  grid.setRange({ r1: sel.r1, c1: sel.c1, r2: sel.r1 + h * reps[0] - 1, c2: sel.c1 + w * reps[1] - 1 }, { r: sel.r1, c: sel.c1 });
  if (clip.cut) { S.clipboard = null; grid.copyRange = null; }
}

function pasteExternal(text, html) {
  let rows = null;
  if (html && /<table/i.test(html)) {
    const doc = new DOMParser().parseFromString(html, 'text/html');
    const t = doc.querySelector('table');
    if (t) rows = [...t.rows].map((tr) => [...tr.cells].map((td) => td.innerText.replace(/ /g, ' ').trim()));
  }
  if (!rows) {
    if (!text) return;
    rows = text.includes('\t') || text.includes('\n') ? io.parseDelimited(text.replace(/\r?\n$/, ''), '\t') : [[text]];
  }
  const s = sheet();
  const sel = grid.sel;
  wb().transact('Colar', () => {
    rows.forEach((row, i) => row.forEach((val, j) => {
      const r = sel.r1 + i, c = sel.c1 + j;
      const old = s.get(r, c);
      if (val.startsWith('=')) { wb().setCell(s, r, c, { v: null, f: val.slice(1), s: old?.s }); return; }
      const p = parseInput(val);
      wb().setCell(s, r, c, p.v === null && !old?.s ? null : { v: p.v, s: p.fmt ? { ...(old?.s || {}), fmt: p.fmt, ...(p.dec !== undefined ? { dec: p.dec } : {}) } : old?.s });
    }));
  });
  grid.setRange({ r1: sel.r1, c1: sel.c1, r2: sel.r1 + rows.length - 1, c2: sel.c1 + Math.max(...rows.map((r) => r.length)) - 1 }, { r: sel.r1, c: sel.c1 });
}

async function pasteFromMenu(valuesOnly = false) {
  try {
    const text = await navigator.clipboard.readText();
    const fake = { preventDefault() {}, clipboardData: { getData: (t) => (t === 'text/plain' ? text : '') } };
    onPaste(fake, valuesOnly);
  } catch {
    if (S.clipboard) pasteInternal(S.clipboard, valuesOnly);
    else toast('Use Ctrl+V para colar.');
  }
}

// ---------------------------------------------------------------------------
// Navegação por teclado

function dataEdge(r, c, dr, dc) {
  const s = sheet();
  const has = (rr, cc) => wb().value(s, rr, cc) !== null;
  const maxR = s.rows.count - 1, maxC = s.cols.count - 1;
  const inb = (rr, cc) => rr >= 0 && cc >= 0 && rr <= maxR && cc <= maxC;
  let nr = r + dr, nc = c + dc;
  if (!inb(nr, nc)) return [r, c];
  if (has(r, c) && has(nr, nc)) {
    while (inb(nr + dr, nc + dc) && has(nr + dr, nc + dc)) { nr += dr; nc += dc; }
    return [nr, nc];
  }
  const u = s.used();
  while (inb(nr, nc) && !has(nr, nc)) {
    if ((dr > 0 && nr > u.r) || (dc > 0 && nc > u.c)) return [dr > 0 ? Math.min(maxR, Math.max(u.r, r)) : nr, dc > 0 ? Math.min(maxC, Math.max(u.c, c)) : nc];
    nr += dr; nc += dc;
  }
  return inb(nr, nc) ? [nr, nc] : [Math.max(0, Math.min(nr - dr, maxR)), Math.max(0, Math.min(nc - dc, maxC))];
}

function currentRegion(r, c) {
  const s = sheet();
  const has = (rr, cc) => rr >= 0 && cc >= 0 && wb().value(s, rr, cc) !== null;
  let r1 = r, r2 = r, c1 = c, c2 = c;
  let grew = true;
  while (grew) {
    grew = false;
    for (let cc = c1 - 1; cc <= c2 + 1; cc++) {
      if (has(r1 - 1, cc)) { r1--; grew = true; break; }
      if (has(r2 + 1, cc)) { r2++; grew = true; break; }
    }
    for (let rr = r1 - 1; rr <= r2 + 1; rr++) {
      if (has(rr, c1 - 1)) { c1--; grew = true; break; }
      if (has(rr, c2 + 1)) { c2++; grew = true; break; }
    }
  }
  return { r1, c1, r2, c2 };
}

const KEYMAP = {
  'Ctrl+Z': () => undo(), 'Ctrl+Y': () => redo(), 'Ctrl+Shift+Z': () => redo(),
  'Ctrl+S': () => save(), 'Ctrl+Shift+S': () => save(S.book, true), 'Ctrl+O': openDialog, 'Ctrl+N': () => newBook(), 'Ctrl+P': () => exportAs('pdf'),
  'Ctrl+B': () => toggleStyle('b'), 'Ctrl+I': () => toggleStyle('i'), 'Ctrl+U': () => toggleStyle('u'), 'Ctrl+5': () => toggleStyle('st'),
  'Ctrl+F': () => findDialog(), 'Ctrl+H': () => findDialog(true), 'Ctrl+K': () => noteDialog(),
  'Ctrl+D': () => fillDir('down'), 'Ctrl+R': () => fillDir('right'),
  'Ctrl+;': () => setCellInput(grid.sel.r, grid.sel.c, fmtDate(Math.floor(Date.now() / 864e5) + 25569 - new Date().getTimezoneOffset() / 1440, 'dd/mm/aaaa')),
  'Ctrl+Shift+L': () => toggleFilter(), 'Alt+=': () => autoSum('SUM'), 'Ctrl+Shift+4': () => setFormat('currency'), 'Ctrl+Shift+5': () => setFormat('percent'),
  'Ctrl+Shift+1': () => setFormat('number'), 'Ctrl+Shift+3': () => setFormat('date'), 'Ctrl+Shift+V': () => pasteFromMenu(true),
  'Ctrl+ ': () => grid.setRange({ r1: 0, r2: sheet().rows.count - 1, c1: grid.sel.c1, c2: grid.sel.c2 }, { r: grid.sel.r, c: grid.sel.c }),
  'Shift+ ': () => grid.setRange({ c1: 0, c2: sheet().cols.count - 1, r1: grid.sel.r1, r2: grid.sel.r2 }, { r: grid.sel.r, c: grid.sel.c }),
  'Ctrl+A': () => { const rg = currentRegion(grid.sel.r, grid.sel.c); const same = rg.r1 === grid.sel.r1 && rg.r2 === grid.sel.r2 && rg.c1 === grid.sel.c1 && rg.c2 === grid.sel.c2; grid.setRange(same || (rg.r1 === rg.r2 && rg.c1 === rg.c2) ? { r1: 0, c1: 0, r2: sheet().rows.count - 1, c2: sheet().cols.count - 1 } : rg, { r: grid.sel.r, c: grid.sel.c }); },
  'Ctrl+PageDown': () => switchSheet(1), 'Ctrl+PageUp': () => switchSheet(-1),
  'Ctrl+=': () => grid.setZoom(grid.zoom * 1.1), 'Ctrl+-': () => grid.setZoom(grid.zoom / 1.1), 'Ctrl+0': () => grid.setZoom(1),
  'Ctrl+1': () => formatCellsDialog(), 'Ctrl+Shift+=': () => insertCells(fullRowsSel() ? 'row' : 'col', 'before'),
  F1: () => helpDialog(), F4: () => {}, 'Shift+F3': () => insertFunctionDialog(), 'Alt+ArrowDown': () => openValidationList(),
  F2: () => editor.begin({ keepValue: true }),
};

const fullRowsSel = () => grid.sel.c1 === 0 && grid.sel.c2 >= sheet().cols.count - 1;
const fullColsSel = () => grid.sel.r1 === 0 && grid.sel.r2 >= sheet().rows.count - 1;

function onSinkKey(e) {
  if (isDialogOpen()) return;
  const combo = comboOf(e);
  if (combo === 'Ctrl+-' && (fullRowsSel() || fullColsSel())) { e.preventDefault(); deleteCells(fullRowsSel() ? 'row' : 'col'); return; }
  const fn = KEYMAP[combo];
  if (fn) { e.preventDefault(); fn(); return; }
  const sel = grid.sel;
  const mod = e.ctrlKey || e.metaKey;
  const nav = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
  if (nav) {
    e.preventDefault();
    const from = e.shiftKey ? grid.selEnd || { r: sel.r, c: sel.c } : { r: sel.r, c: sel.c };
    let [r, c] = mod ? dataEdge(from.r, from.c, nav[0], nav[1]) : [from.r + nav[0], from.c + nav[1]];
    // Pula células mescladas
    if (!mod && !e.shiftKey) {
      const m = sheet().mergeAt(sel.r, sel.c);
      if (m) { if (nav[0] > 0) r = m.r2 + 1; if (nav[1] > 0) c = m.c2 + 1; }
    }
    // Pula linhas/colunas ocultas
    while (r > 0 && r < sheet().rows.count - 1 && sheet().rows.hidden.has(r)) r += nav[0] || 1;
    while (c > 0 && c < sheet().cols.count - 1 && sheet().cols.hidden.has(c)) c += nav[1] || 1;
    grid.select(r, c, e.shiftKey);
    grid.scrollIntoView(r, c);
    return;
  }
  switch (e.key) {
    case 'Enter': e.preventDefault(); moveAfterEdit(e.shiftKey ? [-1, 0] : [1, 0]); return;
    case 'Tab': e.preventDefault(); moveAfterEdit(e.shiftKey ? [0, -1] : [0, 1]); return;
    case 'Delete': e.preventDefault(); clearRange('contents'); return;
    case 'Backspace': e.preventDefault(); clearRange('contents'); editor.begin({ text: '' }); return;
    case 'Escape': grid.copyRange = null; S.clipboard = S.clipboard?.cut ? null : S.clipboard; S.painter = null; document.body.classList.remove('painter'); grid.invalidate(); return;
    case 'Home': e.preventDefault(); grid.select(mod ? 0 : sel.r, 0, e.shiftKey); grid.scrollIntoView(grid.sel.r, 0); return;
    case 'End': if (mod) { e.preventDefault(); const u = sheet().used(); grid.select(Math.max(0, u.r), Math.max(0, u.c), e.shiftKey); grid.scrollIntoView(grid.sel.r, grid.sel.c); } return;
    case 'PageDown': case 'PageUp': {
      e.preventDefault();
      const rows = Math.max(1, Math.floor((grid.h - CH_H) / (24 * grid.zoom)) - 1);
      const r = clamp(sel.r + (e.key === 'PageDown' ? rows : -rows), 0, sheet().rows.count - 1);
      grid.select(r, sel.c, e.shiftKey);
      grid.scrollIntoView(r, sel.c);
      return;
    }
    default:
  }
}

function fillDir(dir) {
  const sel = { ...grid.sel };
  if (dir === 'down') {
    if (sel.r1 === sel.r2) { if (sel.r1 === 0) return; fill({ ...sel, r1: sel.r1 - 1, r2: sel.r1 - 1 }, sel); return; }
    fill({ ...sel, r2: sel.r1 }, sel);
  } else {
    if (sel.c1 === sel.c2) { if (sel.c1 === 0) return; fill({ ...sel, c1: sel.c1 - 1, c2: sel.c1 - 1 }, sel); return; }
    fill({ ...sel, c2: sel.c1 }, sel);
  }
}

function undo() {
  if (editor.active) { editor.cancel(); return; }
  const l = wb().undo();
  if (!l) toast('Nada para desfazer.', 1200);
  restoreAfterHistory();
}

function redo() {
  const l = wb().redo();
  if (!l) toast('Nada para refazer.', 1200);
  restoreAfterHistory();
}

function restoreAfterHistory() {
  grid.updateSpacer();
  grid.invalidate();
  renderAll();
}

// ---------------------------------------------------------------------------
// Atualização de interface

function refresh() {
  grid.updateSpacer();
  grid.invalidate();
  charts.sync();
  renderStatus();
  renderFormulaBar();
  renderTabs();
  refreshToolbar();
}

function renderAll() {
  renderTitle();
  renderBookTabs();
  renderTabs();
  refresh();
}

function selectionChanged() {
  if (S.painter) applyPainter();
  renderFormulaBar();
  renderStatus();
  refreshToolbar();
  positionValidationButton();
}

function renderFormulaBar() {
  const sel = grid.sel;
  const s = sheet();
  const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
  if (document.activeElement !== nameBox) nameBox.value = multi ? `${colName(sel.c1)}${sel.r1 + 1}:${colName(sel.c2)}${sel.r2 + 1}` : `${colName(sel.c)}${sel.r + 1}`;
  if (!editor.active) {
    const cell = s.get(sel.r, sel.c);
    const anchor = wb().spillAnchor(s, sel.r, sel.c);
    if (anchor && !cell) {
      const a = s.get(anchor[0], anchor[1]);
      bar.value = a?.f ? `=${a.f}` : '';
      bar.classList.add('ghost');
    } else {
      bar.value = editText(cell);
      bar.classList.remove('ghost');
    }
  }
}

function renderStatus() {
  const sel = grid.sel;
  const s = sheet();
  const box = $('#sel-stats');
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
    const cell = s.get(sel.r, sel.c);
    const v = wb().value(s, sel.r, sel.c);
    box.textContent = isErr(v) ? `${v.code}${v.msg ? ` — ${v.msg}` : ''}` : cell?.parseError ? `Erro de sintaxe: ${cell.parseError}` : '';
    return;
  }
  let n = 0, cnt = 0, sum = 0, min = Infinity, max = -Infinity;
  const u = s.used();
  const r2 = Math.min(sel.r2, u.r), c2 = Math.min(sel.c2, u.c);
  for (let r = sel.r1; r <= r2; r++) for (let c = sel.c1; c <= c2; c++) {
    if (s.rows.hidden.has(r)) continue;
    const v = wb().value(s, r, c);
    if (v === null || v === '') continue;
    cnt++;
    if (typeof v === 'number') { n++; sum += v; min = Math.min(min, v); max = Math.max(max, v); }
  }
  const f = (x) => new Intl.NumberFormat('pt-BR', { maximumFractionDigits: 4 }).format(x);
  box.replaceChildren(...(n ? [
    el('span', {}, `Soma: ${f(sum)}`), el('span', {}, `Média: ${f(sum / n)}`), el('span', {}, `Mín: ${f(min)}`), el('span', {}, `Máx: ${f(max)}`), el('span', {}, `Contagem: ${cnt}`),
  ] : cnt ? [el('span', {}, `Contagem: ${cnt}`)] : []));
}

function renderTitle() {
  const b = S.book;
  const t = $('#doc-title');
  if (document.activeElement !== t) t.value = b.title;
  $('#save-state').textContent = b.dirty ? 'Alterações não salvas' : b.filePath ? `Salvo · ${fileName(b.filePath)}` : 'Não salvo ainda';
  setTitle(`${b.title}${b.dirty ? ' •' : ''} — Lumina Planilhas`);
  renderBookTabs();
}

function renderBookTabs() {
  const box = $('#book-tabs');
  box.classList.toggle('hidden', S.books.length < 2);
  box.replaceChildren(...S.books.map((b) => el('div', { class: `doc-tab${b === S.book ? ' on' : ''}`, onclick: () => b !== S.book && activate(b), title: b.filePath || b.title },
    el('span', { class: 'nm' }, `${b.title}${b.dirty ? ' •' : ''}`),
    el('button', { class: 'x', html: icon('x', 12), title: 'Fechar', onclick: (e) => { e.stopPropagation(); closeBook(b); } }))));
}

// ---------------------------------------------------------------------------
// Abas de planilha

function renderTabs() {
  const box = $('#sheet-tabs');
  const w = wb();
  box.replaceChildren(
    el('button', { class: 'icon-btn', title: 'Nova planilha', html: icon('plus', 16), onclick: () => { saveView(); w.addSheet(); restoreView(); renderTabs(); } }),
    ...w.sheets.map((s, i) => {
      const tab = el('div', {
        class: `sheet-tab${i === w.active ? ' on' : ''}`, draggable: 'true', title: s.name,
        style: s.tabColor ? { boxShadow: `inset 0 -3px 0 ${s.tabColor}` } : null,
        onclick: () => selectSheet(i),
        ondblclick: () => renameSheet(s),
        oncontextmenu: (e) => { e.preventDefault(); sheetMenu(s, i, e); },
        ondragstart: (e) => e.dataTransfer.setData('text/x-sheet', String(i)),
        ondragover: (e) => e.preventDefault(),
        ondrop: (e) => { e.preventDefault(); const from = +e.dataTransfer.getData('text/x-sheet'); if (!Number.isNaN(from) && from !== i) { w.moveSheet(from, i); renderTabs(); } },
      }, s.name);
      return tab;
    }));
}

function selectSheet(i) {
  const w = wb();
  if (i === w.active) return;
  const pointing = editor.active && editor.isFormula();
  if (editor.active && !pointing) editor.commit();
  saveView();
  w.active = i;
  restoreView();
  charts.sync();
  renderTabs();
  renderFormulaBar();
  if (pointing) editor.focus();
}

function switchSheet(d) {
  const w = wb();
  selectSheet((w.active + d + w.sheets.length) % w.sheets.length);
}

async function renameSheet(s) {
  const name = await prompt('Renomear planilha', 'Nome', s.name);
  if (name === null) return;
  if (!wb().renameSheet(s, name)) toast('Nome inválido ou já existente.');
  renderTabs();
}

function sheetMenu(s, i, e) {
  const w = wb();
  contextMenu([
    { label: 'Renomear…', action: () => renameSheet(s) },
    { label: 'Duplicar', icon: 'copy', action: () => { saveView(); w.duplicateSheet(s); restoreView(); renderTabs(); } },
    { label: 'Excluir', icon: 'trash', disabled: w.sheets.length < 2, action: async () => {
      if (await dialog({ title: 'Excluir planilha', body: el('p', {}, `Excluir "${s.name}"? Você pode desfazer com Ctrl+Z.`), okText: 'Excluir' })) { w.deleteSheet(s); restoreView(); renderTabs(); }
    } },
    '-',
    { label: 'Cor da guia', icon: 'palette', sub: [null, '#e03131', '#f08c00', '#2f9e44', '#1971c2', '#7048e8', '#c2255c', '#495057'].map((c) => ({
      label: c ? '■' : 'Sem cor', action: () => w.transact('Cor da guia', () => { w.snapSheet(s); s.tabColor = c; }),
    })) },
    { label: 'Mostrar linhas de grade', checked: s.showGrid, action: () => w.transact('Linhas de grade', () => { w.snapSheet(s); s.showGrid = !s.showGrid; }) },
    '-',
    { label: 'Mover para a esquerda', disabled: i === 0, action: () => w.moveSheet(i, i - 1) },
    { label: 'Mover para a direita', disabled: i === w.sheets.length - 1, action: () => w.moveSheet(i, i + 1) },
  ], e.clientX, e.clientY);
}

// ---------------------------------------------------------------------------
// Menus de contexto

function cellMenu(e, h) {
  if (h.area === 'colhead' || h.area === 'rowhead') {
    const isCol = h.area === 'colhead';
    contextMenu([
      { label: 'Recortar', key: 'Ctrl+X', action: () => document.execCommand('cut') },
      { label: 'Copiar', key: 'Ctrl+C', action: () => document.execCommand('copy') },
      { label: 'Colar', key: 'Ctrl+V', action: () => pasteFromMenu() },
      '-',
      { label: isCol ? 'Inserir coluna à esquerda' : 'Inserir linha acima', icon: isCol ? 'colInsert' : 'rowInsert', action: () => insertCells(isCol ? 'col' : 'row', 'before') },
      { label: isCol ? 'Inserir coluna à direita' : 'Inserir linha abaixo', action: () => insertCells(isCol ? 'col' : 'row', 'after') },
      { label: isCol ? 'Excluir coluna(s)' : 'Excluir linha(s)', icon: 'trash', action: () => deleteCells(isCol ? 'col' : 'row') },
      '-',
      { label: 'Ocultar', action: () => hideAxis(isCol ? 'col' : 'row', true) },
      { label: 'Reexibir', action: () => hideAxis(isCol ? 'col' : 'row', false) },
      isCol ? { label: 'Ajustar largura', action: () => autofit('col', h.c) } : { label: 'Ajustar altura', action: () => autofit('row', h.r) },
      isCol ? { label: 'Classificar A → Z', icon: 'sortAsc', action: () => sortSel(false) } : null,
      isCol ? { label: 'Classificar Z → A', icon: 'sortDesc', action: () => sortSel(true) } : null,
    ], e.clientX, e.clientY);
    return;
  }
  contextMenu([
    { label: 'Recortar', key: 'Ctrl+X', action: () => { focusGrid(); document.execCommand('cut'); } },
    { label: 'Copiar', key: 'Ctrl+C', action: () => { focusGrid(); document.execCommand('copy'); } },
    { label: 'Colar', key: 'Ctrl+V', action: () => pasteFromMenu() },
    { label: 'Colar somente valores', key: 'Ctrl+Shift+V', action: () => pasteFromMenu(true) },
    '-',
    { label: 'Inserir linha acima', icon: 'rowInsert', action: () => insertCells('row', 'before') },
    { label: 'Inserir coluna à esquerda', icon: 'colInsert', action: () => insertCells('col', 'before') },
    { label: 'Excluir linha(s)', action: () => deleteCells('row') },
    { label: 'Excluir coluna(s)', action: () => deleteCells('col') },
    '-',
    { label: 'Limpar conteúdo', key: 'Delete', action: () => clearRange('contents') },
    { label: 'Limpar formatação', action: () => clearRange('formats') },
    '-',
    { label: 'Classificar A → Z', icon: 'sortAsc', action: () => sortSel(false) },
    { label: 'Classificar Z → A', icon: 'sortDesc', action: () => sortSel(true) },
    { label: 'Filtrar', icon: 'filter', action: toggleFilter },
    '-',
    { label: 'Inserir gráfico', icon: 'chart', action: () => insertChart() },
    { label: 'Formatação condicional…', icon: 'palette', action: cfDialog },
    { label: 'Validação de dados…', icon: 'check', action: validationDialog },
    { label: 'Nota…', icon: 'notes', key: 'Ctrl+K', action: noteDialog },
    { label: 'Formatar células…', icon: 'settings', key: 'Ctrl+1', action: formatCellsDialog },
  ], e.clientX, e.clientY);
}

// ---------------------------------------------------------------------------
// Pincel de formatação, nota e clique em botões de filtro

function startPainter() {
  const st = sheet().get(grid.sel.r, grid.sel.c)?.s;
  S.painter = { style: st ? clone(st) : {} };
  document.body.classList.add('painter');
  toast('Selecione as células que receberão a formatação.');
}

function applyPainter() {
  const p = S.painter;
  if (!p || grid.dragging) return;
  S.painter = null;
  document.body.classList.remove('painter');
  const s = sheet();
  wb().transact('Pincel de formatação', () => {
    for (const [r, c] of selectionCells()) {
      const old = s.get(r, c);
      wb().setCell(s, r, c, { ...(old ? stripCell(old) : { v: null }), s: clone(p.style) });
    }
  });
}

async function noteDialog() {
  const { r, c } = grid.sel;
  const s = sheet();
  const cell = s.get(r, c);
  const ta = el('textarea', { rows: 4, style: { width: '100%' } });
  ta.value = cell?.note || '';
  const ok = await dialog({ title: `Nota em ${colName(c)}${r + 1}`, body: ta, okText: 'Salvar', extraButtons: cell?.note ? [{ label: 'Remover', onClick: (close) => { ta.value = ''; close(true); } }] : [] });
  if (!ok) return;
  wb().transact('Nota', () => wb().setCell(s, r, c, { ...(cell ? stripCell(cell) : { v: null }), note: ta.value.trim() || undefined }));
}

function cellClicked(r, c, e) {
  const s = sheet();
  const f = s.filter;
  if (f && r === f.r1 && c >= f.c1 && c <= f.c2) {
    const rect = grid.cellRect(r, c);
    const [x] = grid.local(e);
    if (x >= rect.x + rect.w - 22) filterPopover(c, e);
  }
}

// ---------------------------------------------------------------------------
// Filtro e classificação

function toggleFilter() {
  const s = sheet();
  if (s.filter) {
    wb().transact('Remover filtro', () => {
      wb().snapSheet(s);
      for (let r = s.filter.r1 + 1; r <= s.filter.r2; r++) s.rows.setHidden(r, false);
      s.filter = null;
    });
    return;
  }
  const rg = grid.sel.r1 === grid.sel.r2 && grid.sel.c1 === grid.sel.c2 ? currentRegion(grid.sel.r, grid.sel.c) : clipRange();
  if (rg.r1 === rg.r2) { toast('Selecione uma tabela com cabeçalho.'); return; }
  wb().transact('Filtro', () => { wb().snapSheet(s); s.filter = { ...rg, criteria: {} }; });
}

function filterPopover(c, e) {
  const s = sheet();
  const f = s.filter;
  const region = currentRegion(f.r1, f.c1);
  const r2 = Math.max(f.r2, region.r2);
  const values = new Map();
  for (let r = f.r1 + 1; r <= r2; r++) {
    const cell = s.get(r, c);
    const v = wb().value(s, r, c);
    const t = v === null ? '(Vazias)' : display(v, cell?.s || {});
    if (!values.has(t)) values.set(t, v);
  }
  const hidden = new Set(f.criteria?.[c] || []);
  const sorted = [...values.entries()].sort((a, b) => compare(a[1], b[1]));
  const search = el('input', { type: 'search', placeholder: 'Pesquisar' });
  const list = el('div', { class: 'filter-list' });
  const checks = new Map();
  const all = el('input', { type: 'checkbox', checked: hidden.size === 0 });
  const draw = () => {
    const q = search.value.toLowerCase();
    list.replaceChildren(...sorted.filter(([t]) => !q || t.toLowerCase().includes(q)).map(([t]) => {
      let cb = checks.get(t);
      if (!cb) { cb = el('input', { type: 'checkbox', checked: !hidden.has(t) }); checks.set(t, cb); }
      return el('label', { class: 'check' }, cb, t);
    }));
  };
  all.addEventListener('change', () => { for (const cb of checks.values()) cb.checked = all.checked; });
  search.addEventListener('input', draw);
  draw();
  const apply = (sortDesc = null) => {
    closePopovers();
    const nextHidden = [...checks.entries()].filter(([, cb]) => !cb.checked).map(([t]) => t);
    wb().transact('Filtrar', () => {
      wb().snapSheet(s);
      f.r2 = r2;
      f.criteria = { ...(f.criteria || {}), [c]: nextHidden };
      if (!nextHidden.length) delete f.criteria[c];
      if (sortDesc !== null) wb().sortRange(s, { r1: f.r1 + 1, c1: f.c1, r2, c2: f.c2 }, [{ col: c, desc: sortDesc }]);
      for (let r = f.r1 + 1; r <= r2; r++) {
        let hide = false;
        for (const [col, list2] of Object.entries(f.criteria)) {
          const v = wb().value(s, r, +col);
          const t = v === null ? '(Vazias)' : display(v, s.get(r, +col)?.s || {});
          if (list2.includes(t)) { hide = true; break; }
        }
        s.rows.setHidden(r, hide);
      }
    });
    grid.updateSpacer();
  };
  popover(el('div', { class: 'filter-pop' },
    el('button', { class: 'btn small block', onclick: () => apply(false), html: `${icon('sortAsc', 14)} Classificar A → Z` }),
    el('button', { class: 'btn small block', onclick: () => apply(true), html: `${icon('sortDesc', 14)} Classificar Z → A` }),
    search, el('label', { class: 'check' }, all, el('b', {}, '(Selecionar tudo)')), list,
    el('div', { class: 'row' }, el('span', { class: 'spacer' }), el('button', { class: 'btn small', onclick: () => closePopovers() }, 'Cancelar'), el('button', { class: 'btn small primary', onclick: () => apply() }, 'OK'))),
  { x: e.clientX - 10, y: e.clientY + 6 });
}

function sortSel(desc) {
  const s = sheet();
  const sel = grid.sel;
  let rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sel.r, sel.c) : clipRange();
  const fullCol = fullColsSel();
  if (fullCol) rg = currentRegion(0, sel.c1);
  // Detecta cabeçalho: primeira linha de texto sobre coluna numérica, ou estilo negrito
  const first = wb().value(s, rg.r1, sel.c);
  const second = wb().value(s, rg.r1 + 1, sel.c);
  const header = rg.r2 > rg.r1 && ((typeof first === 'string' && typeof second !== 'string') || s.get(rg.r1, sel.c)?.s?.b || s.filter?.r1 === rg.r1);
  if (header) rg = { ...rg, r1: rg.r1 + 1 };
  wb().sortRange(s, rg, [{ col: sel.c, desc }]);
}

async function sortDialog() {
  const sel = grid.sel;
  const rg0 = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sel.r, sel.c) : clipRange();
  const s = sheet();
  const hasHeader = el('input', { type: 'checkbox', checked: true });
  const colName2 = (c) => (hasHeader.checked ? `${colName(c)} — ${display(wb().value(s, rg0.r1, c) ?? '', {})}` : `Coluna ${colName(c)}`);
  const levels = el('div', { style: { display: 'flex', flexDirection: 'column', gap: '8px' } });
  const addLevel = (col = sel.c, desc = false) => {
    const cs = select(range(rg0.c1, rg0.c2).map((c) => [c, colName2(c)]), col);
    const ord = select([['asc', 'A → Z / menor → maior'], ['desc', 'Z → A / maior → menor']], desc ? 'desc' : 'asc');
    const row = el('div', { class: 'row' }, el('span', { class: 'muted' }, levels.children.length ? 'Depois por' : 'Classificar por'), cs, ord,
      el('button', { class: 'icon-btn', html: icon('x', 14), onclick: () => row.remove() }));
    levels.append(row);
  };
  addLevel();
  const ok = await dialog({
    title: 'Classificar intervalo', width: 560,
    body: [el('p', { class: 'muted' }, `Intervalo: ${rangeText(rg0)}`), el('label', { class: 'check' }, hasHeader, 'Meus dados têm cabeçalho'), levels,
      el('button', { class: 'btn small', onclick: () => addLevel(), html: `${icon('plus', 14)} Adicionar nível` })],
  });
  if (!ok) return;
  const keys = [...levels.children].map((row) => { const [cs, ord] = row.querySelectorAll('select'); return { col: +cs.value, desc: ord.value === 'desc' }; });
  wb().sortRange(s, hasHeader.checked ? { ...rg0, r1: rg0.r1 + 1 } : rg0, keys);
}

async function removeDuplicates() {
  const sel = grid.sel;
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sel.r, sel.c) : clipRange();
  const s = sheet();
  const hasHeader = el('input', { type: 'checkbox', checked: true });
  const cols = range(rg.c1, rg.c2).map((c) => [c, el('input', { type: 'checkbox', checked: true })]);
  const ok = await dialog({
    title: 'Remover duplicatas', width: 420,
    body: [el('p', { class: 'muted' }, `Intervalo: ${rangeText(rg)}. Considerar as colunas:`), el('label', { class: 'check' }, hasHeader, 'Meus dados têm cabeçalho'),
      el('div', { class: 'grid2' }, ...cols.map(([c, cb]) => el('label', { class: 'check' }, cb, `${colName(c)} ${hasHeader.checked ? `— ${display(wb().value(s, rg.r1, c) ?? '', {})}` : ''}`)))],
  });
  if (!ok) return;
  const keyCols = cols.filter(([, cb]) => cb.checked).map(([c]) => c);
  const start = hasHeader.checked ? rg.r1 + 1 : rg.r1;
  const seen = new Set();
  const keep = [];
  for (let r = start; r <= rg.r2; r++) {
    const k = JSON.stringify(keyCols.map((c) => { const v = wb().value(s, r, c); return typeof v === 'string' ? v.toLowerCase() : v; }));
    if (seen.has(k)) continue;
    seen.add(k);
    keep.push(r);
  }
  const removed = rg.r2 - start + 1 - keep.length;
  wb().transact('Remover duplicatas', () => {
    const rows = keep.map((r) => range(rg.c1, rg.c2).map((c) => clone(stripCell(s.get(r, c)))));
    for (let r = start; r <= rg.r2; r++) for (let c = rg.c1; c <= rg.c2; c++) wb().setCell(s, r, c, null);
    rows.forEach((row, i) => row.forEach((cell, j) => {
      if (cell?.f) cell.f = shiftFormula(cell.f, start + i - keep[i], 0);
      wb().setCell(s, start + i, rg.c1 + j, cell);
    }));
  });
  toast(removed ? `${removed} linha(s) duplicada(s) removida(s).` : 'Nenhuma duplicata encontrada.');
}

// ---------------------------------------------------------------------------
// Soma automática

function autoSum(fn = 'SUM') {
  const s = sheet();
  const sel = grid.sel;
  const pt = { SUM: 'SOMA', AVERAGE: 'MÉDIA', COUNT: 'CONT.NÚM', MAX: 'MÁXIMO', MIN: 'MÍNIMO' }[fn];
  if (sel.r1 === sel.r2 && sel.c1 === sel.c2) {
    // Procura números acima (ou à esquerda)
    let r = sel.r - 1;
    while (r >= 0 && typeof wb().value(s, r, sel.c) !== 'number') r--;
    let r0 = r;
    while (r0 - 1 >= 0 && typeof wb().value(s, r0 - 1, sel.c) === 'number') r0--;
    let ref;
    if (r >= 0) ref = `${colName(sel.c)}${r0 + 1}:${colName(sel.c)}${r + 1}`;
    else {
      let c = sel.c - 1;
      while (c >= 0 && typeof wb().value(s, sel.r, c) !== 'number') c--;
      let c0 = c;
      while (c0 - 1 >= 0 && typeof wb().value(s, sel.r, c0 - 1) === 'number') c0--;
      ref = c >= 0 ? `${colName(c0)}${sel.r + 1}:${colName(c)}${sel.r + 1}` : '';
    }
    editor.begin({ text: `=${pt}(${ref})` });
    return;
  }
  // Seleção de números: total abaixo de cada coluna
  wb().transact('Soma automática', () => {
    for (let c = sel.c1; c <= sel.c2; c++) {
      const r = sel.r2 + 1;
      wb().setCell(s, r, c, { v: null, f: `${pt}(${colName(c)}${sel.r1 + 1}:${colName(c)}${sel.r2 + 1})`, s: s.get(sel.r2, c)?.s });
    }
  });
}

// ---------------------------------------------------------------------------
// Formatação condicional

const CF_TYPES = [
  ['scale3', 'Escala de 3 cores (vermelho → amarelo → verde)'], ['scale2', 'Escala de 2 cores'], ['bar', 'Barras de dados'],
  ['gt', 'Maior que…'], ['lt', 'Menor que…'], ['between', 'Entre…'], ['eq', 'Igual a…'], ['text', 'Texto que contém…'],
  ['dup', 'Valores duplicados'], ['unique', 'Valores únicos'], ['top', 'Primeiros N (maiores)'], ['bottom', 'Últimos N (menores)'],
  ['aboveAvg', 'Acima da média'], ['belowAvg', 'Abaixo da média'], ['blank', 'Células vazias'], ['error', 'Células com erro'],
];
const CF_STYLES = [
  ['Vermelho claro', { bg: '#ffc9c9', color: '#9c0006' }], ['Amarelo', { bg: '#ffec99', color: '#7a5d00' }], ['Verde', { bg: '#c6efce', color: '#006100' }],
  ['Azul', { bg: '#d0ebff', color: '#1c4587' }], ['Texto vermelho', { color: '#c92a2a' }], ['Negrito', { bold: true }],
];

async function cfDialog() {
  const s = sheet();
  const sel = clipRange();
  const type = select(CF_TYPES, 'scale3');
  const a = el('input', { type: 'text', placeholder: 'Valor' });
  const b = el('input', { type: 'text', placeholder: 'e' });
  const styleSel = select(CF_STYLES.map(([l], i) => [i, l]), 0);
  const barColor = el('input', { type: 'color', value: '#5b8cff' });
  const params = el('div', { class: 'row' });
  const sync = () => {
    const t = type.value;
    params.replaceChildren();
    if (['gt', 'lt', 'eq', 'text', 'top', 'bottom'].includes(t)) params.append(a);
    if (t === 'between') params.append(a, el('span', {}, 'e'), b);
    if (t === 'bar') params.append(el('span', { class: 'muted' }, 'Cor'), barColor);
    else if (!t.startsWith('scale')) params.append(el('span', { class: 'muted' }, 'Formato'), styleSel);
    a.placeholder = t === 'top' || t === 'bottom' ? 'N (ex.: 10)' : 'Valor';
  };
  type.addEventListener('change', sync);
  sync();
  const existing = el('div', { class: 'cf-list' });
  const drawExisting = () => existing.replaceChildren(...(s.cf.length ? s.cf.map((rule) => el('div', { class: 'cf-item' },
    el('span', {}, `${CF_TYPES.find(([k]) => k === rule.type)?.[1] || rule.type} — ${rangeText(rule.range)}`),
    el('button', { class: 'icon-btn', html: icon('trash', 14), title: 'Remover regra', onclick: () => { wb().transact('Remover regra', () => { wb().snapSheet(s); s.cf = s.cf.filter((x) => x !== rule); }); drawExisting(); } })))
    : [el('p', { class: 'muted' }, 'Nenhuma regra nesta planilha.')]));
  drawExisting();
  const ok = await dialog({
    title: 'Formatação condicional', width: 520, okText: 'Adicionar regra',
    body: [el('p', { class: 'muted' }, `Aplicar em ${rangeText(sel)}`), field('Regra', type), params, el('div', { class: 'section-label' }, 'Regras existentes'), existing],
  });
  if (!ok) return;
  const num = (x) => { const p = parseInput(x.value); return typeof p.v === 'number' ? p.v : x.value; };
  const rule = { id: `cf${Date.now()}`, type: type.value, range: sel, params: {} };
  if (['gt', 'lt', 'eq', 'text'].includes(rule.type)) rule.params.a = num(a);
  if (rule.type === 'between') { rule.params.a = num(a); rule.params.b = num(b); }
  if (rule.type === 'top' || rule.type === 'bottom') rule.params.n = Math.max(1, parseInt(a.value, 10) || 10);
  if (rule.type === 'bar') rule.params.color = barColor.value;
  if (!rule.type.startsWith('scale') && rule.type !== 'bar') rule.style = CF_STYLES[+styleSel.value][1];
  wb().transact('Formatação condicional', () => { wb().snapSheet(s); s.cf.push(rule); });
}

// ---------------------------------------------------------------------------
// Validação de dados (lista suspensa)

let valBtn = null;
function positionValidationButton() {
  const s = sheet();
  const { r, c } = grid.sel;
  const v = s.validations.find((x) => x.type === 'list' && inRange(x.range, r, c));
  if (!v || editor.active) { valBtn?.remove(); valBtn = null; return; }
  if (!valBtn) {
    valBtn = el('button', { class: 'val-btn', title: 'Escolher da lista (Alt+↓)', html: icon('chevronDown', 14), onmousedown: (e) => { e.preventDefault(); e.stopPropagation(); openValidationList(); } });
    $('#grid').append(valBtn);
  }
  const rect = grid.cellRect(r, c);
  valBtn.style.left = `${rect.x + rect.w + 1}px`;
  valBtn.style.top = `${rect.y}px`;
  valBtn.style.height = `${rect.h}px`;
}

function openValidationList() {
  const s = sheet();
  const { r, c } = grid.sel;
  const v = s.validations.find((x) => x.type === 'list' && inRange(x.range, r, c));
  if (!v || !valBtn) return;
  dropdown(valBtn, v.items.map((it) => ({ label: String(it), action: () => { setCellInput(r, c, String(it)); focusGrid(); } })));
}

async function validationDialog() {
  const s = sheet();
  const sel = clipRange();
  const existing = s.validations.find((x) => inRange(x.range, sel.r1, sel.c1));
  const items = el('input', { type: 'text', value: existing?.items.join('; ') || '', placeholder: 'Ex.: Sim; Não; Talvez   ou   =A1:A10' });
  const ok = await dialog({
    title: 'Validação de dados', width: 480,
    body: [el('p', { class: 'muted' }, `Lista suspensa para ${rangeText(sel)}.`), field('Itens (separados por ";") ou intervalo', items)],
    extraButtons: existing ? [{ label: 'Remover', onClick: (close) => { wb().transact('Remover validação', () => { wb().snapSheet(s); s.validations = s.validations.filter((x) => x !== existing); }); close(false); } }] : [],
  });
  if (!ok) return;
  let list;
  const t = items.value.trim();
  if (t.startsWith('=')) {
    try {
      const tok = tokenize(t.slice(1))[0];
      const ref = tok.v;
      list = [];
      for (let r = ref.a.row; r <= (ref.b?.row ?? ref.a.row); r++) for (let c = ref.a.col; c <= (ref.b?.col ?? ref.a.col); c++) {
        const v = wb().value(ref.sheet ? wb().byName(ref.sheet) : s, r, c);
        if (v !== null && v !== '') list.push(display(v, {}));
      }
    } catch { toast('Intervalo inválido.'); return; }
  } else list = t.split(/[;\n]/).map((x) => x.trim()).filter(Boolean);
  if (!list.length) return;
  wb().transact('Validação de dados', () => {
    wb().snapSheet(s);
    s.validations = s.validations.filter((x) => x !== existing);
    s.validations.push({ range: sel, type: 'list', items: list });
  });
  positionValidationButton();
}

// ---------------------------------------------------------------------------
// Gráficos

function insertChart(type = 'column') {
  const s = sheet();
  const sel = grid.sel;
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sel.r, sel.c) : clipRange();
  if (rg.r1 === rg.r2 && rg.c1 === rg.c2) { toast('Selecione os dados do gráfico (com cabeçalhos).'); return; }
  const pos = defaultChartPlacement(grid, rg);
  const ch = { id: `ch${Date.now()}`, type, title: '', range: rg, seriesIn: 'cols', legend: 'bottom', sheetId: s.id, ...pos };
  wb().transact('Inserir gráfico', () => { wb().snapSheet(s); s.charts.push(ch); });
  charts.selected = ch.id;
}

function updateChart(id, data, label = 'Editar gráfico') {
  const s = sheet();
  wb().transact(label, () => {
    wb().snapSheet(s);
    const ch = s.charts.find((c) => c.id === id);
    if (ch) Object.assign(ch, data);
  });
}

function chartMenu(id, e) {
  const s = sheet();
  contextMenu([
    { label: 'Editar gráfico…', icon: 'chart', action: () => chartDialog(id) },
    { label: 'Tipo', sub: CHART_TYPES.map(([k, l]) => ({ label: l, action: () => updateChart(id, { type: k }) })) },
    { label: 'Exportar como imagem (PNG)', icon: 'download', action: async () => {
      const img = charts.image(id);
      if (img) await saveFile(Uint8Array.from(atob(img.split(',')[1]), (ch) => ch.charCodeAt(0)), { name: 'grafico.png', filters: [{ name: 'PNG', extensions: ['png'] }] });
    } },
    '-',
    { label: 'Excluir gráfico', icon: 'trash', action: () => wb().transact('Excluir gráfico', () => { wb().snapSheet(s); s.charts = s.charts.filter((c) => c.id !== id); }) },
  ], e.clientX, e.clientY);
}

async function chartDialog(id) {
  const s = sheet();
  const ch = s.charts.find((c) => c.id === id);
  if (!ch) return;
  const type = select(CHART_TYPES, ch.type);
  const title = el('input', { type: 'text', value: ch.title || '', placeholder: 'Título do gráfico' });
  const rng = el('input', { type: 'text', value: rangeText(ch.range) });
  const seriesIn = select([['cols', 'Séries em colunas'], ['rows', 'Séries em linhas']], ch.seriesIn);
  const legend = select([['bottom', 'Embaixo'], ['top', 'Em cima'], ['right', 'À direita'], ['none', 'Sem legenda']], ch.legend || 'bottom');
  const stacked = el('input', { type: 'checkbox', checked: !!ch.stacked });
  const ok = await dialog({
    title: 'Editar gráfico', width: 460,
    body: [el('div', { class: 'grid2' }, field('Tipo', type), field('Legenda', legend)), field('Título', title),
      el('div', { class: 'grid2' }, field('Intervalo de dados', rng), field('Organização', seriesIn)),
      el('label', { class: 'check' }, stacked, 'Empilhado (colunas, barras, área)')],
  });
  if (!ok) return;
  let range2 = ch.range;
  const m = /^([A-Za-z]+\d+)(?::([A-Za-z]+\d+))?$/.exec(rng.value.trim().replace(/\$/g, ''));
  if (m) {
    const a = parseAddr(m[1]), b = parseAddr(m[2] || m[1]);
    range2 = norm(a.r, a.c, b.r, b.c);
  }
  updateChart(id, { type: type.value, title: title.value, range: range2, seriesIn: seriesIn.value, legend: legend.value, stacked: stacked.checked, pick: m ? undefined : ch.pick });
}

// ---------------------------------------------------------------------------
// Tabela dinâmica (com fórmulas, atualiza sozinha)

async function pivotDialog() {
  const s = sheet();
  const sel = grid.sel;
  const rg = sel.r1 === sel.r2 && sel.c1 === sel.c2 ? currentRegion(sel.r, sel.c) : clipRange();
  if (rg.r2 <= rg.r1) { toast('Selecione uma tabela com cabeçalho e dados.'); return; }
  const fields = range(rg.c1, rg.c2).map((c) => [c, String(display(wb().value(s, rg.r1, c) ?? colName(c), {}))]);
  const rows = select(fields, rg.c1);
  const cols = select([['', '(nenhuma)'], ...fields], '');
  const numericCol = fields.find(([c]) => typeof wb().value(s, rg.r1 + 1, c) === 'number')?.[0] ?? rg.c2;
  const vals = select(fields, numericCol);
  const agg = select([['SUMIFS', 'Soma'], ['COUNTIFS', 'Contagem'], ['AVERAGEIFS', 'Média'], ['MAXIFS', 'Máximo'], ['MINIFS', 'Mínimo']], 'SUMIFS');
  const ok = await dialog({
    title: 'Tabela dinâmica', width: 480, okText: 'Criar',
    body: [el('p', { class: 'muted' }, `Fonte: ${s.name}!${rangeText(rg)}. A tabela usa fórmulas e se atualiza quando os dados mudam.`),
      el('div', { class: 'grid2' }, field('Linhas', rows), field('Colunas', cols), field('Valores', vals), field('Resumir por', agg))],
  });
  if (!ok) return;
  const rc = +rows.value, cc = cols.value === '' ? null : +cols.value, vc = +vals.value, fn = agg.value;
  const uniq = (c) => {
    const seen = new Map();
    for (let r = rg.r1 + 1; r <= rg.r2; r++) {
      const v = wb().value(s, r, c);
      if (v === null || v === '') continue;
      const k = typeof v === 'string' ? v.toLowerCase() : v;
      if (!seen.has(k)) seen.set(k, v);
    }
    return [...seen.values()].sort(compare);
  };
  const rowVals = uniq(rc);
  const colVals = cc === null ? [null] : uniq(cc);
  const src = (c) => `${quote(s.name)}!$${colName(c)}$${rg.r1 + 2}:$${colName(c)}$${rg.r2 + 1}`;
  const pt = { SUMIFS: 'SOMASES', COUNTIFS: 'CONT.SES', AVERAGEIFS: 'MÉDIASES', MAXIFS: 'MÁXIMOSES', MINIFS: 'MÍNIMOSES' }[fn];
  const w = wb();
  w.transact('Tabela dinâmica', () => {
    w.snapWorkbook();
    const ns = new (s.constructor)(w.uniqueName('Dinâmica'));
    w.sheets.splice(w.active + 1, 0, ns);
    w.active = w.active + 1;
    const head = { b: true, bg: '#2f9e6e', color: '#ffffff', ha: 'center' };
    const lit = (v) => (typeof v === 'number' ? String(v).replace('.', ',') : `"${String(v).replace(/"/g, '""')}"`);
    ns.set(0, 0, { v: `${fields.find(([c]) => c === rc)[1]}${cc !== null ? ` / ${fields.find(([c]) => c === cc)[1]}` : ''}`, s: head });
    colVals.forEach((cv, j) => ns.set(0, j + 1, { v: cv === null ? `${{ SUMIFS: 'Soma', COUNTIFS: 'Contagem', AVERAGEIFS: 'Média', MAXIFS: 'Máximo', MINIFS: 'Mínimo' }[fn]} de ${fields.find(([c]) => c === vc)[1]}` : cv, s: head }));
    const totalCol = colVals.length + 1;
    if (cc !== null) ns.set(0, totalCol, { v: 'Total', s: head });
    rowVals.forEach((rv, i) => {
      const r = i + 1;
      ns.set(r, 0, { v: rv, s: { b: true } });
      colVals.forEach((cv, j) => {
        const crit = `${src(rc)};${lit(rv)}${cv !== null ? `;${src(cc)};${lit(cv)}` : ''}`;
        const f = fn === 'COUNTIFS' ? `${pt}(${crit})` : `${pt}(${src(vc)};${crit})`;
        ns.set(r, j + 1, { v: null, f, s: { fmt: fn === 'COUNTIFS' ? 'general' : 'number', dec: 2 } });
      });
      if (cc !== null) ns.set(r, totalCol, { v: null, f: `SOMA(B${r + 1}:${colName(totalCol - 1)}${r + 1})`, s: { b: true, fmt: 'number', dec: 2 } });
    });
    const tr = rowVals.length + 1;
    ns.set(tr, 0, { v: 'Total geral', s: { b: true, bd: { t: { w: 'medium' } } } });
    for (let j = 1; j <= (cc !== null ? totalCol : 1); j++) {
      ns.set(tr, j, { v: null, f: fn === 'AVERAGEIFS' ? `MÉDIA(${colName(j)}2:${colName(j)}${tr})` : fn === 'MAXIFS' ? `MÁXIMO(${colName(j)}2:${colName(j)}${tr})` : fn === 'MINIFS' ? `MÍNIMO(${colName(j)}2:${colName(j)}${tr})` : `SOMA(${colName(j)}2:${colName(j)}${tr})`, s: { b: true, fmt: 'number', dec: 2, bd: { t: { w: 'medium' } } } });
    }
    ns.cols.setSize(0, 180);
    ns.freeze = { r: 1, c: 1 };
  });
  restoreView();
  renderTabs();
}

const quote = (n) => (/^[A-Za-z_][\w.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);

// ---------------------------------------------------------------------------
// Localizar e substituir

async function findDialog(replace = false) {
  const s = sheet();
  const q = el('input', { type: 'search', placeholder: 'Localizar' });
  const rep = el('input', { type: 'text', placeholder: 'Substituir por' });
  const mc = el('input', { type: 'checkbox' });
  const whole = el('input', { type: 'checkbox' });
  const inFormulas = el('input', { type: 'checkbox' });
  const status = el('span', { class: 'muted' });
  let matches = [], idx = -1;
  const search = () => {
    matches = [];
    const needle = mc.checked ? q.value : q.value.toLowerCase();
    if (!needle) { status.textContent = ''; return; }
    for (const [k, cell] of [...s.cells].sort((a, b) => a[0] - b[0])) {
      const [r, c] = RC(k);
      let t = inFormulas.checked && cell.f ? `=${cell.f}` : display(wb().value(s, r, c), cell.s || {});
      if (!mc.checked) t = t.toLowerCase();
      if (whole.checked ? t === needle : t.includes(needle)) matches.push([r, c]);
    }
    status.textContent = matches.length ? `${matches.length} resultado(s)` : 'Nada encontrado';
    idx = -1;
  };
  const next = () => {
    if (!matches.length) search();
    if (!matches.length) return;
    idx = (idx + 1) % matches.length;
    const [r, c] = matches[idx];
    grid.select(r, c);
    grid.scrollIntoView(r, c);
    status.textContent = `${idx + 1} de ${matches.length}`;
  };
  const replaceIn = (r, c) => {
    const cell = s.get(r, c);
    const src = cell?.f ? `=${cell.f}` : editText(cell);
    const re = new RegExp(q.value.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), mc.checked ? 'g' : 'gi');
    const out = whole.checked ? rep.value : src.replace(re, rep.value);
    if (out !== src) setCellInput(r, c, out);
  };
  [q, mc, whole, inFormulas].forEach((x) => x.addEventListener(x === q ? 'input' : 'change', search));
  q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); e.stopPropagation(); next(); } });
  await dialog({
    title: replace ? 'Localizar e substituir' : 'Localizar', width: 460, okText: null, cancelText: 'Fechar',
    body: [q, replace ? rep : null, el('div', { class: 'row' }, el('label', { class: 'check' }, mc, 'Diferenciar maiúsculas'), el('label', { class: 'check' }, whole, 'Célula inteira'), el('label', { class: 'check' }, inFormulas, 'Procurar em fórmulas')), status],
    extraButtons: [
      { label: 'Próximo', onClick: () => next() },
      ...(replace ? [
        { label: 'Substituir', onClick: () => { if (idx < 0) next(); if (idx >= 0) { replaceIn(...matches[idx]); search(); next(); } } },
        { label: 'Substituir tudo', onClick: () => {
          search();
          const list = [...matches];
          wb().transact('Substituir tudo', () => list.forEach(([r, c]) => replaceIn(r, c)));
          toast(`${list.length} célula(s) alterada(s).`);
          search();
        } },
      ] : []),
    ],
  });
  focusGrid();
}

// ---------------------------------------------------------------------------
// Inserir função

async function insertFunctionDialog() {
  const search = el('input', { type: 'search', placeholder: 'Pesquisar funções (ex.: soma, data, procura)' });
  const cat = select([['', 'Todas as categorias'], ...CATEGORIES.map((c) => [c, c])], '');
  const list = el('div', { class: 'fn-list' });
  const info = el('div', { class: 'fn-info' });
  let chosen = null;
  const all = Object.values(FUNCS).map((f) => ({ ...f, label: f.pt.split('|')[0] })).sort((a, b) => a.label.localeCompare(b.label, 'pt-BR'));
  const norm2 = (x) => x.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();
  const draw = () => {
    const q = norm2(search.value);
    const items = all.filter((f) => (!cat.value || f.cat === cat.value) && (!q || norm2(`${f.label} ${f.name} ${f.d}`).includes(q)));
    list.replaceChildren(...items.map((f) => el('div', {
      class: `fn-row${chosen === f ? ' on' : ''}`,
      onclick: () => { chosen = f; draw(); },
      ondblclick: () => { chosen = f; done?.(true); },
    }, el('b', {}, f.label), el('span', { class: 'muted' }, f.name !== f.label ? ` (${f.name})` : ''))));
    info.innerHTML = chosen ? `<code>${escapeHtml(chosen.s)}</code><p>${escapeHtml(chosen.d)}</p><small class="muted">${escapeHtml(chosen.cat)}</small>` : '<p class="muted">Escolha uma função para ver como usar.</p>';
  };
  search.addEventListener('input', draw);
  cat.addEventListener('change', draw);
  draw();
  let done;
  const ok = await dialog({
    title: 'Inserir função', wide: true, okText: 'Inserir',
    body: [el('div', { class: 'row' }, search, cat), el('div', { class: 'fn-pick' }, list, info)],
    onOpen: (_b, close) => { done = close; },
  });
  if (!ok || !chosen) { focusGrid(); return; }
  if (editor.active) {
    const elx = editor.el;
    const p = elx.selectionStart;
    elx.value = `${elx.value.slice(0, p)}${chosen.label}(${elx.value.slice(p)}`;
    editor.input.value = elx.value;
    bar.value = elx.value;
    elx.focus();
    elx.setSelectionRange(p + chosen.label.length + 1, p + chosen.label.length + 1);
    editor.update();
  } else editor.begin({ text: `=${chosen.label}(` });
}

// ---------------------------------------------------------------------------
// Formatar células (Ctrl+1)

async function formatCellsDialog() {
  const cell = sheet().get(grid.sel.r, grid.sel.c);
  const st = cell?.s || {};
  const fmt = select([...FORMATS, ['custom', 'Personalizado']], st.fmt || 'general');
  const dec = el('input', { type: 'number', min: 0, max: 12, value: st.dec ?? 2 });
  const cur = select([['BRL', 'Real (R$)'], ['USD', 'Dólar (US$)'], ['EUR', 'Euro (€)']], st.cur || 'BRL');
  const code = el('input', { type: 'text', value: st.code || '#.##0,00', placeholder: 'Ex.: #.##0,00 "kg"' });
  const preview = el('div', { class: 'fmt-preview' });
  const sample = wb().value(sheet(), grid.sel.r, grid.sel.c);
  const upd = () => {
    const s2 = { fmt: fmt.value, dec: +dec.value, cur: cur.value, code: code.value };
    preview.textContent = `Exemplo: ${display(typeof sample === 'number' ? sample : 1234.5678, s2)}`;
  };
  [fmt, dec, cur, code].forEach((x) => x.addEventListener('input', upd));
  upd();
  const ok = await dialog({
    title: 'Formatar células', width: 460,
    body: [el('div', { class: 'grid2' }, field('Categoria', fmt), field('Casas decimais', dec)), el('div', { class: 'grid2' }, field('Moeda', cur), field('Código personalizado', code)), preview],
  });
  if (ok) setFormat(fmt.value, { dec: +dec.value, cur: cur.value, code: fmt.value === 'custom' ? code.value : null });
  focusGrid();
}

// ---------------------------------------------------------------------------
// Cabeçalho, menus e barra de ferramentas

function exportItems() {
  return [
    { label: 'Excel (.xlsx)', icon: 'sheet', action: () => exportAs('xlsx') },
    { label: 'CSV (.csv)', icon: 'text', action: () => exportAs('csv') },
    { label: 'PDF (.pdf)', icon: 'pdf', key: 'Ctrl+P', action: () => exportAs('pdf') },
  ];
}

function setupHeader() {
  $('#badge').innerHTML = icon('sheet', 20);
  const t = $('#doc-title');
  t.addEventListener('input', () => { S.book.title = t.value || 'Planilha sem título'; S.book.dirty = true; $('#save-state').textContent = 'Alterações não salvas'; renderBookTabs(); });
  t.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); focusGrid(); } });
  t.addEventListener('blur', renderTitle);
  const exp = el('button', { class: 'btn', html: `${icon('download', 16)} Exportar`, 'data-popover-anchor': '' });
  exp.addEventListener('click', () => dropdown(exp, exportItems()));
  $('#header-actions').append(exp, el('button', { class: 'btn primary', html: `${icon('save', 16)} Salvar`, onclick: () => save() }));
}

function setupMenus() {
  menubar($('#menus'), [
    ['Arquivo', () => [
      { label: 'Nova planilha', icon: 'plus', key: 'Ctrl+N', action: () => newBook() },
      { label: 'Novo a partir de modelo', icon: 'layout', sub: Object.entries(SHEET_TEMPLATES).map(([id, t]) => ({ label: id === 'blank' ? 'Em branco' : t.title, action: () => newBook(id) })) },
      { label: 'Abrir…', icon: 'open', key: 'Ctrl+O', action: openDialog },
      '-',
      { label: 'Salvar', icon: 'save', key: 'Ctrl+S', action: () => save() },
      { label: 'Salvar como…', key: 'Ctrl+Shift+S', action: () => save(S.book, true) },
      { label: 'Exportar', icon: 'download', sub: exportItems() },
      '-',
      { label: 'Fechar', action: () => closeBook() },
    ]],
    ['Editar', () => [
      { label: 'Desfazer', icon: 'undo', key: 'Ctrl+Z', action: undo, disabled: !wb().undoStack.length },
      { label: 'Refazer', icon: 'redo', key: 'Ctrl+Y', action: redo, disabled: !wb().redoStack.length },
      '-',
      { label: 'Recortar', key: 'Ctrl+X', action: () => { focusGrid(); document.execCommand('cut'); } },
      { label: 'Copiar', key: 'Ctrl+C', action: () => { focusGrid(); document.execCommand('copy'); } },
      { label: 'Colar', key: 'Ctrl+V', action: () => pasteFromMenu() },
      { label: 'Colar somente valores', key: 'Ctrl+Shift+V', action: () => pasteFromMenu(true) },
      '-',
      { label: 'Preencher abaixo', key: 'Ctrl+D', action: () => fillDir('down') },
      { label: 'Preencher à direita', key: 'Ctrl+R', action: () => fillDir('right') },
      { label: 'Limpar', sub: [
        { label: 'Conteúdo', key: 'Delete', action: () => clearRange('contents') },
        { label: 'Formatos', action: () => clearRange('formats') },
        { label: 'Tudo', action: () => clearRange('all') },
      ] },
      '-',
      { label: 'Localizar…', icon: 'search', key: 'Ctrl+F', action: () => findDialog() },
      { label: 'Localizar e substituir…', key: 'Ctrl+H', action: () => findDialog(true) },
    ]],
    ['Exibir', () => [
      { label: 'Congelar', icon: 'freeze', sub: [
        { label: 'Primeira linha', action: () => freeze('row') },
        { label: 'Primeira coluna', action: () => freeze('col') },
        { label: 'Até a célula selecionada', action: () => freeze('sel') },
        { label: 'Descongelar', action: () => freeze('none') },
      ] },
      { label: 'Linhas de grade', checked: sheet().showGrid, action: () => wb().transact('Linhas de grade', () => { wb().snapSheet(sheet()); sheet().showGrid = !sheet().showGrid; }) },
      '-',
      { label: 'Ampliar', icon: 'zoomIn', key: 'Ctrl+=', action: () => grid.setZoom(grid.zoom * 1.1) },
      { label: 'Reduzir', icon: 'zoomOut', key: 'Ctrl+-', action: () => grid.setZoom(grid.zoom / 1.1) },
      { label: '100%', key: 'Ctrl+0', action: () => grid.setZoom(1) },
    ]],
    ['Inserir', () => [
      { label: 'Linhas acima', icon: 'rowInsert', action: () => insertCells('row', 'before') },
      { label: 'Linhas abaixo', action: () => insertCells('row', 'after') },
      { label: 'Colunas à esquerda', icon: 'colInsert', action: () => insertCells('col', 'before') },
      { label: 'Colunas à direita', action: () => insertCells('col', 'after') },
      '-',
      { label: 'Gráfico', icon: 'chart', sub: CHART_TYPES.map(([k, l]) => ({ label: l, action: () => insertChart(k) })) },
      { label: 'Tabela dinâmica…', icon: 'pivot', action: pivotDialog },
      { label: 'Função…', icon: 'function', key: 'Shift+F3', action: insertFunctionDialog },
      { label: 'Nota…', icon: 'notes', key: 'Ctrl+K', action: noteDialog },
      '-',
      { label: 'Nova planilha', icon: 'plus', action: () => { saveView(); wb().addSheet(); restoreView(); renderTabs(); } },
      { label: 'Data de hoje', icon: 'calendar', key: 'Ctrl+;', action: () => KEYMAP['Ctrl+;']() },
    ]],
    ['Formatar', () => [
      { label: 'Formatar células…', icon: 'settings', key: 'Ctrl+1', action: formatCellsDialog },
      { label: 'Número', sub: FORMATS.map(([k, l]) => ({ label: l, action: () => setFormat(k) })) },
      '-',
      { label: 'Negrito', icon: 'bold', key: 'Ctrl+B', action: () => toggleStyle('b') },
      { label: 'Itálico', icon: 'italic', key: 'Ctrl+I', action: () => toggleStyle('i') },
      { label: 'Sublinhado', icon: 'underline', key: 'Ctrl+U', action: () => toggleStyle('u') },
      { label: 'Tachado', icon: 'strike', key: 'Ctrl+5', action: () => toggleStyle('st') },
      '-',
      { label: 'Mesclar e centralizar', icon: 'merge', action: () => mergeCells(true) },
      { label: 'Quebrar texto', icon: 'wrap', action: () => toggleStyle('wrap') },
      { label: 'Formatação condicional…', icon: 'palette', action: cfDialog },
      { label: 'Pincel de formatação', icon: 'brush', action: startPainter },
      '-',
      { label: 'Ajustar largura das colunas', action: () => { for (let c = grid.sel.c1; c <= Math.min(grid.sel.c2, sheet().used().c); c++) autofit('col', c); } },
      { label: 'Limpar formatação', icon: 'clearFormat', action: () => clearRange('formats') },
    ]],
    ['Dados', () => [
      { label: 'Classificar A → Z', icon: 'sortAsc', action: () => sortSel(false) },
      { label: 'Classificar Z → A', icon: 'sortDesc', action: () => sortSel(true) },
      { label: 'Classificar por vários critérios…', action: sortDialog },
      { label: sheet().filter ? 'Remover filtro' : 'Criar filtro', icon: 'filter', key: 'Ctrl+Shift+L', action: toggleFilter },
      '-',
      { label: 'Remover duplicatas…', action: removeDuplicates },
      { label: 'Validação de dados…', icon: 'check', action: validationDialog },
      { label: 'Tabela dinâmica…', icon: 'pivot', action: pivotDialog },
      '-',
      { label: 'Soma automática', icon: 'sigma', key: 'Alt+=', action: () => autoSum('SUM') },
      { label: 'Recalcular tudo', action: () => { wb().recalc(); refresh(); toast('Recalculado.'); } },
    ]],
    ['Ajuda', () => [
      { label: 'Atalhos e recursos', icon: 'info', key: 'F1', action: helpDialog },
      { label: 'Lista de funções…', icon: 'function', action: insertFunctionDialog },
    ]],
  ]);
}

const tb = {};
function setupToolbar() {
  const tbar = $('#toolbar');
  const colorBtn = (ic, title, key) => {
    const b = toolButton(ic, title, () => colorPopover(b, (c) => { applyStyle({ [key]: c }); if (c) b.querySelector('.swatch-bar').style.background = c; focusGrid(); }, { allowNone: true, noneLabel: key === 'bg' ? 'Sem preenchimento' : 'Automática' }));
    b.append(el('span', { class: 'swatch-bar', style: { background: key === 'bg' ? '#ffe066' : '#e03131' } }));
    b.setAttribute('data-popover-anchor', '');
    return b;
  };
  tb.font = select([['', 'Calibri'], ['Arial, sans-serif', 'Arial'], ['"Segoe UI", sans-serif', 'Segoe UI'], ['Verdana, sans-serif', 'Verdana'], ['Georgia, serif', 'Georgia'], ['"Times New Roman", serif', 'Times New Roman'], ['"Courier New", monospace', 'Courier New'], ['Consolas, monospace', 'Consolas']], '', { class: 'tb-select', title: 'Fonte', style: { width: '112px' } });
  tb.font.addEventListener('change', () => { applyStyle({ font: tb.font.value || null }); focusGrid(); });
  tb.size = select([8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48].map((n) => [n, n]), 11, { class: 'tb-select', title: 'Tamanho', style: { width: '58px' } });
  tb.size.addEventListener('change', () => { applyStyle({ size: +tb.size.value === 11 ? null : +tb.size.value }); focusGrid(); });
  tb.fmt = select([...FORMATS, ['custom', 'Personalizado…']], 'general', { class: 'tb-select', title: 'Formato de número', style: { width: '118px' } });
  tb.fmt.addEventListener('change', () => { if (tb.fmt.value === 'custom') formatCellsDialog(); else setFormat(tb.fmt.value); focusGrid(); });
  const borders = toolButton('border', 'Bordas', () => dropdown(borders, [
    ['all', 'Todas as bordas'], ['outer', 'Borda externa'], ['thick', 'Borda externa espessa'], ['bottom', 'Borda inferior'], ['top', 'Borda superior'],
    ['left', 'Borda esquerda'], ['right', 'Borda direita'], ['doubleBottom', 'Borda inferior grossa'], ['none', 'Sem bordas'],
  ].map(([k, l]) => ({ label: l, action: () => setBorders(k) }))));
  borders.setAttribute('data-popover-anchor', '');
  const sigma = toolButton('sigma', 'Soma automática (Alt+=)', () => autoSum('SUM'));
  const sigmaMore = toolButton('chevronDown', 'Outras funções', () => dropdown(sigmaMore, [
    ['SUM', 'Soma'], ['AVERAGE', 'Média'], ['COUNT', 'Contar números'], ['MAX', 'Máximo'], ['MIN', 'Mínimo'],
  ].map(([k, l]) => ({ label: l, action: () => autoSum(k) })).concat(['-', { label: 'Mais funções…', icon: 'function', action: insertFunctionDialog }])));
  sigmaMore.setAttribute('data-popover-anchor', '');
  sigmaMore.style.minWidth = '18px';
  sigmaMore.style.padding = '0 2px';
  const chartBtn = toolButton('chart', 'Inserir gráfico', () => dropdown(chartBtn, CHART_TYPES.map(([k, l]) => ({ label: l, action: () => insertChart(k) }))));
  chartBtn.setAttribute('data-popover-anchor', '');
  const freezeBtn = toolButton('freeze', 'Congelar painéis', () => dropdown(freezeBtn, [
    { label: 'Congelar primeira linha', action: () => freeze('row') }, { label: 'Congelar primeira coluna', action: () => freeze('col') },
    { label: 'Congelar até a seleção', action: () => freeze('sel') }, { label: 'Descongelar', action: () => freeze('none') },
  ]));
  freezeBtn.setAttribute('data-popover-anchor', '');
  const alignBtn = (ha, ic, title) => toolButton(ic, title, () => applyStyle({ ha: sheet().get(grid.sel.r, grid.sel.c)?.s?.ha === ha ? null : ha }));
  tb.b = toolButton('bold', 'Negrito', () => toggleStyle('b'), { key: 'Ctrl+B' });
  tb.i = toolButton('italic', 'Itálico', () => toggleStyle('i'), { key: 'Ctrl+I' });
  tb.u = toolButton('underline', 'Sublinhado', () => toggleStyle('u'), { key: 'Ctrl+U' });
  tb.st = toolButton('strike', 'Tachado', () => toggleStyle('st'));
  tb.left = alignBtn('left', 'alignLeft', 'Alinhar à esquerda');
  tb.center = alignBtn('center', 'alignCenter', 'Centralizar');
  tb.right = alignBtn('right', 'alignRight', 'Alinhar à direita');
  tb.wrap = toolButton('wrap', 'Quebrar texto', () => toggleStyle('wrap'));
  tb.merge = toolButton('merge', 'Mesclar e centralizar', () => mergeCells(true));
  tb.filter = toolButton('filter', 'Filtro (Ctrl+Shift+L)', toggleFilter);
  tbar.append(
    toolButton('undo', 'Desfazer', undo, { key: 'Ctrl+Z' }), toolButton('redo', 'Refazer', redo, { key: 'Ctrl+Y' }), toolButton('brush', 'Pincel de formatação', startPainter),
    sep(), tb.font, tb.size,
    sep(), tb.b, tb.i, tb.u, tb.st, colorBtn('textColor', 'Cor do texto', 'color'), colorBtn('bucket', 'Cor de preenchimento', 'bg'), borders,
    sep(), tb.left, tb.center, tb.right,
    toolButton('alignTop', 'Alinhar em cima', () => applyStyle({ va: 'top' })), toolButton('alignMiddle', 'Alinhar ao meio', () => applyStyle({ va: 'middle' })), toolButton('alignBottom', 'Alinhar embaixo', () => applyStyle({ va: null })),
    tb.wrap, tb.merge,
    sep(), tb.fmt, toolButton('currency', 'Moeda (R$)', () => setFormat('currency')), toolButton('percent', 'Porcentagem', () => setFormat('percent')),
    toolButton('decDec', 'Diminuir casas decimais', () => stepDecimals(-1)), toolButton('decInc', 'Aumentar casas decimais', () => stepDecimals(1)),
    sep(), sigma, sigmaMore, toolButton('sortAsc', 'Classificar A → Z', () => sortSel(false)), toolButton('sortDesc', 'Classificar Z → A', () => sortSel(true)), tb.filter,
    sep(), chartBtn, toolButton('pivot', 'Tabela dinâmica', pivotDialog), toolButton('palette', 'Formatação condicional', cfDialog), freezeBtn,
  );
  for (const b of tbar.querySelectorAll('select')) b.addEventListener('mousedown', () => { if (editor.active) editor.commit(); });
}

function refreshToolbar() {
  const st = sheet().get(grid.sel.r, grid.sel.c)?.s || {};
  tb.b.classList.toggle('on', !!st.b);
  tb.i.classList.toggle('on', !!st.i);
  tb.u.classList.toggle('on', !!st.u);
  tb.st.classList.toggle('on', !!st.st);
  tb.wrap.classList.toggle('on', !!st.wrap);
  tb.left.classList.toggle('on', st.ha === 'left');
  tb.center.classList.toggle('on', st.ha === 'center');
  tb.right.classList.toggle('on', st.ha === 'right');
  tb.merge.classList.toggle('on', !!sheet().mergeAt(grid.sel.r, grid.sel.c));
  tb.filter.classList.toggle('on', !!sheet().filter);
  tb.fmt.value = st.fmt && FORMATS.some(([k]) => k === st.fmt) ? st.fmt : st.fmt === 'custom' ? 'custom' : 'general';
  tb.size.value = st.size || 11;
  tb.font.value = st.font || '';
}

function helpDialog() {
  const rows = [
    ['Digitar', 'Começa a editar a célula (Enter confirma, Esc cancela, Alt+Enter quebra linha)'], ['F2 / duplo clique', 'Editar mantendo o conteúdo'],
    ['=SOMA(…)', 'Fórmulas em português ou inglês (SOMA = SUM, PROCV = VLOOKUP…). Use ; ou , como separador'],
    ['Clique durante a fórmula', 'Insere a referência da célula (arraste para intervalos; troque de aba para outras planilhas)'],
    ['A1#', 'Referência ao resultado espalhado de uma matriz dinâmica (FILTRO, ÚNICO, SEQUÊNCIA…)'],
    ['Alça de preenchimento', 'Arraste o quadradinho para copiar ou continuar séries (1, 2, 3… / jan, fev… / Item 1, Item 2…)'],
    ['Ctrl+Setas', 'Ir até o fim dos dados'], ['Shift+Setas', 'Estender a seleção'], ['Ctrl+A', 'Selecionar a tabela atual (de novo: tudo)'],
    ['Ctrl+C / X / V', 'Copiar, recortar e colar (também de/para Excel e Google Sheets)'], ['Ctrl+Shift+V', 'Colar somente valores'],
    ['Ctrl+D / Ctrl+R', 'Preencher abaixo / à direita'], ['Ctrl+Enter', 'Escrever o mesmo valor em toda a seleção'], ['Alt+=', 'Soma automática'],
    ['Ctrl+;', 'Data de hoje'], ['Ctrl+1', 'Formatar células'], ['Ctrl+Shift+4 / 5', 'Moeda / porcentagem'], ['Ctrl+Shift+L', 'Filtro'],
    ['Ctrl+F / Ctrl+H', 'Localizar / substituir'], ['Shift+F3', 'Inserir função'], ['Ctrl+PgUp / PgDn', 'Trocar de planilha'], ['Ctrl+roda', 'Zoom'],
  ];
  dialog({ title: 'Atalhos e recursos', cancelText: null, wide: true, body: el('div', { class: 'help-grid' }, ...rows.flatMap(([k, v]) => [el('kbd', {}, k), el('span', {}, v)])) });
}

// ---------------------------------------------------------------------------
// Caixa de nome

function setupNameBox() {
  nameBox.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { renderFormulaBar(); focusGrid(); return; }
    if (e.key !== 'Enter') return;
    e.preventDefault();
    const t = nameBox.value.trim().replace(/\$/g, '');
    let m = /^([A-Za-z]{1,3})(\d+)(?::([A-Za-z]{1,3})(\d+))?$/.exec(t);
    if (m) {
      const r1 = +m[2] - 1, c1 = colIndex(m[1]);
      const r2 = m[3] ? +m[4] - 1 : r1, c2 = m[3] ? colIndex(m[3]) : c1;
      grid.setRange(norm(r1, c1, r2, c2), { r: r1, c: c1 });
      grid.scrollIntoView(r1, c1);
      focusGrid();
      return;
    }
    m = /^([A-Za-z]{1,3}):([A-Za-z]{1,3})$/.exec(t);
    if (m) { const c1 = colIndex(m[1]), c2 = colIndex(m[2]); grid.setRange({ r1: 0, r2: sheet().rows.count - 1, c1: Math.min(c1, c2), c2: Math.max(c1, c2) }); focusGrid(); return; }
    toast('Endereço inválido. Ex.: B7 ou A1:D20');
  });
}

// ---------------------------------------------------------------------------
// Inicialização

function setupSink() {
  sink.addEventListener('keydown', onSinkKey);
  sink.addEventListener('input', (e) => {
    if (e.isComposing) return;
    const text = sink.value;
    sink.value = '';
    if (text) editor.begin({ text });
  });
  sink.addEventListener('compositionend', () => {
    const text = sink.value;
    sink.value = '';
    if (text) editor.begin({ text });
  });
  sink.addEventListener('copy', (e) => onCopy(e, false));
  sink.addEventListener('cut', (e) => onCopy(e, true));
  sink.addEventListener('paste', (e) => onPaste(e));
  document.addEventListener('keydown', (e) => {
    // Atalhos globais quando o foco está fora da grade (ex.: barra de fórmulas sem edição)
    if (isDialogOpen() || e.target === sink || editor.active) return;
    if (e.target instanceof HTMLInputElement || e.target instanceof HTMLTextAreaElement || e.target instanceof HTMLSelectElement) {
      const c = comboOf(e);
      if (c === 'Ctrl+S') { e.preventDefault(); save(); }
      return;
    }
    const fn = KEYMAP[comboOf(e)];
    if (fn) { e.preventDefault(); fn(); }
  });
  $('#zoom-out').addEventListener('click', () => grid.setZoom(grid.zoom / 1.1));
  $('#zoom-in').addEventListener('click', () => grid.setZoom(grid.zoom * 1.1));
  $('#zoom-val').addEventListener('click', () => grid.setZoom(1));
  $('#fx-btn').addEventListener('click', insertFunctionDialog);
  window.addEventListener('focus', () => { if (!editor.active && !isDialogOpen() && !(document.activeElement instanceof HTMLInputElement)) focusGrid(); });
  grid.root.addEventListener('mousedown', () => setTimeout(() => { if (!editor.active) focusGrid(); }, 0));
}

registerApp({
  openFiles,
  newDocument: (tpl) => newBook(tpl),
  unsaved: () => S.books.filter((b) => b.dirty).map((b) => b.title),
  beforeClose: async () => {
    for (const b of [...S.books]) if (b.dirty && !(await askSave(b))) return false;
    return true;
  },
  activated: () => { grid.readTheme(); focusGrid(); },
  themeChanged: () => { grid.readTheme(); charts.sync(); },
});

initTheme();
window.addEventListener('storage', (e) => { if (e.key === 'lumina.theme') setTimeout(() => { grid.readTheme(); charts.sync(); }, 50); });
setupHeader();
setupMenus();
setupToolbar();
setupNameBox();
setupSink();
S.books.push(makeBook());
activate(S.books[0]);
focusGrid();
standaloneBoot({ openFiles });
const onResize = debounce(() => charts.position(), 50);
window.addEventListener('resize', onResize);
window.__sheets = { S, grid, editor, charts, io, setCellInput, newBook, openFiles, save, exportAs, fill, pivotDialog, insertChart };
void RH_W;
void KEY;
