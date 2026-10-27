// Modelo da pasta de trabalho: planilhas, células, recálculo, desfazer e operações estruturais.
import {
  ERR, err, isErr, isMatrix, RangeRef, toMatrix, compare, colName,
} from './values.js';
import {
  parse, evaluate, tokenize, MAX_ROW, MAX_COL, shiftFormula, adjustForStructure, renameSheetInFormula,
} from './formula.js';
import { FUNCS } from './functions.js';

export const KEY = (r, c) => r * 16384 + c;
export const RC = (k) => [Math.floor(k / 16384), k % 16384];
const clone = (x) => (x === undefined ? undefined : structuredClone(x));

// ---------------------------------------------------------------------------
// Eixo (linhas ou colunas) com tamanhos variáveis esparsos

export class Axis {
  constructor(def, count) {
    this.def = def;
    this.count = count;
    this.sizes = new Map();
    this.hidden = new Set();
    this.dirty = true;
  }

  size(i) {
    if (this.hidden.has(i)) return 0;
    return this.sizes.get(i) ?? this.def;
  }

  setSize(i, s) {
    if (Math.round(s) === this.def) this.sizes.delete(i);
    else this.sizes.set(i, Math.max(2, Math.round(s)));
    this.dirty = true;
  }

  setHidden(i, h) {
    if (h) this.hidden.add(i); else this.hidden.delete(i);
    this.dirty = true;
  }

  rebuild() {
    const keys = new Set([...this.sizes.keys(), ...this.hidden]);
    this.keys = [...keys].sort((a, b) => a - b);
    this.cum = [];
    let acc = 0;
    for (const k of this.keys) {
      acc += this.size(k) - this.def;
      this.cum.push(acc);
    }
    this.dirty = false;
  }

  // Posição (px) do início do índice i
  pos(i) {
    if (this.dirty) this.rebuild();
    let lo = 0, hi = this.keys.length;
    while (lo < hi) {
      const m = (lo + hi) >> 1;
      if (this.keys[m] < i) lo = m + 1; else hi = m;
    }
    return i * this.def + (lo > 0 ? this.cum[lo - 1] : 0);
  }

  indexAt(px) {
    if (px <= 0) return 0;
    let lo = 0, hi = this.count - 1;
    while (lo < hi) {
      const m = (lo + hi + 1) >> 1;
      if (this.pos(m) <= px) lo = m; else hi = m - 1;
    }
    return lo;
  }

  total() {
    return this.pos(this.count);
  }

  shift(at, n) {
    const remap = (m) => {
      const out = new Map();
      for (const [k, v] of m) {
        if (k < at) out.set(k, v);
        else if (n > 0) out.set(k + n, v);
        else if (k >= at - n) out.set(k + n, v);
      }
      return out;
    };
    this.sizes = remap(this.sizes);
    this.hidden = new Set([...remap(new Map([...this.hidden].map((k) => [k, 1]))).keys()]);
    this.dirty = true;
  }

  toJSON() {
    return { def: this.def, count: this.count, sizes: [...this.sizes], hidden: [...this.hidden] };
  }

  static from(j) {
    const a = new Axis(j.def, j.count);
    a.sizes = new Map(j.sizes || []);
    a.hidden = new Set(j.hidden || []);
    return a;
  }
}

// ---------------------------------------------------------------------------
// Planilha

let sheetSeq = 1;
export class Sheet {
  constructor(name) {
    this.id = sheetSeq++;
    this.name = name;
    this.cells = new Map();
    this.rows = new Axis(24, 1000);
    this.cols = new Axis(100, 26);
    this.merges = [];
    this.freeze = { r: 0, c: 0 };
    this.cf = [];
    this.charts = [];
    this.filter = null;
    this.validations = [];
    this.tabColor = null;
    this.showGrid = true;
    this._used = null;
  }

  get(r, c) {
    return this.cells.get(KEY(r, c));
  }

  set(r, c, cell) {
    const k = KEY(r, c);
    if (!cell || (cell.v == null && !cell.f && (!cell.s || !Object.keys(cell.s).length) && !cell.note)) this.cells.delete(k);
    else this.cells.set(k, cell);
    this._used = null;
    if (r >= this.rows.count - 50) this.rows.count = Math.min(MAX_ROW + 1, Math.max(this.rows.count, r + 200));
    if (c >= this.cols.count - 5) this.cols.count = Math.min(MAX_COL + 1, Math.max(this.cols.count, c + 10));
  }

  used() {
    if (!this._used) {
      let mr = -1, mc = -1;
      for (const [k, cell] of this.cells) {
        if (cell.v == null && !cell.f) continue;
        const r = Math.floor(k / 16384), c = k % 16384;
        if (r > mr) mr = r;
        if (c > mc) mc = c;
        if (cell.spillRows) { mr = Math.max(mr, r + cell.spillRows - 1); mc = Math.max(mc, c + cell.spillCols - 1); }
      }
      this._used = { r: mr, c: mc };
    }
    return this._used;
  }

  mergeAt(r, c) {
    return this.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2) || null;
  }

  toJSON() {
    return {
      id: this.id, name: this.name,
      cells: [...this.cells].map(([k, c]) => [k, { v: c.v, f: c.f, s: c.s, note: c.note, forceText: c.forceText }]),
      rows: this.rows.toJSON(), cols: this.cols.toJSON(), merges: clone(this.merges), freeze: clone(this.freeze),
      cf: this.cf.map(({ _stats, ...rule }) => clone(rule)),
      charts: clone(this.charts), filter: clone(this.filter), validations: clone(this.validations), tabColor: this.tabColor, showGrid: this.showGrid,
    };
  }

  load(j) {
    this.name = j.name;
    this.cells = new Map(j.cells.map(([k, c]) => [k, clone(c)]));
    this.rows = Axis.from(j.rows);
    this.cols = Axis.from(j.cols);
    this.merges = clone(j.merges || []);
    this.freeze = clone(j.freeze || { r: 0, c: 0 });
    this.cf = clone(j.cf || []);
    this.charts = clone(j.charts || []);
    this.filter = clone(j.filter || null);
    this.validations = clone(j.validations || []);
    this.tabColor = j.tabColor || null;
    this.showGrid = j.showGrid !== false;
    this._used = null;
    return this;
  }

  static from(j) {
    const s = new Sheet(j.name);
    if (j.id) { s.id = j.id; sheetSeq = Math.max(sheetSeq, j.id + 1); }
    return s.load(j);
  }
}

// ---------------------------------------------------------------------------
// Pasta de trabalho

export class Workbook {
  constructor() {
    this.sheets = [new Sheet('Planilha1')];
    this.active = 0;
    this.names = {};
    this.undoStack = [];
    this.redoStack = [];
    this.tx = null;
    this.spills = new Map();
    this.listeners = [];
    this.version = 0;
  }

  get sheet() { return this.sheets[this.active]; }

  byName(name) {
    const n = String(name).toLowerCase();
    return this.sheets.find((s) => s.name.toLowerCase() === n) || null;
  }

  byId(id) { return this.sheets.find((s) => s.id === id) || null; }

  onChange(fn) { this.listeners.push(fn); }
  emit(kind) { this.version++; for (const fn of this.listeners) fn(kind); }

  // ---- Cálculo --------------------------------------------------------------

  recalc() {
    for (const s of this.sheets) for (const c of s.cells.values()) if (c.f) { c.spillRows = 0; c.spillCols = 0; }
    this.spills = new Map();
    this.pass();
    if (this.layoutSpills()) this.pass();
    for (const s of this.sheets) s._used = null;
    this.computeCf();
  }

  pass() {
    this.state = new Map();
    for (const s of this.sheets) for (const [k, c] of s.cells) if (c.f) this.compute(s, k, c);
  }

  compute(sheet, key, cell) {
    const sk = `${sheet.id}:${key}`;
    const st = this.state.get(sk);
    if (st === 2) return cell.val;
    if (st === 1) return err(ERR.CIRC, 'Referência circular');
    this.state.set(sk, 1);
    if (!cell.ast) {
      try {
        cell.ast = parse(cell.f);
        cell.parseError = null;
        cell.autoFmt = cell.ast.type === 'func' ? FUNCS[cell.ast.name]?.fmt || null : null;
      } catch (e) {
        cell.ast = { type: 'err', v: ERR.NAME };
        cell.parseError = e.message;
      }
    }
    const [r, c] = RC(key);
    let v;
    try {
      v = evaluate(cell.ast, this.ctx(sheet, r, c));
    } catch (e) {
      v = isErr(e) ? e : err(ERR.VALUE, e.message);
    }
    if (v instanceof RangeRef) v = v.rows === 1 && v.cols === 1 ? v.matrix()[0][0] : v.matrix();
    if (isMatrix(v)) {
      const m = toMatrix(v);
      v = m.length === 1 && m[0].length === 1 ? m[0][0] : m;
      if (isMatrix(v) && cell.spillBlocked) v = err(ERR.SPILL, 'Há células ocupadas no caminho do resultado');
    }
    if (v === undefined) v = null;
    cell.val = v;
    this.state.set(sk, 2);
    return v;
  }

  // Reserva células para resultados matriciais; retorna true se algum foi espalhado.
  layoutSpills() {
    let any = false;
    for (const s of this.sheets) {
      const map = new Map();
      for (const [k, cell] of s.cells) {
        if (!cell.f || !isMatrix(cell.val)) continue;
        const [r, c] = RC(k);
        const m = cell.val, rows = m.length, cols = m[0]?.length || 0;
        let blocked = r + rows - 1 > MAX_ROW || c + cols - 1 > MAX_COL;
        for (let i = 0; i < rows && !blocked; i++) {
          for (let j = 0; j < cols; j++) {
            if (!i && !j) continue;
            const kk = KEY(r + i, c + j);
            const other = s.cells.get(kk);
            if ((other && (other.v != null || other.f)) || map.has(kk)) { blocked = true; break; }
          }
        }
        cell.spillBlocked = blocked;
        if (blocked) continue;
        cell.spillRows = rows;
        cell.spillCols = cols;
        for (let i = 0; i < rows; i++) for (let j = 0; j < cols; j++) if (i || j) map.set(KEY(r + i, c + j), { a: k, i, j });
        any = true;
      }
      this.spills.set(s.id, map);
    }
    return any;
  }

  // Valor exibido / referenciado de uma célula.
  value(sheet, r, c) {
    const k = KEY(r, c);
    const cell = sheet.cells.get(k);
    if (cell) {
      if (cell.f) {
        const v = this.state ? this.compute(sheet, k, cell) : cell.val;
        return isMatrix(v) ? v[0]?.[0] ?? null : v ?? null;
      }
      return cell.v ?? null;
    }
    const sp = this.spills.get(sheet.id)?.get(k);
    if (sp) {
      const anchor = sheet.cells.get(sp.a);
      const m = this.state ? this.compute(sheet, sp.a, anchor) : anchor?.val;
      return isMatrix(m) ? m[sp.i]?.[sp.j] ?? null : null;
    }
    return null;
  }

  spillAnchor(sheet, r, c) {
    const sp = this.spills.get(sheet.id)?.get(KEY(r, c));
    return sp ? RC(sp.a) : null;
  }

  ctx(sheet, row, col) {
    const wb = this;
    const resolve = (name) => {
      if (!name) return sheet;
      const s = wb.byName(name);
      if (!s) throw err(ERR.REF, `Planilha não encontrada: ${name}`);
      return s;
    };
    return {
      sheet, row, col, locals: null,
      cellValue: (name, r, c) => {
        const s = resolve(name);
        return wb.value(s, r, c);
      },
      rangeValues: (rng) => {
        const s = resolve(rng.sheet);
        const u = s.used();
        const r2 = Math.min(rng.r2, Math.max(u.r, rng.r1));
        const c2 = Math.min(rng.c2, Math.max(u.c, rng.c1));
        const out = [];
        for (let r = rng.r1; r <= r2; r++) {
          const rowv = [];
          for (let c = rng.c1; c <= c2; c++) rowv.push(wb.value(s, r, c));
          out.push(rowv);
        }
        return out.length ? out : [[null]];
      },
      spillRange: (name, r, c) => {
        const s = resolve(name);
        const cell = s.get(r, c);
        if (!cell?.f) return err(ERR.REF);
        const v = wb.state ? wb.compute(s, KEY(r, c), cell) : cell.val;
        if (!isMatrix(v)) return v;
        return new RangeRef(name, r, c, r + v.length - 1, c + v[0].length - 1, this.ctx(sheet, row, col));
      },
      indirect: (text) => wb.indirect(text, sheet, row, col),
      definedName: (name) => {
        const ref = wb.names[name.toUpperCase()];
        return ref ? wb.indirect(ref, sheet, row, col) : undefined;
      },
    };
  }

  indirect(text, sheet, row, col) {
    let toks;
    try { toks = tokenize(text); } catch { return err(ERR.REF); }
    if (toks.length !== 1 || toks[0].t !== 'ref') return err(ERR.REF);
    const r = toks[0].v;
    const ctx = this.ctx(sheet, row, col);
    if (r.kind === 'cell') return ctx.cellValue(r.sheet, r.a.row, r.a.col);
    if (r.kind === 'range') return new RangeRef(r.sheet, r.a.row, r.a.col, r.b.row, r.b.col, ctx);
    if (r.kind === 'cols') return new RangeRef(r.sheet, 0, r.a.col, MAX_ROW, r.b.col, ctx);
    return new RangeRef(r.sheet, r.a.row, 0, r.b.row, MAX_COL, ctx);
  }

  // ---- Formatação condicional (pré-cálculo) -------------------------------

  computeCf() {
    for (const s of this.sheets) {
      for (const rule of s.cf) {
        const nums = [];
        const counts = new Map();
        const { r1, c1, r2, c2 } = rule.range;
        const u = s.used();
        for (let r = r1; r <= Math.min(r2, u.r); r++) {
          for (let c = c1; c <= Math.min(c2, u.c); c++) {
            const v = this.value(s, r, c);
            if (typeof v === 'number') nums.push(v);
            if (v !== null && v !== '') {
              const key = typeof v === 'string' ? v.toLowerCase() : String(v);
              counts.set(key, (counts.get(key) || 0) + 1);
            }
          }
        }
        const sorted = [...nums].sort((a, b) => a - b);
        const n = Math.max(1, Math.min(sorted.length, rule.params?.n || 10));
        const pctN = rule.params?.percent ? Math.max(1, Math.round((sorted.length * n) / 100)) : n;
        rule._stats = {
          min: sorted[0] ?? 0, max: sorted[sorted.length - 1] ?? 0,
          mid: sorted.length ? sorted[Math.floor((sorted.length - 1) / 2)] : 0,
          avg: nums.length ? nums.reduce((a, b) => a + b, 0) / nums.length : 0,
          counts,
          top: sorted[sorted.length - pctN] ?? Infinity,
          bottom: sorted[pctN - 1] ?? -Infinity,
        };
      }
    }
  }

  // ---- Transações e desfazer ---------------------------------------------------

  transact(label, fn) {
    if (this.tx) { fn(); return; }
    const tx = { label, cells: new Map(), sheets: new Map(), wb: null, activeBefore: this.active };
    this.tx = tx;
    try {
      fn();
    } finally {
      this.tx = null;
    }
    if (!tx.cells.size && !tx.sheets.size && !tx.wb) return;
    for (const rec of tx.cells.values()) rec.after = clone(stripCell(rec.sheet.get(...RC(rec.key))));
    for (const rec of tx.sheets.values()) rec.after = this.byId(rec.id)?.toJSON() ?? null;
    if (tx.wb) tx.wb.after = this.snapshot();
    tx.activeAfter = this.active;
    this.undoStack.push(tx);
    if (this.undoStack.length > 200) this.undoStack.shift();
    this.redoStack = [];
    this.recalc();
    this.emit('edit');
  }

  setCell(sheet, r, c, cell) {
    if (this.tx && !this.tx.wb && !this.tx.sheets.has(sheet.id)) {
      const k = `${sheet.id}:${KEY(r, c)}`;
      if (!this.tx.cells.has(k)) this.tx.cells.set(k, { sheet, key: KEY(r, c), before: clone(stripCell(sheet.get(r, c))) });
    }
    if (cell) { cell.ast = null; cell.val = undefined; cell.spillBlocked = false; }
    sheet.set(r, c, cell);
  }

  snapSheet(sheet) {
    if (this.tx && !this.tx.wb && !this.tx.sheets.has(sheet.id)) this.tx.sheets.set(sheet.id, { id: sheet.id, before: sheet.toJSON() });
  }

  snapWorkbook() {
    if (this.tx && !this.tx.wb) this.tx.wb = { before: this.snapshot() };
  }

  snapshot() {
    return { sheets: this.sheets.map((s) => s.toJSON()), active: this.active, names: { ...this.names } };
  }

  restoreSnapshot(snap) {
    this.sheets = snap.sheets.map((j) => Sheet.from(j));
    this.active = Math.min(snap.active, this.sheets.length - 1);
    this.names = { ...snap.names };
  }

  applyTx(tx, dir) {
    const key = dir === 'undo' ? 'before' : 'after';
    if (tx.wb) {
      this.restoreSnapshot(tx.wb[key]);
    } else {
      for (const rec of tx.sheets.values()) this.byId(rec.id)?.load(rec[key]);
      for (const rec of tx.cells.values()) {
        const s = this.byId(rec.sheet.id) || rec.sheet;
        const [r, c] = RC(rec.key);
        const v = clone(rec[key]);
        s.set(r, c, v ? { ...v, ast: null } : null);
      }
    }
    this.active = Math.min(dir === 'undo' ? tx.activeBefore : tx.activeAfter, this.sheets.length - 1);
    this.recalc();
    this.emit('undo');
  }

  undo() {
    const tx = this.undoStack.pop();
    if (!tx) return null;
    this.applyTx(tx, 'undo');
    this.redoStack.push(tx);
    return tx.label;
  }

  redo() {
    const tx = this.redoStack.pop();
    if (!tx) return null;
    this.applyTx(tx, 'redo');
    this.undoStack.push(tx);
    return tx.label;
  }

  // ---- Operações estruturais -------------------------------------------------

  insert(sheet, axis, at, count) {
    this.transact(count > 0 ? (axis === 'row' ? 'Inserir linhas' : 'Inserir colunas') : (axis === 'row' ? 'Excluir linhas' : 'Excluir colunas'), () => {
      for (const s of this.sheets) this.snapSheet(s);
      const isRow = axis === 'row';
      const moved = new Map();
      for (const [k, cell] of sheet.cells) {
        let [r, c] = RC(k);
        const p = isRow ? r : c;
        if (count < 0 && p >= at && p < at - count) continue;
        if (p >= at) { if (isRow) r += count; else c += count; }
        moved.set(KEY(r, c), cell);
      }
      sheet.cells = moved;
      sheet._used = null;
      (isRow ? sheet.rows : sheet.cols).shift(at, count);
      if (count > 0) (isRow ? sheet.rows : sheet.cols).count += count;
      // Fórmulas em todas as planilhas
      for (const s of this.sheets) {
        for (const cell of s.cells.values()) {
          if (!cell.f) continue;
          const nf = adjustForStructure(cell.f, { targetSheet: sheet.name, formulaSheet: s.name, axis, at, count });
          if (nf !== cell.f) { cell.f = nf; cell.ast = null; }
        }
      }
      const adj = (rg) => adjustRange(rg, isRow, at, count);
      sheet.merges = sheet.merges.map(adj).filter(Boolean);
      sheet.cf = sheet.cf.map((x) => ({ ...x, range: adj(x.range) })).filter((x) => x.range);
      sheet.validations = sheet.validations.map((x) => ({ ...x, range: adj(x.range) })).filter((x) => x.range);
      for (const ch of sheet.charts) ch.range = adj(ch.range) || ch.range;
      if (sheet.filter) sheet.filter = null;
    });
  }

  addSheet(name, index = this.sheets.length) {
    let s;
    this.transact('Nova planilha', () => {
      this.snapWorkbook();
      s = new Sheet(name || this.uniqueName('Planilha'));
      this.sheets.splice(index, 0, s);
      this.active = index;
    });
    return s;
  }

  uniqueName(base) {
    let i = this.sheets.length + 1;
    while (this.byName(`${base}${i}`)) i++;
    return `${base}${i}`;
  }

  renameSheet(sheet, name) {
    name = name.trim().replace(/[\\/?*[\]:]/g, '').slice(0, 31);
    if (!name || (this.byName(name) && this.byName(name) !== sheet)) return false;
    this.transact('Renomear planilha', () => {
      this.snapWorkbook();
      const old = sheet.name;
      sheet.name = name;
      for (const s of this.sheets) for (const cell of s.cells.values()) {
        if (cell.f && cell.f.includes(old)) { cell.f = renameSheetInFormula(cell.f, old, name); cell.ast = null; }
      }
    });
    return true;
  }

  deleteSheet(sheet) {
    if (this.sheets.length <= 1) return false;
    this.transact('Excluir planilha', () => {
      this.snapWorkbook();
      const i = this.sheets.indexOf(sheet);
      this.sheets.splice(i, 1);
      this.active = Math.max(0, Math.min(this.active, this.sheets.length - 1));
    });
    return true;
  }

  duplicateSheet(sheet) {
    this.transact('Duplicar planilha', () => {
      this.snapWorkbook();
      const j = sheet.toJSON();
      delete j.id;
      const s = Sheet.from(j);
      let n = 2;
      while (this.byName(`${sheet.name} (${n})`)) n++;
      s.name = `${sheet.name} (${n})`;
      const i = this.sheets.indexOf(sheet) + 1;
      this.sheets.splice(i, 0, s);
      this.active = i;
    });
  }

  moveSheet(from, to) {
    this.transact('Mover planilha', () => {
      this.snapWorkbook();
      const [s] = this.sheets.splice(from, 1);
      this.sheets.splice(to, 0, s);
      this.active = to;
    });
  }

  // Ordena as linhas de um intervalo por uma ou mais colunas.
  sortRange(sheet, rg, keys) {
    this.transact('Classificar', () => {
      this.snapSheet(sheet);
      const rows = [];
      for (let r = rg.r1; r <= rg.r2; r++) {
        const cells = [];
        for (let c = rg.c1; c <= rg.c2; c++) cells.push(sheet.get(r, c));
        rows.push({ r, cells, keys: keys.map((k) => this.value(sheet, r, k.col)) });
      }
      rows.sort((a, b) => {
        for (let i = 0; i < keys.length; i++) {
          const x = a.keys[i], y = b.keys[i];
          const ex = x === null || x === '', ey = y === null || y === '';
          if (ex !== ey) return ex ? 1 : -1;
          const c = compare(x, y);
          if (c) return keys[i].desc ? -c : c;
        }
        return a.r - b.r;
      });
      rows.forEach((row, i) => {
        const tr = rg.r1 + i;
        row.cells.forEach((cell, j) => {
          const c = rg.c1 + j;
          let n = cell ? clone(stripCell(cell)) : null;
          if (n?.f && row.r !== tr) n.f = shiftFormula(n.f, tr - row.r, 0);
          sheet.set(tr, c, n);
        });
      });
    });
  }

  // ---- Arquivo ---------------------------------------------------------------

  toJSON() {
    return { app: 'Lumina Planilhas', format: 1, ...this.snapshot() };
  }

  static from(j) {
    if (j.format !== 1 || !Array.isArray(j.sheets)) throw new Error('Arquivo .lumsheet inválido.');
    const wb = new Workbook();
    wb.restoreSnapshot(j);
    wb.recalc();
    return wb;
  }
}

function stripCell(c) {
  if (!c) return null;
  const out = { v: c.v ?? null };
  if (c.f) out.f = c.f;
  if (c.s && Object.keys(c.s).length) out.s = c.s;
  if (c.note) out.note = c.note;
  if (c.forceText) out.forceText = true;
  return out;
}

export function adjustRange(rg, isRow, at, count) {
  if (!rg) return null;
  const n = { ...rg };
  const a = isRow ? 'r1' : 'c1', b = isRow ? 'r2' : 'c2';
  if (count > 0) {
    if (n[a] >= at) n[a] += count;
    if (n[b] >= at) n[b] += count;
    return n;
  }
  const del = -count, end = at + del - 1;
  if (n[a] >= at && n[b] <= end) return null;
  if (n[a] > end) n[a] -= del; else if (n[a] >= at) n[a] = at;
  if (n[b] > end) n[b] -= del; else if (n[b] >= at) n[b] = at - 1;
  return n[b] >= n[a] ? n : null;
}

export const rangeText = (rg) => (rg.r1 === rg.r2 && rg.c1 === rg.c2 ? `${colName(rg.c1)}${rg.r1 + 1}` : `${colName(rg.c1)}${rg.r1 + 1}:${colName(rg.c2)}${rg.r2 + 1}`);
export { stripCell, clone };
