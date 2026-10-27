// Tipos de valor, erros, coerções e utilidades do motor de cálculo.

export class Err {
  constructor(code, msg = '') {
    this.code = code;
    this.msg = msg;
  }
  toString() { return this.code; }
}

export const ERR = {
  DIV0: '#DIV/0!', NA: '#N/A', NAME: '#NAME?', REF: '#REF!', VALUE: '#VALUE!', NUM: '#NUM!', NULL: '#NULL!',
  SPILL: '#SPILL!', CIRC: '#CIRC!', CALC: '#CALC!',
};
export const ERROR_CODES = Object.values(ERR);

export const err = (code, msg) => new Err(code, msg);
export const isErr = (v) => v instanceof Err;
export const isMatrix = (v) => Array.isArray(v);
export const isBlank = (v) => v === null || v === undefined || v === '';

// Referência a um intervalo (avaliada sob demanda).
export class RangeRef {
  constructor(sheet, r1, c1, r2, c2, ctx) {
    this.sheet = sheet;
    this.r1 = Math.min(r1, r2); this.c1 = Math.min(c1, c2);
    this.r2 = Math.max(r1, r2); this.c2 = Math.max(c1, c2);
    this.ctx = ctx;
  }
  get rows() { return this.r2 - this.r1 + 1; }
  get cols() { return this.c2 - this.c1 + 1; }
  matrix() {
    if (!this._m) this._m = this.ctx.rangeValues(this);
    return this._m;
  }
  // Itera somente células existentes (rápido para colunas inteiras)
  values() { return this.matrix().flat(); }
}

export const toMatrix = (v) => {
  if (v instanceof RangeRef) return v.matrix();
  if (isMatrix(v)) return isMatrix(v[0]) ? v : [v];
  return [[v]];
};

export const flat = (args) => {
  const out = [];
  for (const a of args) {
    if (a instanceof RangeRef) out.push(...a.values());
    else if (isMatrix(a)) for (const row of toMatrix(a)) out.push(...row);
    else out.push(a);
  }
  return out;
};

// Converte resultado em escalar (interseção implícita: canto superior esquerdo).
export const scalar = (v) => {
  if (v instanceof RangeRef) return v.matrix()[0]?.[0] ?? null;
  if (isMatrix(v)) return toMatrix(v)[0]?.[0] ?? null;
  return v;
};

// ---- Números e textos no padrão brasileiro ---------------------------------

export function parseNumberText(s) {
  if (typeof s !== 'string') return null;
  let t = s.trim().replace(/ /g, ' ');
  if (!t) return null;
  let pct = false;
  if (t.endsWith('%')) { pct = true; t = t.slice(0, -1).trim(); }
  t = t.replace(/^R\$\s?|^US\$\s?|^\$\s?|^€\s?/, '');
  let neg = false;
  if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }
  let n = null;
  if (/^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) n = +t.replace(/\./g, '').replace(',', '.');
  else if (/^[+-]?\d+,\d+$/.test(t) || /^[+-]?,\d+$/.test(t)) n = +t.replace(',', '.');
  else if (/^[+-]?(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(t)) n = +t;
  if (n === null || Number.isNaN(n)) return null;
  if (neg) n = -n;
  return pct ? n / 100 : n;
}

const genFmt = new Intl.NumberFormat('pt-BR', { maximumSignificantDigits: 15, useGrouping: false });
export function numberToText(n) {
  if (!Number.isFinite(n)) return n > 0 ? '∞' : '-∞';
  if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
  const a = Math.abs(n);
  if (a !== 0 && (a < 1e-9 || a >= 1e15)) return n.toExponential(5).replace('.', ',').replace('e+', 'E+').replace('e-', 'E-');
  return genFmt.format(+n.toPrecision(15));
}

export function toNum(v) {
  if (typeof v === 'number') return v;
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v instanceof Err) return v;
  if (typeof v === 'string') {
    const n = parseNumberText(v);
    if (n !== null) return n;
    const d = parseDateText(v);
    if (d) return d.serial;
    return err(ERR.VALUE, `"${v}" não é um número`);
  }
  if (v instanceof RangeRef || isMatrix(v)) return toNum(scalar(v));
  return err(ERR.VALUE);
}

export function toStr(v) {
  if (v === null || v === undefined) return '';
  if (typeof v === 'string') return v;
  if (typeof v === 'number') return numberToText(v);
  if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
  if (v instanceof Err) return v;
  if (v instanceof RangeRef || isMatrix(v)) return toStr(scalar(v));
  return String(v);
}

export function toBool(v) {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null || v === undefined || v === '') return false;
  if (v instanceof Err) return v;
  if (typeof v === 'string') {
    const u = v.trim().toUpperCase();
    if (u === 'TRUE' || u === 'VERDADEIRO') return true;
    if (u === 'FALSE' || u === 'FALSO') return false;
    return err(ERR.VALUE);
  }
  return toBool(scalar(v));
}

// Ordem do Excel: números < textos < lógicos; vazio conta como 0 / "".
const typeRank = (v) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : v === null || v === undefined ? -1 : 3);
const collator = new Intl.Collator('pt-BR', { sensitivity: 'base', numeric: false });
export function compare(a, b) {
  if (a === null || a === undefined) a = typeof b === 'string' ? '' : typeof b === 'boolean' ? false : 0;
  if (b === null || b === undefined) b = typeof a === 'string' ? '' : typeof a === 'boolean' ? false : 0;
  const ra = typeRank(a), rb = typeRank(b);
  if (ra !== rb) return ra - rb;
  if (typeof a === 'string') return collator.compare(a, b);
  if (a === b) return 0;
  return a < b ? -1 : 1;
}

// ---- Critérios (CONT.SE, SOMASE…) -----------------------------------------

export function criteria(c) {
  if (c instanceof Err) return () => false;
  if (typeof c === 'number' || typeof c === 'boolean') return (v) => compare(v, c) === 0 && typeRank(v) === typeRank(c);
  const s = toStr(c);
  const m = /^(<=|>=|<>|=|<|>)?(.*)$/s.exec(s);
  const op = m[1] || '=';
  const rhs = m[2];
  const num = parseNumberText(rhs);
  const dt = num === null ? parseDateText(rhs) : null;
  const target = num !== null ? num : dt ? dt.serial : rhs;
  if (typeof target === 'number') {
    return (v) => {
      if (typeof v !== 'number') {
        if (op === '<>') return true;
        return false;
      }
      return op === '=' ? v === target : op === '<>' ? v !== target : op === '<' ? v < target : op === '>' ? v > target : op === '<=' ? v <= target : v >= target;
    };
  }
  if (op === '=' || op === '<>') {
    if (rhs === '' && op === '=') return (v) => isBlank(v);
    if (rhs === '' && op === '<>') return (v) => !isBlank(v);
    const re = wildcard(rhs);
    return op === '=' ? (v) => re.test(toStr(v)) : (v) => !re.test(toStr(v));
  }
  return (v) => {
    if (typeof v !== 'string') return false;
    const c2 = collator.compare(v, rhs);
    return op === '<' ? c2 < 0 : op === '>' ? c2 > 0 : op === '<=' ? c2 <= 0 : c2 >= 0;
  };
}

export function wildcard(pattern, full = true) {
  let src = '';
  for (let i = 0; i < pattern.length; i++) {
    const ch = pattern[i];
    if (ch === '~' && i + 1 < pattern.length) { src += pattern[++i].replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); continue; }
    if (ch === '*') src += '.*';
    else if (ch === '?') src += '.';
    else src += ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }
  return new RegExp(full ? `^${src}$` : src, 'is');
}

// ---- Datas (número de série estilo Excel) ----------------------------------

export const EPOCH = 25569; // 01/01/1970
export const serialFromParts = (y, m, d, hh = 0, mm = 0, ss = 0) => Date.UTC(y, m - 1, d, hh, mm, ss) / 864e5 + EPOCH;
export const dateFromSerial = (s) => new Date(Math.round((s - EPOCH) * 864e5));
export const todaySerial = () => {
  const n = new Date();
  return serialFromParts(n.getFullYear(), n.getMonth() + 1, n.getDate());
};
export const nowSerial = () => {
  const n = new Date();
  return serialFromParts(n.getFullYear(), n.getMonth() + 1, n.getDate(), n.getHours(), n.getMinutes(), n.getSeconds());
};

// Reconhece dd/mm/aaaa, dd/mm, aaaa-mm-dd e horas hh:mm(:ss).
export function parseDateText(s) {
  if (typeof s !== 'string') return null;
  const t = s.trim();
  let m = /^(\d{1,2})[/.-](\d{1,2})(?:[/.-](\d{2,4}))?(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (m) {
    const d = +m[1], mo = +m[2];
    let y = m[3] ? +m[3] : new Date().getFullYear();
    if (y < 100) y += y < 50 ? 2000 : 1900;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return null;
    const hasTime = m[4] !== undefined;
    return { serial: serialFromParts(y, mo, d, +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)), fmt: hasTime ? 'datetime' : 'date' };
  }
  m = /^(\d{4})-(\d{1,2})-(\d{1,2})(?:[T\s](\d{1,2}):(\d{2})(?::(\d{2}))?)?$/.exec(t);
  if (m) return { serial: serialFromParts(+m[1], +m[2], +m[3], +(m[4] || 0), +(m[5] || 0), +(m[6] || 0)), fmt: m[4] ? 'datetime' : 'date' };
  m = /^(\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
  if (m && +m[1] < 48 && +m[2] < 60) return { serial: (+m[1] * 3600 + +m[2] * 60 + +(m[3] || 0)) / 86400, fmt: 'time' };
  return null;
}

// ---- Endereços --------------------------------------------------------------

export function colName(c) {
  let s = '';
  c += 1;
  while (c > 0) {
    const r = (c - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    c = Math.floor((c - 1) / 26);
  }
  return s;
}

export function colIndex(name) {
  let n = 0;
  for (const ch of name.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}

export const addr = (r, c) => `${colName(c)}${r + 1}`;

export function parseAddr(a) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(a.trim());
  if (!m) return null;
  return { r: +m[2] - 1, c: colIndex(m[1]) };
}

export function quoteSheet(name) {
  return /^[A-Za-z_À-ɏ][\wÀ-ɏ.]*$/.test(name) && !/^[A-Za-z]{1,3}\d+$/.test(name) ? name : `'${name.replace(/'/g, "''")}'`;
}
