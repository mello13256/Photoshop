// Executor de macros do Lumina Planilhas (Web Worker isolado: sem acesso à página, rede ou arquivos).
'use strict';

// Bloqueia rede e carregamento de scripts dentro da macro
for (const k of ['fetch', 'XMLHttpRequest', 'WebSocket', 'EventSource', 'importScripts', 'indexedDB', 'caches', 'BroadcastChannel']) {
  try { Object.defineProperty(self, k, { value: undefined, configurable: false, writable: false }); } catch { /* ignora */ }
}

let book = null; // { sheets: [{ name, cells: Map("r,c" -> { v, f }) }], active, sel }
let ops = [];
let pending = null;

const colIndex = (s) => { let c = 0; for (const ch of s.toUpperCase()) c = c * 26 + ch.charCodeAt(0) - 64; return c - 1; };
const colName = (c) => { let s = ''; for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s; return s; };

function sheetOf(name) {
  if (name === undefined || name === null) return book.sheets[book.active];
  const s = book.sheets.find((x) => x.name.toLowerCase() === String(name).toLowerCase());
  if (!s) throw new Error(`Planilha não encontrada: ${name}`);
  return s;
}

// "A1", "B2:C10", "Planilha2!A1:B3", "A:A", "3:3"
function parseRef(ref) {
  if (typeof ref !== 'string') throw new Error('Use uma referência em texto, como "A1" ou "A1:B10".');
  let sheet;
  let r = ref.trim();
  const bang = r.lastIndexOf('!');
  if (bang > 0) { sheet = r.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"); r = r.slice(bang + 1); }
  const s = sheetOf(sheet);
  const one = (x) => {
    const m = /^\$?([A-Za-z]{1,3})?\$?(\d+)?$/.exec(x);
    if (!m || (!m[1] && !m[2])) throw new Error(`Referência inválida: ${ref}`);
    return { c: m[1] ? colIndex(m[1]) : null, r: m[2] ? +m[2] - 1 : null };
  };
  const [a, b] = r.split(':');
  const pa = one(a), pb = b ? one(b) : pa;
  const used = usedOf(s);
  return {
    s,
    r1: pa.r ?? 0, c1: pa.c ?? 0,
    r2: pb.r ?? (pa.r === null ? used.r : pa.r), c2: pb.c ?? (pa.c === null ? used.c : pa.c),
  };
}

function usedOf(s) {
  let r = 0, c = 0;
  for (const k of s.cells.keys()) { const [rr, cc] = k.split(',').map(Number); if (rr > r) r = rr; if (cc > c) c = cc; }
  return { r, c };
}

function getValue(s, r, c) {
  const cell = s.cells.get(`${r},${c}`);
  return cell ? cell.v ?? null : null;
}

function setValue(s, r, c, value) {
  let v = value;
  if (v instanceof Date) v = v.getTime() / 86400000 + 25569 - v.getTimezoneOffset() / 1440;
  if (v !== null && v !== undefined && typeof v === 'object') v = String(v);
  if (typeof v === 'number' && !Number.isFinite(v)) v = '#NÚM!';
  const isFormula = typeof v === 'string' && v.startsWith('=') && v.length > 1;
  if (v === undefined || v === null || v === '') s.cells.delete(`${r},${c}`);
  else s.cells.set(`${r},${c}`, isFormula ? { v: null, f: v.slice(1) } : { v });
  ops.push({ op: 'set', sheet: s.name, r, c, value: v ?? null });
}

const api = {
  cel(ref) { const g = parseRef(ref); return getValue(g.s, g.r1, g.c1); },
  definir(ref, valor) {
    const g = parseRef(ref);
    if (Array.isArray(valor)) return api.definirIntervalo(ref, valor);
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) setValue(g.s, r, c, valor);
    return valor;
  },
  formula(ref, f) { return api.definir(ref, String(f).startsWith('=') ? f : `=${f}`); },
  intervalo(ref) {
    const g = parseRef(ref);
    const out = [];
    for (let r = g.r1; r <= g.r2; r++) { const row = []; for (let c = g.c1; c <= g.c2; c++) row.push(getValue(g.s, r, c)); out.push(row); }
    return out;
  },
  definirIntervalo(ref, matriz) {
    const g = parseRef(ref);
    if (!Array.isArray(matriz)) throw new Error('definirIntervalo espera uma matriz, como [[1, 2], [3, 4]].');
    const rows = Array.isArray(matriz[0]) ? matriz : matriz.map((x) => [x]);
    rows.forEach((row, i) => row.forEach((v, j) => setValue(g.s, g.r1 + i, g.c1 + j, v)));
  },
  limpar(ref) { const g = parseRef(ref); for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) setValue(g.s, r, c, null); },
  formatar(ref, estilo = {}) { const g = parseRef(ref); ops.push({ op: 'format', sheet: g.s.name, range: { r1: g.r1, c1: g.c1, r2: g.r2, c2: g.c2 }, style: estilo }); },
  ultimaLinha(col, planilha) {
    const s = sheetOf(planilha);
    let max = -1;
    const c = col === undefined ? null : typeof col === 'number' ? col - 1 : colIndex(col);
    for (const k of s.cells.keys()) { const [rr, cc] = k.split(',').map(Number); if ((c === null || cc === c) && rr > max) max = rr; }
    return max + 1;
  },
  ultimaColuna(linha, planilha) {
    const s = sheetOf(planilha);
    let max = -1;
    for (const k of s.cells.keys()) { const [rr, cc] = k.split(',').map(Number); if ((linha === undefined || rr === linha - 1) && cc > max) max = cc; }
    return max + 1;
  },
  endereco(linha, coluna) { return `${colName(coluna - 1)}${linha}`; },
  coluna(n) { return colName(n - 1); },
  planilhas() { return book.sheets.map((s) => s.name); },
  ativar(nome) { const i = book.sheets.indexOf(sheetOf(nome)); book.active = i; ops.push({ op: 'activate', sheet: book.sheets[i].name }); },
  novaPlanilha(nome) {
    const name = String(nome || `Planilha${book.sheets.length + 1}`);
    if (!book.sheets.some((s) => s.name.toLowerCase() === name.toLowerCase())) book.sheets.push({ name, cells: new Map() });
    ops.push({ op: 'addSheet', sheet: name });
    return name;
  },
  log(...args) { self.postMessage({ type: 'log', text: args.map(fmt).join(' ') }); },
  alerta(msg) { self.postMessage({ type: 'alert', text: fmt(msg) }); },
  async recalcular() {
    const batch = ops;
    ops = [];
    self.postMessage({ type: 'flush', ops: batch });
    const snap = await new Promise((resolve) => { pending = resolve; });
    load(snap);
  },
  hoje() { const d = new Date(); return Math.floor((d.getTime() - d.getTimezoneOffset() * 60000) / 86400000) + 25569; },
  data(ano, mes, dia) { return Date.UTC(ano, mes - 1, dia) / 86400000 + 25569; },
};
// Nomes em inglês
Object.assign(api, {
  get: api.cel, set: api.definir, range: api.intervalo, setRange: api.definirIntervalo, clear: api.limpar, format: api.formatar,
  lastRow: api.ultimaLinha, lastColumn: api.ultimaColuna, address: api.endereco, sheets: api.planilhas, activate: api.ativar,
  addSheet: api.novaPlanilha, alert: api.alerta, recalc: api.recalcular, celula: api.cel, valor: api.cel,
});

function fmt(v) {
  if (v === null || v === undefined) return String(v);
  if (typeof v === 'object') { try { return JSON.stringify(v); } catch { return String(v); } }
  return String(v);
}

function load(snap) {
  book = {
    active: snap.active,
    sel: snap.sel,
    sheets: snap.sheets.map((s) => ({ name: s.name, cells: new Map(s.cells) })),
  };
}

self.onmessage = async (e) => {
  const msg = e.data;
  if (msg.type === 'flushed') { const p = pending; pending = null; p?.(msg.snapshot); return; }
  if (msg.type !== 'run') return;
  load(msg.snapshot);
  ops = [];
  const names = Object.keys(api);
  const selecao = msg.snapshot.sel;
  try {
    const AsyncFunction = Object.getPrototypeOf(async () => {}).constructor;
    const fn = new AsyncFunction(...names, 'selecao', 'selection', 'console', `"use strict";\n${msg.code}`);
    const consoleShim = { log: api.log, info: api.log, warn: api.log, error: api.log };
    await fn(...names.map((n) => api[n]), selecao, selecao, consoleShim);
    self.postMessage({ type: 'done', ops });
  } catch (err) {
    const line = /<anonymous>:(\d+)/.exec(err.stack || '')?.[1];
    self.postMessage({ type: 'error', ops, message: err.message, line: line ? +line - 3 : null });
  }
};
