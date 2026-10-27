// Fórmulas: análise léxica, sintática, avaliação e ajuste de referências.
import {
  Err, ERR, err, isErr, isMatrix, RangeRef, toNum, toStr, compare, toMatrix, colIndex, colName, quoteSheet, ERROR_CODES,
} from './values.js';
import { FUNCS, resolveName } from './functions.js';

export const MAX_ROW = 1048575;
export const MAX_COL = 16383;

// ---------------------------------------------------------------------------
// Tokens

const SHEET = "(?:'((?:[^']|'')+)'|([A-Za-z_\\u00C0-\\u024F][\\w\\u00C0-\\u024F.]*))!";
const CELL = '(\\$?)([A-Za-z]{1,3})(\\$?)(\\d+)';
const REF_RE = new RegExp(`(?:${SHEET})?(?:${CELL}(?::${CELL})?(#)?|(\\$?)([A-Za-z]{1,3}):(\\$?)([A-Za-z]{1,3})|(\\$?)(\\d+):(\\$?)(\\d+))`, 'y');
const IDENT_RE = /[A-Za-z_À-ɏ][\wÀ-ɏ.]*/y;
const NUM_RE = /(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?/y;
const OPS = ['<=', '>=', '<>', '+', '-', '*', '/', '^', '&', '=', '<', '>', '%'];

// Vírgula decimal (padrão brasileiro) quando a fórmula usa ";" como separador ou não tem funções.
function usesDecimalComma(src) {
  const bare = src.replace(/"(?:[^"]|"")*"/g, '').replace(/'(?:[^']|'')*'/g, '');
  return bare.includes(';') || !bare.includes('(');
}
const NUM_PT_RE = /(\d+(?:,\d+)?|,\d+)([eE][+-]?\d+)?/y;

export function tokenize(src) {
  const toks = [];
  let i = 0;
  const n = src.length;
  const ptDec = usesDecimalComma(src);
  while (i < n) {
    const ch = src[i];
    if (/\s/.test(ch)) { i++; continue; }
    const start = i;
    if (ch === '"') {
      let s = '';
      i++;
      while (i < n) {
        if (src[i] === '"') {
          if (src[i + 1] === '"') { s += '"'; i += 2; continue; }
          break;
        }
        s += src[i++];
      }
      if (i >= n) throw new SyntaxError('Texto sem aspas de fechamento');
      i++;
      toks.push({ t: 'str', v: s, s: start, e: i });
      continue;
    }
    if (ch === '#') {
      const code = ERROR_CODES.find((c) => src.startsWith(c, i) || src.toUpperCase().startsWith(c, i));
      if (code) { i += code.length; toks.push({ t: 'err', v: code, s: start, e: i }); continue; }
    }
    // Referências (inclusive com nome de planilha) — testadas antes de números por causa de 1:3
    REF_RE.lastIndex = i;
    let m = REF_RE.exec(src);
    if (m && m[0]) {
      const after = src[REF_RE.lastIndex] || '';
      if (!/[\w(.!\u00C0-\u024F]/.test(after)) {
        i = REF_RE.lastIndex;
        toks.push({ t: 'ref', v: refFromMatch(m), s: start, e: i });
        continue;
      }
    }
    if (/\d/.test(ch) || (ch === '.' && /\d/.test(src[i + 1] || '')) || (ptDec && ch === ',' && /\d/.test(src[i + 1] || '') && !/[\w)]/.test(src[i - 1] || ''))) {
      const re = ptDec && /^(\d+,\d|,\d)/.test(src.slice(i)) ? NUM_PT_RE : NUM_RE;
      re.lastIndex = i;
      m = re.exec(src);
      i = re.lastIndex;
      toks.push({ t: 'num', v: parseFloat(m[0].replace(',', '.')), s: start, e: i });
      continue;
    }
    IDENT_RE.lastIndex = i;
    m = IDENT_RE.exec(src);
    if (m) {
      i = IDENT_RE.lastIndex;
      let j = i;
      while (j < n && src[j] === ' ') j++;
      if (src[j] === '(') toks.push({ t: 'func', v: m[0], s: start, e: i });
      else toks.push({ t: 'name', v: m[0], s: start, e: i });
      continue;
    }
    if (ch === "'") throw new SyntaxError('Nome de planilha inválido');
    const op = OPS.find((o) => src.startsWith(o, i));
    if (op) { i += op.length; toks.push({ t: 'op', v: op, s: start, e: i }); continue; }
    if ('(),;{}'.includes(ch)) { i++; toks.push({ t: ch, v: ch, s: start, e: i }); continue; }
    throw new SyntaxError(`Caractere inesperado: ${ch}`);
  }
  return toks;
}

function refFromMatch(m) {
  const sheet = m[1] !== undefined ? m[1].replace(/''/g, "'") : m[2] !== undefined ? m[2] : null;
  if (m[4] !== undefined) {
    const a = { col: colIndex(m[4]), row: +m[6] - 1, colAbs: !!m[3], rowAbs: !!m[5] };
    if (m[8] !== undefined) {
      const b = { col: colIndex(m[8]), row: +m[10] - 1, colAbs: !!m[7], rowAbs: !!m[9] };
      return { sheet, kind: 'range', a, b };
    }
    return { sheet, kind: 'cell', a, spill: !!m[11] };
  }
  if (m[13] !== undefined) {
    return { sheet, kind: 'cols', a: { col: colIndex(m[13]), colAbs: !!m[12] }, b: { col: colIndex(m[15]), colAbs: !!m[14] } };
  }
  return { sheet, kind: 'rows', a: { row: +m[17] - 1, rowAbs: !!m[16] }, b: { row: +m[19] - 1, rowAbs: !!m[18] } };
}

export function refText(r) {
  const pre = r.sheet ? `${quoteSheet(r.sheet)}!` : '';
  const cell = (p) => `${p.colAbs ? '$' : ''}${colName(p.col)}${p.rowAbs ? '$' : ''}${p.row + 1}`;
  if (r.kind === 'cell') return `${pre}${cell(r.a)}${r.spill ? '#' : ''}`;
  if (r.kind === 'range') return `${pre}${cell(r.a)}:${cell(r.b)}`;
  if (r.kind === 'cols') return `${pre}${r.a.colAbs ? '$' : ''}${colName(r.a.col)}:${r.b.colAbs ? '$' : ''}${colName(r.b.col)}`;
  return `${pre}${r.a.rowAbs ? '$' : ''}${r.a.row + 1}:${r.b.rowAbs ? '$' : ''}${r.b.row + 1}`;
}

// ---------------------------------------------------------------------------
// Análise sintática (Pratt)

const BIN_PREC = { '=': 1, '<>': 1, '<': 1, '>': 1, '<=': 1, '>=': 1, '&': 2, '+': 3, '-': 3, '*': 4, '/': 4, '^': 5 };

export function parse(src) {
  const toks = tokenize(src);
  let p = 0;
  const peek = () => toks[p];
  const next = () => toks[p++];
  const expect = (t) => {
    const k = next();
    if (!k || k.t !== t) throw new SyntaxError(`Esperado "${t}"`);
    return k;
  };

  const primary = () => {
    const k = next();
    if (!k) throw new SyntaxError('Fórmula incompleta');
    switch (k.t) {
      case 'num': return { type: 'num', v: k.v };
      case 'str': return { type: 'str', v: k.v };
      case 'err': return { type: 'err', v: k.v };
      case 'ref': return refNode(k.v);
      case 'name': {
        const u = k.v.toUpperCase();
        if (u === 'TRUE' || u === 'VERDADEIRO') return { type: 'bool', v: true };
        if (u === 'FALSE' || u === 'FALSO') return { type: 'bool', v: false };
        return { type: 'name', name: k.v };
      }
      case 'func': {
        expect('(');
        const args = [];
        if (peek()?.t !== ')') {
          for (;;) {
            if (peek()?.t === ',' || peek()?.t === ';') args.push({ type: 'blank' });
            else if (peek()?.t === ')') { args.push({ type: 'blank' }); break; }
            else args.push(expr(0));
            const sepTok = peek();
            if (sepTok?.t === ',' || sepTok?.t === ';') { next(); continue; }
            break;
          }
        }
        expect(')');
        return { type: 'func', name: resolveName(k.v), raw: k.v, args };
      }
      case '(': {
        const e = expr(0);
        expect(')');
        return e;
      }
      case '{': {
        const rows = [[]];
        for (;;) {
          let neg = false;
          if (peek()?.t === 'op' && peek().v === '-') { next(); neg = true; }
          const v = next();
          if (!v) throw new SyntaxError('Matriz incompleta');
          let val;
          if (v.t === 'num') val = neg ? -v.v : v.v;
          else if (v.t === 'str') val = v.v;
          else if (v.t === 'name' && /^(TRUE|FALSE|VERDADEIRO|FALSO)$/i.test(v.v)) val = /^(TRUE|VERDADEIRO)$/i.test(v.v);
          else if (v.t === 'err') val = err(v.v);
          else throw new SyntaxError('Matriz constante inválida');
          rows[rows.length - 1].push(val);
          const s = next();
          if (s?.t === '}') break;
          if (s?.t === ';') rows.push([]);
          else if (s?.t !== ',') throw new SyntaxError('Matriz constante inválida');
        }
        return { type: 'arr', rows };
      }
      case 'op':
        if (k.v === '-' || k.v === '+') return { type: 'un', op: k.v, a: expr(6) };
        throw new SyntaxError(`Operador inesperado: ${k.v}`);
      default:
        throw new SyntaxError(`Símbolo inesperado: ${k.v}`);
    }
  };

  const expr = (minPrec) => {
    let left = primary();
    for (;;) {
      const k = peek();
      if (!k || k.t !== 'op') break;
      if (k.v === '%') { next(); left = { type: 'pct', a: left }; continue; }
      const prec = BIN_PREC[k.v];
      if (prec === undefined || prec < minPrec) break;
      next();
      const right = expr(prec + 1);
      left = { type: 'bin', op: k.v, a: left, b: right };
    }
    return left;
  };

  const ast = expr(0);
  if (p < toks.length) throw new SyntaxError(`Símbolo inesperado: ${toks[p].v}`);
  return ast;
}

function refNode(r) {
  if (r.kind === 'cell') return { type: 'ref', sheet: r.sheet, row: r.a.row, col: r.a.col, spill: r.spill };
  if (r.kind === 'range') return { type: 'range', sheet: r.sheet, r1: r.a.row, c1: r.a.col, r2: r.b.row, c2: r.b.col };
  if (r.kind === 'cols') return { type: 'range', sheet: r.sheet, r1: 0, c1: r.a.col, r2: MAX_ROW, c2: r.b.col, full: 'cols' };
  return { type: 'range', sheet: r.sheet, r1: r.a.row, c1: 0, r2: r.b.row, c2: MAX_COL, full: 'rows' };
}

// ---------------------------------------------------------------------------
// Avaliação

// ctx: { sheet, row, col, cellValue(sheet, r, c), rangeValues(range), spillRange(sheet, r, c), locals }
export function evaluate(node, ctx) {
  switch (node.type) {
    case 'num': case 'str': case 'bool': return node.v;
    case 'err': return err(node.v);
    case 'blank': return null;
    case 'ref':
      if (node.spill) return ctx.spillRange(node.sheet, node.row, node.col);
      return ctx.cellValue(node.sheet, node.row, node.col);
    case 'range':
      return new RangeRef(node.sheet, node.r1, node.c1, node.r2, node.c2, ctx);
    case 'name': {
      for (let s = ctx.locals; s; s = s.parent) if (s.vars.has(node.name.toUpperCase())) return s.vars.get(node.name.toUpperCase());
      const named = ctx.definedName?.(node.name);
      if (named !== undefined) return named;
      return err(ERR.NAME, `Nome desconhecido: ${node.name}`);
    }
    case 'arr': return node.rows.map((r) => r.slice());
    case 'un': {
      const v = evaluate(node.a, ctx);
      if (node.op === '+') return v;
      return lift1(v, (x) => { const n = toNum(x); return isErr(n) ? n : -n; });
    }
    case 'pct':
      return lift1(evaluate(node.a, ctx), (x) => { const n = toNum(x); return isErr(n) ? n : n / 100; });
    case 'bin':
      return lift2(evaluate(node.a, ctx), evaluate(node.b, ctx), BIN[node.op]);
    case 'func':
      return callFunction(node, ctx);
    default:
      return err(ERR.VALUE);
  }
}

const arith = (f) => (a, b) => {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  const x = toNum(a);
  if (isErr(x)) return x;
  const y = toNum(b);
  if (isErr(y)) return y;
  return f(x, y);
};
const cmp = (f) => (a, b) => {
  if (isErr(a)) return a;
  if (isErr(b)) return b;
  return f(compare(a, b));
};
const BIN = {
  '+': arith((x, y) => x + y),
  '-': arith((x, y) => x - y),
  '*': arith((x, y) => x * y),
  '/': arith((x, y) => (y === 0 ? err(ERR.DIV0) : x / y)),
  '^': arith((x, y) => { const r = x ** y; return Number.isFinite(r) ? r : err(ERR.NUM); }),
  '&': (a, b) => (isErr(a) ? a : isErr(b) ? b : toStr(a) + toStr(b)),
  '=': cmp((c) => c === 0), '<>': cmp((c) => c !== 0), '<': cmp((c) => c < 0), '>': cmp((c) => c > 0), '<=': cmp((c) => c <= 0), '>=': cmp((c) => c >= 0),
};

const isArr = (v) => v instanceof RangeRef || isMatrix(v);

export function lift1(v, f) {
  if (!isArr(v)) return f(v);
  return toMatrix(v).map((row) => row.map(f));
}

// Operação elemento a elemento com expansão de linha/coluna únicas.
export function lift2(a, b, f) {
  if (!isArr(a) && !isArr(b)) return f(a, b);
  const A = toMatrix(a), B = toMatrix(b);
  const rows = Math.max(A.length, B.length), cols = Math.max(A[0]?.length || 0, B[0]?.length || 0);
  const get = (M, r, c) => {
    const rr = M.length === 1 ? 0 : r, cc = M[0].length === 1 ? 0 : c;
    return rr < M.length && cc < M[0].length ? M[rr][cc] : err(ERR.NA);
  };
  const out = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) row.push(f(get(A, r, c), get(B, r, c)));
    out.push(row);
  }
  return out;
}

function callFunction(node, ctx) {
  const def = FUNCS[node.name];
  if (!def) return err(ERR.NAME, `Função desconhecida: ${node.raw}`);
  const n = node.args.length;
  if (n < def.a[0] || n > def.a[1]) return err(ERR.VALUE, `${node.raw}: número de argumentos inválido`);
  if (def.lazy) return def.fn(node.args, ctx, evaluate);
  const args = node.args.map((a) => {
    if (a.type === 'ref' && !a.spill && def.refs) return new RangeRef(a.sheet, a.row, a.col, a.row, a.col, ctx);
    return evaluate(a, ctx);
  });
  if (def.ew) {
    const idx = args.findIndex((x) => isArr(x));
    if (idx >= 0) return elementwise(def, args, ctx);
  }
  if (!def.errs) for (const a of args) if (isErr(a)) return a;
  try {
    const r = def.fn(args, ctx);
    if (typeof r === 'number' && !Number.isFinite(r)) return err(ERR.NUM);
    return r === undefined ? null : r;
  } catch (e) {
    if (e instanceof Err) return e;
    console.warn(e);
    return err(ERR.VALUE, e.message);
  }
}

function elementwise(def, args, ctx) {
  const mats = args.map((a) => (isArr(a) ? toMatrix(a) : null));
  const rows = Math.max(...mats.map((m) => (m ? m.length : 1)));
  const cols = Math.max(...mats.map((m) => (m ? m[0].length : 1)));
  const out = [];
  for (let r = 0; r < rows; r++) {
    const row = [];
    for (let c = 0; c < cols; c++) {
      const a = args.map((x, i) => {
        const m = mats[i];
        if (!m) return x;
        const rr = m.length === 1 ? 0 : r, cc = m[0].length === 1 ? 0 : c;
        return m[rr]?.[cc] ?? err(ERR.NA);
      });
      const e = def.errs ? null : a.find(isErr);
      row.push(e || def.fn(a, ctx));
    }
    out.push(row);
  }
  return out;
}

// ---------------------------------------------------------------------------
// Reescrita de fórmulas (copiar/colar, inserir/excluir linhas, exportar)

function rewrite(src, fnRef, fnTok) {
  const toks = tokenize(src);
  let out = '';
  let last = 0;
  for (const t of toks) {
    let rep = null;
    if (t.t === 'ref') rep = fnRef?.(t.v);
    else if (fnTok) rep = fnTok(t, src.slice(t.s, t.e));
    if (rep === null || rep === undefined) continue;
    out += src.slice(last, t.s) + rep;
    last = t.e;
  }
  return out + src.slice(last);
}

const safeRewrite = (f, ...a) => {
  try { return rewrite(f, ...a); } catch { return f; }
};

// Desloca referências relativas (fórmula sem "=").
export function shiftFormula(f, dr, dc) {
  return safeRewrite(f, (r) => {
    const n = structuredClone(r);
    const mv = (p) => {
      if (p.row !== undefined && !p.rowAbs) p.row += dr;
      if (p.col !== undefined && !p.colAbs) p.col += dc;
      return (p.row === undefined || (p.row >= 0 && p.row <= MAX_ROW)) && (p.col === undefined || (p.col >= 0 && p.col <= MAX_COL));
    };
    const ok = mv(n.a) & (n.b ? mv(n.b) : true);
    return ok ? refText(n) : '#REF!';
  });
}

// Ajusta referências quando linhas/colunas são inseridas (count>0) ou excluídas (count<0).
export function adjustForStructure(f, { targetSheet, formulaSheet, axis, at, count }) {
  return safeRewrite(f, (r) => {
    const sheet = r.sheet ?? formulaSheet;
    if (sheet !== targetSheet) return null;
    const key = axis === 'row' ? 'row' : 'col';
    const n = structuredClone(r);
    const pts = [n.a, n.b].filter(Boolean).filter((p) => p[key] !== undefined);
    if (!pts.length) return null;
    if (count > 0) {
      for (const p of pts) if (p[key] >= at) p[key] += count;
      return refText(n);
    }
    const del = -count, end = at + del - 1;
    if (pts.length === 1) {
      const p = pts[0];
      if (p[key] >= at && p[key] <= end) return '#REF!';
      if (p[key] > end) p[key] -= del;
      return refText(n);
    }
    let [p1, p2] = pts;
    if (p1[key] > p2[key]) [p1, p2] = [p2, p1];
    if (p1[key] >= at && p2[key] <= end) return '#REF!';
    if (p1[key] > end) p1[key] -= del;
    else if (p1[key] >= at) p1[key] = at;
    if (p2[key] > end) p2[key] -= del;
    else if (p2[key] >= at) p2[key] = at - 1;
    return refText(n);
  });
}

export function renameSheetInFormula(f, oldName, newName) {
  return safeRewrite(f, (r) => (r.sheet === oldName ? refText({ ...r, sheet: newName }) : null));
}

// Converte para a sintaxe do Excel em inglês (nomes de função e separador ",").
export function toExcelFormula(f) {
  let depthBrace = 0;
  return safeRewrite(f, null, (t, text) => {
    if (t.t === 'func') return FUNCS[resolveName(t.v)] ? resolveName(t.v) : text;
    if (t.t === 'num' && text.includes(',')) return text.replace(',', '.');
    if (t.t === 'name' && /^VERDADEIRO$/i.test(t.v)) return 'TRUE';
    if (t.t === 'name' && /^FALSO$/i.test(t.v)) return 'FALSE';
    if (t.t === '{') depthBrace++;
    if (t.t === '}') depthBrace--;
    if (t.t === ';' && !depthBrace) return ',';
    return null;
  });
}

// Converte para a escrita em português (SOMA, PROCV, ";" e vírgula decimal).
export function toPortugueseFormula(f) {
  let depthBrace = 0;
  return safeRewrite(f, null, (t, text) => {
    if (t.t === 'func') { const d = FUNCS[resolveName(t.v)]; return d?.pt ? d.pt.split('|')[0] : text; }
    if (t.t === 'name' && /^TRUE$/i.test(t.v)) return 'VERDADEIRO';
    if (t.t === 'name' && /^FALSE$/i.test(t.v)) return 'FALSO';
    if (t.t === 'num' && text.includes('.')) return text.replace('.', ',');
    if (t.t === '{') depthBrace++;
    if (t.t === '}') depthBrace--;
    if (t.t === ',' && !depthBrace) return ';';
    return null;
  });
}

// Lista de referências na fórmula (para destacar com cores durante a edição).
export function formulaRefs(f) {
  try {
    return tokenize(f).filter((t) => t.t === 'ref').map((t) => ({ ...t.v, s: t.s, e: t.e }));
  } catch {
    return [];
  }
}

export { Err, ERR };
