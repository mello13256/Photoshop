// Biblioteca de funções (nomes em inglês e português, como no Excel).
import {
  Err, ERR, err, isErr, isMatrix, isBlank, RangeRef, toNum, toStr, toBool, toMatrix, flat, scalar, compare, criteria, wildcard,
  serialFromParts, dateFromSerial, todaySerial, nowSerial, parseDateText, numberToText,
} from './values.js';
import { formatCode } from './format.js';

export const FUNCS = {};
const ALIAS = {};

const strip = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();

export function resolveName(name) {
  const u = name.toUpperCase();
  if (FUNCS[u]) return u;
  return ALIAS[strip(name)] || ALIAS[u] || u;
}

// def(nomeEN, nomePT, categoria, sintaxe(pt), descrição, [min,max], fn, opções)
function def(name, pt, cat, syntax, desc, a, fn, opts = {}) {
  FUNCS[name] = { name, pt, cat, s: syntax, d: desc, a, fn, ...opts };
  if (pt) for (const p of pt.split('|')) ALIAS[strip(p)] = name;
  ALIAS[strip(name)] = name;
}

const MANY = Infinity;
const fail = (code, msg) => { throw err(code, msg); };
const num = (v) => { const n = toNum(v); if (isErr(n)) throw n; return n; };
const str = (v) => { const s = toStr(v); if (isErr(s)) throw s; return s; };
const bool = (v) => { const b = toBool(v); if (isErr(b)) throw b; return b; };
const int = (v) => Math.trunc(num(v));
const opt = (args, i, d) => (args[i] === undefined || args[i] === null ? d : args[i]);

// Números para funções estatísticas: em intervalos, ignora textos e lógicos.
function numbers(args, { includeA = false } = {}) {
  const out = [];
  for (const a of args) {
    if (a instanceof RangeRef || isMatrix(a)) {
      for (const v of flat([a])) {
        if (isErr(v)) throw v;
        if (typeof v === 'number') out.push(v);
        else if (includeA && typeof v === 'boolean') out.push(v ? 1 : 0);
        else if (includeA && typeof v === 'string') out.push(0);
      }
    } else if (a !== null && a !== undefined) {
      out.push(num(a));
    }
  }
  return out;
}

const sum = (xs) => xs.reduce((s, x) => s + x, 0);
const mean = (xs) => (xs.length ? sum(xs) / xs.length : fail(ERR.DIV0));
function variance(xs, sample) {
  if (xs.length < (sample ? 2 : 1)) fail(ERR.DIV0);
  const m = mean(xs);
  return sum(xs.map((x) => (x - m) ** 2)) / (xs.length - (sample ? 1 : 0));
}
const matrixOf = (v) => toMatrix(v);
const sameShape = (a, b) => a.length === b.length && a[0].length === b[0].length;

// Quando o critério é uma matriz, calcula para cada elemento (resultado espalhado).
function perCriterion(args, idx, fn) {
  const c = args[idx];
  if (!(c instanceof RangeRef || isMatrix(c))) return fn(args);
  const m = toMatrix(c);
  if (m.length === 1 && m[0].length === 1) return fn(args.map((x, i) => (i === idx ? m[0][0] : x)));
  return m.map((row) => row.map((v) => {
    try { return fn(args.map((x, i) => (i === idx ? v : x))); } catch (e) { if (e instanceof Err) return e; throw e; }
  }));
}

function ifsPairs(args, start) {
  const pairs = [];
  for (let i = start; i < args.length; i += 2) {
    const m = matrixOf(args[i]);
    pairs.push({ m, test: criteria(scalar(args[i + 1])) });
  }
  return pairs;
}
function ifsMatch(pairs, r, c) {
  return pairs.every((p) => p.test(p.m[r]?.[c] ?? null));
}

function roundTo(x, d, mode = 'round') {
  const f = 10 ** d;
  const v = x * f;
  const e = Number.EPSILON * Math.max(1, Math.abs(v)) * 4;
  let r;
  if (mode === 'up') r = Math.sign(v) * Math.ceil(Math.abs(v) - e);
  else if (mode === 'down') r = Math.sign(v) * Math.floor(Math.abs(v) + e);
  else r = Math.sign(v) * Math.round(Math.abs(v) + e);
  return r / f;
}

// ---------------------------------------------------------------------------
// Matemática

const M = 'Matemática';
def('SUM', 'SOMA', M, 'SOMA(núm1; [núm2]; …)', 'Soma todos os números.', [1, MANY], (a) => sum(numbers(a)));
def('PRODUCT', 'MULT', M, 'MULT(núm1; [núm2]; …)', 'Multiplica todos os números.', [1, MANY], (a) => numbers(a).reduce((p, x) => p * x, 1));
def('SUMSQ', 'SOMAQUAD', M, 'SOMAQUAD(núm1; …)', 'Soma dos quadrados.', [1, MANY], (a) => sum(numbers(a).map((x) => x * x)));
def('SUMPRODUCT', 'SOMARPRODUTO', M, 'SOMARPRODUTO(matriz1; [matriz2]; …)', 'Soma dos produtos de matrizes correspondentes.', [1, MANY], (a) => {
  const ms = a.map(matrixOf);
  if (!ms.every((m) => sameShape(m, ms[0]))) fail(ERR.VALUE);
  let s = 0;
  for (let r = 0; r < ms[0].length; r++) for (let c = 0; c < ms[0][0].length; c++) {
    let p = 1;
    for (const m of ms) { const v = m[r][c]; if (isErr(v)) throw v; p *= typeof v === 'number' ? v : typeof v === 'boolean' ? +v : 0; }
    s += p;
  }
  return s;
});
def('SUMIF', 'SOMASE', M, 'SOMASE(intervalo; critério; [intervalo_soma])', 'Soma as células que atendem a um critério.', [2, 3], (args) => perCriterion(args, 1, (a) => {
  const rng = matrixOf(a[0]), test = criteria(scalar(a[1])), s = a[2] !== undefined ? matrixOf(a[2]) : rng;
  let t = 0;
  rng.forEach((row, r) => row.forEach((v, c) => { if (test(v)) { const x = s[r]?.[c]; if (typeof x === 'number') t += x; } }));
  return t;
}), { errs: true });
def('SUMIFS', 'SOMASES', M, 'SOMASES(intervalo_soma; intervalo1; critério1; …)', 'Soma com vários critérios.', [3, MANY], (a) => {
  const s = matrixOf(a[0]), pairs = ifsPairs(a, 1);
  let t = 0;
  s.forEach((row, r) => row.forEach((x, c) => { if (typeof x === 'number' && ifsMatch(pairs, r, c)) t += x; }));
  return t;
}, { errs: true });
def('ROUND', 'ARRED', M, 'ARRED(núm; casas)', 'Arredonda para o número de casas decimais.', [1, 2], (a) => roundTo(num(a[0]), int(opt(a, 1, 0))), { ew: true });
def('ROUNDUP', 'ARREDONDAR.PARA.CIMA', M, 'ARREDONDAR.PARA.CIMA(núm; casas)', 'Arredonda afastando de zero.', [1, 2], (a) => roundTo(num(a[0]), int(opt(a, 1, 0)), 'up'), { ew: true });
def('ROUNDDOWN', 'ARREDONDAR.PARA.BAIXO', M, 'ARREDONDAR.PARA.BAIXO(núm; casas)', 'Arredonda em direção a zero.', [1, 2], (a) => roundTo(num(a[0]), int(opt(a, 1, 0)), 'down'), { ew: true });
def('INT', 'INT', M, 'INT(núm)', 'Arredonda para baixo até o inteiro.', [1, 1], (a) => Math.floor(num(a[0])), { ew: true });
def('TRUNC', 'TRUNCAR', M, 'TRUNCAR(núm; [casas])', 'Remove a parte decimal.', [1, 2], (a) => roundTo(num(a[0]), int(opt(a, 1, 0)), 'down'), { ew: true });
def('MOD', 'MOD', M, 'MOD(núm; divisor)', 'Resto da divisão.', [2, 2], (a) => { const d = num(a[1]); if (d === 0) fail(ERR.DIV0); const n = num(a[0]); return n - d * Math.floor(n / d); }, { ew: true });
def('QUOTIENT', 'QUOCIENTE', M, 'QUOCIENTE(numerador; denominador)', 'Parte inteira da divisão.', [2, 2], (a) => { const d = num(a[1]); if (d === 0) fail(ERR.DIV0); return Math.trunc(num(a[0]) / d); }, { ew: true });
def('ABS', 'ABS', M, 'ABS(núm)', 'Valor absoluto.', [1, 1], (a) => Math.abs(num(a[0])), { ew: true });
def('SIGN', 'SINAL', M, 'SINAL(núm)', 'Sinal do número (-1, 0 ou 1).', [1, 1], (a) => Math.sign(num(a[0])), { ew: true });
def('SQRT', 'RAIZ', M, 'RAIZ(núm)', 'Raiz quadrada.', [1, 1], (a) => { const n = num(a[0]); if (n < 0) fail(ERR.NUM); return Math.sqrt(n); }, { ew: true });
def('POWER', 'POTÊNCIA', M, 'POTÊNCIA(núm; expoente)', 'Eleva a uma potência.', [2, 2], (a) => num(a[0]) ** num(a[1]), { ew: true });
def('EXP', 'EXP', M, 'EXP(núm)', 'e elevado a núm.', [1, 1], (a) => Math.exp(num(a[0])), { ew: true });
def('LN', 'LN', M, 'LN(núm)', 'Logaritmo natural.', [1, 1], (a) => { const n = num(a[0]); if (n <= 0) fail(ERR.NUM); return Math.log(n); }, { ew: true });
def('LOG', 'LOG', M, 'LOG(núm; [base])', 'Logaritmo na base indicada (padrão 10).', [1, 2], (a) => { const n = num(a[0]), b = num(opt(a, 1, 10)); if (n <= 0 || b <= 0 || b === 1) fail(ERR.NUM); return Math.log(n) / Math.log(b); }, { ew: true });
def('LOG10', 'LOG10', M, 'LOG10(núm)', 'Logaritmo na base 10.', [1, 1], (a) => { const n = num(a[0]); if (n <= 0) fail(ERR.NUM); return Math.log10(n); }, { ew: true });
def('PI', 'PI', M, 'PI()', 'O número π.', [0, 0], () => Math.PI);
def('RAND', 'ALEATÓRIO', M, 'ALEATÓRIO()', 'Número aleatório entre 0 e 1.', [0, 0], () => Math.random(), { volatile: true });
def('RANDBETWEEN', 'ALEATÓRIOENTRE', M, 'ALEATÓRIOENTRE(inferior; superior)', 'Inteiro aleatório no intervalo.', [2, 2], (a) => { const lo = Math.ceil(num(a[0])), hi = Math.floor(num(a[1])); return lo + Math.floor(Math.random() * (hi - lo + 1)); }, { volatile: true });
const ceilTo = (x, s) => (s === 0 ? 0 : Math.ceil(x / s - 1e-12) * s);
const floorTo = (x, s) => (s === 0 ? 0 : Math.floor(x / s + 1e-12) * s);
def('CEILING', 'TETO|TETO.MAT|CEILING.MATH', M, 'TETO(núm; [múltiplo])', 'Arredonda para cima até o múltiplo.', [1, 2], (a) => ceilTo(num(a[0]), Math.abs(num(opt(a, 1, 1)))), { ew: true });
def('FLOOR', 'ARREDMULTB|ARREDMULTB.MAT|FLOOR.MATH', M, 'ARREDMULTB(núm; [múltiplo])', 'Arredonda para baixo até o múltiplo.', [1, 2], (a) => floorTo(num(a[0]), Math.abs(num(opt(a, 1, 1)))), { ew: true });
def('MROUND', 'MARRED', M, 'MARRED(núm; múltiplo)', 'Arredonda para o múltiplo mais próximo.', [2, 2], (a) => { const m = num(a[1]); return m === 0 ? 0 : Math.round(num(a[0]) / m) * m; }, { ew: true });
def('EVEN', 'PAR', M, 'PAR(núm)', 'Arredonda para o próximo par.', [1, 1], (a) => { const n = num(a[0]); const r = Math.ceil(Math.abs(n) / 2) * 2; return Math.sign(n) * r; }, { ew: true });
def('ODD', 'ÍMPAR', M, 'ÍMPAR(núm)', 'Arredonda para o próximo ímpar.', [1, 1], (a) => { const n = num(a[0]); let r = Math.ceil(Math.abs(n)); if (r % 2 === 0) r++; return (Math.sign(n) || 1) * r; }, { ew: true });
const gcd2 = (x, y) => { while (y) [x, y] = [y, x % y]; return x; };
def('GCD', 'MDC', M, 'MDC(núm1; …)', 'Máximo divisor comum.', [1, MANY], (a) => numbers(a).map(Math.trunc).reduce(gcd2));
def('LCM', 'MMC', M, 'MMC(núm1; …)', 'Mínimo múltiplo comum.', [1, MANY], (a) => numbers(a).map(Math.trunc).reduce((x, y) => (x && y ? Math.abs(x * y) / gcd2(x, y) : 0)));
const fact = (n) => { if (n < 0) fail(ERR.NUM); let r = 1; for (let i = 2; i <= n; i++) r *= i; return r; };
def('FACT', 'FATORIAL', M, 'FATORIAL(núm)', 'Fatorial.', [1, 1], (a) => fact(Math.trunc(num(a[0]))), { ew: true });
def('COMBIN', 'COMBIN', M, 'COMBIN(n; k)', 'Combinações de n elementos k a k.', [2, 2], (a) => { const n = int(a[0]), k = int(a[1]); if (k > n || k < 0) fail(ERR.NUM); let r = 1; for (let i = 1; i <= k; i++) r = (r * (n - k + i)) / i; return Math.round(r); }, { ew: true });
def('PERMUT', 'PERMUT', M, 'PERMUT(n; k)', 'Permutações de n elementos k a k.', [2, 2], (a) => { const n = int(a[0]), k = int(a[1]); if (k > n || k < 0) fail(ERR.NUM); let r = 1; for (let i = 0; i < k; i++) r *= n - i; return r; }, { ew: true });
for (const [en, pt, f, d] of [['SIN', 'SEN', Math.sin, 'Seno'], ['COS', 'COS', Math.cos, 'Cosseno'], ['TAN', 'TAN', Math.tan, 'Tangente'], ['ASIN', 'ASEN', Math.asin, 'Arco seno'], ['ACOS', 'ACOS', Math.acos, 'Arco cosseno'], ['ATAN', 'ATAN', Math.atan, 'Arco tangente']]) {
  def(en, pt, M, `${pt}(núm)`, `${d} (radianos).`, [1, 1], (a) => { const r = f(num(a[0])); if (Number.isNaN(r)) fail(ERR.NUM); return r; }, { ew: true });
}
def('ATAN2', 'ATAN2', M, 'ATAN2(x; y)', 'Arco tangente de x e y.', [2, 2], (a) => Math.atan2(num(a[1]), num(a[0])), { ew: true });
def('DEGREES', 'GRAUS', M, 'GRAUS(ângulo)', 'Radianos para graus.', [1, 1], (a) => (num(a[0]) * 180) / Math.PI, { ew: true });
def('RADIANS', 'RADIANOS', M, 'RADIANOS(ângulo)', 'Graus para radianos.', [1, 1], (a) => (num(a[0]) * Math.PI) / 180, { ew: true });
def('SUBTOTAL', 'SUBTOTAL', M, 'SUBTOTAL(função; ref1; …)', 'Subtotal (1=MÉDIA, 2=CONT.NÚM, 3=CONT.VALORES, 4=MÁX, 5=MÍN, 9=SOMA…).', [2, MANY], (a, ctx) => {
  const map = { 1: 'AVERAGE', 2: 'COUNT', 3: 'COUNTA', 4: 'MAX', 5: 'MIN', 6: 'PRODUCT', 7: 'STDEV', 8: 'STDEV.P', 9: 'SUM', 10: 'VAR', 11: 'VAR.P' };
  const f = FUNCS[map[int(a[0]) % 100]];
  if (!f) fail(ERR.VALUE);
  return f.fn(a.slice(1), ctx);
});

// ---------------------------------------------------------------------------
// Estatística

const ST = 'Estatística';
def('AVERAGE', 'MÉDIA', ST, 'MÉDIA(núm1; [núm2]; …)', 'Média aritmética.', [1, MANY], (a) => mean(numbers(a)));
def('AVERAGEA', 'MÉDIAA', ST, 'MÉDIAA(valor1; …)', 'Média contando textos como 0 e lógicos.', [1, MANY], (a) => mean(numbers(a, { includeA: true })));
def('AVERAGEIF', 'MÉDIASE', ST, 'MÉDIASE(intervalo; critério; [intervalo_média])', 'Média das células que atendem ao critério.', [2, 3], (args) => perCriterion(args, 1, (a) => {
  const rng = matrixOf(a[0]), test = criteria(scalar(a[1])), s = a[2] !== undefined ? matrixOf(a[2]) : rng;
  const xs = [];
  rng.forEach((row, r) => row.forEach((v, c) => { if (test(v) && typeof s[r]?.[c] === 'number') xs.push(s[r][c]); }));
  return mean(xs);
}), { errs: true });
def('AVERAGEIFS', 'MÉDIASES', ST, 'MÉDIASES(intervalo_média; intervalo1; critério1; …)', 'Média com vários critérios.', [3, MANY], (a) => {
  const s = matrixOf(a[0]), pairs = ifsPairs(a, 1), xs = [];
  s.forEach((row, r) => row.forEach((x, c) => { if (typeof x === 'number' && ifsMatch(pairs, r, c)) xs.push(x); }));
  return mean(xs);
}, { errs: true });
def('MEDIAN', 'MED', ST, 'MED(núm1; …)', 'Mediana.', [1, MANY], (a) => {
  const xs = numbers(a).sort((x, y) => x - y);
  if (!xs.length) fail(ERR.NUM);
  const m = xs.length >> 1;
  return xs.length % 2 ? xs[m] : (xs[m - 1] + xs[m]) / 2;
});
def('MODE', 'MODO|MODO.ÚNICO|MODE.SNGL', ST, 'MODO(núm1; …)', 'Valor mais frequente.', [1, MANY], (a) => {
  const cnt = new Map();
  let best = null, bc = 1;
  for (const x of numbers(a)) { const c = (cnt.get(x) || 0) + 1; cnt.set(x, c); if (c > bc) { bc = c; best = x; } }
  return best === null ? fail(ERR.NA) : best;
});
def('MIN', 'MÍNIMO', ST, 'MÍNIMO(núm1; …)', 'Menor valor.', [1, MANY], (a) => { const xs = numbers(a); return xs.length ? Math.min(...xs) : 0; });
def('MAX', 'MÁXIMO', ST, 'MÁXIMO(núm1; …)', 'Maior valor.', [1, MANY], (a) => { const xs = numbers(a); return xs.length ? Math.max(...xs) : 0; });
def('MINIFS', 'MÍNIMOSES', ST, 'MÍNIMOSES(intervalo_mín; intervalo1; critério1; …)', 'Menor valor com critérios.', [3, MANY], (a) => {
  const s = matrixOf(a[0]), pairs = ifsPairs(a, 1), xs = [];
  s.forEach((row, r) => row.forEach((x, c) => { if (typeof x === 'number' && ifsMatch(pairs, r, c)) xs.push(x); }));
  return xs.length ? Math.min(...xs) : 0;
}, { errs: true });
def('MAXIFS', 'MÁXIMOSES', ST, 'MÁXIMOSES(intervalo_máx; intervalo1; critério1; …)', 'Maior valor com critérios.', [3, MANY], (a) => {
  const s = matrixOf(a[0]), pairs = ifsPairs(a, 1), xs = [];
  s.forEach((row, r) => row.forEach((x, c) => { if (typeof x === 'number' && ifsMatch(pairs, r, c)) xs.push(x); }));
  return xs.length ? Math.max(...xs) : 0;
}, { errs: true });
def('COUNT', 'CONT.NÚM', ST, 'CONT.NÚM(valor1; …)', 'Conta células com números.', [1, MANY], (a) => {
  let n = 0;
  for (const x of a) {
    if (x instanceof RangeRef || isMatrix(x)) n += flat([x]).filter((v) => typeof v === 'number').length;
    else if (typeof x === 'number' || (typeof x === 'string' && !isErr(toNum(x)) && x !== '')) n++;
  }
  return n;
}, { errs: true });
def('COUNTA', 'CONT.VALORES', ST, 'CONT.VALORES(valor1; …)', 'Conta células não vazias.', [1, MANY], (a) => flat(a).filter((v) => !isBlank(v)).length, { errs: true });
def('COUNTBLANK', 'CONTAR.VAZIO', ST, 'CONTAR.VAZIO(intervalo)', 'Conta células vazias.', [1, 1], (a) => flat(a).filter((v) => isBlank(v)).length, { errs: true });
def('COUNTIF', 'CONT.SE', ST, 'CONT.SE(intervalo; critério)', 'Conta células que atendem ao critério.', [2, 2], (args) => perCriterion(args, 1, (a) => { const t = criteria(scalar(a[1])); return flat([a[0]]).filter((v) => t(v)).length; }), { errs: true });
def('COUNTIFS', 'CONT.SES', ST, 'CONT.SES(intervalo1; critério1; …)', 'Conta com vários critérios.', [2, MANY], (a) => {
  const pairs = ifsPairs(a, 0);
  const m = pairs[0].m;
  let n = 0;
  m.forEach((row, r) => row.forEach((_, c) => { if (ifsMatch(pairs, r, c)) n++; }));
  return n;
}, { errs: true });
def('LARGE', 'MAIOR', ST, 'MAIOR(matriz; k)', 'k-ésimo maior valor.', [2, 2], (a) => { const xs = numbers([a[0]]).sort((x, y) => y - x); const k = int(a[1]); if (k < 1 || k > xs.length) fail(ERR.NUM); return xs[k - 1]; });
def('SMALL', 'MENOR', ST, 'MENOR(matriz; k)', 'k-ésimo menor valor.', [2, 2], (a) => { const xs = numbers([a[0]]).sort((x, y) => x - y); const k = int(a[1]); if (k < 1 || k > xs.length) fail(ERR.NUM); return xs[k - 1]; });
def('RANK', 'ORDEM|ORDEM.EQ|RANK.EQ', ST, 'ORDEM(núm; ref; [ordem])', 'Posição do número na lista.', [2, 3], (a) => {
  const x = num(a[0]), xs = numbers([a[1]]), asc = num(opt(a, 2, 0)) !== 0;
  if (!xs.includes(x)) fail(ERR.NA);
  return 1 + xs.filter((v) => (asc ? v < x : v > x)).length;
});
function percentile(xs, p) {
  xs.sort((x, y) => x - y);
  if (!xs.length || p < 0 || p > 1) fail(ERR.NUM);
  const h = (xs.length - 1) * p, lo = Math.floor(h);
  return xs[lo] + (h - lo) * ((xs[lo + 1] ?? xs[lo]) - xs[lo]);
}
def('PERCENTILE', 'PERCENTIL|PERCENTIL.INC|PERCENTILE.INC', ST, 'PERCENTIL(matriz; k)', 'k-ésimo percentil (0 a 1).', [2, 2], (a) => percentile(numbers([a[0]]), num(a[1])));
def('QUARTILE', 'QUARTIL|QUARTIL.INC|QUARTILE.INC', ST, 'QUARTIL(matriz; quarto)', 'Quartil (0 a 4).', [2, 2], (a) => percentile(numbers([a[0]]), int(a[1]) / 4));
def('STDEV', 'DESVPAD|DESVPAD.A|STDEV.S', ST, 'DESVPAD(núm1; …)', 'Desvio padrão da amostra.', [1, MANY], (a) => Math.sqrt(variance(numbers(a), true)));
def('STDEV.P', 'DESVPAD.P|DESVPADP|STDEVP', ST, 'DESVPAD.P(núm1; …)', 'Desvio padrão da população.', [1, MANY], (a) => Math.sqrt(variance(numbers(a), false)));
def('VAR', 'VAR|VAR.A|VAR.S', ST, 'VAR(núm1; …)', 'Variância da amostra.', [1, MANY], (a) => variance(numbers(a), true));
def('VAR.P', 'VAR.P|VARP', ST, 'VAR.P(núm1; …)', 'Variância da população.', [1, MANY], (a) => variance(numbers(a), false));
def('GEOMEAN', 'MÉDIA.GEOMÉTRICA', ST, 'MÉDIA.GEOMÉTRICA(núm1; …)', 'Média geométrica.', [1, MANY], (a) => { const xs = numbers(a); if (xs.some((x) => x <= 0)) fail(ERR.NUM); return Math.exp(sum(xs.map(Math.log)) / xs.length); });
function pairsXY(ys, xs) {
  const Y = flat([ys]), X = flat([xs]);
  if (Y.length !== X.length) fail(ERR.NA);
  const out = [];
  for (let i = 0; i < Y.length; i++) if (typeof Y[i] === 'number' && typeof X[i] === 'number') out.push([X[i], Y[i]]);
  if (out.length < 2) fail(ERR.DIV0);
  return out;
}
function regression(ys, xs) {
  const P = pairsXY(ys, xs);
  const mx = mean(P.map((p) => p[0])), my = mean(P.map((p) => p[1]));
  let sxy = 0, sxx = 0, syy = 0;
  for (const [x, y] of P) { sxy += (x - mx) * (y - my); sxx += (x - mx) ** 2; syy += (y - my) ** 2; }
  if (sxx === 0) fail(ERR.DIV0);
  const slope = sxy / sxx;
  return { slope, intercept: my - slope * mx, r: sxy / Math.sqrt(sxx * syy) };
}
def('CORREL', 'CORREL', ST, 'CORREL(matriz1; matriz2)', 'Coeficiente de correlação.', [2, 2], (a) => regression(a[0], a[1]).r);
def('SLOPE', 'INCLINAÇÃO', ST, 'INCLINAÇÃO(ys; xs)', 'Inclinação da regressão linear.', [2, 2], (a) => regression(a[0], a[1]).slope);
def('INTERCEPT', 'INTERCEPÇÃO', ST, 'INTERCEPÇÃO(ys; xs)', 'Interseção da regressão linear.', [2, 2], (a) => regression(a[0], a[1]).intercept);
def('FORECAST', 'PREVISÃO|PREVISÃO.LINEAR|FORECAST.LINEAR', ST, 'PREVISÃO(x; ys; xs)', 'Previsão por regressão linear.', [3, 3], (a) => { const g = regression(a[1], a[2]); return g.intercept + g.slope * num(a[0]); });

// ---------------------------------------------------------------------------
// Lógica

const L = 'Lógica';
def('IF', 'SE', L, 'SE(teste; valor_se_verdadeiro; [valor_se_falso])', 'Retorna um valor se a condição for verdadeira e outro se for falsa.', [1, 3], (args, ctx, ev) => {
  const t = ev(args[0], ctx);
  if (isErr(t)) return t;
  if (t instanceof RangeRef || isMatrix(t)) {
    const yes = args[1] ? ev(args[1], ctx) : true, no = args[2] ? ev(args[2], ctx) : false;
    return toMatrix(t).map((row, r) => row.map((v, c) => {
      const b = toBool(v);
      if (isErr(b)) return b;
      const pick = b ? yes : no;
      return pick instanceof RangeRef || isMatrix(pick) ? toMatrix(pick)[r]?.[c] ?? err(ERR.NA) : pick;
    }));
  }
  const b = toBool(t);
  if (isErr(b)) return b;
  if (b) return args[1] && args[1].type !== 'blank' ? ev(args[1], ctx) : args[1] ? 0 : true;
  return args[2] && args[2].type !== 'blank' ? ev(args[2], ctx) : args[2] ? 0 : false;
}, { lazy: true });
def('IFS', 'SES', L, 'SES(teste1; valor1; [teste2; valor2]; …)', 'Retorna o valor da primeira condição verdadeira.', [2, MANY], (args, ctx, ev) => {
  for (let i = 0; i + 1 < args.length; i += 2) {
    const b = toBool(scalar(ev(args[i], ctx)));
    if (isErr(b)) return b;
    if (b) return ev(args[i + 1], ctx);
  }
  return err(ERR.NA);
}, { lazy: true });
def('IFERROR', 'SEERRO', L, 'SEERRO(valor; valor_se_erro)', 'Retorna outro valor quando há erro.', [2, 2], (args, ctx, ev) => {
  const v = ev(args[0], ctx);
  if (isErr(v)) return ev(args[1], ctx);
  if (isMatrix(v)) { const alt = ev(args[1], ctx); return v.map((row) => row.map((x) => (isErr(x) ? scalar(alt) : x))); }
  return v;
}, { lazy: true });
def('IFNA', 'SENÃODISP', L, 'SENÃODISP(valor; valor_se_nd)', 'Retorna outro valor quando o resultado é #N/D.', [2, 2], (args, ctx, ev) => {
  const v = ev(args[0], ctx);
  return isErr(v) && v.code === ERR.NA ? ev(args[1], ctx) : v;
}, { lazy: true });
const logicVals = (a) => {
  const out = [];
  for (const x of a) {
    if (x instanceof RangeRef || isMatrix(x)) { for (const v of flat([x])) { if (isErr(v)) throw v; if (typeof v === 'boolean' || typeof v === 'number') out.push(!!v); } }
    else out.push(bool(x));
  }
  if (!out.length) fail(ERR.VALUE);
  return out;
};
def('AND', 'E', L, 'E(lógico1; [lógico2]; …)', 'VERDADEIRO se todos os argumentos forem verdadeiros.', [1, MANY], (a) => logicVals(a).every(Boolean));
def('OR', 'OU', L, 'OU(lógico1; [lógico2]; …)', 'VERDADEIRO se algum argumento for verdadeiro.', [1, MANY], (a) => logicVals(a).some(Boolean));
def('XOR', 'OUEXCLUSIVO|XOU', L, 'OUEXCLUSIVO(lógico1; …)', 'Ou exclusivo.', [1, MANY], (a) => logicVals(a).filter(Boolean).length % 2 === 1);
def('NOT', 'NÃO', L, 'NÃO(lógico)', 'Inverte o valor lógico.', [1, 1], (a) => !bool(a[0]), { ew: true });
def('TRUE', 'VERDADEIRO', L, 'VERDADEIRO()', 'Valor lógico verdadeiro.', [0, 0], () => true);
def('FALSE', 'FALSO', L, 'FALSO()', 'Valor lógico falso.', [0, 0], () => false);
def('SWITCH', 'PARÂMETRO', L, 'PARÂMETRO(expressão; valor1; resultado1; …; [padrão])', 'Compara com uma lista de valores.', [3, MANY], (args, ctx, ev) => {
  const v = scalar(ev(args[0], ctx));
  if (isErr(v)) return v;
  let i = 1;
  for (; i + 1 < args.length; i += 2) if (compare(v, scalar(ev(args[i], ctx))) === 0) return ev(args[i + 1], ctx);
  return i < args.length ? ev(args[i], ctx) : err(ERR.NA);
}, { lazy: true });
def('LET', 'LET', L, 'LET(nome1; valor1; …; cálculo)', 'Atribui nomes a resultados intermediários.', [3, MANY], (args, ctx, ev) => {
  if (args.length % 2 === 0) return err(ERR.VALUE);
  const scope = { vars: new Map(), parent: ctx.locals };
  const c2 = { ...ctx, locals: scope };
  for (let i = 0; i + 1 < args.length; i += 2) {
    if (args[i].type !== 'name') return err(ERR.NAME);
    scope.vars.set(args[i].name.toUpperCase(), ev(args[i + 1], c2));
  }
  return ev(args[args.length - 1], c2);
}, { lazy: true });

// ---------------------------------------------------------------------------
// Pesquisa e referência

const LK = 'Pesquisa';
function matchEq(v, x) {
  if (typeof v === 'string' && typeof x === 'string' && /[*?~]/.test(v)) return wildcard(v).test(x);
  return compare(v, x) === 0 && (typeof v === typeof x || (v === null && x === null));
}
function approxIndex(list, v, desc = false) {
  let idx = -1;
  for (let i = 0; i < list.length; i++) {
    const c = compare(list[i], v);
    if (!desc && c <= 0 && typeof list[i] === typeof v) idx = i;
    else if (desc && c >= 0 && typeof list[i] === typeof v) idx = i;
    else if (!desc && c > 0) break;
  }
  return idx;
}
def('VLOOKUP', 'PROCV', LK, 'PROCV(valor; tabela; coluna; [procurar_intervalo])', 'Procura na primeira coluna e retorna da coluna indicada.', [3, 4], (a) => {
  const v = scalar(a[0]), t = matrixOf(a[1]), c = int(a[2]), approx = a[3] === undefined ? true : bool(a[3]);
  if (c < 1 || c > t[0].length) fail(ERR.REF);
  const col = t.map((r) => r[0]);
  const i = approx ? approxIndex(col, v) : col.findIndex((x) => matchEq(v, x));
  return i < 0 ? err(ERR.NA) : t[i][c - 1] ?? null;
});
def('HLOOKUP', 'PROCH', LK, 'PROCH(valor; tabela; linha; [procurar_intervalo])', 'Procura na primeira linha e retorna da linha indicada.', [3, 4], (a) => {
  const v = scalar(a[0]), t = matrixOf(a[1]), r = int(a[2]), approx = a[3] === undefined ? true : bool(a[3]);
  if (r < 1 || r > t.length) fail(ERR.REF);
  const row = t[0];
  const i = approx ? approxIndex(row, v) : row.findIndex((x) => matchEq(v, x));
  return i < 0 ? err(ERR.NA) : t[r - 1][i] ?? null;
});
function xfind(list, v, mode, search) {
  const order = [...list.keys()];
  if (search === -1) order.reverse();
  if (mode === 0 || mode === 2) {
    for (const i of order) if (mode === 2 ? matchEq(v, list[i]) : compare(v, list[i]) === 0 && (typeof v === typeof list[i])) return i;
    return -1;
  }
  let best = -1;
  for (const i of order) {
    const c = compare(list[i], v);
    if (c === 0 && typeof list[i] === typeof v) return i;
    if (mode === -1 && c < 0 && typeof list[i] === typeof v && (best < 0 || compare(list[i], list[best]) > 0)) best = i;
    if (mode === 1 && c > 0 && typeof list[i] === typeof v && (best < 0 || compare(list[i], list[best]) < 0)) best = i;
  }
  return best;
}
def('XLOOKUP', 'PROCX', LK, 'PROCX(valor; matriz_procura; matriz_retorno; [se_não_encontrado]; [modo_corresp]; [modo_pesquisa])', 'Procura moderna em qualquer direção.', [3, 6], (a) => {
  const v = scalar(a[0]), look = matrixOf(a[1]), ret = matrixOf(a[2]);
  const vertical = look.length > 1 || look[0].length === 1;
  const list = vertical ? look.map((r) => r[0]) : look[0];
  const i = xfind(list, v, int(opt(a, 4, 0)), int(opt(a, 5, 1)));
  if (i < 0) return a[3] !== undefined && a[3] !== null ? a[3] : err(ERR.NA);
  if (vertical) return ret[i] ? (ret[i].length === 1 ? ret[i][0] : [ret[i].slice()]) : err(ERR.REF);
  return ret.length === 1 ? ret[0][i] : ret.map((r) => [r[i]]);
}, { errs: true });
def('MATCH', 'CORRESP', LK, 'CORRESP(valor; matriz; [tipo])', 'Posição do valor em uma lista (1 = menor ou igual, 0 = exato, -1 = maior ou igual).', [2, 3], (a) => {
  const v = scalar(a[0]), list = flat([a[1]]), t = int(opt(a, 2, 1));
  const i = t === 0 ? list.findIndex((x) => matchEq(v, x)) : approxIndex(list, v, t < 0);
  return i < 0 ? err(ERR.NA) : i + 1;
});
def('XMATCH', 'CORRESPX', LK, 'CORRESPX(valor; matriz; [modo_corresp]; [modo_pesquisa])', 'Posição do valor (versão moderna).', [2, 4], (a) => {
  const i = xfind(flat([a[1]]), scalar(a[0]), int(opt(a, 2, 0)), int(opt(a, 3, 1)));
  return i < 0 ? err(ERR.NA) : i + 1;
});
def('LOOKUP', 'PROC', LK, 'PROC(valor; vetor_procura; [vetor_resultado])', 'Procura aproximada em vetor.', [2, 3], (a) => {
  const list = flat([a[1]]), res = a[2] !== undefined ? flat([a[2]]) : list;
  const i = approxIndex(list, scalar(a[0]));
  return i < 0 ? err(ERR.NA) : res[i];
});
def('INDEX', 'ÍNDICE', LK, 'ÍNDICE(matriz; linha; [coluna])', 'Valor na posição indicada.', [2, 3], (a) => {
  const m = matrixOf(a[0]);
  let r = int(a[1]), c = int(opt(a, 2, 0));
  if (m.length === 1 && a[2] === undefined) { c = r; r = 1; }
  if (r < 0 || c < 0 || r > m.length || c > m[0].length) fail(ERR.REF);
  if (r === 0 && c === 0) return m;
  if (r === 0) return m.map((row) => [row[c - 1]]);
  if (c === 0) return m[0].length === 1 ? m[r - 1][0] : [m[r - 1].slice()];
  return m[r - 1][c - 1];
});
def('CHOOSE', 'ESCOLHER', LK, 'ESCOLHER(índice; valor1; [valor2]; …)', 'Escolhe um valor da lista pelo índice.', [2, MANY], (args, ctx, ev) => {
  const i = toNum(scalar(ev(args[0], ctx)));
  if (isErr(i)) return i;
  const k = Math.trunc(i);
  if (k < 1 || k >= args.length) return err(ERR.VALUE);
  return ev(args[k], ctx);
}, { lazy: true });
def('ROW', 'LIN', LK, 'LIN([ref])', 'Número da linha.', [0, 1], (a, ctx) => (a[0] instanceof RangeRef ? a[0].r1 + 1 : ctx.row + 1), { refs: true });
def('COLUMN', 'COL', LK, 'COL([ref])', 'Número da coluna.', [0, 1], (a, ctx) => (a[0] instanceof RangeRef ? a[0].c1 + 1 : ctx.col + 1), { refs: true });
def('ROWS', 'LINS', LK, 'LINS(matriz)', 'Quantidade de linhas.', [1, 1], (a) => (a[0] instanceof RangeRef ? a[0].rows : matrixOf(a[0]).length), { refs: true });
def('COLUMNS', 'COLS', LK, 'COLS(matriz)', 'Quantidade de colunas.', [1, 1], (a) => (a[0] instanceof RangeRef ? a[0].cols : matrixOf(a[0])[0].length), { refs: true });
def('OFFSET', 'DESLOC', LK, 'DESLOC(ref; linhas; colunas; [altura]; [largura])', 'Referência deslocada.', [3, 5], (a, ctx) => {
  const ref = a[0];
  if (!(ref instanceof RangeRef)) fail(ERR.VALUE);
  const r1 = ref.r1 + int(a[1]), c1 = ref.c1 + int(a[2]);
  const h = a[3] !== undefined ? int(a[3]) : ref.rows, w = a[4] !== undefined ? int(a[4]) : ref.cols;
  if (r1 < 0 || c1 < 0 || h < 1 || w < 1) fail(ERR.REF);
  const rr = new RangeRef(ref.sheet, r1, c1, r1 + h - 1, c1 + w - 1, ctx);
  return h === 1 && w === 1 ? rr.matrix()[0][0] : rr;
}, { refs: true, volatile: true });
def('INDIRECT', 'INDIRETO', LK, 'INDIRETO(texto_ref)', 'Referência a partir de um texto.', [1, 2], (a, ctx) => ctx.indirect(str(a[0])), { volatile: true });
def('TRANSPOSE', 'TRANSPOR', LK, 'TRANSPOR(matriz)', 'Troca linhas por colunas.', [1, 1], (a) => { const m = matrixOf(a[0]); return m[0].map((_, c) => m.map((r) => r[c])); });

// ---------------------------------------------------------------------------
// Matrizes dinâmicas

const DA = 'Matrizes dinâmicas';
def('FILTER', 'FILTRO', DA, 'FILTRO(matriz; incluir; [se_vazio])', 'Filtra linhas (ou colunas) que atendem à condição.', [2, 3], (a) => {
  const m = matrixOf(a[0]), inc = matrixOf(a[1]);
  let out;
  if (inc.length === m.length && inc[0].length === 1) out = m.filter((_, r) => toBool(inc[r][0]) === true);
  else if (inc.length === 1 && inc[0].length === m[0].length) {
    const keep = inc[0].map((v) => toBool(v) === true);
    out = m.map((row) => row.filter((_, c) => keep[c]));
    if (!out[0].length) out = [];
  } else fail(ERR.VALUE);
  if (!out.length) return a[2] !== undefined ? a[2] : err(ERR.CALC, 'FILTRO sem resultados');
  return out;
}, { errs: true });
def('SORT', 'CLASSIFICAR', DA, 'CLASSIFICAR(matriz; [índice]; [ordem]; [por_coluna])', 'Ordena uma matriz.', [1, 4], (a) => {
  const m = matrixOf(a[0]).map((r) => r.slice()), idx = int(opt(a, 1, 1)) - 1, ord = num(opt(a, 2, 1)) < 0 ? -1 : 1, byCol = a[3] !== undefined && bool(a[3]);
  if (byCol) {
    const t = m[0].map((_, c) => m.map((r) => r[c]));
    t.sort((x, y) => ord * compare(x[idx], y[idx]));
    return t[0].map((_, c) => t.map((r) => r[c]));
  }
  return m.sort((x, y) => ord * compare(x[idx], y[idx]));
});
def('SORTBY', 'CLASSIFICARPOR', DA, 'CLASSIFICARPOR(matriz; por1; [ordem1]; …)', 'Ordena pela ordem de outra matriz.', [2, MANY], (a) => {
  const m = matrixOf(a[0]);
  const keys = [];
  for (let i = 1; i < a.length; i += 2) keys.push({ v: flat([a[i]]), o: num(opt(a, i + 1, 1)) < 0 ? -1 : 1 });
  const idx = m.map((_, i) => i);
  idx.sort((x, y) => { for (const k of keys) { const c = compare(k.v[x], k.v[y]); if (c) return k.o * c; } return 0; });
  return idx.map((i) => m[i]);
});
def('UNIQUE', 'ÚNICO', DA, 'ÚNICO(matriz; [por_coluna]; [exatamente_uma_vez])', 'Valores únicos.', [1, 3], (a) => {
  const m = matrixOf(a[0]), once = a[2] !== undefined && bool(a[2]);
  const key = (row) => JSON.stringify(row.map((v) => (typeof v === 'string' ? v.toLowerCase() : v)));
  const cnt = new Map();
  for (const r of m) cnt.set(key(r), (cnt.get(key(r)) || 0) + 1);
  const seen = new Set(), out = [];
  for (const r of m) {
    const k = key(r);
    if (seen.has(k) || (once && cnt.get(k) > 1)) continue;
    seen.add(k);
    out.push(r);
  }
  return out.length ? out : err(ERR.CALC);
});
def('SEQUENCE', 'SEQUÊNCIA', DA, 'SEQUÊNCIA(linhas; [colunas]; [início]; [incremento])', 'Gera uma sequência de números.', [1, 4], (a) => {
  const rows = int(a[0]), cols = int(opt(a, 1, 1)), start = num(opt(a, 2, 1)), step = num(opt(a, 3, 1));
  if (rows < 1 || cols < 1 || rows * cols > 1e6) fail(ERR.VALUE);
  return Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => start + (r * cols + c) * step));
});
def('RANDARRAY', 'MATRIZALEATÓRIA', DA, 'MATRIZALEATÓRIA([linhas]; [colunas]; [mín]; [máx]; [inteiro])', 'Matriz de números aleatórios.', [0, 5], (a) => {
  const rows = int(opt(a, 0, 1)), cols = int(opt(a, 1, 1)), lo = num(opt(a, 2, 0)), hi = num(opt(a, 3, 1)), whole = a[4] !== undefined && bool(a[4]);
  return Array.from({ length: rows }, () => Array.from({ length: cols }, () => (whole ? lo + Math.floor(Math.random() * (hi - lo + 1)) : lo + Math.random() * (hi - lo))));
}, { volatile: true });
def('TAKE', 'PEGAR', DA, 'PEGAR(matriz; linhas; [colunas])', 'Pega linhas/colunas do início (ou do fim, se negativo).', [2, 3], (a) => {
  let m = matrixOf(a[0]);
  const r = int(a[1]), c = a[2] !== undefined ? int(a[2]) : null;
  m = r >= 0 ? m.slice(0, r) : m.slice(r);
  if (c !== null) m = m.map((row) => (c >= 0 ? row.slice(0, c) : row.slice(c)));
  return m.length && m[0].length ? m : err(ERR.CALC);
});
def('DROP', 'DESCARTAR|SOLTAR', DA, 'DESCARTAR(matriz; linhas; [colunas])', 'Remove linhas/colunas do início (ou do fim).', [2, 3], (a) => {
  let m = matrixOf(a[0]);
  const r = int(a[1]), c = a[2] !== undefined ? int(a[2]) : 0;
  m = r >= 0 ? m.slice(r) : m.slice(0, r);
  if (c) m = m.map((row) => (c >= 0 ? row.slice(c) : row.slice(0, c)));
  return m.length && m[0].length ? m : err(ERR.CALC);
});
def('VSTACK', 'EMPILHARV', DA, 'EMPILHARV(matriz1; …)', 'Empilha matrizes verticalmente.', [1, MANY], (a) => {
  const ms = a.map(matrixOf), w = Math.max(...ms.map((m) => m[0].length));
  return ms.flatMap((m) => m.map((r) => [...r, ...Array(w - r.length).fill(err(ERR.NA))]));
});
def('HSTACK', 'EMPILHARH', DA, 'EMPILHARH(matriz1; …)', 'Junta matrizes lado a lado.', [1, MANY], (a) => {
  const ms = a.map(matrixOf), h = Math.max(...ms.map((m) => m.length));
  return Array.from({ length: h }, (_, r) => ms.flatMap((m) => m[r] || Array(m[0].length).fill(err(ERR.NA))));
});
def('CHOOSECOLS', 'ESCOLHERCOLS', DA, 'ESCOLHERCOLS(matriz; col1; …)', 'Escolhe colunas.', [2, MANY], (a) => {
  const m = matrixOf(a[0]);
  return m.map((row) => a.slice(1).map((x) => { const i = int(x); return row[i > 0 ? i - 1 : row.length + i] ?? err(ERR.VALUE); }));
});
def('CHOOSEROWS', 'ESCOLHERLINS', DA, 'ESCOLHERLINS(matriz; lin1; …)', 'Escolhe linhas.', [2, MANY], (a) => {
  const m = matrixOf(a[0]);
  return a.slice(1).map((x) => { const i = int(x); return m[i > 0 ? i - 1 : m.length + i] || fail(ERR.VALUE); });
});
def('TOCOL', 'PARACOL', DA, 'PARACOL(matriz)', 'Transforma em uma coluna.', [1, 3], (a) => flat([a[0]]).filter((v) => !isBlank(v)).map((v) => [v]));
def('TOROW', 'PARALIN', DA, 'PARALIN(matriz)', 'Transforma em uma linha.', [1, 3], (a) => [flat([a[0]]).filter((v) => !isBlank(v))]);

// ---------------------------------------------------------------------------
// Texto

const T = 'Texto';
def('CONCATENATE', 'CONCATENAR', T, 'CONCATENAR(texto1; [texto2]; …)', 'Junta textos.', [1, MANY], (a) => a.map(str).join(''));
def('CONCAT', 'CONCAT', T, 'CONCAT(texto1; …)', 'Junta textos e intervalos.', [1, MANY], (a) => flat(a).map(str).join(''));
def('TEXTJOIN', 'UNIRTEXTO', T, 'UNIRTEXTO(delimitador; ignorar_vazio; texto1; …)', 'Junta textos com um delimitador.', [3, MANY], (a) => {
  const d = str(a[0]), skip = bool(a[1]);
  return flat(a.slice(2)).filter((v) => !(skip && isBlank(v))).map(str).join(d);
});
def('LEFT', 'ESQUERDA', T, 'ESQUERDA(texto; [n])', 'Primeiros caracteres.', [1, 2], (a) => str(a[0]).slice(0, Math.max(0, int(opt(a, 1, 1)))), { ew: true });
def('RIGHT', 'DIREITA', T, 'DIREITA(texto; [n])', 'Últimos caracteres.', [1, 2], (a) => { const s = str(a[0]), n = int(opt(a, 1, 1)); return n <= 0 ? '' : s.slice(-n); }, { ew: true });
def('MID', 'EXT.TEXTO', T, 'EXT.TEXTO(texto; início; n)', 'Parte do texto.', [3, 3], (a) => str(a[0]).substr(Math.max(0, int(a[1]) - 1), Math.max(0, int(a[2]))), { ew: true });
def('LEN', 'NÚM.CARACT', T, 'NÚM.CARACT(texto)', 'Quantidade de caracteres.', [1, 1], (a) => str(a[0]).length, { ew: true });
def('UPPER', 'MAIÚSCULA', T, 'MAIÚSCULA(texto)', 'Converte para maiúsculas.', [1, 1], (a) => str(a[0]).toLocaleUpperCase('pt-BR'), { ew: true });
def('LOWER', 'MINÚSCULA', T, 'MINÚSCULA(texto)', 'Converte para minúsculas.', [1, 1], (a) => str(a[0]).toLocaleLowerCase('pt-BR'), { ew: true });
def('PROPER', 'PRI.MAIÚSCULA', T, 'PRI.MAIÚSCULA(texto)', 'Primeira letra de cada palavra em maiúscula.', [1, 1], (a) => str(a[0]).toLocaleLowerCase('pt-BR').replace(/(^|[^\p{L}])(\p{L})/gu, (m, x, y) => x + y.toLocaleUpperCase('pt-BR')), { ew: true });
def('TRIM', 'ARRUMAR', T, 'ARRUMAR(texto)', 'Remove espaços extras.', [1, 1], (a) => str(a[0]).replace(/ +/g, ' ').trim(), { ew: true });
def('CLEAN', 'TIRAR', T, 'TIRAR(texto)', 'Remove caracteres não imprimíveis.', [1, 1], (a) => str(a[0]).replace(/[\x00-\x1f]/g, ''), { ew: true });
def('SUBSTITUTE', 'SUBSTITUIR', T, 'SUBSTITUIR(texto; antigo; novo; [ocorrência])', 'Substitui texto.', [3, 4], (a) => {
  const s = str(a[0]), o = str(a[1]), n = str(a[2]);
  if (!o) return s;
  if (a[3] === undefined) return s.split(o).join(n);
  const k = int(a[3]);
  let idx = -1;
  for (let i = 0; i < k; i++) { idx = s.indexOf(o, idx + 1); if (idx < 0) return s; }
  return s.slice(0, idx) + n + s.slice(idx + o.length);
}, { ew: true });
def('REPLACE', 'MUDAR', T, 'MUDAR(texto; início; n; novo)', 'Troca parte do texto pela posição.', [4, 4], (a) => { const s = str(a[0]), st = int(a[1]) - 1; return s.slice(0, st) + str(a[3]) + s.slice(st + int(a[2])); }, { ew: true });
def('FIND', 'PROCURAR', T, 'PROCURAR(procurado; texto; [início])', 'Posição do texto (diferencia maiúsculas).', [2, 3], (a) => { const i = str(a[1]).indexOf(str(a[0]), int(opt(a, 2, 1)) - 1); return i < 0 ? err(ERR.VALUE) : i + 1; }, { ew: true });
def('SEARCH', 'LOCALIZAR', T, 'LOCALIZAR(procurado; texto; [início])', 'Posição do texto (aceita * e ?).', [2, 3], (a) => {
  const s = str(a[1]), st = int(opt(a, 2, 1)) - 1;
  const m = wildcard(str(a[0]), false).exec(s.slice(st));
  return m ? st + m.index + 1 : err(ERR.VALUE);
}, { ew: true });
def('TEXT', 'TEXTO', T, 'TEXTO(valor; formato)', 'Formata um número como texto (ex.: "R$ #.##0,00", "dd/mm/aaaa").', [2, 2], (a) => {
  const v = a[0];
  const n = typeof v === 'number' ? v : toNum(v);
  return isErr(n) ? str(v) : formatCode(n, str(a[1]));
}, { ew: true });
def('VALUE', 'VALOR', T, 'VALOR(texto)', 'Converte texto em número.', [1, 1], (a) => num(a[0]), { ew: true });
def('NUMBERVALUE', 'VALORNUMÉRICO', T, 'VALORNUMÉRICO(texto; [sep_decimal]; [sep_grupo])', 'Converte texto em número com separadores definidos.', [1, 3], (a) => {
  let s = str(a[0]);
  const d = a[1] !== undefined ? str(a[1]) : ',', g = a[2] !== undefined ? str(a[2]) : '.';
  s = s.split(g).join('').replace(d, '.');
  const n = parseFloat(s);
  return Number.isNaN(n) ? err(ERR.VALUE) : n;
});
def('REPT', 'REPT', T, 'REPT(texto; n)', 'Repete o texto.', [2, 2], (a) => str(a[0]).repeat(Math.max(0, int(a[1]))), { ew: true });
def('EXACT', 'EXATO', T, 'EXATO(texto1; texto2)', 'Compara textos exatamente.', [2, 2], (a) => str(a[0]) === str(a[1]), { ew: true });
def('CHAR', 'CARACT', T, 'CARACT(núm)', 'Caractere pelo código.', [1, 1], (a) => String.fromCharCode(int(a[0])), { ew: true });
def('CODE', 'CÓDIGO', T, 'CÓDIGO(texto)', 'Código do primeiro caractere.', [1, 1], (a) => str(a[0]).charCodeAt(0) || err(ERR.VALUE), { ew: true });
def('UNICHAR', 'CARACTUNICODE', T, 'CARACTUNICODE(núm)', 'Caractere Unicode.', [1, 1], (a) => String.fromCodePoint(int(a[0])), { ew: true });
def('UNICODE', 'UNICODE', T, 'UNICODE(texto)', 'Código Unicode.', [1, 1], (a) => str(a[0]).codePointAt(0) ?? err(ERR.VALUE), { ew: true });
def('TEXTBEFORE', 'TEXTOANTES', T, 'TEXTOANTES(texto; delimitador; [ocorrência])', 'Texto antes do delimitador.', [2, 3], (a) => {
  const s = str(a[0]), d = str(a[1]), n = int(opt(a, 2, 1));
  const parts = s.split(d);
  if (parts.length <= Math.abs(n)) return err(ERR.NA);
  return n > 0 ? parts.slice(0, n).join(d) : parts.slice(0, parts.length + n).join(d);
}, { ew: true });
def('TEXTAFTER', 'TEXTODEPOIS', T, 'TEXTODEPOIS(texto; delimitador; [ocorrência])', 'Texto depois do delimitador.', [2, 3], (a) => {
  const s = str(a[0]), d = str(a[1]), n = int(opt(a, 2, 1));
  const parts = s.split(d);
  if (parts.length <= Math.abs(n)) return err(ERR.NA);
  return n > 0 ? parts.slice(n).join(d) : parts.slice(parts.length + n).join(d);
}, { ew: true });
def('TEXTSPLIT', 'DIVIDIRTEXTO', T, 'DIVIDIRTEXTO(texto; delim_coluna; [delim_linha])', 'Divide o texto em várias células.', [2, 3], (a) => {
  const s = str(a[0]), cd = str(a[1]), rd = a[2] !== undefined && a[2] !== null ? str(a[2]) : null;
  const rows = rd ? s.split(rd) : [s];
  const m = rows.map((r) => (cd ? r.split(cd) : [r]));
  const w = Math.max(...m.map((r) => r.length));
  return m.map((r) => [...r, ...Array(w - r.length).fill(err(ERR.NA))]);
});
def('T', 'T', T, 'T(valor)', 'Retorna o texto ou vazio.', [1, 1], (a) => (typeof a[0] === 'string' ? a[0] : ''));
def('N', 'N', T, 'N(valor)', 'Converte em número.', [1, 1], (a) => (typeof a[0] === 'number' ? a[0] : typeof a[0] === 'boolean' ? +a[0] : 0));
def('FIXED', 'DEF.NÚM.DEC', T, 'DEF.NÚM.DEC(núm; [casas]; [sem_separador])', 'Formata com casas decimais.', [1, 3], (a) => {
  const d = int(opt(a, 1, 2)), noSep = a[2] !== undefined && bool(a[2]);
  return new Intl.NumberFormat('pt-BR', { minimumFractionDigits: Math.max(0, d), maximumFractionDigits: Math.max(0, d), useGrouping: !noSep }).format(roundTo(num(a[0]), d));
}, { ew: true });

// ---------------------------------------------------------------------------
// Data e hora

const D = 'Data e hora';
const dparts = (v) => { const d = dateFromSerial(num(v)); return { y: d.getUTCFullYear(), m: d.getUTCMonth() + 1, d: d.getUTCDate(), wd: d.getUTCDay(), date: d }; };
def('TODAY', 'HOJE', D, 'HOJE()', 'Data de hoje.', [0, 0], () => todaySerial(), { volatile: true, fmt: 'date' });
def('NOW', 'AGORA', D, 'AGORA()', 'Data e hora atuais.', [0, 0], () => nowSerial(), { volatile: true, fmt: 'datetime' });
def('DATE', 'DATA', D, 'DATA(ano; mês; dia)', 'Cria uma data.', [3, 3], (a) => { let y = int(a[0]); if (y < 1900) y += 1900; return serialFromParts(y, int(a[1]), int(a[2])); }, { ew: true, fmt: 'date' });
def('TIME', 'TEMPO', D, 'TEMPO(hora; minuto; segundo)', 'Cria uma hora.', [3, 3], (a) => ((int(a[0]) * 3600 + int(a[1]) * 60 + int(a[2])) / 86400) % 1, { ew: true, fmt: 'time' });
def('YEAR', 'ANO', D, 'ANO(data)', 'Ano da data.', [1, 1], (a) => dparts(a[0]).y, { ew: true });
def('MONTH', 'MÊS', D, 'MÊS(data)', 'Mês da data (1 a 12).', [1, 1], (a) => dparts(a[0]).m, { ew: true });
def('DAY', 'DIA', D, 'DIA(data)', 'Dia do mês.', [1, 1], (a) => dparts(a[0]).d, { ew: true });
def('HOUR', 'HORA', D, 'HORA(valor)', 'Hora (0 a 23).', [1, 1], (a) => Math.floor(((num(a[0]) % 1) * 24) + 1e-9), { ew: true });
def('MINUTE', 'MINUTO', D, 'MINUTO(valor)', 'Minuto.', [1, 1], (a) => Math.floor(((num(a[0]) * 1440) % 60) + 1e-9), { ew: true });
def('SECOND', 'SEGUNDO', D, 'SEGUNDO(valor)', 'Segundo.', [1, 1], (a) => Math.round((num(a[0]) * 86400) % 60) % 60, { ew: true });
def('WEEKDAY', 'DIA.DA.SEMANA', D, 'DIA.DA.SEMANA(data; [tipo])', 'Dia da semana (1 = domingo).', [1, 2], (a) => {
  const wd = dparts(a[0]).wd, t = int(opt(a, 1, 1));
  return t === 2 ? ((wd + 6) % 7) + 1 : t === 3 ? (wd + 6) % 7 : wd + 1;
}, { ew: true });
function isoWeek(date) {
  const d = new Date(Date.UTC(date.getUTCFullYear(), date.getUTCMonth(), date.getUTCDate()));
  const day = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - day);
  const y0 = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - y0) / 864e5 + 1) / 7);
}
def('WEEKNUM', 'NÚMSEMANA', D, 'NÚMSEMANA(data; [tipo])', 'Número da semana no ano.', [1, 2], (a) => {
  const p = dparts(a[0]);
  const jan1 = new Date(Date.UTC(p.y, 0, 1));
  const start = int(opt(a, 1, 1)) === 2 ? 1 : 0;
  const offset = (jan1.getUTCDay() - start + 7) % 7;
  return Math.floor(((p.date - jan1) / 864e5 + offset) / 7) + 1;
}, { ew: true });
def('ISOWEEKNUM', 'NÚMSEMANAISO', D, 'NÚMSEMANAISO(data)', 'Semana ISO.', [1, 1], (a) => isoWeek(dparts(a[0]).date), { ew: true });
const addMonths = (s, n) => { const p = dparts(s); const last = new Date(Date.UTC(p.y, p.m - 1 + n + 1, 0)).getUTCDate(); return serialFromParts(p.y, p.m + n, Math.min(p.d, last)); };
def('EDATE', 'DATAM', D, 'DATAM(data; meses)', 'Data somando meses.', [2, 2], (a) => addMonths(a[0], int(a[1])), { ew: true, fmt: 'date' });
def('EOMONTH', 'FIMMÊS', D, 'FIMMÊS(data; meses)', 'Último dia do mês.', [2, 2], (a) => { const p = dparts(a[0]); return serialFromParts(p.y, p.m + int(a[1]) + 1, 0); }, { ew: true, fmt: 'date' });
def('DATEDIF', 'DATADIF', D, 'DATADIF(início; fim; unidade)', 'Diferença entre datas ("D", "M", "Y", "MD", "YM", "YD").', [3, 3], (a) => {
  const s = dparts(a[0]), e = dparts(a[1]), u = str(a[2]).toUpperCase();
  if (num(a[1]) < num(a[0])) fail(ERR.NUM);
  let months = (e.y - s.y) * 12 + (e.m - s.m);
  if (e.d < s.d) months--;
  switch (u) {
    case 'D': return Math.floor(num(a[1])) - Math.floor(num(a[0]));
    case 'M': return months;
    case 'Y': case 'A': return Math.floor(months / 12);
    case 'YM': case 'AM': return months % 12;
    case 'MD': { const prev = new Date(Date.UTC(e.y, e.m - 1, 0)).getUTCDate(); return e.d >= s.d ? e.d - s.d : prev - s.d + e.d; }
    case 'YD': case 'AD': { let y = e.y; if (e.m < s.m || (e.m === s.m && e.d < s.d)) y--; return Math.floor(num(a[1]) - serialFromParts(y, s.m, s.d)); }
    default: return err(ERR.NUM);
  }
});
def('DAYS', 'DIAS', D, 'DIAS(data_final; data_inicial)', 'Dias entre duas datas.', [2, 2], (a) => Math.floor(num(a[0])) - Math.floor(num(a[1])), { ew: true });
const isWork = (s, hol) => { const wd = dateFromSerial(s).getUTCDay(); return wd !== 0 && wd !== 6 && !hol.has(Math.floor(s)); };
def('NETWORKDAYS', 'DIATRABALHOTOTAL', D, 'DIATRABALHOTOTAL(início; fim; [feriados])', 'Dias úteis entre datas.', [2, 3], (a) => {
  let s = Math.floor(num(a[0])), e = Math.floor(num(a[1]));
  const hol = new Set(a[2] !== undefined ? numbers([a[2]]).map(Math.floor) : []);
  const sign = e < s ? -1 : 1;
  if (sign < 0) [s, e] = [e, s];
  let n = 0;
  for (let d = s; d <= e; d++) if (isWork(d, hol)) n++;
  return n * sign;
});
def('WORKDAY', 'DIATRABALHO', D, 'DIATRABALHO(início; dias; [feriados])', 'Data após N dias úteis.', [2, 3], (a) => {
  let d = Math.floor(num(a[0]));
  let n = int(a[1]);
  const hol = new Set(a[2] !== undefined ? numbers([a[2]]).map(Math.floor) : []);
  const step = n < 0 ? -1 : 1;
  while (n !== 0) { d += step; if (isWork(d, hol)) n -= step; }
  return d;
}, { fmt: 'date' });
def('DATEVALUE', 'DATA.VALOR|DATAVALOR', D, 'DATA.VALOR(texto)', 'Converte texto em data.', [1, 1], (a) => { const d = parseDateText(str(a[0])); return d ? Math.floor(d.serial) : err(ERR.VALUE); }, { ew: true, fmt: 'date' });
def('TIMEVALUE', 'VALOR.TEMPO', D, 'VALOR.TEMPO(texto)', 'Converte texto em hora.', [1, 1], (a) => { const d = parseDateText(str(a[0])); return d ? d.serial % 1 : err(ERR.VALUE); }, { ew: true, fmt: 'time' });

// ---------------------------------------------------------------------------
// Informações

const I = 'Informações';
def('ISBLANK', 'ÉCÉL.VAZIA', I, 'ÉCÉL.VAZIA(valor)', 'Verifica se está vazio.', [1, 1], (a) => isBlank(a[0]), { ew: true, errs: true });
def('ISNUMBER', 'ÉNÚM', I, 'ÉNÚM(valor)', 'Verifica se é número.', [1, 1], (a) => typeof a[0] === 'number', { ew: true, errs: true });
def('ISTEXT', 'ÉTEXTO', I, 'ÉTEXTO(valor)', 'Verifica se é texto.', [1, 1], (a) => typeof a[0] === 'string', { ew: true, errs: true });
def('ISNONTEXT', 'ÉNÃOTEXTO', I, 'ÉNÃOTEXTO(valor)', 'Verifica se não é texto.', [1, 1], (a) => typeof a[0] !== 'string', { ew: true, errs: true });
def('ISLOGICAL', 'ÉLÓGICO', I, 'ÉLÓGICO(valor)', 'Verifica se é lógico.', [1, 1], (a) => typeof a[0] === 'boolean', { ew: true, errs: true });
def('ISERROR', 'ÉERROS', I, 'ÉERROS(valor)', 'Verifica se é qualquer erro.', [1, 1], (a) => isErr(a[0]), { ew: true, errs: true });
def('ISERR', 'ÉERRO', I, 'ÉERRO(valor)', 'Erro diferente de #N/D.', [1, 1], (a) => isErr(a[0]) && a[0].code !== ERR.NA, { ew: true, errs: true });
def('ISNA', 'É.NÃO.DISP', I, 'É.NÃO.DISP(valor)', 'Verifica se é #N/D.', [1, 1], (a) => isErr(a[0]) && a[0].code === ERR.NA, { ew: true, errs: true });
def('ISEVEN', 'ÉPAR', I, 'ÉPAR(núm)', 'Verifica se é par.', [1, 1], (a) => Math.trunc(num(a[0])) % 2 === 0, { ew: true });
def('ISODD', 'ÉIMPAR|ÉÍMPAR', I, 'ÉIMPAR(núm)', 'Verifica se é ímpar.', [1, 1], (a) => Math.abs(Math.trunc(num(a[0])) % 2) === 1, { ew: true });
def('NA', 'NÃO.DISP', I, 'NÃO.DISP()', 'Retorna o erro #N/D.', [0, 0], () => err(ERR.NA));
def('TYPE', 'TIPO', I, 'TIPO(valor)', 'Tipo do valor (1 número, 2 texto, 4 lógico, 16 erro, 64 matriz).', [1, 1], (a) => {
  const v = a[0];
  return typeof v === 'number' || v === null ? 1 : typeof v === 'string' ? 2 : typeof v === 'boolean' ? 4 : isErr(v) ? 16 : 64;
}, { errs: true });

// ---------------------------------------------------------------------------
// Financeiro

const FI = 'Financeiro';
function pmt(rate, nper, pv, fv = 0, type = 0) {
  if (rate === 0) return -(pv + fv) / nper;
  const f = (1 + rate) ** nper;
  return (-rate * (pv * f + fv)) / ((1 + rate * type) * (f - 1));
}
function fv(rate, nper, p, pv = 0, type = 0) {
  if (rate === 0) return -(pv + p * nper);
  const f = (1 + rate) ** nper;
  return -(pv * f + (p * (1 + rate * type) * (f - 1)) / rate);
}
def('PMT', 'PGTO', FI, 'PGTO(taxa; nper; vp; [vf]; [tipo])', 'Valor da parcela de um empréstimo.', [3, 5], (a) => pmt(num(a[0]), num(a[1]), num(a[2]), num(opt(a, 3, 0)), num(opt(a, 4, 0))), { fmt: 'currency' });
def('FV', 'VF', FI, 'VF(taxa; nper; pgto; [vp]; [tipo])', 'Valor futuro de um investimento.', [3, 5], (a) => fv(num(a[0]), num(a[1]), num(a[2]), num(opt(a, 3, 0)), num(opt(a, 4, 0))), { fmt: 'currency' });
def('PV', 'VP', FI, 'VP(taxa; nper; pgto; [vf]; [tipo])', 'Valor presente.', [3, 5], (a) => {
  const r = num(a[0]), n = num(a[1]), p = num(a[2]), f = num(opt(a, 3, 0)), t = num(opt(a, 4, 0));
  if (r === 0) return -(f + p * n);
  return (-f - (p * (1 + r * t) * ((1 + r) ** n - 1)) / r) / (1 + r) ** n;
}, { fmt: 'currency' });
def('NPER', 'NPER', FI, 'NPER(taxa; pgto; vp; [vf]; [tipo])', 'Número de períodos.', [3, 5], (a) => {
  const r = num(a[0]), p = num(a[1]), v = num(a[2]), f = num(opt(a, 3, 0)), t = num(opt(a, 4, 0));
  if (r === 0) return -(v + f) / p;
  const z = p * (1 + r * t) / r;
  return Math.log((z - f) / (z + v)) / Math.log(1 + r);
});
def('RATE', 'TAXA', FI, 'TAXA(nper; pgto; vp; [vf]; [tipo]; [estimativa])', 'Taxa de juros por período.', [3, 6], (a) => {
  const n = num(a[0]), p = num(a[1]), v = num(a[2]), f = num(opt(a, 3, 0)), t = num(opt(a, 4, 0));
  let r = num(opt(a, 5, 0.1));
  for (let i = 0; i < 100; i++) {
    const g = (x) => v * (1 + x) ** n + p * (1 + x * t) * (((1 + x) ** n - 1) / (x || 1e-12)) + f;
    const y = g(r), dy = (g(r + 1e-7) - y) / 1e-7;
    const nr = r - y / dy;
    if (!Number.isFinite(nr)) break;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  return Number.isFinite(r) ? r : err(ERR.NUM);
}, { fmt: 'percent' });
def('NPV', 'VPL', FI, 'VPL(taxa; valor1; …)', 'Valor presente líquido.', [2, MANY], (a) => { const r = num(a[0]); return numbers(a.slice(1)).reduce((s, x, i) => s + x / (1 + r) ** (i + 1), 0); }, { fmt: 'currency' });
def('IRR', 'TIR', FI, 'TIR(valores; [estimativa])', 'Taxa interna de retorno.', [1, 2], (a) => {
  const xs = numbers([a[0]]);
  let r = num(opt(a, 1, 0.1));
  for (let i = 0; i < 200; i++) {
    let f = 0, df = 0;
    xs.forEach((x, k) => { f += x / (1 + r) ** k; df -= (k * x) / (1 + r) ** (k + 1); });
    const nr = r - f / df;
    if (!Number.isFinite(nr)) break;
    if (Math.abs(nr - r) < 1e-10) return nr;
    r = nr;
  }
  return err(ERR.NUM);
}, { fmt: 'percent' });
def('IPMT', 'IPGTO', FI, 'IPGTO(taxa; período; nper; vp; [vf]; [tipo])', 'Juros de uma parcela.', [4, 6], (a) => {
  const r = num(a[0]), per = num(a[1]), n = num(a[2]), v = num(a[3]), f = num(opt(a, 4, 0)), t = num(opt(a, 5, 0));
  const p = pmt(r, n, v, f, t);
  const bal = fv(r, per - 1, p, v, t);
  return t && per === 1 ? 0 : bal * r * (t ? 1 / (1 + r) : 1);
}, { fmt: 'currency' });
def('PPMT', 'PPGTO', FI, 'PPGTO(taxa; período; nper; vp; [vf]; [tipo])', 'Amortização de uma parcela.', [4, 6], (a, ctx) => {
  const p = pmt(num(a[0]), num(a[2]), num(a[3]), num(opt(a, 4, 0)), num(opt(a, 5, 0)));
  return p - FUNCS.IPMT.fn(a, ctx);
}, { fmt: 'currency' });

export const CATEGORIES = [...new Set(Object.values(FUNCS).map((f) => f.cat))];
export { Err, numberToText };
