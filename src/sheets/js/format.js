// Formatos de número e interpretação do que o usuário digita (padrão brasileiro).
import {
  isErr, dateFromSerial, parseNumberText, parseDateText, numberToText,
} from './values.js';

export const FORMATS = [
  ['general', 'Geral'], ['number', 'Número'], ['currency', 'Moeda (R$)'], ['accounting', 'Contábil'], ['percent', 'Porcentagem'],
  ['date', 'Data'], ['datetime', 'Data e hora'], ['time', 'Hora'], ['longdate', 'Data por extenso'], ['scientific', 'Científico'], ['text', 'Texto'],
];

const DEFAULT_DEC = { number: 2, currency: 2, accounting: 2, percent: 0, scientific: 2 };
const nfCache = new Map();
function nf(opts) {
  const k = JSON.stringify(opts);
  if (!nfCache.has(k)) nfCache.set(k, new Intl.NumberFormat('pt-BR', opts));
  return nfCache.get(k);
}

const MONTHS = ['janeiro', 'fevereiro', 'março', 'abril', 'maio', 'junho', 'julho', 'agosto', 'setembro', 'outubro', 'novembro', 'dezembro'];
const DAYS = ['domingo', 'segunda-feira', 'terça-feira', 'quarta-feira', 'quinta-feira', 'sexta-feira', 'sábado'];
const pad = (n, w = 2) => String(n).padStart(w, '0');

// Formato "Geral": até 10 dígitos significativos, como o Excel numa coluna padrão.
const genDisp = new Intl.NumberFormat('pt-BR', { maximumSignificantDigits: 10, useGrouping: false });
function generalDisplay(v) {
  const a = Math.abs(v);
  if (Number.isInteger(v) && a < 1e11) return String(v);
  if (a !== 0 && (a < 1e-9 || a >= 1e11)) return v.toExponential(4).replace('.', ',').replace('e+', 'E+').replace('e-', 'E-');
  return genDisp.format(v);
}

// Texto exibido na célula.
export function display(v, style = {}) {
  if (v === null || v === undefined) return '';
  if (isErr(v)) return v.code;
  if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
  if (Array.isArray(v)) return display(v[0]?.[0], style);
  if (typeof v === 'string') return v;
  const fmt = style.fmt || 'general';
  const dec = style.dec ?? DEFAULT_DEC[fmt] ?? 2;
  switch (fmt) {
    case 'number': return nf({ minimumFractionDigits: dec, maximumFractionDigits: dec }).format(v);
    case 'currency': return nf({ style: 'currency', currency: style.cur || 'BRL', minimumFractionDigits: dec, maximumFractionDigits: dec }).format(v);
    case 'accounting': {
      const s = nf({ minimumFractionDigits: dec, maximumFractionDigits: dec }).format(Math.abs(v));
      const sym = { BRL: 'R$', USD: 'US$', EUR: '€' }[style.cur || 'BRL'];
      return v < 0 ? `${sym} (${s})` : v === 0 ? `${sym} -` : `${sym} ${s}`;
    }
    case 'percent': return nf({ style: 'percent', minimumFractionDigits: dec, maximumFractionDigits: dec }).format(v);
    case 'scientific': return v.toExponential(dec).replace('.', ',').replace('e+', 'E+').replace('e-', 'E-');
    case 'date': return fmtDate(v, 'dd/mm/aaaa');
    case 'datetime': return fmtDate(v, 'dd/mm/aaaa hh:mm');
    case 'time': return fmtDate(v, 'hh:mm:ss');
    case 'longdate': return fmtDate(v, 'dddd, d "de" mmmm "de" aaaa');
    case 'custom': return formatCode(v, style.code || 'Geral');
    case 'text': return numberToText(v);
    default: return generalDisplay(v);
  }
}

export function fmtDate(serial, pattern) {
  if (typeof serial !== 'number' || !Number.isFinite(serial)) return String(serial);
  const d = dateFromSerial(serial);
  const Y = d.getUTCFullYear(), M = d.getUTCMonth(), D = d.getUTCDate(), W = d.getUTCDay();
  const h = d.getUTCHours(), mi = d.getUTCMinutes(), s = d.getUTCSeconds();
  let out = '';
  const p = pattern;
  for (let i = 0; i < p.length;) {
    if (p[i] === '"') { const j = p.indexOf('"', i + 1); out += p.slice(i + 1, j < 0 ? p.length : j); i = j < 0 ? p.length : j + 1; continue; }
    const rest = p.slice(i).toLowerCase();
    const tok = ['dddd', 'ddd', 'dd', 'd', 'mmmm', 'mmm', 'mm', 'm', 'aaaa', 'yyyy', 'aa', 'yy', 'hh', 'h', 'ss', 's', 'nn'].find((t) => rest.startsWith(t));
    if (!tok) { out += p[i++]; continue; }
    const prev = out.slice(-3);
    const isMinute = (tok === 'mm' || tok === 'm') && (/[h]{1,2}[:\s]*$/i.test(p.slice(0, i)) || /^:?s/i.test(p.slice(i + tok.length)) || /:\s*$/.test(prev));
    switch (tok) {
      case 'dddd': out += DAYS[W]; break;
      case 'ddd': out += DAYS[W].slice(0, 3); break;
      case 'dd': out += pad(D); break;
      case 'd': out += D; break;
      case 'mmmm': out += MONTHS[M]; break;
      case 'mmm': out += MONTHS[M].slice(0, 3); break;
      case 'mm': out += isMinute ? pad(mi) : pad(M + 1); break;
      case 'm': out += isMinute ? mi : M + 1; break;
      case 'aaaa': case 'yyyy': out += Y; break;
      case 'aa': case 'yy': out += pad(Y % 100); break;
      case 'hh': out += pad(h); break;
      case 'h': out += h; break;
      case 'ss': out += pad(s); break;
      case 's': out += s; break;
      case 'nn': out += pad(mi); break;
      default:
    }
    i += tok.length;
  }
  return out;
}

// Códigos de formato do Excel (subconjunto): 0, #,##0.00, 0%, datas, literais e seções ;
export function formatCode(v, code) {
  if (typeof v !== 'number') return v === null ? '' : String(v);
  if (/^(geral|general)$/i.test(code)) return numberToText(v);
  const sections = splitSections(code);
  let sec = sections[0];
  let val = v;
  if (v < 0 && sections[1] !== undefined) { sec = sections[1]; val = -v; }
  else if (v === 0 && sections[2] !== undefined) sec = sections[2];
  const stripped = sec.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '');
  if (/[dyaHhs]|mm?m?m?(?![^"]*")/i.test(stripped) && !/[0#]/.test(stripped)) return fmtDate(v, sec.replace(/\[[^\]]*\]/g, ''));
  // Número: decide separadores pelo padrão (pt-BR "#.##0,00" ou en "#,##0.00")
  const ptStyle = /\d|#/.test(stripped) && /,[0#]+(?![^]*\.)/.test(stripped) && (/\.#|\.0{3}|#\.#/.test(stripped) || !/\./.test(stripped));
  const decSep = ptStyle ? ',' : '.';
  const grpSep = ptStyle ? '.' : ',';
  const pct = (stripped.match(/%/g) || []).length;
  val *= 100 ** pct;
  const m = new RegExp(`[0#?]+(?:[${grpSep === '.' ? '\\.' : ','}][0#?]+)*(?:\\${decSep}[0#?]*)?`).exec(stripped);
  if (!m) return sec.replace(/"([^"]*)"/g, '$1');
  const numPat = m[0];
  const [intPat, decPat = ''] = numPat.split(decSep);
  const decimals = (decPat.match(/[0#?]/g) || []).length;
  const minDec = (decPat.match(/0/g) || []).length;
  const grouping = intPat.includes(grpSep);
  const minInt = (intPat.replace(new RegExp(`\\${grpSep}`, 'g'), '').match(/0/g) || []).length;
  const s = nf({ minimumFractionDigits: minDec, maximumFractionDigits: decimals, useGrouping: grouping, minimumIntegerDigits: Math.max(1, minInt) }).format(sections[1] !== undefined ? val : val);
  // Reconstrói com os literais em volta do padrão numérico
  const idx = sec.indexOf(numPat);
  const pre = idx >= 0 ? sec.slice(0, idx) : '';
  const post = idx >= 0 ? sec.slice(idx + numPat.length) : '';
  const lit = (t) => t.replace(/\[[^\]]*\]/g, '').replace(/"([^"]*)"/g, '$1').replace(/\\(.)/g, '$1').replace(/_./g, ' ');
  return lit(pre) + (minInt === 0 && Math.abs(val) < 1 && s.startsWith('0') && !intPat.includes('0') ? s.slice(1) : s) + lit(post);
}

function splitSections(code) {
  const out = [];
  let cur = '', q = false;
  for (const ch of code) {
    if (ch === '"') q = !q;
    if (ch === ';' && !q) { out.push(cur); cur = ''; continue; }
    cur += ch;
  }
  out.push(cur);
  return out;
}

// Interpreta texto digitado: retorna { v, fmt? }
export function parseInput(text) {
  if (text === null || text === undefined) return { v: null };
  const s = String(text);
  if (!s.trim()) return { v: null };
  if (s.startsWith("'")) return { v: s.slice(1), forceText: true };
  const u = s.trim().toUpperCase();
  if (u === 'VERDADEIRO' || u === 'TRUE') return { v: true };
  if (u === 'FALSO' || u === 'FALSE') return { v: false };
  const t = s.trim();
  if (/^[+-]?(R\$|US\$|€)\s?/.test(t) || /^\(?R\$/.test(t)) {
    const n = parseNumberText(t);
    if (n !== null) return { v: n, fmt: 'currency', cur: t.includes('US$') ? 'USD' : t.includes('€') ? 'EUR' : 'BRL' };
  }
  if (/%$/.test(t)) {
    const n = parseNumberText(t);
    if (n !== null) return { v: n, fmt: 'percent', dec: (t.split(/[,.]/)[1] || '').replace('%', '').length };
  }
  const n = parseNumberText(t);
  if (n !== null) return { v: n, fmt: /^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t) ? 'number' : undefined, dec: /^[+-]?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t) ? (t.split(',')[1] || '').length : undefined };
  const d = parseDateText(t);
  if (d) return { v: d.serial, fmt: d.fmt };
  return { v: s };
}

// Texto usado ao editar a célula (sem formatação de exibição, exceto datas).
export function editText(cell) {
  if (!cell) return '';
  if (cell.f) return `=${cell.f}`;
  const v = cell.v;
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'VERDADEIRO' : 'FALSO';
  if (typeof v === 'number') {
    const fmt = cell.s?.fmt;
    if (fmt === 'date') return fmtDate(v, 'dd/mm/aaaa');
    if (fmt === 'datetime') return fmtDate(v, 'dd/mm/aaaa hh:mm');
    if (fmt === 'time') return fmtDate(v, 'hh:mm:ss');
    if (fmt === 'percent') return `${numberToText(+(v * 100).toPrecision(15))}%`;
    return numberToText(v);
  }
  if (typeof v === 'string' && cell.forceText) return `'${v}`;
  return String(v);
}
