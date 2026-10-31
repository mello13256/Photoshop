// VBA de arquivos .xlsm: leitura do vbaProject.bin (Compound File Binary + compressão MS-OVBA)
// e conversão aproximada dos módulos para macros Lumina (JavaScript).

// ---- Compound File Binary ---------------------------------------------------------

function readCfb(buf) {
  const dv = new DataView(buf.buffer, buf.byteOffset, buf.byteLength);
  const sig = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  if (!sig.every((b, i) => buf[i] === b)) throw new Error('vbaProject.bin inválido');
  const u16 = (o) => dv.getUint16(o, true);
  const u32 = (o) => dv.getUint32(o, true);
  const secSize = 1 << u16(0x1e);
  const miniSize = 1 << u16(0x20);
  const firstDir = u32(0x30);
  const cutoff = u32(0x38);
  const firstMiniFat = u32(0x3c);
  const nMiniFat = u32(0x40);
  let difatSec = u32(0x44);
  const nDifat = u32(0x48);
  const sectorOff = (n) => (n + 1) * secSize;
  // DIFAT -> lista de setores da FAT
  const fatSectors = [];
  for (let i = 0; i < 109; i++) { const s = u32(0x4c + i * 4); if (s < 0xfffffffa) fatSectors.push(s); }
  for (let k = 0; k < nDifat && difatSec < 0xfffffffa; k++) {
    const off = sectorOff(difatSec);
    for (let i = 0; i < secSize / 4 - 1; i++) { const s = u32(off + i * 4); if (s < 0xfffffffa) fatSectors.push(s); }
    difatSec = u32(off + secSize - 4);
  }
  const fat = [];
  for (const s of fatSectors) { const off = sectorOff(s); for (let i = 0; i < secSize / 4; i++) fat.push(u32(off + i * 4)); }
  const chain = (start, table) => {
    const out = [];
    const seen = new Set();
    for (let s = start; s < 0xfffffffa && !seen.has(s) && s < table.length; s = table[s]) { seen.add(s); out.push(s); }
    return out;
  };
  const readChain = (start, size) => {
    const out = new Uint8Array(chain(start, fat).length * secSize);
    chain(start, fat).forEach((s, i) => out.set(buf.subarray(sectorOff(s), sectorOff(s) + secSize), i * secSize));
    return size === undefined ? out : out.subarray(0, size);
  };
  // Diretório
  const dirBytes = readChain(firstDir);
  const entries = [];
  for (let o = 0; o + 128 <= dirBytes.length; o += 128) {
    const e = new DataView(dirBytes.buffer, dirBytes.byteOffset + o, 128);
    const nameLen = e.getUint16(64, true);
    let name = '';
    for (let i = 0; i < Math.max(0, nameLen / 2 - 1); i++) name += String.fromCharCode(e.getUint16(i * 2, true));
    entries.push({ name, type: e.getUint8(66), left: e.getUint32(68, true), right: e.getUint32(72, true), child: e.getUint32(76, true), start: e.getUint32(116, true), size: e.getUint32(120, true) });
  }
  const root = entries[0];
  const miniFat = [];
  if (nMiniFat) { const mf = readChain(firstMiniFat); for (let i = 0; i < mf.length / 4; i++) miniFat.push(new DataView(mf.buffer, mf.byteOffset).getUint32(i * 4, true)); }
  const miniStream = root ? readChain(root.start, root.size) : new Uint8Array();
  const readStream = (e) => {
    if (e.size < cutoff) {
      const secs = chain(e.start, miniFat);
      const out = new Uint8Array(secs.length * miniSize);
      secs.forEach((s, i) => out.set(miniStream.subarray(s * miniSize, (s + 1) * miniSize), i * miniSize));
      return out.subarray(0, e.size);
    }
    return readChain(e.start, e.size);
  };
  // Caminhos completos
  const files = new Map();
  const walk = (id, prefix) => {
    if (id === 0xffffffff || id >= entries.length) return;
    const e = entries[id];
    walk(e.left, prefix);
    const p = prefix ? `${prefix}/${e.name}` : e.name;
    if (e.type === 2) files.set(p.toLowerCase(), () => readStream(e));
    if (e.type === 1 || e.type === 5) walk(e.child, p);
    walk(e.right, prefix);
  };
  walk(root.child, '');
  return files;
}

// ---- MS-OVBA ----------------------------------------------------------------------

function decompress(buf, start = 0) {
  if (buf[start] !== 1) throw new Error('fluxo VBA inválido');
  const out = [];
  let pos = start + 1;
  while (pos + 1 < buf.length) {
    const header = buf[pos] | (buf[pos + 1] << 8);
    const chunkEnd = Math.min(buf.length, pos + (header & 0x0fff) + 3);
    pos += 2;
    const chunkStart = out.length;
    if (!(header & 0x8000)) {
      for (let i = 0; i < 4096 && pos < buf.length; i++) out.push(buf[pos++]);
      continue;
    }
    while (pos < chunkEnd) {
      const flags = buf[pos++];
      for (let bit = 0; bit < 8 && pos < chunkEnd; bit++) {
        if (!(flags & (1 << bit))) { out.push(buf[pos++]); continue; }
        const token = buf[pos] | (buf[pos + 1] << 8);
        pos += 2;
        const d = out.length - chunkStart;
        let bits = 4;
        while ((1 << bits) < d) bits++;
        const lenMask = 0xffff >> bits;
        const length = (token & lenMask) + 3;
        const offset = (token >> (16 - bits)) + 1;
        const from = out.length - offset;
        for (let i = 0; i < length; i++) out.push(out[from + i]);
      }
    }
  }
  return new Uint8Array(out);
}

export function extractVba(bin) {
  const files = readCfb(bin);
  const dirKey = [...files.keys()].find((k) => /(^|\/)vba\/dir$/.test(k));
  if (!dirKey) return [];
  const base = dirKey.slice(0, -3); // "…/vba/"
  const dir = decompress(files.get(dirKey)());
  const dv = new DataView(dir.buffer);
  let pos = 0, codePage = 1252;
  const modules = [];
  let cur = null;
  while (pos + 6 <= dir.length) {
    const id = dv.getUint16(pos, true);
    let size = dv.getUint32(pos + 2, true);
    if (id === 0x0009) size = 6; // PROJECTVERSION tem tamanho fixo
    const data = dir.subarray(pos + 6, pos + 6 + size);
    pos += 6 + size;
    if (id === 0x0003) codePage = new DataView(data.buffer, data.byteOffset, 2).getUint16(0, true);
    else if (id === 0x0019) { cur = { name: decode(data, codePage), stream: null, offset: 0, type: 'module' }; modules.push(cur); }
    else if (id === 0x001a && cur) cur.stream = decode(data, codePage);
    else if (id === 0x0031 && cur) cur.offset = new DataView(data.buffer, data.byteOffset, 4).getUint32(0, true);
    else if (id === 0x0022 && cur) cur.type = 'class';
  }
  const out = [];
  for (const m of modules) {
    const reader = files.get(`${base}${(m.stream || m.name).toLowerCase()}`);
    if (!reader) continue;
    try {
      const raw = reader();
      const text = decode(decompress(raw, m.offset), codePage)
        .split(/\r?\n/).filter((l) => !/^Attribute VB_/i.test(l)).join('\n').trim();
      if (text) out.push({ name: m.name, type: m.type, code: text });
    } catch { /* módulo ilegível */ }
  }
  return out;
}

function decode(bytes, codePage) {
  const enc = { 1252: 'windows-1252', 65001: 'utf-8', 1250: 'windows-1250', 1251: 'windows-1251', 932: 'shift_jis', 936: 'gbk', 950: 'big5' }[codePage] || 'windows-1252';
  try { return new TextDecoder(enc).decode(bytes); } catch { return new TextDecoder('windows-1252').decode(bytes); }
}

// ---- Conversão VBA -> JavaScript ---------------------------------------------------------

const SHIMS = `// ---- Compatibilidade com VBA (gerado automaticamente) ----
const __rng = (ref) => {
  const o = {
    get Value() { return cel(ref); }, set Value(v) { definir(ref, v === true ? 'VERDADEIRO' : v === false ? 'FALSO' : v); },
    get Formula() { return cel(ref); }, set Formula(f) { formula(ref, f); },
    get Text() { return String(cel(ref) ?? ''); },
    get Address() { return ref; },
    get Row() { return +(/\\d+/.exec(ref.split('!').pop())[0]); },
    get Column() { let c = 0; for (const ch of /[A-Z]+/.exec(ref.split('!').pop())[0]) c = c * 26 + ch.charCodeAt(0) - 64; return c; },
    Offset(r, c) { const p = ref.includes('!') ? ref.split('!')[0] + '!' : ''; return __rng(p + endereco(o.Row + r, o.Column + c)); },
    ClearContents() { limpar(ref); }, Clear() { limpar(ref); },
    Font: { set Bold(v) { formatar(ref, { negrito: v }); }, set Italic(v) { formatar(ref, { italico: v }); }, set Color(v) { formatar(ref, { cor: __hex(v) }); }, set Size(v) { formatar(ref, { tamanho: v }); } },
    Interior: { set Color(v) { formatar(ref, { fundo: __hex(v) }); } },
    set NumberFormat(v) { formatar(ref, { formato: 'custom', code: v }); },
    get Cells() { return __cells(ref); },
    get Count() { return __cells(ref).length; },
    [Symbol.iterator]() { return __cells(ref)[Symbol.iterator](); },
  };
  return o;
};
const __cells = (ref) => {
  const p = ref.includes('!') ? ref.split('!')[0] + '!' : '';
  const r = ref.split('!').pop();
  const [a, b = a] = r.split(':');
  const pos = (x) => { const m = /([A-Z]+)(\\d+)/.exec(x); let c = 0; for (const ch of m[1]) c = c * 26 + ch.charCodeAt(0) - 64; return [+m[2], c]; };
  const [r1, c1] = pos(a), [r2, c2] = pos(b);
  const out = [];
  for (let i = r1; i <= r2; i++) for (let j = c1; j <= c2; j++) out.push(__rng(p + endereco(i, j)));
  return out;
};
const __sheet = (name) => ({ Range: (r) => __rng(name + '!' + r), Cells: (r, c) => __rng(name + '!' + endereco(r, c)), Name: name, Activate() { ativar(name); } });
const Range = (r, r2) => __rng(r2 ? r.split(':')[0] + ':' + r2 : r);
const Cells = (r, c) => __rng(endereco(r, c));
const Worksheets = (n) => __sheet(typeof n === 'number' ? planilhas()[n - 1] : n);
const Sheets = Worksheets;
const ActiveSheet = { Range, Cells };
const __hex = (v) => typeof v === 'number' ? '#' + [v & 255, (v >> 8) & 255, (v >> 16) & 255].map((x) => x.toString(16).padStart(2, '0')).join('') : v;
const RGB = (r, g, b) => r + g * 256 + b * 65536;
const MsgBox = (m) => alerta(m);
const Len = (s) => String(s ?? '').length, UCase = (s) => String(s ?? '').toUpperCase(), LCase = (s) => String(s ?? '').toLowerCase();
const Left = (s, n) => String(s ?? '').slice(0, n), Right = (s, n) => String(s ?? '').slice(-n), Mid = (s, i, n) => String(s ?? '').substr(i - 1, n ?? undefined);
const Trim = (s) => String(s ?? '').trim(), InStr = (a, b, c) => (c === undefined ? String(a).indexOf(b) + 1 : String(b).indexOf(c, a - 1) + 1);
const Replace = (s, a, b) => String(s ?? '').split(a).join(b), CStr = (v) => String(v ?? ''), CInt = (v) => Math.round(+v), CLng = CInt, CDbl = (v) => +v, Val = (v) => parseFloat(v) || 0;
const Int = Math.floor, Abs = Math.abs, Round = (v, d = 0) => Math.round(v * 10 ** d) / 10 ** d, Sqr = Math.sqrt;
const IsEmpty = (v) => v === null || v === undefined || v === '', IsNumeric = (v) => v !== '' && v !== null && !isNaN(+v);
const Now = () => hoje(), Date_ = () => hoje(), Year = (d) => new Date((d - 25569) * 86400000).getUTCFullYear(), Month = (d) => new Date((d - 25569) * 86400000).getUTCMonth() + 1, Day = (d) => new Date((d - 25569) * 86400000).getUTCDate();
const vbCrLf = '\\n', vbNewLine = '\\n', vbTab = '\\t';
const WorksheetFunction = { Sum: (...a) => a.flatMap((x) => (x && x.Cells ? x.Cells.map((c) => c.Value) : [x])).filter((v) => typeof v === 'number').reduce((s, v) => s + v, 0) };
const Application = { WorksheetFunction, ScreenUpdating: true, Calculate: () => recalcular() };
// ---------------------------------------------------------------
`;

export function vbaToJs(vba) {
  // Junta linhas continuadas com " _"
  const lines = vba.replace(/ _\r?\n\s*/g, ' ').split(/\r?\n/);
  const out = [];
  const subs = [];
  let fnName = null;
  let selDepth = 0;
  const conv = (code) => convertExpr(code);
  for (let raw of lines) {
    const indent = /^\s*/.exec(raw)[0];
    raw = raw.trim();
    if (!raw) { out.push(''); continue; }
    const { code, comment } = splitComment(raw);
    const c = code.trim();
    const tail = comment ? ` // ${comment}` : '';
    let m;
    const push = (s) => out.push(indent + s + tail);
    if (!c) { out.push(`${indent}//${comment ? ` ${comment}` : ''}`); continue; }
    if (/^Option\s/i.test(c)) continue;
    if ((m = /^(?:(?:Public|Private|Friend)\s+)?(?:Static\s+)?(Sub|Function)\s+(\w+)\s*\(([^)]*)\)/i.exec(c))) {
      const params = m[3].split(',').map((p) => p.replace(/\b(Optional|ByVal|ByRef|ParamArray)\b/gi, '').replace(/\s+As\s+[\w.]+/i, '').replace(/=.*$/, '').trim()).filter(Boolean);
      fnName = m[1].toLowerCase() === 'function' ? m[2] : null;
      subs.push({ name: m[2], params: params.length, kind: m[1].toLowerCase() });
      push(`async function ${m[2]}(${params.join(', ')}) {${fnName ? ` let ${fnName}__ = null;` : ''}`);
      continue;
    }
    if (/^End\s+(Sub|Function)$/i.test(c)) { out.push(`${indent}${fnName ? `  return ${fnName}__;\n${indent}` : ''}}`); fnName = null; continue; }
    if (/^Exit\s+(Sub|Function)$/i.test(c)) { push(fnName ? `return ${fnName}__;` : 'return;'); continue; }
    if (/^Exit\s+(For|Do)$/i.test(c)) { push('break;'); continue; }
    if ((m = /^Dim\s+(.+)$/i.exec(c)) || (m = /^(?:Private|Public|Static)\s+(?!Sub|Function)(.+)$/i.exec(c))) {
      const vars = m[1].split(',').map((v) => v.replace(/\s+As\s+(New\s+)?[\w.]+/i, '').replace(/\(.*\)/, '').trim()).filter(Boolean);
      push(`let ${vars.join(', ')};`);
      continue;
    }
    if ((m = /^Const\s+(\w+)(?:\s+As\s+\w+)?\s*=\s*(.+)$/i.exec(c))) { push(`const ${m[1]} = ${conv(m[2])};`); continue; }
    if ((m = /^If\s+(.+)\s+Then\s*$/i.exec(c))) { push(`if (${convCond(m[1])}) {`); continue; }
    if ((m = /^If\s+(.+?)\s+Then\s+(.+?)(?:\s+Else\s+(.+))?$/i.exec(c))) {
      push(`if (${convCond(m[1])}) { ${convStmt(m[2], fnName)} }${m[3] ? ` else { ${convStmt(m[3], fnName)} }` : ''}`);
      continue;
    }
    if ((m = /^ElseIf\s+(.+)\s+Then$/i.exec(c))) { push(`} else if (${convCond(m[1])}) {`); continue; }
    if (/^Else$/i.test(c)) { push('} else {'); continue; }
    if (/^End\s+If$/i.test(c)) { push('}'); continue; }
    if ((m = /^For\s+(\w+)\s*=\s*(.+?)\s+To\s+(.+?)(?:\s+Step\s+(.+))?$/i.exec(c))) {
      const step = m[4] ? conv(m[4]) : '1';
      const neg = /^-/.test(step.trim());
      push(`for (let ${m[1]} = ${conv(m[2])}; ${m[1]} ${neg ? '>=' : '<='} ${conv(m[3])}; ${m[1]} += ${step}) {`);
      continue;
    }
    if ((m = /^For\s+Each\s+(\w+)\s+In\s+(.+)$/i.exec(c))) { push(`for (const ${m[1]} of ${conv(m[2])}) {`); continue; }
    if (/^Next(\s+\w+)?$/i.test(c)) { push('}'); continue; }
    if ((m = /^Do\s+While\s+(.+)$/i.exec(c)) || (m = /^While\s+(.+)$/i.exec(c))) { push(`while (${convCond(m[1])}) {`); continue; }
    if ((m = /^Do\s+Until\s+(.+)$/i.exec(c))) { push(`while (!(${convCond(m[1])})) {`); continue; }
    if (/^Do$/i.test(c)) { push('do {'); continue; }
    if ((m = /^Loop\s+While\s+(.+)$/i.exec(c))) { push(`} while (${convCond(m[1])});`); continue; }
    if ((m = /^Loop\s+Until\s+(.+)$/i.exec(c))) { push(`} while (!(${convCond(m[1])}));`); continue; }
    if (/^(Loop|Wend)$/i.test(c)) { push('}'); continue; }
    if ((m = /^Select\s+Case\s+(.+)$/i.exec(c))) { selDepth++; push(`{ const __sel${selDepth} = ${conv(m[1])}; if (false) {`); continue; }
    if ((m = /^Case\s+Else$/i.exec(c))) { push('} else {'); continue; }
    if ((m = /^Case\s+(.+)$/i.exec(c))) {
      const v = `__sel${selDepth}`;
      const parts = splitArgs(m[1]).map((p) => {
        let k;
        if ((k = /^Is\s*(<>|<=|>=|<|>|=)\s*(.+)$/i.exec(p))) return `${v} ${k[1] === '=' ? '===' : k[1] === '<>' ? '!==' : k[1]} ${conv(k[2])}`;
        if ((k = /^(.+?)\s+To\s+(.+)$/i.exec(p))) return `(${v} >= ${conv(k[1])} && ${v} <= ${conv(k[2])})`;
        return `${v} === ${conv(p)}`;
      });
      push(`} else if (${parts.join(' || ')}) {`);
      continue;
    }
    if (/^End\s+Select$/i.test(c)) { selDepth = Math.max(0, selDepth - 1); push('} }'); continue; }
    if (/^(On\s+Error|Application\.(ScreenUpdating|DisplayAlerts|EnableEvents|Calculation))/i.test(c)) { push(`// ${c}`); continue; }
    if (/^End\s+With$|^With\s/i.test(c)) { push(`// ${c} (blocos With não são convertidos automaticamente)`); continue; }
    push(convStmt(c, fnName));
  }
  const entry = subs.find((s) => s.kind === 'sub' && s.params === 0);
  return `// Convertido automaticamente do VBA. Revise antes de executar.\n${SHIMS}\n${out.join('\n')}\n\n${entry ? `await ${entry.name}();` : '// Nenhuma Sub sem parâmetros para executar.'}\n`;
}

function convStmt(c, fnName) {
  let m;
  if ((m = /^Call\s+(\w+)\s*(?:\((.*)\))?$/i.exec(c))) return `await ${m[1]}(${m[2] ? convertExpr(m[2]) : ''});`;
  if ((m = /^(MsgBox|Debug\.Print)\s*(.*)$/i.exec(c))) {
    // Debug.Print aceita ";" e "," entre itens; MsgBox usa só a mensagem
    const args = splitArgs(m[2].replace(/;(?=(?:[^"]*"[^"]*")*[^"]*$)/g, ',')).filter(Boolean);
    const list = /^Debug/i.test(m[1]) ? args : args.slice(0, 1);
    return `${/^Debug/i.test(m[1]) ? 'log' : 'alerta'}(${list.map((x) => convertExpr(x)).join(', ')});`;
  }
  if ((m = /^Set\s+(.+?)\s*=\s*(.+)$/i.exec(c))) return `${convertExpr(m[1])} = ${convertExpr(m[2])};`;
  if ((m = /^(?:Let\s+)?([\w.]+(?:\([^=]*\))?(?:\.[\w]+)*)\s*=\s*(.+)$/i.exec(c))) {
    const lhs = fnName && m[1].toLowerCase() === fnName.toLowerCase() ? `${fnName}__` : convertExpr(m[1]);
    return `${lhs} = ${convertExpr(m[2])};`;
  }
  if (/^\w+$/.test(c)) return `await ${c}();`;
  if ((m = /^(\w+)\s+(.+)$/.exec(c)) && !/[=(]/.test(c.split(' ')[0])) return `await ${m[1]}(${convertExpr(m[2])});`;
  return `${convertExpr(c)};`;
}

function convCond(c) {
  // Em condições, "=" compara
  return convertExpr(c, true);
}

function convertExpr(src, cond = false) {
  return mapCode(src, (code) => {
    let s = code;
    s = s.replace(/\bAnd\b/gi, '&&').replace(/\bOr\b/gi, '||').replace(/\bNot\s+/gi, '!').replace(/\bMod\b/gi, '%')
      .replace(/\bTrue\b/gi, 'true').replace(/\bFalse\b/gi, 'false').replace(/\bNothing\b/gi, 'null').replace(/\bEmpty\b/g, "''")
      .replace(/\bIs\s+null\b/gi, '=== null').replace(/&/g, '+').replace(/\+\+/g, '&&')
      .replace(/<>/g, '!==').replace(/\bDate\b(?!\s*\()/g, 'Date_()').replace(/\bActiveSheet\./gi, '')
      .replace(/\bActiveCell\b/gi, 'Range(selecao.split(":")[0])')
      .replace(/\bApplication\.WorksheetFunction\./gi, 'WorksheetFunction.')
      .replace(/(\w+)\.Value\b/g, '$1.Value');
    if (cond) s = s.replace(/([^<>!=])=(?!=)/g, '$1===');
    // Restaura && que virou "+&" por conversão de &
    return s.replace(/\+&/g, '&&');
  });
}

// Aplica fn somente fora de strings; converte strings VBA ("a""b") para JS.
function mapCode(src, fn) {
  let out = '', buf = '';
  for (let i = 0; i < src.length; i++) {
    const ch = src[i];
    if (ch === '"') {
      out += fn(buf);
      buf = '';
      let j = i + 1, str = '';
      while (j < src.length) {
        if (src[j] === '"' && src[j + 1] === '"') { str += '"'; j += 2; continue; }
        if (src[j] === '"') break;
        str += src[j++];
      }
      out += JSON.stringify(str);
      i = j;
    } else buf += ch;
  }
  return out + fn(buf);
}

function splitComment(line) {
  let inStr = false;
  for (let i = 0; i < line.length; i++) {
    if (line[i] === '"') inStr = !inStr;
    else if (!inStr && line[i] === "'") return { code: line.slice(0, i), comment: line.slice(i + 1).trim() };
    else if (!inStr && /^Rem\s/i.test(line.slice(i)) && (i === 0 || /\s|:/.test(line[i - 1]))) return { code: line.slice(0, i), comment: line.slice(i + 4).trim() };
  }
  return { code: line, comment: '' };
}

function splitArgs(s) {
  const out = [];
  let depth = 0, inStr = false, cur = '';
  for (const ch of s) {
    if (ch === '"') inStr = !inStr;
    if (!inStr && ch === '(') depth++;
    if (!inStr && ch === ')') depth--;
    if (!inStr && depth === 0 && ch === ',') { out.push(cur.trim()); cur = ''; continue; }
    cur += ch;
  }
  if (cur.trim()) out.push(cur.trim());
  return out;
}
