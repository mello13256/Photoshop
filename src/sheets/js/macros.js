// Macros do Lumina Planilhas: editor, gravador e execução isolada em Web Worker.
import { el, dialog, toast, select, escapeHtml } from '../../shared/ui.js';
import { colName } from './values.js';
import { KEY } from './model.js';
import { parseInput } from './format.js';

export const MACRO_EXAMPLES = [
  {
    name: 'Numerar linhas',
    code: `// Escreve 1, 2, 3… na coluna A até a última linha usada da coluna B
const fim = ultimaLinha('B');
for (let i = 2; i <= fim; i++) {
  definir('A' + i, i - 1);
}
alerta('Linhas numeradas: ' + (fim - 1));`,
  },
  {
    name: 'Destacar valores acima da média',
    code: `// Pinta de verde os números da seleção acima da média
const dados = intervalo(selecao).flat().filter((v) => typeof v === 'number');
const media = dados.reduce((a, b) => a + b, 0) / dados.length;
const [ini] = selecao.split(':');
const linha0 = +ini.replace(/\\D/g, ''), col0 = ini.replace(/\\d/g, '');
intervalo(selecao).forEach((linha, i) => linha.forEach((v, j) => {
  if (typeof v === 'number' && v > media) {
    formatar(endereco(linha0 + i, coluna0(col0) + j), { fundo: '#c6efce', cor: '#006100', negrito: true });
  }
}));
function coluna0(letras) { let c = 0; for (const ch of letras) c = c * 26 + ch.charCodeAt(0) - 64; return c; }
log('Média:', media.toFixed(2));`,
  },
  {
    name: 'Relatório em nova planilha',
    code: `// Cria uma planilha "Resumo" com totais por coluna da planilha atual
const origem = planilhas()[0];
const ultima = ultimaLinha(undefined, origem);
const colunas = ultimaColuna(1, origem);
const nome = novaPlanilha('Resumo');
definir(nome + '!A1', 'Coluna');
definir(nome + '!B1', 'Total');
for (let c = 1; c <= colunas; c++) {
  const titulo = cel(origem + '!' + endereco(1, c));
  definir(nome + '!A' + (c + 1), titulo ?? coluna(c));
  definir(nome + '!B' + (c + 1), '=SOMA(' + origem + '!' + coluna(c) + '2:' + coluna(c) + ultima + ')');
}
formatar(nome + '!A1:B1', { negrito: true, fundo: '#dbe4ff' });
await recalcular();
log('Resumo criado com', colunas, 'colunas.');`,
  },
];

// ---------------------------------------------------------------------------
// Execução

let worker = null;

export function runMacro(app, macro, { onLog } = {}) {
  const wb = app.wb();
  return new Promise((resolve) => {
    worker?.terminate();
    worker = new Worker(new URL('./macro-worker.js', import.meta.url));
    let timer = null;
    const arm = () => { clearTimeout(timer); timer = setTimeout(() => finish({ error: 'A macro passou de 30 segundos e foi interrompida.' }), 30000); };
    const alerts = [];
    let applied = 0;
    const apply = (ops) => {
      if (!ops?.length) return;
      applyOps(app, ops, macro.name);
      applied += ops.length;
    };
    const finish = (res) => {
      clearTimeout(timer);
      worker?.terminate();
      worker = null;
      resolve({ ...res, alerts, applied });
    };
    worker.onmessage = (e) => {
      const m = e.data;
      if (m.type === 'log') onLog?.(m.text);
      else if (m.type === 'alert') { alerts.push(m.text); onLog?.(`⚠ ${m.text}`); }
      else if (m.type === 'flush') { apply(m.ops); arm(); worker.postMessage({ type: 'flushed', snapshot: snapshot(app) }); }
      else if (m.type === 'done') { apply(m.ops); finish({ ok: true }); }
      else if (m.type === 'error') { apply(m.ops); finish({ error: m.message, line: m.line }); }
    };
    worker.onerror = (e) => { e.preventDefault(); finish({ error: e.message || 'Erro de sintaxe na macro.' }); };
    arm();
    worker.postMessage({ type: 'run', code: macro.code, snapshot: snapshot(app) });
    void wb;
  });
}

export function stopMacro() {
  worker?.terminate();
  worker = null;
}

function snapshot(app) {
  const wb = app.wb();
  const sel = app.grid().sel;
  return {
    active: wb.active,
    sel: sel.r1 === sel.r2 && sel.c1 === sel.c2 ? `${colName(sel.c1)}${sel.r1 + 1}` : `${colName(sel.c1)}${sel.r1 + 1}:${colName(sel.c2)}${sel.r2 + 1}`,
    sheets: wb.sheets.map((s) => {
      const cells = [];
      const spills = wb.spills?.get(s.id);
      const keys = new Set([...s.cells.keys(), ...(spills ? spills.keys() : [])]);
      for (const k of keys) {
        const r = Math.floor(k / 16384), c = k % 16384;
        let v = wb.value(s, r, c);
        if (v && typeof v === 'object') v = v.code || String(v);
        const cell = s.cells.get(k);
        if (v === null && !cell?.f) continue;
        cells.push([`${r},${c}`, cell?.f ? { v, f: cell.f } : { v }]);
      }
      return { name: s.name, cells };
    }),
  };
}

const STYLE_KEYS = { negrito: 'b', bold: 'b', italico: 'i', italic: 'i', sublinhado: 'u', underline: 'u', cor: 'color', color: 'color', fundo: 'bg', background: 'bg', tamanho: 'size', size: 'size', fonte: 'font', font: 'font', alinhamento: 'ha', align: 'ha', formato: 'fmt', format: 'fmt', casas: 'dec', decimals: 'dec', quebrar: 'wrap', wrap: 'wrap' };
const FMT_ALIAS = { moeda: 'currency', numero: 'number', número: 'number', porcentagem: 'percent', percentual: 'percent', data: 'date', hora: 'time', texto: 'text', geral: 'general' };

function applyOps(app, ops, label) {
  const wb = app.wb();
  wb.transact(`Macro: ${label}`, () => {
    for (const op of ops) {
      if (op.op === 'addSheet') {
        if (!wb.sheets.some((s) => s.name.toLowerCase() === op.sheet.toLowerCase())) {
          wb.snapWorkbook();
          const s = wb.addSheet(op.sheet);
          if (s && s.name !== op.sheet) s.name = op.sheet;
        }
        continue;
      }
      const s = wb.sheets.find((x) => x.name.toLowerCase() === op.sheet.toLowerCase());
      if (!s) continue;
      if (op.op === 'activate') { wb.active = wb.sheets.indexOf(s); continue; }
      if (op.op === 'set') {
        const old = s.get(op.r, op.c);
        const cell = { v: null, s: old?.s };
        if (old?.note) cell.note = old.note;
        if (old?.link) cell.link = old.link;
        if (typeof op.value === 'string' && op.value.startsWith('=')) cell.f = op.value.slice(1);
        else if (typeof op.value === 'string') {
          const p = parseInput(op.value);
          cell.v = typeof p.v === 'number' && /^-?[\d.,]+%?$/.test(op.value.trim()) ? p.v : op.value;
        } else cell.v = op.value;
        wb.setCell(s, op.r, op.c, cell);
      } else if (op.op === 'format') {
        const st = {};
        for (const [k, v] of Object.entries(op.style || {})) {
          const key = STYLE_KEYS[k] || k;
          st[key] = key === 'fmt' ? FMT_ALIAS[v] || v : v;
        }
        const { r1, c1, r2, c2 } = op.range;
        for (let r = r1; r <= r2; r++) for (let c = c1; c <= c2; c++) {
          const old = s.get(r, c);
          const ns = { ...(old?.s || {}), ...st };
          for (const k of Object.keys(ns)) if (ns[k] === false || ns[k] === null) delete ns[k];
          wb.setCell(s, r, c, { ...(old ? { v: old.v, f: old.f, note: old.note, link: old.link, forceText: old.forceText } : { v: null }), s: ns });
        }
      }
    }
  });
  void KEY;
}

// ---------------------------------------------------------------------------
// Gravador

export class Recorder {
  constructor() { this.lines = null; }
  get active() { return !!this.lines; }
  start() { this.lines = []; }
  stop() { const l = this.lines || []; this.lines = null; return l.join('\n'); }
  cell(sheetName, r, c, text, multiSheet) {
    if (!this.lines) return;
    const ref = `${multiSheet ? `${sheetName}!` : ''}${colName(c)}${r + 1}`;
    const n = typeof text === 'string' ? parseInput(text) : { v: text };
    const val = typeof text === 'string' && !text.startsWith('=') && typeof n.v === 'number' && String(n.v) === text.trim() ? n.v : text;
    this.lines.push(`definir(${JSON.stringify(ref)}, ${JSON.stringify(val)});`);
  }
  style(rangeTxt, style) {
    if (!this.lines) return;
    const map = { b: 'negrito', i: 'italico', u: 'sublinhado', color: 'cor', bg: 'fundo', size: 'tamanho', font: 'fonte', ha: 'alinhamento', fmt: 'formato', dec: 'casas', wrap: 'quebrar' };
    const out = {};
    for (const [k, v] of Object.entries(style)) if (map[k]) out[map[k]] = v ?? false;
    if (Object.keys(out).length) this.lines.push(`formatar(${JSON.stringify(rangeTxt)}, ${JSON.stringify(out)});`);
  }
}

// ---------------------------------------------------------------------------
// Diálogo

export async function macrosDialog(app, { initialCode = null, vbaModules = null } = {}) {
  const wb = app.wb();
  wb.macros ||= [];
  let current = null;
  const list = el('div', { class: 'macro-list' });
  const name = el('input', { type: 'text', placeholder: 'Nome da macro' });
  const code = el('textarea', { class: 'macro-code', spellcheck: false, rows: 18, placeholder: '// Escreva sua macro em JavaScript.\n// Ex.: definir("A1", "Olá"); alerta(cel("A1"));' });
  const out = el('pre', { class: 'macro-out' });
  const log = (t) => { out.textContent += `${t}\n`; out.scrollTop = out.scrollHeight; };
  code.addEventListener('keydown', (e) => {
    if (e.key === 'Tab') {
      e.preventDefault();
      const { selectionStart: a, selectionEnd: b } = code;
      code.setRangeText('  ', a, b, 'end');
    }
    if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); run(); }
  });

  const draw = () => {
    list.replaceChildren(
      ...wb.macros.map((m) => el('button', { class: `macro-item${m === current ? ' on' : ''}`, onclick: () => pick(m) }, m.name || 'Sem nome')),
      el('button', { class: 'macro-item add', onclick: () => pick(null) }, '+ Nova macro'),
    );
  };
  const commitEdit = () => {
    const n = name.value.trim() || 'Macro sem nome';
    if (!code.value.trim()) return null;
    wb.transact('Editar macro', () => {
      wb.snapWorkbook();
      if (current && wb.macros.includes(current)) Object.assign(current, { name: n, code: code.value });
      else { current = { id: `m${Date.now().toString(36)}`, name: n, code: code.value }; wb.macros.push(current); }
    });
    // A transação recria o array; reaponta para o objeto atual
    current = wb.macros.find((m) => m.name === n && m.code === code.value) || current;
    app.dirty();
    draw();
    return current;
  };
  const pick = (m) => { current = m; name.value = m?.name || ''; code.value = m?.code || ''; out.textContent = ''; draw(); };
  const run = async () => {
    const m = commitEdit() || (code.value.trim() ? { name: name.value || 'Macro', code: code.value } : null);
    if (!m) { toast('Escreva o código da macro.'); return; }
    out.textContent = '';
    log(`▶ Executando "${m.name}"…`);
    const res = await runMacro(app, m, { onLog: log });
    if (res.error) log(`✖ Erro${res.line ? ` na linha ${res.line}` : ''}: ${res.error}`);
    else log(`✔ Concluída (${res.applied} alteraç${res.applied === 1 ? 'ão' : 'ões'}). Ctrl+Z desfaz.`);
    app.refresh();
  };
  const examples = select([['', 'Exemplos…'], ...MACRO_EXAMPLES.map((x, i) => [String(i), x.name])], '', {
    onchange: (e) => { const x = MACRO_EXAMPLES[+e.target.value]; if (x) { pick(null); name.value = x.name; code.value = x.code; } e.target.value = ''; },
  });
  const help = el('details', { class: 'macro-help' }, el('summary', {}, 'Referência rápida'), el('div', { html: `
<code>cel("A1")</code> lê um valor · <code>definir("A1", valor)</code> escreve (texto iniciado por <code>=</code> vira fórmula)<br>
<code>intervalo("A1:C10")</code> → matriz · <code>definirIntervalo("A1", [[1,2],[3,4]])</code> · <code>limpar("A1:B5")</code><br>
<code>formatar("A1:B1", { negrito: true, cor: "#c00", fundo: "#ffe", formato: "moeda", casas: 2 })</code><br>
<code>ultimaLinha("A")</code>, <code>ultimaColuna(1)</code>, <code>endereco(2, 3)</code> → "C2", <code>coluna(3)</code> → "C"<br>
<code>planilhas()</code>, <code>novaPlanilha("Resumo")</code>, <code>ativar("Resumo")</code>, <code>"Planilha2!A1"</code> acessa outra aba<br>
<code>selecao</code> (ex.: "B2:D9") · <code>log(...)</code> · <code>alerta(msg)</code> · <code>await recalcular()</code> atualiza fórmulas no meio da macro<br>
<span class="muted">As macros rodam isoladas: não acessam a internet nem seus arquivos. Tudo o que fazem pode ser desfeito com Ctrl+Z.</span>` }));
  const vba = vbaModules?.length ? el('details', { class: 'macro-help', open: true },
    el('summary', {}, `Código VBA do arquivo (${vbaModules.length} módulo${vbaModules.length > 1 ? 's' : ''})`),
    el('p', { class: 'muted' }, 'O VBA não roda no Lumina, mas fica preservado ao salvar em .xlsm. Use "Converter" para gerar uma macro equivalente em JavaScript (revise o resultado).'),
    ...vbaModules.map((mod) => el('div', { class: 'vba-mod' },
      el('div', { class: 'row' }, el('b', {}, mod.name), el('button', { class: 'btn small', onclick: async () => {
        const { vbaToJs } = await import('./vba.js');
        pick(null); name.value = `${mod.name} (convertida)`; code.value = vbaToJs(mod.code);
      } }, 'Converter para macro Lumina')),
      el('pre', { class: 'vba-code', html: escapeHtml(mod.code) })))) : null;

  if (initialCode !== null) { name.value = 'Macro gravada'; code.value = initialCode; }
  else if (wb.macros[0]) pick(wb.macros[0]);
  draw();
  await dialog({
    title: 'Macros', wide: true, okText: 'Fechar', cancelText: null,
    body: el('div', { class: 'macro-dlg' },
      el('div', { class: 'macro-side' }, list),
      el('div', { class: 'macro-main' },
        el('div', { class: 'row' }, name, examples,
          el('button', { class: 'btn primary', onclick: run }, '▶ Executar'),
          el('button', { class: 'btn', onclick: () => { if (commitEdit()) toast('Macro salva na planilha.'); } }, 'Salvar'),
          el('button', { class: 'btn', onclick: () => {
            if (!current) { pick(null); return; }
            wb.transact('Excluir macro', () => { wb.snapWorkbook(); wb.macros = wb.macros.filter((m) => m.id !== current.id); });
            app.dirty();
            pick(wb.macros[0] || null);
          } }, 'Excluir')),
        code, out, help, vba)),
  });
  stopMacro();
}
