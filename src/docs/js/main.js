// Lumina Texto — aplicação principal.
import {
  el, $, $$, toast, dialog, prompt, field, select, menubar, toolButton, sep, shortcuts, popover, closePopovers,
  colorPopover, dropdown, debounce, clamp, escapeHtml,
} from '../../shared/ui.js';
import { icon } from '../../shared/icons.js';
import {
  pickFiles, saveFile, exportPdf, isNative, registerApp, standaloneBoot, setTitle, addRecent, extOf, baseName,
  initTheme, pickImageDataUrl, fmtDate, fileName,
} from '../../shared/native.js';
import { Editor, History, placeCaret, paragraph, cleanHTML } from './editor.js';
import * as io from './io.js';
import { TEMPLATES } from './templates.js';

const editorEl = $('#editor');
const paper = $('#paper');
const scroller = $('#scroller');
const S = { docs: [], doc: null, zoom: 1, outline: true, side: null, darkPaper: false, seq: 1 };
const FONTS = [
  ['Calibri, Carlito, "Segoe UI", Arial, sans-serif', 'Calibri'], ['Arial, Helvetica, sans-serif', 'Arial'],
  ['"Segoe UI Variable", "Segoe UI", Inter, system-ui, sans-serif', 'Segoe UI'], ['Verdana, sans-serif', 'Verdana'],
  ['Tahoma, sans-serif', 'Tahoma'], ['"Trebuchet MS", sans-serif', 'Trebuchet MS'], ['Georgia, serif', 'Georgia'],
  ['"Times New Roman", Times, serif', 'Times New Roman'], ['Garamond, "EB Garamond", serif', 'Garamond'],
  ['"Book Antiqua", Palatino, serif', 'Palatino'], ['"Courier New", monospace', 'Courier New'],
  ['Consolas, "Cascadia Code", monospace', 'Consolas'], ['"Comic Sans MS", cursive', 'Comic Sans MS'], ['Impact, sans-serif', 'Impact'],
];
const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 28, 32, 36, 40, 48, 60, 72, 96];
const BLOCKS = [['p', 'Texto normal'], ['title', 'Título do documento'], ['subtitle', 'Subtítulo'], ['h1', 'Título 1'], ['h2', 'Título 2'], ['h3', 'Título 3'], ['h4', 'Título 4'], ['blockquote', 'Citação'], ['pre', 'Código']];

const editor = new Editor(editorEl, { onChange, onSelection: () => selectionChanged() });
editor.sanitize = io.sanitize;

// ---------------------------------------------------------------------------
// Documentos

function makeDoc({ title, html, page, style, versions, filePath = null, ext = 'lumdoc' } = {}) {
  return {
    id: S.seq++,
    title: title || 'Documento sem título',
    html: html || '<p><br></p>',
    history: null,
    page: { ...io.defaultPage(), ...(page || {}), margins: { ...io.defaultPage().margins, ...(page?.margins || {}) } },
    style: { ...io.defaultStyle(), ...(style || {}) },
    versions: versions || [],
    filePath,
    ext,
    dirty: false,
    scroll: 0,
  };
}

function isPristine(d) {
  return d && !d.dirty && !d.filePath && /^(<p><br><\/p>)?$/.test((d === S.doc ? editorEl.innerHTML : d.html).trim()) && d.title === 'Documento sem título';
}

function storeCurrent() {
  if (!S.doc) return;
  S.doc.html = editorEl.innerHTML;
  S.doc.history = editor.history;
  S.doc.scroll = scroller.scrollTop;
}

function addDoc(doc) {
  storeCurrent();
  if (S.docs.length === 1 && isPristine(S.docs[0])) S.docs = [];
  S.docs.push(doc);
  activate(doc, false);
}

function activate(doc, store = true) {
  if (store) storeCurrent();
  S.doc = doc;
  editor.setContent(doc.html, doc.history);
  doc.history = editor.history;
  applyLook();
  renderAll();
  requestAnimationFrame(() => { scroller.scrollTop = doc.scroll || 0; });
  updateToc();
  editorEl.focus({ preventScroll: true });
}

async function closeDoc(doc = S.doc) {
  if (doc.dirty && !(await askSave(doc))) return false;
  forgetRecovery(doc);
  const i = S.docs.indexOf(doc);
  S.docs.splice(i, 1);
  if (!S.docs.length) S.docs.push(makeDoc());
  activate(S.docs[Math.min(i, S.docs.length - 1)], false);
  return true;
}

async function askSave(doc) {
  if (doc !== S.doc) activate(doc);
  let choice = 'cancel';
  const ok = await dialog({
    title: 'Salvar alterações?',
    body: el('p', {}, `O documento "${doc.title}" tem alterações não salvas.`),
    okText: 'Salvar',
    extraButtons: [{ label: 'Não salvar', onClick: (close) => { choice = 'discard'; close(false); } }],
  });
  if (ok) return save(doc);
  return choice === 'discard';
}

function newDocument(tplId = 'blank') {
  const t = TEMPLATES[tplId] || TEMPLATES.blank;
  const page = io.defaultPage();
  if (t.margins) page.margins = { ...t.margins };
  const doc = makeDoc({ title: t.title === 'Documento sem título' ? t.title : `${t.title}`, html: t.html.trim(), page, style: io.defaultStyle(t.preset) });
  addDoc(doc);
  if (tplId !== 'blank') updateToc();
  placeCaret(editorEl.querySelector('p, h1') || editorEl, true);
}

async function openFiles(files) {
  await io.ensurePurify();
  for (const f of files) {
    try {
      const data = await io.importFile(f);
      const ext = extOf(f.name);
      const doc = makeDoc({ ...data, filePath: f.path, ext });
      addDoc(doc);
      addRecent({ app: 'docs', name: f.name, path: f.path });
      if (ext === 'docx') toast('Documento do Word aberto. Ao salvar, ele será gravado de volta em .docx.');
    } catch (err) {
      console.error(err);
      toast(`Não foi possível abrir ${f.name}: ${err.message}`);
    }
  }
}

async function openDialog() {
  const files = await pickFiles([
    { name: 'Documentos', extensions: ['lumdoc', 'docx', 'md', 'markdown', 'txt', 'html', 'htm'] },
    { name: 'Documento Lumina', extensions: ['lumdoc'] }, { name: 'Word', extensions: ['docx'] },
    { name: 'Markdown', extensions: ['md', 'markdown'] }, { name: 'Texto', extensions: ['txt'] }, { name: 'HTML', extensions: ['html', 'htm'] },
  ]);
  if (files.length) openFiles(files);
}

const cleanStr = (html) => {
  const d = document.createElement('div');
  d.innerHTML = html;
  return cleanHTML(d);
};

function currentHtml(doc = S.doc) {
  return doc === S.doc ? editor.html() : cleanStr(doc.html);
}

// Salva no formato nativo (ou de volta no formato original quando é editável).
async function save(doc = S.doc, saveAs = false) {
  storeCurrent();
  const html = currentHtml(doc);
  const keepFormat = !saveAs && doc.filePath && ['docx', 'md', 'markdown', 'html', 'htm', 'txt'].includes(doc.ext);
  try {
    let path;
    if (keepFormat) {
      path = await writeFormat(doc, html, doc.ext, doc.filePath);
    } else {
      const last = doc.versions[doc.versions.length - 1];
      if (!last || last.html !== html) doc.versions.push({ date: Date.now(), title: doc.title, html, auto: true });
      path = await saveFile(io.serialize(doc, html), {
        name: `${doc.title}.lumdoc`,
        filters: [{ name: 'Documento Lumina', extensions: ['lumdoc'] }],
        path: !saveAs && doc.ext === 'lumdoc' ? doc.filePath : null,
      });
      if (path) doc.ext = 'lumdoc';
    }
    if (!path) return false;
    if (isNative) doc.filePath = path;
    doc.dirty = false;
    forgetRecovery(doc);
    addRecent({ app: 'docs', name: fileName(path), path: isNative ? path : null });
    toast(isNative ? `Salvo: ${fileName(path)}` : 'Documento baixado.');
    renderAll();
    return true;
  } catch (err) {
    console.error(err);
    toast(`Erro ao salvar: ${err.message}`);
    return false;
  }
}

async function writeFormat(doc, html, fmt, path = null) {
  const name = `${doc.title}.${fmt === 'markdown' ? 'md' : fmt}`;
  const filters = (n, e) => [{ name: n, extensions: [e] }];
  switch (fmt) {
    case 'docx':
      return saveFile(await io.exportDocx(doc, html), { name, filters: filters('Documento do Word', 'docx'), path });
    case 'md': case 'markdown':
      return saveFile(io.htmlToMd(html), { name, filters: filters('Markdown', 'md'), path });
    case 'html': case 'htm':
      return saveFile(await io.standaloneHtml(doc, html), { name, filters: filters('Página HTML', 'html'), path });
    case 'txt':
      return saveFile(io.plainText(html), { name, filters: filters('Texto', 'txt'), path });
    default:
      throw new Error('Formato desconhecido');
  }
}

async function exportAs(fmt) {
  storeCurrent();
  const html = currentHtml();
  try {
    if (fmt === 'pdf') {
      const doc = S.doc;
      const out = await exportPdf(await io.standaloneHtml(doc, html, { forPrint: true }), {
        name: `${doc.title}.pdf`,
        options: doc.page.pageNumbers ? {
          displayHeaderFooter: true,
          headerTemplate: '<span></span>',
          footerTemplate: '<div style="width:100%;text-align:center;font-size:9px;color:#777;font-family:Arial"><span class="pageNumber"></span></div>',
        } : {},
      });
      if (out) toast(isNative ? `PDF exportado: ${fileName(out)}` : 'Use "Salvar como PDF" na janela de impressão.');
      return;
    }
    const out = await writeFormat(S.doc, html, fmt);
    if (out) toast(`Exportado: ${fileName(out)}`);
  } catch (err) {
    console.error(err);
    toast(`Erro ao exportar: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// Aparência da página

function applyLook() {
  const d = S.doc;
  const [w, h] = io.pageDims(d.page);
  const m = d.page.margins;
  paper.style.setProperty('--page-w', `${w}cm`);
  paper.style.setProperty('--page-h', `${h}cm`);
  paper.style.setProperty('--m-top', `${m.top}cm`);
  paper.style.setProperty('--m-right', `${m.right}cm`);
  paper.style.setProperty('--m-bottom', `${m.bottom}cm`);
  paper.style.setProperty('--m-left', `${m.left}cm`);
  paper.style.zoom = S.zoom;
  paper.classList.toggle('dark-paper', S.darkPaper);
  io.applyStyleVars(editorEl, d.style);
}

// ---------------------------------------------------------------------------
// Alterações, sumário, estrutura e estatísticas

function onChange() {
  if (S.doc && !S.doc.dirty) {
    S.doc.dirty = true;
    renderTitle();
    renderTabs();
  }
  editorEl.classList.toggle('empty', isEmpty());
  slowUpdate();
  checkSlash();
}

const isEmpty = () => editorEl.childElementCount <= 1 && !editorEl.textContent.replace(/​/g, '').trim() && !editorEl.querySelector('img, table, hr, .toc');

const slowUpdate = debounce(() => {
  updateToc();
  renderOutline();
  renderStatus();
  scheduleRecovery();
}, 300);

function headingId(h) {
  if (!h.id) h.id = `h-${Math.random().toString(36).slice(2, 9)}`;
  return h.id;
}

function updateToc() {
  const tocs = editorEl.querySelectorAll('.toc');
  if (!tocs.length) return;
  const heads = editor.headings();
  for (const toc of tocs) {
    const title = toc.querySelector('.toc-title')?.outerHTML || '<div class="toc-title">Sumário</div>';
    const links = heads.map((h) => `<a href="#${headingId(h)}" class="l${h.tagName[1]}">${escapeHtml(h.textContent.replace(/​/g, ''))}</a>`).join('');
    const next = title + (links || '<a class="l1" style="color:#9aa3b2">Adicione títulos (Título 1, 2, 3) para gerar o sumário.</a>');
    if (toc.innerHTML !== next) toc.innerHTML = next;
    toc.contentEditable = 'false';
  }
}

editorEl.addEventListener('click', (e) => {
  const a = e.target.closest('.toc a[href]');
  if (a) {
    e.preventDefault();
    document.getElementById(a.getAttribute('href').slice(1))?.scrollIntoView({ behavior: 'smooth', block: 'start' });
  }
});

function renderOutline() {
  const box = $('#outline');
  box.classList.toggle('collapsed', !S.outline);
  if (!S.outline) return;
  const heads = editor.headings();
  box.replaceChildren(
    el('div', { class: 'panel-head' }, 'Estrutura', el('button', { class: 'icon-btn', title: 'Ocultar', html: icon('x', 14), onclick: () => { S.outline = false; renderOutline(); } })),
    el('div', { class: 'panel-body', style: { gap: '1px' } },
      heads.length ? heads.map((h) => el('button', {
        class: `outline-item l${h.tagName[1]}`, title: h.textContent,
        onclick: () => { h.scrollIntoView({ behavior: 'smooth', block: 'start' }); placeCaret(h, false); },
      }, h.textContent.replace(/​/g, '') || '(sem texto)'))
        : el('div', { class: 'outline-empty' }, 'Os títulos do documento aparecerão aqui. Use Título 1, 2 ou 3, ou digite "# " no início de uma linha.')));
}

function stats() {
  const text = editor.text();
  const words = (text.match(/[\p{L}\p{N}'’-]+/gu) || []).length;
  const pageH = io.pageDims(S.doc.page)[1] * 37.795;
  const pages = Math.max(1, Math.ceil((paper.offsetHeight - 1) / pageH));
  return {
    words,
    chars: text.replace(/\n/g, '').length,
    charsNoSpace: text.replace(/\s/g, '').length,
    paragraphs: editorEl.querySelectorAll('p, li, h1, h2, h3, h4, blockquote').length,
    pages,
    reading: Math.max(1, Math.round(words / 200)),
    speaking: Math.max(1, Math.round(words / 130)),
  };
}

function renderStatus() {
  if (!S.doc) return;
  const s = stats();
  const zoom = el('input', { type: 'range', min: 50, max: 200, step: 10, value: Math.round(S.zoom * 100) });
  zoom.addEventListener('input', () => setZoom(zoom.value / 100));
  $('#status').replaceChildren(
    el('span', {}, `${s.pages} ${s.pages === 1 ? 'página' : 'páginas'}`),
    el('button', { class: 'status-btn', title: 'Contagem de palavras', onclick: wordCount }, `${s.words.toLocaleString('pt-BR')} palavras`),
    el('span', {}, `${s.chars.toLocaleString('pt-BR')} caracteres`),
    el('span', {}, `Leitura: ${s.reading} min`),
    el('span', { class: 'spacer' }),
    el('span', {}, 'Português (Brasil)'),
    el('button', { class: 'status-btn', title: 'Reduzir', onclick: () => setZoom(S.zoom - 0.1) }, '−'),
    zoom,
    el('button', { class: 'status-btn', title: 'Ampliar', onclick: () => setZoom(S.zoom + 0.1) }, '+'),
    el('button', { class: 'status-btn', title: 'Voltar para 100%', onclick: () => setZoom(1) }, `${Math.round(S.zoom * 100)}%`));
}

function setZoom(z) {
  S.zoom = clamp(Math.round(z * 10) / 10, 0.5, 2);
  paper.style.zoom = S.zoom;
  renderStatus();
}

function wordCount() {
  const s = stats();
  const rows = [['Páginas', s.pages], ['Palavras', s.words], ['Caracteres (com espaços)', s.chars], ['Caracteres (sem espaços)', s.charsNoSpace],
    ['Parágrafos', s.paragraphs], ['Tempo de leitura', `${s.reading} min`], ['Tempo de fala', `${s.speaking} min`]];
  dialog({
    title: 'Contagem de palavras', cancelText: null, width: 380,
    body: el('div', { class: 'stats-grid' }, ...rows.flatMap(([k, v]) => [el('span', {}, k), el('b', {}, typeof v === 'number' ? v.toLocaleString('pt-BR') : v)])),
  });
}

// ---------------------------------------------------------------------------
// Cabeçalho, abas e menus

function renderTitle() {
  const d = S.doc;
  const t = $('#doc-title');
  if (document.activeElement !== t) t.value = d.title;
  $('#save-state').textContent = d.dirty ? 'Alterações não salvas' : d.filePath ? `Salvo · ${fileName(d.filePath)}` : 'Não salvo ainda';
  setTitle(`${d.title}${d.dirty ? ' •' : ''} — Lumina Texto`);
}

function renderTabs() {
  const box = $('#tabs');
  box.classList.toggle('hidden', S.docs.length < 2);
  box.replaceChildren(...S.docs.map((d) => el('div', {
    class: `doc-tab${d === S.doc ? ' on' : ''}`, title: d.filePath || d.title,
    onclick: () => d !== S.doc && activate(d),
    onauxclick: (e) => { if (e.button === 1) closeDoc(d); },
  }, el('span', { class: 'nm' }, `${d.title}${d.dirty ? ' •' : ''}`),
  el('button', { class: 'x', title: 'Fechar', html: icon('x', 12), onclick: (e) => { e.stopPropagation(); closeDoc(d); } }))));
}

function renderAll() {
  renderTitle();
  renderTabs();
  renderOutline();
  renderStatus();
  renderSide();
  refreshToolbar();
  editorEl.classList.toggle('empty', isEmpty());
}

function setupHeader() {
  $('#badge').innerHTML = icon('doc', 20);
  const t = $('#doc-title');
  t.addEventListener('input', () => {
    S.doc.title = t.value || 'Documento sem título';
    S.doc.dirty = true;
    renderTabs();
    $('#save-state').textContent = 'Alterações não salvas';
  });
  t.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); editorEl.focus(); } });
  t.addEventListener('blur', () => renderTitle());
  const exp = el('button', { class: 'btn', html: `${icon('download', 16)} Exportar`, 'data-popover-anchor': '' });
  exp.addEventListener('click', () => dropdown(exp, exportItems()));
  $('#header-actions').append(
    el('button', { class: 'icon-btn', title: 'Modo foco (F11)', html: icon('focus', 18), onclick: toggleFocus }),
    exp,
    el('button', { class: 'btn primary', html: `${icon('save', 16)} Salvar`, onclick: () => save() }));
  document.body.append(el('button', { class: 'btn focus-exit', onclick: toggleFocus, html: `${icon('x', 14)} Sair do modo foco` }));
}

const exportItems = () => [
  { label: 'PDF (.pdf)', icon: 'pdf', action: () => exportAs('pdf'), key: 'Ctrl+P' },
  { label: 'Word (.docx)', icon: 'doc', action: () => exportAs('docx') },
  { label: 'Página web (.html)', icon: 'code', action: () => exportAs('html') },
  { label: 'Markdown (.md)', icon: 'text', action: () => exportAs('md') },
  { label: 'Texto simples (.txt)', icon: 'alignLeft', action: () => exportAs('txt') },
];

function setupMenus() {
  menubar($('#menus'), [
    ['Arquivo', () => [
      { label: 'Novo documento', icon: 'plus', key: 'Ctrl+N', action: () => newDocument() },
      { label: 'Novo a partir de modelo', icon: 'layout', sub: Object.entries(TEMPLATES).map(([id, t]) => ({ label: id === 'blank' ? 'Em branco' : t.title, action: () => newDocument(id) })) },
      { label: 'Abrir…', icon: 'open', key: 'Ctrl+O', action: openDialog },
      '-',
      { label: 'Salvar', icon: 'save', key: 'Ctrl+S', action: () => save() },
      { label: 'Salvar como…', key: 'Ctrl+Shift+S', action: () => save(S.doc, true) },
      { label: 'Exportar', icon: 'download', sub: exportItems() },
      '-',
      { label: 'Histórico de versões…', icon: 'history', action: versionsDialog },
      { label: 'Configurar página…', icon: 'settings', action: pageSetup },
      { label: 'Imprimir / PDF', icon: 'printer', key: 'Ctrl+P', action: () => exportAs('pdf') },
      '-',
      { label: 'Fechar documento', key: 'Ctrl+W', action: () => closeDoc() },
    ]],
    ['Editar', () => [
      { label: 'Desfazer', icon: 'undo', key: 'Ctrl+Z', action: () => editor.undo() },
      { label: 'Refazer', icon: 'redo', key: 'Ctrl+Y', action: () => editor.redo() },
      '-',
      { label: 'Recortar', key: 'Ctrl+X', action: () => { editor.focus(); document.execCommand('cut'); } },
      { label: 'Copiar', key: 'Ctrl+C', action: () => { editor.focus(); document.execCommand('copy'); } },
      { label: 'Colar sem formatação', key: 'Ctrl+Shift+V', action: pastePlain },
      { label: 'Selecionar tudo', key: 'Ctrl+A', action: () => { editor.focus(); document.execCommand('selectAll'); } },
      '-',
      { label: 'Localizar', icon: 'search', key: 'Ctrl+F', action: () => openSide('find') },
      { label: 'Localizar e substituir', key: 'Ctrl+H', action: () => openSide('find', true) },
    ]],
    ['Inserir', () => [
      { label: 'Imagem…', icon: 'image', action: insertImage },
      { label: 'Tabela', icon: 'table', sub: [2, 3, 4, 5].map((c) => ({ label: `${c} colunas × 3 linhas`, action: () => editor.insertTable(3, c) })) },
      { label: 'Link…', icon: 'link', key: 'Ctrl+K', action: insertLink },
      '-',
      { label: 'Sumário automático', icon: 'outline', action: insertToc },
      { label: 'Destaque', icon: 'sparkles', sub: CALLOUTS.map(([k, ic, l]) => ({ label: `${ic} ${l}`, action: () => insertCallout(k, ic) })) },
      { label: 'Lista de tarefas', icon: 'checklist', action: () => editor.list('check') },
      { label: 'Citação', icon: 'quote', action: () => editor.setBlock('blockquote') },
      { label: 'Bloco de código', icon: 'code', action: () => editor.setBlock('pre') },
      '-',
      { label: 'Linha horizontal', icon: 'hr', action: () => editor.insertBlock(document.createElement('hr'), { focusInside: false }) },
      { label: 'Quebra de página', icon: 'pageBreak', key: 'Ctrl+Enter', action: insertPageBreak },
      { label: 'Data e hora', icon: 'calendar', sub: dateItems() },
      { label: 'Caracteres especiais…', action: specialChars },
    ]],
    ['Formatar', () => [
      { label: 'Negrito', icon: 'bold', key: 'Ctrl+B', action: () => editor.exec('bold') },
      { label: 'Itálico', icon: 'italic', key: 'Ctrl+I', action: () => editor.exec('italic') },
      { label: 'Sublinhado', icon: 'underline', key: 'Ctrl+U', action: () => editor.exec('underline') },
      { label: 'Tachado', icon: 'strike', key: 'Alt+Shift+5', action: () => editor.exec('strikeThrough') },
      { label: 'Sobrescrito', icon: 'sup', key: 'Ctrl+.', action: () => editor.exec('superscript') },
      { label: 'Subscrito', icon: 'sub', key: 'Ctrl+,', action: () => editor.exec('subscript') },
      '-',
      { label: 'Estilos de parágrafo', sub: BLOCKS.map(([v, l], i) => ({ label: l, key: i <= 5 && v.startsWith('h') ? `Ctrl+Alt+${v[1]}` : v === 'p' ? 'Ctrl+Alt+0' : undefined, action: () => editor.setBlock(v) })) },
      { label: 'Alinhamento', sub: [['left', 'Esquerda', 'Ctrl+L'], ['center', 'Centralizado', 'Ctrl+E'], ['right', 'Direita', 'Ctrl+R'], ['justify', 'Justificado', 'Ctrl+J']].map(([a, l, k]) => ({ label: l, key: k, icon: `align${a[0].toUpperCase()}${a.slice(1)}`, action: () => editor.setAlign(a) })) },
      { label: 'Espaçamento entre linhas', sub: LINE_SPACINGS.map(([v, l]) => ({ label: l, action: () => editor.setBlockStyle('lineHeight', v) })) },
      { label: 'Maiúsculas e minúsculas', sub: [['upper', 'MAIÚSCULAS'], ['lower', 'minúsculas'], ['title', 'Primeira Letra Maiúscula'], ['sentence', 'Frase normal']].map(([k, l]) => ({ label: l, action: () => changeCase(k) })) },
      '-',
      { label: 'Limpar formatação', icon: 'clearFormat', key: 'Ctrl+\\', action: () => editor.clearFormat() },
      { label: 'Estilo do documento…', icon: 'palette', action: () => openSide('style') },
    ]],
    ['Exibir', () => [
      { label: 'Estrutura do documento', icon: 'outline', checked: S.outline, action: () => { S.outline = !S.outline; renderOutline(); } },
      { label: 'Papel escuro (modo noturno)', icon: 'moon', checked: S.darkPaper, action: () => { S.darkPaper = !S.darkPaper; applyLook(); } },
      { label: 'Modo foco', icon: 'focus', key: 'F11', action: toggleFocus },
      '-',
      { label: 'Ampliar', icon: 'zoomIn', action: () => setZoom(S.zoom + 0.1) },
      { label: 'Reduzir', icon: 'zoomOut', action: () => setZoom(S.zoom - 0.1) },
      { label: 'Tamanho real (100%)', action: () => setZoom(1) },
    ]],
    ['Ferramentas', () => [
      { label: 'Contagem de palavras', icon: 'sigma', key: 'Ctrl+Shift+C', action: wordCount },
      { label: 'Localizar e substituir', icon: 'search', key: 'Ctrl+H', action: () => openSide('find', true) },
      { label: 'Verificação ortográfica', checked: editorEl.spellcheck, action: () => { editorEl.spellcheck = !editorEl.spellcheck; editorEl.blur(); editorEl.focus(); } },
      { label: 'Histórico de versões…', icon: 'history', action: versionsDialog },
    ]],
    ['Ajuda', () => [
      { label: 'Atalhos e recursos', icon: 'info', key: 'F1', action: helpDialog },
    ]],
  ]);
}

// ---------------------------------------------------------------------------
// Barra de ferramentas

const LINE_SPACINGS = [['1', 'Simples (1,0)'], ['1.15', '1,15'], ['1.5', '1,5'], ['2', 'Duplo (2,0)'], ['2.5', '2,5'], ['3', '3,0'], ['', 'Padrão do documento']];
const CALLOUTS = [['', '💡', 'Informação'], ['warning', '⚠️', 'Atenção'], ['success', '✅', 'Sucesso'], ['danger', '⛔', 'Importante']];
const tb = {};

function setupToolbar() {
  const bar = $('#toolbar');
  tb.style = select(BLOCKS, 'p', { class: 'tb-select style-sel', title: 'Estilo do parágrafo' });
  tb.style.addEventListener('change', () => editor.setBlock(tb.style.value));
  tb.font = select([['', '(Fonte do documento)'], ...FONTS.map(([v, l]) => [v, l])], '', { class: 'tb-select font-sel', title: 'Fonte' });
  tb.font.addEventListener('change', () => { if (tb.font.value) editor.setFontFamily(tb.font.value); });
  tb.size = el('input', { class: 'tb-select size-sel', list: 'size-list', title: 'Tamanho da fonte (pt)', value: '11' });
  const sizeList = el('datalist', { id: 'size-list' }, ...SIZES.map((s) => el('option', { value: s })));
  tb.size.addEventListener('change', () => {
    const v = parseFloat(String(tb.size.value).replace(',', '.'));
    if (v > 0 && v < 400) editor.setFontSize(v);
  });
  tb.size.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); tb.size.dispatchEvent(new Event('change')); editor.focus(); } });
  for (const x of [tb.style, tb.font, tb.size]) x.addEventListener('mousedown', () => editor.rememberRange());
  const stepSize = (dir) => {
    const cur = parseFloat(tb.size.value) || 11;
    const next = dir > 0 ? SIZES.find((s) => s > cur) || cur + 4 : [...SIZES].reverse().find((s) => s < cur) || Math.max(1, cur - 1);
    editor.setFontSize(next);
    tb.size.value = next;
  };

  const colorBtn = (ic, title, onPick, opts) => {
    const b = toolButton(ic, title, () => colorPopover(b, (c) => { editor.focus(); onPick(c); if (c) b.querySelector('.swatch-bar').style.background = c; }, opts));
    b.setAttribute('data-popover-anchor', '');
    b.append(el('span', { class: 'swatch-bar' }));
    b.addEventListener('mousedown', () => editor.rememberRange());
    return b;
  };
  tb.color = colorBtn('textColor', 'Cor do texto', (c) => editor.setColor(c || '#1f2328'), { allowNone: true, noneLabel: 'Automática' });
  tb.color.querySelector('.swatch-bar').style.background = '#e03131';
  tb.hl = colorBtn('highlight', 'Cor de realce', (c) => editor.setHighlight(c), { allowNone: true, noneLabel: 'Sem realce' });
  tb.hl.querySelector('.swatch-bar').style.background = '#ffe066';

  const lineBtn = toolButton('alignJustify', 'Espaçamento entre linhas', () => dropdown(lineBtn, LINE_SPACINGS.map(([v, l]) => ({ label: l, action: () => editor.setBlockStyle('lineHeight', v) }))), { label: '' });
  lineBtn.innerHTML = `${icon('chevronDown', 12)}${icon('alignLeft', 18)}`;
  lineBtn.setAttribute('data-popover-anchor', '');
  const tableBtn = toolButton('table', 'Inserir tabela', () => tablePicker(tableBtn));
  tableBtn.setAttribute('data-popover-anchor', '');
  const plusBtn = toolButton('plus', 'Inserir', () => dropdown(plusBtn, [
    { label: 'Sumário automático', icon: 'outline', action: insertToc },
    { label: 'Destaque (caixa colorida)', icon: 'sparkles', action: () => insertCallout('', '💡') },
    { label: 'Bloco de código', icon: 'code', action: () => editor.setBlock('pre') },
    { label: 'Citação', icon: 'quote', action: () => editor.setBlock('blockquote') },
    { label: 'Linha horizontal', icon: 'hr', action: () => editor.insertBlock(document.createElement('hr'), { focusInside: false }) },
    { label: 'Quebra de página', icon: 'pageBreak', action: insertPageBreak },
    { label: 'Data de hoje', icon: 'calendar', action: () => editor.insertHTML(escapeHtml(new Date().toLocaleDateString('pt-BR'))) },
    { label: 'Caracteres especiais…', icon: 'more', action: specialChars },
  ]), { label: 'Inserir' });
  plusBtn.setAttribute('data-popover-anchor', '');

  tb.bold = toolButton('bold', 'Negrito', () => editor.exec('bold'), { key: 'Ctrl+B' });
  tb.italic = toolButton('italic', 'Itálico', () => editor.exec('italic'), { key: 'Ctrl+I' });
  tb.underline = toolButton('underline', 'Sublinhado', () => editor.exec('underline'), { key: 'Ctrl+U' });
  tb.strike = toolButton('strike', 'Tachado', () => editor.exec('strikeThrough'));
  tb.left = toolButton('alignLeft', 'Alinhar à esquerda', () => editor.setAlign('left'), { key: 'Ctrl+L' });
  tb.center = toolButton('alignCenter', 'Centralizar', () => editor.setAlign('center'), { key: 'Ctrl+E' });
  tb.right = toolButton('alignRight', 'Alinhar à direita', () => editor.setAlign('right'), { key: 'Ctrl+R' });
  tb.justify = toolButton('alignJustify', 'Justificar', () => editor.setAlign('justify'), { key: 'Ctrl+J' });
  tb.ul = toolButton('list', 'Lista com marcadores', () => editor.list('bullet'), { key: 'Ctrl+Shift+8' });
  tb.ol = toolButton('listOrdered', 'Lista numerada', () => editor.list('number'), { key: 'Ctrl+Shift+7' });
  tb.check = toolButton('checklist', 'Lista de tarefas', () => editor.list('check'), { key: 'Ctrl+Shift+9' });

  bar.append(
    toolButton('undo', 'Desfazer', () => editor.undo(), { key: 'Ctrl+Z' }),
    toolButton('redo', 'Refazer', () => editor.redo(), { key: 'Ctrl+Y' }),
    sep(), tb.style, sep(), tb.font,
    toolButton('zoomOut', 'Diminuir fonte', () => stepSize(-1)), tb.size, sizeList, toolButton('zoomIn', 'Aumentar fonte', () => stepSize(1)),
    sep(), tb.bold, tb.italic, tb.underline, tb.strike, tb.color, tb.hl,
    sep(), toolButton('link', 'Inserir link', insertLink, { key: 'Ctrl+K' }), toolButton('image', 'Inserir imagem', insertImage), tableBtn,
    sep(), tb.left, tb.center, tb.right, tb.justify, lineBtn,
    sep(), tb.ul, tb.ol, tb.check, toolButton('outdent', 'Diminuir recuo', () => editor.indent(-1)), toolButton('indent', 'Aumentar recuo', () => editor.indent(1)),
    sep(), toolButton('clearFormat', 'Limpar formatação', () => editor.clearFormat(), { key: 'Ctrl+\\' }), plusBtn,
    el('span', { class: 'spacer' }),
    toolButton('palette', 'Estilo do documento', () => openSide('style'), { label: 'Estilos' }),
    toolButton('search', 'Localizar e substituir', () => openSide('find'), { key: 'Ctrl+F' }),
    toolButton('outline', 'Estrutura', () => { S.outline = !S.outline; renderOutline(); }));
}

const refreshToolbar = () => requestAnimationFrame(() => {
  if (!editor.hasFocus()) return;
  const q = (c) => { try { return document.queryCommandState(c); } catch { return false; } };
  for (const [k, c] of [['bold', 'bold'], ['italic', 'italic'], ['underline', 'underline'], ['strike', 'strikeThrough'], ['left', 'justifyLeft'], ['center', 'justifyCenter'], ['right', 'justifyRight'], ['justify', 'justifyFull'], ['ol', 'insertOrderedList']]) {
    tb[k].classList.toggle('on', q(c));
  }
  const inUl = editor.closest('UL');
  tb.ul.classList.toggle('on', !!inUl && !inUl.classList.contains('checklist'));
  tb.check.classList.toggle('on', !!inUl?.classList.contains('checklist'));
  tb.style.value = editor.blockType();
  let node = editor.range()?.startContainer;
  if (node?.nodeType === 3) node = node.parentElement;
  if (node && editorEl.contains(node)) {
    const cs = getComputedStyle(node);
    const pt = Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2;
    if (document.activeElement !== tb.size) tb.size.value = String(pt).replace('.', ',');
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const match = FONTS.find(([v]) => v.split(',')[0].replace(/["']/g, '').trim() === fam);
    const docFam = (S.doc.style.font || '').split(',')[0].replace(/["']/g, '').trim();
    tb.font.value = match && fam !== docFam ? match[0] : '';
  }
});

// ---------------------------------------------------------------------------
// Inserções

async function insertImage() {
  editor.rememberRange();
  const src = await pickImageDataUrl();
  if (src) { editor.focus(); editor.insertImage(src); }
}

async function insertLink() {
  editor.rememberRange();
  const existing = editor.closest('A');
  const sel = getSelection().toString();
  const url = el('input', { type: 'url', value: existing?.getAttribute('href') || '', placeholder: 'https://exemplo.com' });
  const text = el('input', { type: 'text', value: existing?.textContent || sel });
  const ok = await dialog({
    title: existing ? 'Editar link' : 'Inserir link', width: 440,
    body: [field('Endereço (URL)', url), field('Texto exibido', text)],
    extraButtons: existing ? [{ label: 'Remover link', onClick: (close) => { close(false); editor.focus(); document.execCommand('unlink'); editor.changed(true); } }] : [],
  });
  if (!ok || !url.value.trim()) return;
  let href = url.value.trim();
  if (!/^(https?:|mailto:|#|tel:)/i.test(href)) href = /@/.test(href) && !href.includes('/') ? `mailto:${href}` : `https://${href}`;
  editor.focus();
  if (existing) {
    existing.setAttribute('href', href);
    if (text.value && text.value !== existing.textContent) existing.textContent = text.value;
    editor.changed(true);
  } else if (sel) {
    editor.insertLink(href);
  } else {
    editor.insertLink(href, text.value || href);
  }
}

function insertToc() {
  const toc = el('div', { class: 'toc', contenteditable: 'false' }, el('div', { class: 'toc-title' }, 'Sumário'));
  editor.insertBlock(toc, { focusInside: false });
  updateToc();
}

function insertCallout(kind, emoji) {
  const c = el('div', { class: `callout ${kind}`.trim() },
    el('span', { class: 'callout-icon', contenteditable: 'false' }, emoji),
    el('div', { class: 'callout-body' }, el('p', { html: '<br>' })));
  editor.insertBlock(c);
}

function insertPageBreak() {
  editor.insertBlock(el('div', { class: 'page-break', contenteditable: 'false' }), { focusInside: false });
}

function dateItems() {
  const d = new Date();
  return [
    d.toLocaleDateString('pt-BR'),
    d.toLocaleDateString('pt-BR', { day: 'numeric', month: 'long', year: 'numeric' }),
    d.toLocaleDateString('pt-BR', { weekday: 'long', day: 'numeric', month: 'long', year: 'numeric' }),
    `${d.toLocaleDateString('pt-BR')} ${d.toLocaleTimeString('pt-BR', { hour: '2-digit', minute: '2-digit' })}`,
  ].map((s) => ({ label: s, action: () => editor.insertHTML(escapeHtml(s)) }));
}

function tablePicker(anchor) {
  editor.rememberRange();
  let rows = 0, cols = 0;
  const cells = [];
  const label = el('div', { class: 'table-label' }, 'Escolha o tamanho');
  const grid = el('div', { class: 'table-grid' });
  for (let r = 0; r < 8; r++) for (let c = 0; c < 10; c++) {
    const s = el('span');
    s.addEventListener('mouseenter', () => {
      rows = r + 1; cols = c + 1;
      cells.forEach((x) => x.el.classList.toggle('on', x.r < rows && x.c < cols));
      label.textContent = `${cols} × ${rows}`;
    });
    s.addEventListener('click', () => { closePopovers(); editor.focus(); editor.insertTable(rows, cols); });
    cells.push({ el: s, r, c });
    grid.append(s);
  }
  popover(el('div', { class: 'table-picker' }, grid, label), anchor);
}

const SPECIAL = '©®™§¶†‡•·…–—‘’“”«»€£¥¢$°±×÷≠≤≥≈∞√∑∏∫∂πΩµαβγδεθλσφψω→←↑↓↔⇒⇔✓✔✗✘★☆♥♦♣♠☐☑☒½¼¾⅓⅔¹²³ªº№℃‰¿¡';
function specialChars() {
  editor.rememberRange();
  let close;
  const grid = el('div', { class: 'chars-grid' }, ...[...SPECIAL].map((ch) => el('button', {
    title: `U+${ch.codePointAt(0).toString(16).toUpperCase().padStart(4, '0')}`,
    onclick: () => { close(true); editor.focus(); document.execCommand('insertText', false, ch); },
  }, ch)));
  dialog({ title: 'Caracteres especiais', body: grid, okText: null, cancelText: 'Fechar', width: 520, onOpen: (_b, c) => { close = c; } });
}

function changeCase(kind) {
  editor.focus();
  const s = getSelection();
  const t = s.toString();
  if (!t) { toast('Selecione o texto primeiro.'); return; }
  const out = kind === 'upper' ? t.toLocaleUpperCase('pt-BR')
    : kind === 'lower' ? t.toLocaleLowerCase('pt-BR')
      : kind === 'title' ? t.toLocaleLowerCase('pt-BR').replace(/(^|[\s(“"'-])(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('pt-BR'))
        : t.toLocaleLowerCase('pt-BR').replace(/(^\s*|[.!?]\s+)(\p{L})/gu, (m, a, b) => a + b.toLocaleUpperCase('pt-BR'));
  document.execCommand('insertText', false, out);
  editor.changed(true);
}

async function pastePlain() {
  try {
    const text = await navigator.clipboard.readText();
    editor.focus();
    document.execCommand('insertText', false, text);
    editor.changed(true);
  } catch {
    toast('Não foi possível ler a área de transferência.');
  }
}

// ---------------------------------------------------------------------------
// Comandos de barra "/"

const SLASH = [
  ['p', 'Texto', 'Parágrafo simples', 'text', () => editor.setBlock('p')],
  ['h1', 'Título 1', 'Seção principal', 'text', () => editor.setBlock('h1')],
  ['h2', 'Título 2', 'Subseção', 'text', () => editor.setBlock('h2')],
  ['h3', 'Título 3', 'Tópico menor', 'text', () => editor.setBlock('h3')],
  ['ul', 'Lista com marcadores', 'Itens simples', 'list', () => editor.list('bullet')],
  ['ol', 'Lista numerada', 'Passos em ordem', 'listOrdered', () => editor.list('number')],
  ['todo', 'Lista de tarefas', 'Caixas de seleção', 'checklist', () => editor.list('check')],
  ['table', 'Tabela', 'Grade 3 × 3', 'table', () => editor.insertTable(3, 3)],
  ['image', 'Imagem', 'Do computador', 'image', insertImage],
  ['quote', 'Citação', 'Destaque de fala', 'quote', () => editor.setBlock('blockquote')],
  ['code', 'Código', 'Bloco monoespaçado', 'code', () => editor.setBlock('pre')],
  ['callout', 'Destaque', 'Caixa de informação', 'sparkles', () => insertCallout('', '💡')],
  ['warning', 'Aviso', 'Caixa de atenção', 'sparkles', () => insertCallout('warning', '⚠️')],
  ['toc', 'Sumário', 'Gerado dos títulos', 'outline', insertToc],
  ['hr', 'Divisor', 'Linha horizontal', 'hr', () => editor.insertBlock(document.createElement('hr'), { focusInside: false })],
  ['page', 'Quebra de página', 'Começa nova página', 'pageBreak', insertPageBreak],
  ['date', 'Data de hoje', new Date().toLocaleDateString('pt-BR'), 'calendar', () => editor.insertHTML(escapeHtml(new Date().toLocaleDateString('pt-BR')))],
];
const slash = { el: null, items: [], index: 0, block: null };
const norm = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase();

function checkSlash() {
  const r = editor.range();
  const b = editor.currentBlock();
  const t = b?.textContent || '';
  if (!r || !r.collapsed || !b || b.tagName !== 'P' || !/^\/[\p{L}\d ]{0,20}$/u.test(t) || /\s{2}/.test(t)) { hideSlash(); return; }
  const q = norm(t.slice(1).trim());
  slash.items = SLASH.filter(([id, name, desc]) => !q || norm(`${id} ${name} ${desc}`).includes(q));
  slash.block = b;
  slash.index = Math.min(slash.index, Math.max(0, slash.items.length - 1));
  const rect = r.getBoundingClientRect();
  if (!slash.el) {
    slash.el = el('div', { class: 'slash-menu' });
    document.body.append(slash.el);
  }
  slash.el.replaceChildren(...(slash.items.length ? slash.items.map(([, name, desc, ic], i) => el('div', {
    class: `slash-item${i === slash.index ? ' on' : ''}`,
    onmousedown: (e) => { e.preventDefault(); runSlash(i); },
    onmouseenter: () => { slash.index = i; slash.el.querySelectorAll('.slash-item').forEach((x, j) => x.classList.toggle('on', j === i)); },
  }, el('span', { class: 'si-icon', html: icon(ic, 18) }), el('div', {}, el('div', { class: 'si-name' }, name), el('div', { class: 'si-desc' }, desc))))
    : [el('div', { class: 'slash-empty' }, 'Nenhum comando encontrado')]));
  const top = rect.bottom + 6 + 340 > innerHeight ? Math.max(8, rect.top - 346) : rect.bottom + 6;
  Object.assign(slash.el.style, { left: `${Math.min(rect.left, innerWidth - 300)}px`, top: `${top}px` });
  slash.el.querySelector('.slash-item.on')?.scrollIntoView({ block: 'nearest' });
}

function hideSlash() {
  slash.el?.remove();
  slash.el = null;
  slash.index = 0;
}

function runSlash(i) {
  const item = slash.items[i];
  const b = slash.block;
  hideSlash();
  if (!item || !b) return;
  b.innerHTML = '<br>';
  placeCaret(b, true);
  item[4]();
}

editorEl.addEventListener('keydown', (e) => {
  if (!slash.el) return;
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
    e.preventDefault();
    e.stopImmediatePropagation();
    const n = slash.items.length || 1;
    slash.index = (slash.index + (e.key === 'ArrowDown' ? 1 : -1) + n) % n;
    checkSlash();
  } else if (e.key === 'Enter' || e.key === 'Tab') {
    if (!slash.items.length) return;
    e.preventDefault();
    e.stopImmediatePropagation();
    runSlash(slash.index);
  } else if (e.key === 'Escape') {
    e.preventDefault();
    hideSlash();
  }
}, true);

// ---------------------------------------------------------------------------
// Menus flutuantes: seleção, imagem e tabela

let bubble = null;
function selectionChanged() {
  editor.rememberRange();
  refreshToolbar();
  markActiveBlock();
  const r = editor.range();
  if (r && !r.collapsed && r.toString().trim() && !document.body.classList.contains('focus-mode')) showBubble(r);
  else hideBubble();
  updateTableTools();
}

function showBubble(r) {
  const rect = r.getBoundingClientRect();
  if (!bubble) {
    bubble = el('div', { class: 'bubble' },
      toolButton('bold', 'Negrito', () => editor.exec('bold')),
      toolButton('italic', 'Itálico', () => editor.exec('italic')),
      toolButton('underline', 'Sublinhado', () => editor.exec('underline')),
      toolButton('strike', 'Tachado', () => editor.exec('strikeThrough')),
      toolButton('code', 'Código', () => wrapCode()),
      toolButton('link', 'Link', insertLink),
      toolButton('textColor', 'Cor', function onColor() { colorPopover(this, (c) => { editor.focus(); editor.setColor(c || '#1f2328'); }, { allowNone: true, noneLabel: 'Automática' }); }),
      toolButton('highlight', 'Realce', function onHl() { colorPopover(this, (c) => { editor.focus(); editor.setHighlight(c); }, { allowNone: true, noneLabel: 'Sem realce' }); }),
      toolButton('text', 'Título 2', () => editor.setBlock('h2')),
      toolButton('clearFormat', 'Limpar formatação', () => editor.clearFormat()));
    bubble.querySelectorAll('.tb-btn').forEach((b) => b.setAttribute('data-popover-anchor', ''));
    document.body.append(bubble);
  }
  const w = bubble.offsetWidth;
  const top = rect.top - 46 < 8 ? rect.bottom + 8 : rect.top - 46;
  bubble.style.left = `${clamp(rect.left + rect.width / 2 - w / 2, 8, innerWidth - w - 8)}px`;
  bubble.style.top = `${top}px`;
}

function hideBubble() {
  bubble?.remove();
  bubble = null;
}

function wrapCode() {
  const r = editor.range();
  if (!r || r.collapsed) return;
  const code = document.createElement('code');
  code.textContent = r.toString();
  r.deleteContents();
  r.insertNode(code);
  editor.changed(true);
}

function markActiveBlock() {
  const top = editor.topBlock(editor.range()?.startContainer);
  if (top === S.activeBlock) return;
  S.activeBlock?.classList.remove('active-block');
  S.activeBlock = top;
  if (top && document.body.classList.contains('focus-mode')) top.classList.add('active-block');
}

let imgTools = null, imgHandle = null, selectedImg = null;
editorEl.addEventListener('mousedown', (e) => {
  if (e.target.tagName === 'IMG') {
    selectImage(e.target);
  } else if (selectedImg && !e.target.closest('.img-tools')) {
    deselectImage();
  }
});

function selectImage(img) {
  deselectImage();
  selectedImg = img;
  img.classList.add('img-selected');
  const btn = (label, title, fn) => el('button', { class: 'tb-btn', title, onmousedown: (e) => e.preventDefault(), onclick: () => { fn(); positionImageTools(); editor.changed(true); } }, label);
  imgTools = el('div', { class: 'img-tools' },
    el('button', { class: 'tb-btn', title: 'À esquerda (texto ao redor)', html: icon('hAlignLeft', 16), onmousedown: (e) => e.preventDefault(), onclick: () => { setImgAlign('align-left'); } }),
    el('button', { class: 'tb-btn', title: 'Centralizada', html: icon('hAlignCenter', 16), onmousedown: (e) => e.preventDefault(), onclick: () => { setImgAlign('align-center'); } }),
    el('button', { class: 'tb-btn', title: 'À direita (texto ao redor)', html: icon('hAlignRight', 16), onmousedown: (e) => e.preventDefault(), onclick: () => { setImgAlign('align-right'); } }),
    el('span', { class: 'tb-sep' }),
    ...[25, 50, 75, 100].map((p) => btn(`${p}%`, `Largura ${p}%`, () => { img.style.width = `${p}%`; img.removeAttribute('width'); })),
    el('span', { class: 'tb-sep' }),
    btn('Alt', 'Texto alternativo', async () => {
      const v = await prompt('Texto alternativo', 'Descreva a imagem (acessibilidade)', img.alt);
      if (v !== null) { img.alt = v; editor.changed(true); }
    }),
    el('button', { class: 'tb-btn', title: 'Excluir imagem', html: icon('trash', 16), onmousedown: (e) => e.preventDefault(), onclick: () => { const p = img.parentElement; img.remove(); if (p && p.tagName === 'P' && !p.textContent.trim() && !p.querySelector('img')) p.innerHTML = '<br>'; deselectImage(); editor.changed(true); } }));
  imgHandle = el('div', { class: 'img-handle', title: 'Arraste para redimensionar' });
  imgHandle.addEventListener('pointerdown', (e) => {
    e.preventDefault();
    imgHandle.setPointerCapture(e.pointerId);
    const startX = e.clientX, startW = img.getBoundingClientRect().width / S.zoom;
    const maxW = editorEl.getBoundingClientRect().width / S.zoom;
    const move = (ev) => {
      img.style.width = `${Math.round(clamp(startW + (ev.clientX - startX) / S.zoom, 24, maxW))}px`;
      positionImageTools();
    };
    const up = () => { imgHandle.removeEventListener('pointermove', move); editor.changed(true); };
    imgHandle.addEventListener('pointermove', move);
    imgHandle.addEventListener('pointerup', up, { once: true });
  });
  document.body.append(imgTools, imgHandle);
  imgTools.style.position = imgHandle.style.position = 'fixed';
  positionImageTools();
}

function setImgAlign(cls) {
  if (!selectedImg) return;
  selectedImg.classList.remove('align-left', 'align-center', 'align-right');
  selectedImg.classList.add(cls);
  positionImageTools();
  editor.changed(true);
}

function positionImageTools() {
  if (!selectedImg || !imgTools) return;
  const r = selectedImg.getBoundingClientRect();
  imgTools.style.left = `${clamp(r.left, 8, innerWidth - imgTools.offsetWidth - 8)}px`;
  imgTools.style.top = `${Math.max(8, r.top - 42)}px`;
  imgHandle.style.left = `${r.right - 7}px`;
  imgHandle.style.top = `${r.bottom - 7}px`;
}

function deselectImage() {
  selectedImg?.classList.remove('img-selected');
  selectedImg = null;
  imgTools?.remove();
  imgHandle?.remove();
  imgTools = imgHandle = null;
}

let tableTools = null, toolsTable = null;
function updateTableTools() {
  const cell = editor.hasFocus() ? editor.closest('TD') || editor.closest('TH') : null;
  const table = cell?.closest('table');
  if (!table) { tableTools?.remove(); tableTools = null; toolsTable = null; return; }
  if (!tableTools || toolsTable !== table) {
    tableTools?.remove();
    toolsTable = table;
    const b = (label, op, title) => el('button', { class: 'tb-btn', title, onmousedown: (e) => e.preventDefault(), onclick: () => { editor.tableOp(op); updateTableTools(); } }, label);
    tableTools = el('div', { class: 'table-tools' },
      b('+ Linha ↑', 'rowAbove', 'Inserir linha acima'), b('+ Linha ↓', 'rowBelow', 'Inserir linha abaixo'),
      b('+ Coluna ←', 'colLeft', 'Inserir coluna à esquerda'), b('+ Coluna →', 'colRight', 'Inserir coluna à direita'),
      el('span', { class: 'tb-sep' }),
      b('− Linha', 'delRow', 'Excluir linha'), b('− Coluna', 'delCol', 'Excluir coluna'),
      el('span', { class: 'tb-sep' }),
      b('Cabeçalho', 'header', 'Alternar linha de cabeçalho'), b('Listrada', 'striped', 'Linhas alternadas'),
      el('button', { class: 'tb-btn', title: 'Excluir tabela', html: icon('trash', 15), onmousedown: (e) => e.preventDefault(), onclick: () => { editor.tableOp('delTable'); updateTableTools(); } }));
    tableTools.style.position = 'fixed';
    document.body.append(tableTools);
  }
  const r = table.getBoundingClientRect();
  tableTools.style.left = `${clamp(r.left, 8, innerWidth - tableTools.offsetWidth - 8)}px`;
  tableTools.style.top = `${Math.max(8, r.top - 40)}px`;
}

scroller.addEventListener('scroll', () => {
  positionImageTools();
  if (tableTools) updateTableTools();
  if (bubble) { const r = editor.range(); if (r && !r.collapsed) showBubble(r); }
  if (slash.el) checkSlash();
}, { passive: true });

// ---------------------------------------------------------------------------
// Painel lateral: localizar/substituir e estilo do documento

const find = { q: '', rep: '', caseSens: false, word: false, regex: false, matches: [], index: -1 };

function openSide(kind, focusReplace = false) {
  S.side = S.side === kind && !focusReplace ? null : kind;
  if (kind === 'find' && S.side) {
    const sel = getSelection().toString();
    if (sel && sel.length < 100 && !sel.includes('\n')) find.q = sel;
  }
  renderSide();
  if (S.side === 'find') setTimeout(() => $(focusReplace ? '#find-rep' : '#find-q')?.focus(), 0);
  if (!S.side) clearHighlights();
}

function renderSide() {
  const side = $('#side');
  side.classList.toggle('hidden', !S.side);
  if (!S.side) return;
  const head = (title) => el('div', { class: 'panel-head' }, title, el('button', { class: 'icon-btn', title: 'Fechar', html: icon('x', 14), onclick: () => openSide(S.side) }));
  if (S.side === 'find') {
    const q = el('input', { id: 'find-q', type: 'search', placeholder: 'Localizar', value: find.q });
    const rep = el('input', { id: 'find-rep', type: 'text', placeholder: 'Substituir por', value: find.rep });
    const count = el('span', { class: 'find-count' });
    const opt = (key, label) => {
      const c = el('input', { type: 'checkbox', checked: find[key] });
      c.addEventListener('change', () => { find[key] = c.checked; runFind(); });
      return el('label', { class: 'check' }, c, label);
    };
    const runFind = () => {
      find.q = q.value;
      computeMatches();
      count.textContent = find.q ? (find.matches.length ? `${find.index + 1} de ${find.matches.length}` : 'Nenhum resultado') : '';
    };
    q.addEventListener('input', runFind);
    q.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); step(e.shiftKey ? -1 : 1); } });
    rep.addEventListener('input', () => { find.rep = rep.value; });
    const step = (d) => {
      if (!find.matches.length) return;
      find.index = (find.index + d + find.matches.length) % find.matches.length;
      showCurrent();
      count.textContent = `${find.index + 1} de ${find.matches.length}`;
    };
    side.replaceChildren(head('Localizar e substituir'), el('div', { class: 'panel-body' },
      q,
      el('div', { class: 'row' }, count, el('span', { class: 'spacer' }),
        el('button', { class: 'icon-btn', title: 'Anterior (Shift+Enter)', html: icon('chevronLeft', 16), onclick: () => step(-1) }),
        el('button', { class: 'icon-btn', title: 'Próximo (Enter)', html: icon('chevronRight', 16), onclick: () => step(1) })),
      rep,
      el('div', { class: 'row' },
        el('button', { class: 'btn small', onclick: () => { replaceCurrent(); runFind(); } }, 'Substituir'),
        el('button', { class: 'btn small', onclick: () => { const n = replaceAll(); runFind(); toast(`${n} substituição(ões) feita(s).`); } }, 'Substituir tudo')),
      opt('caseSens', 'Diferenciar maiúsculas'), opt('word', 'Palavra inteira'), opt('regex', 'Expressão regular')));
    runFind();
    return;
  }
  if (S.side === 'style') {
    const st = S.doc.style;
    const set = (k, v) => { st[k] = v; st.preset = 'custom'; io.applyStyleVars(editorEl, st); markDirty(); renderSide(); };
    const presets = el('div', { class: 'preset-list' }, ...Object.entries(io.STYLE_PRESETS).map(([id, p]) => el('button', {
      class: `preset${st.preset === id ? ' on' : ''}`,
      onclick: () => {
        S.doc.style = io.defaultStyle(id);
        if (id === 'abnt') S.doc.page.margins = { top: 3, right: 2, bottom: 2, left: 3 };
        applyLook(); markDirty(); renderSide();
      },
    }, el('b', { style: { fontFamily: p.font } }, p.label), el('small', {}, `${p.font.split(',')[0].replace(/"/g, '')} ${p.size} pt · entrelinha ${String(p.line).replace('.', ',')}`))));
    const fontSel = (k) => {
      const s = select([['', '(igual ao texto)'], ...FONTS], st[k], {});
      if (k === 'font') s.firstChild.remove();
      s.addEventListener('change', () => set(k, s.value));
      return s;
    };
    const num = (k, step, min, max) => {
      const i = el('input', { type: 'number', step, min, max, value: st[k] });
      i.addEventListener('change', () => set(k, clamp(parseFloat(i.value) || 0, min, max)));
      return i;
    };
    const color = (k) => { const c = el('input', { type: 'color', value: st[k] }); c.addEventListener('change', () => set(k, c.value)); return c; };
    const align = select([['left', 'Esquerda'], ['justify', 'Justificado'], ['center', 'Centralizado'], ['right', 'Direita']], st.align);
    align.addEventListener('change', () => set('align', align.value));
    side.replaceChildren(head('Estilo do documento'), el('div', { class: 'panel-body' },
      el('small', { class: 'muted' }, 'Muda a aparência do documento inteiro de uma vez.'),
      presets,
      el('div', { class: 'panel-head', style: { padding: '10px 0 0' } }, 'Personalizar'),
      field('Fonte do texto', fontSel('font')), field('Fonte dos títulos', fontSel('headFont')),
      el('div', { class: 'grid2' }, field('Tamanho (pt)', num('size', 0.5, 6, 30)), field('Entrelinha', num('line', 0.05, 0.8, 3))),
      el('div', { class: 'grid2' }, field('Espaço após (pt)', num('space', 1, 0, 40)), field('Recuo 1ª linha (cm)', num('indent', 0.25, 0, 5))),
      field('Alinhamento', align),
      el('div', { class: 'grid2' }, field('Cor dos títulos', color('headColor')), field('Cor de destaque', color('accent')))));
  }
}

function markDirty() {
  if (!S.doc.dirty) { S.doc.dirty = true; renderTitle(); renderTabs(); }
  scheduleRecovery();
}

function buildRegex() {
  if (!find.q) return null;
  let src = find.regex ? find.q : find.q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  if (find.word) src = `(?<![\\p{L}\\p{N}_])${src}(?![\\p{L}\\p{N}_])`;
  try { return new RegExp(src, `gu${find.caseSens ? '' : 'i'}`); } catch { return null; }
}

function computeMatches() {
  find.matches = [];
  const re = buildRegex();
  if (re) {
    const walker = document.createTreeWalker(editorEl, NodeFilter.SHOW_TEXT, { acceptNode: (n) => (n.parentElement.closest('.toc') ? NodeFilter.FILTER_REJECT : NodeFilter.FILTER_ACCEPT) });
    for (let n = walker.nextNode(); n; n = walker.nextNode()) {
      re.lastIndex = 0;
      let m;
      while ((m = re.exec(n.textContent))) {
        if (!m[0].length) { re.lastIndex++; continue; }
        const r = document.createRange();
        r.setStart(n, m.index);
        r.setEnd(n, m.index + m[0].length);
        find.matches.push(r);
      }
    }
  }
  find.index = find.matches.length ? 0 : -1;
  paintHighlights();
}

function paintHighlights() {
  if (!window.CSS?.highlights) return;
  CSS.highlights.set('lumina-find', new Highlight(...find.matches));
  const cur = find.matches[find.index];
  if (cur) CSS.highlights.set('lumina-find-current', new Highlight(cur));
  else CSS.highlights.delete('lumina-find-current');
}

function clearHighlights() {
  if (!window.CSS?.highlights) return;
  CSS.highlights.delete('lumina-find');
  CSS.highlights.delete('lumina-find-current');
}

function showCurrent() {
  const r = find.matches[find.index];
  if (!r) return;
  paintHighlights();
  const rect = r.getBoundingClientRect();
  const sr = scroller.getBoundingClientRect();
  if (rect.top < sr.top + 40 || rect.bottom > sr.bottom - 40) scroller.scrollTop += rect.top - sr.top - sr.height / 3;
}

function replaceText(r, text) {
  const s = getSelection();
  s.removeAllRanges();
  s.addRange(r);
  document.execCommand('insertText', false, text);
}

function replaceCurrent() {
  const r = find.matches[find.index];
  if (!r) return;
  editor.focus();
  replaceText(r, find.rep);
  editor.changed(true);
}

function replaceAll() {
  const list = [...find.matches].reverse();
  for (const r of list) {
    r.deleteContents();
    if (find.rep) r.insertNode(document.createTextNode(find.rep));
  }
  editorEl.normalize();
  if (list.length) editor.changed(true);
  return list.length;
}

// ---------------------------------------------------------------------------
// Configurar página, versões, foco e ajuda

async function pageSetup() {
  const d = S.doc;
  const p = { ...d.page, margins: { ...d.page.margins } };
  const size = select(Object.entries(io.PAGE_SIZES).map(([k, v]) => [k, v[0]]), p.size);
  const orient = select([['portrait', 'Retrato'], ['landscape', 'Paisagem']], p.orientation);
  const m = {};
  const marginInput = (k) => { m[k] = el('input', { type: 'number', step: 0.1, min: 0, max: 10, value: p.margins[k] }); return m[k]; };
  const nums = el('input', { type: 'checkbox', checked: p.pageNumbers });
  const presets = el('div', { class: 'row' },
    el('span', { class: 'muted' }, 'Margens:'),
    ...[['Normal', [2.5, 2.5, 2.5, 2.5]], ['ABNT', [3, 2, 2, 3]], ['Estreita', [1.27, 1.27, 1.27, 1.27]], ['Larga', [2.54, 5.08, 2.54, 5.08]]].map(([l, v]) =>
      el('button', { class: 'btn small', onclick: () => { ['top', 'right', 'bottom', 'left'].forEach((k, i) => { m[k].value = v[i]; }); } }, l)));
  const ok = await dialog({
    title: 'Configurar página', width: 460,
    body: [el('div', { class: 'grid2' }, field('Tamanho do papel', size), field('Orientação', orient)), presets,
      el('div', { class: 'grid2' }, field('Superior (cm)', marginInput('top')), field('Inferior (cm)', marginInput('bottom')),
        field('Esquerda (cm)', marginInput('left')), field('Direita (cm)', marginInput('right'))),
      el('label', { class: 'check' }, nums, 'Numerar páginas no PDF e no Word')],
  });
  if (!ok) return;
  d.page = {
    size: size.value, orientation: orient.value, pageNumbers: nums.checked,
    margins: Object.fromEntries(['top', 'right', 'bottom', 'left'].map((k) => [k, clamp(parseFloat(m[k].value) || 0, 0, 10)])),
  };
  applyLook();
  markDirty();
  renderStatus();
}

async function versionsDialog() {
  const d = S.doc;
  storeCurrent();
  const list = el('div', { class: 'versions' });
  const draw = () => {
    if (!d.versions.length) { list.replaceChildren(el('p', { class: 'muted' }, 'Nenhuma versão ainda. Versões são criadas a cada salvamento ou manualmente.')); return; }
    list.replaceChildren(...[...d.versions].reverse().map((v) => {
      const words = (io.plainText(v.html).match(/[\p{L}\p{N}]+/gu) || []).length;
      return el('div', { class: 'version' },
        el('div', { class: 'vt' }, el('b', {}, v.name || (v.auto ? 'Salvamento' : 'Versão')), el('small', {}, `${fmtDate(new Date(v.date))} · ${words} palavras`)),
        el('button', { class: 'btn small', onclick: () => previewVersion(v) }, 'Ver'),
        el('button', { class: 'btn small', onclick: () => { restoreVersion(v); close(false); } }, 'Restaurar'));
    }));
  };
  let close;
  draw();
  dialog({
    title: 'Histórico de versões', width: 520, okText: null, cancelText: 'Fechar',
    body: [el('p', { class: 'muted' }, 'As versões ficam guardadas dentro do arquivo .lumdoc.'), list],
    extraButtons: [{ label: 'Salvar versão atual…', onClick: async () => {
      const name = await prompt('Nova versão', 'Nome da versão', `Versão ${d.versions.length + 1}`);
      if (name === null) return;
      d.versions.push({ date: Date.now(), name, title: d.title, html: currentHtml() });
      markDirty();
      draw();
    } }],
    onOpen: (_b, c) => { close = c; },
  });
}

function previewVersion(v) {
  const box = el('div', { class: 'doc-content', style: { maxHeight: '60vh', overflow: 'auto', background: '#fff', color: '#1f2328', padding: '24px', borderRadius: '8px', border: '1px solid var(--line)' }, html: io.sanitize(v.html) });
  io.applyStyleVars(box, S.doc.style);
  box.style.maxHeight = '60vh';
  box.style.overflow = 'auto';
  box.style.padding = '24px';
  box.style.background = '#fff';
  dialog({ title: `${v.name || 'Versão'} — ${fmtDate(new Date(v.date))}`, body: box, wide: true, okText: 'Restaurar esta versão', cancelText: 'Fechar' })
    .then((ok) => ok && restoreVersion(v));
}

function restoreVersion(v) {
  S.doc.versions.push({ date: Date.now(), name: 'Antes de restaurar', title: S.doc.title, html: currentHtml() });
  editorEl.innerHTML = v.html;
  editor.ensureBlock();
  editor.changed(true);
  updateToc();
  renderAll();
  toast('Versão restaurada. Você pode desfazer com Ctrl+Z.');
}

function toggleFocus() {
  const on = document.body.classList.toggle('focus-mode');
  S.activeBlock = null;
  editorEl.querySelectorAll('.active-block').forEach((n) => n.classList.remove('active-block'));
  if (on) { markActiveBlock(); hideBubble(); toast('Modo foco: Esc para sair.'); }
  editor.focus();
}

function helpDialog() {
  const rows = [
    ['/', 'Menu de blocos (títulos, listas, tabela, imagem, sumário…)'],
    ['# espaço', 'Título 1 (## Título 2, ### Título 3)'], ['- espaço', 'Lista com marcadores'], ['1. espaço', 'Lista numerada'],
    ['[] espaço', 'Lista de tarefas'], ['> espaço', 'Citação'], ['--- Enter', 'Linha horizontal'], ['``` Enter', 'Bloco de código'],
    ['**texto**', 'Negrito (também *itálico*, `código`, ~~tachado~~)'],
    ['Ctrl+B / I / U', 'Negrito, itálico, sublinhado'], ['Ctrl+Alt+1…3', 'Títulos'], ['Ctrl+Alt+0', 'Texto normal'],
    ['Ctrl+E / L / R / J', 'Centralizar, esquerda, direita, justificar'], ['Ctrl+Shift+7 / 8 / 9', 'Lista numerada / marcadores / tarefas'],
    ['Ctrl+K', 'Inserir link'], ['Ctrl+Enter', 'Quebra de página'], ['Ctrl+F / Ctrl+H', 'Localizar / substituir'],
    ['Ctrl+Shift+V', 'Colar sem formatação'], ['Ctrl+P', 'Exportar PDF'], ['F11', 'Modo foco'], ['Tab (tabela)', 'Próxima célula / nova linha'],
    ['Ctrl+clique', 'Abrir link'], ['Ctrl+roda', 'Zoom'],
  ];
  dialog({
    title: 'Atalhos e recursos', cancelText: null, width: 560,
    body: el('div', { class: 'stats-grid', style: { gridTemplateColumns: 'auto 1fr', fontSize: '13px' } }, ...rows.flatMap(([k, v]) => [el('kbd', {}, k), el('span', {}, v)])),
  });
}

// ---------------------------------------------------------------------------
// Recuperação automática

const RECOVERY_KEY = 'lumina.docs.recovery.v1';
const scheduleRecovery = debounce(() => {
  try {
    storeCurrent();
    const dirty = S.docs.filter((d) => d.dirty).map((d) => ({ id: d.id, title: d.title, html: d === S.doc ? editor.html() : d.html, page: d.page, style: d.style, date: Date.now() }));
    if (dirty.length) localStorage.setItem(RECOVERY_KEY, JSON.stringify(dirty));
    else localStorage.removeItem(RECOVERY_KEY);
  } catch { /* cota excedida ou indisponível */ }
}, 2000);

function forgetRecovery() {
  scheduleRecovery();
}

async function offerRecovery() {
  let list = [];
  try { list = JSON.parse(localStorage.getItem(RECOVERY_KEY) || '[]'); } catch { list = []; }
  if (!list.length) return;
  localStorage.removeItem(RECOVERY_KEY);
  const ok = await dialog({
    title: 'Recuperar documentos?',
    body: [el('p', {}, 'Encontramos alterações que não foram salvas da última vez:'), el('ul', {}, ...list.map((d) => el('li', {}, `${d.title} — ${fmtDate(new Date(d.date))}`)))],
    okText: 'Recuperar', cancelText: 'Descartar',
  });
  if (!ok) return;
  for (const d of list) {
    const doc = makeDoc({ title: d.title, html: d.html, page: d.page, style: d.style });
    doc.dirty = true;
    addDoc(doc);
  }
}

// ---------------------------------------------------------------------------
// Atalhos e inicialização

function setupShortcuts() {
  shortcuts({
    'Ctrl+N': () => newDocument(),
    'Ctrl+O': openDialog,
    'Ctrl+S': () => save(),
    'Ctrl+Shift+S': () => save(S.doc, true),
    'Ctrl+P': () => exportAs('pdf'),
    'Ctrl+W': () => closeDoc(),
    'Ctrl+F': () => openSide('find'),
    'Ctrl+H': () => openSide('find', true),
    'Ctrl+K': insertLink,
    'Ctrl+E': () => editor.setAlign('center'),
    'Ctrl+L': () => editor.setAlign('left'),
    'Ctrl+R': () => editor.setAlign('right'),
    'Ctrl+J': () => editor.setAlign('justify'),
    'Ctrl+Alt+1': () => editor.setBlock('h1'),
    'Ctrl+Alt+2': () => editor.setBlock('h2'),
    'Ctrl+Alt+3': () => editor.setBlock('h3'),
    'Ctrl+Alt+0': () => editor.setBlock('p'),
    'Ctrl+Shift+7': () => editor.list('number'),
    'Ctrl+Shift+8': () => editor.list('bullet'),
    'Ctrl+Shift+9': () => editor.list('check'),
    'Ctrl+Enter': insertPageBreak,
    'Ctrl+\\': () => editor.clearFormat(),
    'Ctrl+.': () => editor.exec('superscript'),
    'Ctrl+,': () => editor.exec('subscript'),
    'Alt+Shift+5': () => editor.exec('strikeThrough'),
    'Ctrl+Shift+V': pastePlain,
    'Ctrl+Shift+C': wordCount,
    F11: toggleFocus,
    F1: helpDialog,
    Escape: () => {
      if (document.body.classList.contains('focus-mode')) { toggleFocus(); return true; }
      if (selectedImg) { deselectImage(); return true; }
      return false;
    },
  }, { allowInInputs: ['Ctrl+S', 'Ctrl+F', 'Ctrl+H'] });
  scroller.addEventListener('wheel', (e) => {
    if (!e.ctrlKey) return;
    e.preventDefault();
    setZoom(S.zoom + (e.deltaY < 0 ? 0.1 : -0.1));
  }, { passive: false });
  scroller.addEventListener('mousedown', (e) => {
    // Clique abaixo do texto coloca o cursor no final.
    if (e.target === scroller || e.target === paper) {
      e.preventDefault();
      placeCaret(editorEl.lastElementChild || editorEl, false);
    }
  });
}

registerApp({
  openFiles,
  newDocument: (tpl) => newDocument(tpl),
  unsaved: () => S.docs.filter((d) => d.dirty).map((d) => d.title),
  beforeClose: async () => {
    for (const d of [...S.docs]) if (d.dirty && !(await askSave(d))) return false;
    return true;
  },
  activated: () => editorEl.focus({ preventScroll: true }),
});

initTheme();
setupHeader();
setupMenus();
setupToolbar();
setupShortcuts();
io.ensurePurify().catch(() => {});
S.docs.push(makeDoc());
activate(S.docs[0], false);
standaloneBoot({ openFiles });
setTimeout(offerRecovery, 400);
window.__docs = { S, editor, io, save, exportAs, newDocument, openFiles };
