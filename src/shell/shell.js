// Casca da suíte: barra lateral, página inicial e roteamento de arquivos entre os apps.
import { el, $, toast, dialog, debounce } from '../shared/ui.js';
import { icon } from '../shared/icons.js';
import {
  native, isNative, pickFiles, readPath, extOf, getThemePref, setThemePref, initTheme, fmtDate,
} from '../shared/native.js';

export const APPS = {
  photo: {
    name: 'Foto', full: 'Lumina Foto', icon: 'photo', url: 'photo/index.html',
    color: 'linear-gradient(135deg,#4aa8ff,#a66bff)', solid: '#7a7cff',
    desc: 'Edição de imagens com camadas, filtros e retoque.',
    exts: ['lumina', 'png', 'jpg', 'jpeg', 'webp', 'gif', 'bmp', 'svg', 'ico', 'avif'],
  },
  docs: {
    name: 'Texto', full: 'Lumina Texto', icon: 'doc', url: 'docs/index.html',
    color: 'linear-gradient(135deg,#2f6bff,#4f9bff)', solid: '#3b6cf6',
    desc: 'Documentos com estilos, tabelas, ABNT e exportação para Word e PDF.',
    exts: ['lumdoc', 'docx', 'md', 'markdown', 'txt', 'html', 'htm'],
  },
  sheets: {
    name: 'Planilhas', full: 'Lumina Planilhas', icon: 'sheet', url: 'sheets/index.html',
    color: 'linear-gradient(135deg,#1f9d6b,#3fcf8e)', solid: '#2f9e6e',
    desc: 'Fórmulas em português, gráficos, tabelas dinâmicas e Excel.',
    exts: ['lumsheet', 'xlsx', 'csv', 'tsv'],
  },
  slides: {
    name: 'Slides', full: 'Lumina Apresentações', icon: 'slides', url: 'slides/index.html',
    color: 'linear-gradient(135deg,#ff7a45,#ff4d8d)', solid: '#f0603f',
    desc: 'Apresentações modernas com temas, animações e PowerPoint.',
    exts: ['lumslides', 'pptx'],
  },
};

const TEMPLATES = {
  docs: [
    ['blank', 'Documento em branco', ''],
    ['abnt', 'Trabalho acadêmico', 'Normas ABNT'],
    ['report', 'Relatório', 'Com sumário'],
    ['letter', 'Carta formal', ''],
    ['resume', 'Currículo', 'Moderno'],
    ['meeting', 'Ata de reunião', ''],
  ],
  sheets: [
    ['blank', 'Planilha em branco', ''],
    ['budget', 'Orçamento pessoal', 'Receitas e despesas'],
    ['tasks', 'Lista de tarefas', 'Com status'],
    ['sales', 'Controle de vendas', 'Com gráfico'],
  ],
  slides: [
    ['blank', 'Apresentação em branco', 'Tema Claro'],
    ['aurora', 'Aurora', 'Gradiente vibrante'],
    ['midnight', 'Meia-noite', 'Escuro elegante'],
    ['coral', 'Coral', 'Quente e moderno'],
    ['forest', 'Floresta', 'Natural'],
    ['pitch', 'Pitch de startup', '8 slides prontos'],
  ],
  photo: [
    ['new', 'Nova imagem', 'Escolher tamanho'],
    ['open', 'Abrir imagem…', 'PNG, JPG, WebP…'],
  ],
};

const SLIDE_THEMES = {
  blank: ['#ffffff', '#1d2129', '#3b6cf6'],
  aurora: ['linear-gradient(135deg,#6a5cff,#ff5ea8)', '#ffffff', '#ffe066'],
  midnight: ['linear-gradient(160deg,#0f172a,#1e293b)', '#e2e8f0', '#38bdf8'],
  coral: ['linear-gradient(135deg,#ff9a76,#ff6a88)', '#ffffff', '#fff1e6'],
  forest: ['linear-gradient(160deg,#0b3d2e,#1f6f4a)', '#ecfdf5', '#a7f3d0'],
  pitch: ['#0b0b0f', '#ffffff', '#7c5cff'],
};

const RECENTS_KEY = 'lumina.recents.v1';
const state = { current: 'home', frames: {}, homeTab: 'docs', search: '' };

// ---------------------------------------------------------------------------
// Barra lateral

function renderRail() {
  const rail = $('#rail');
  const btn = (id, label, ic, solid) => el('button', {
    class: `rail-btn${state.current === id ? ' on' : ''}${state.frames[id] ? ' loaded' : ''}`,
    title: label,
    onclick: () => show(id),
    html: `${id === 'home' ? icon(ic, 22) : `<span style="color:${solid}">${icon(ic, 22)}</span>`}<span>${label}</span><i class="dot"></i>`,
  });
  const pref = getThemePref();
  rail.replaceChildren(
    el('img', { class: 'rail-logo', src: 'shared/logo.svg', alt: 'Lumina' }),
    btn('home', 'Início', 'home'),
    ...Object.entries(APPS).map(([id, a]) => btn(id, a.name, a.icon, a.solid)),
    el('span', { class: 'spacer' }),
    el('button', {
      class: 'rail-btn', title: 'Tema: sistema → claro → escuro',
      html: `${icon(pref === 'dark' ? 'moon' : 'sun', 20)}<span>${{ system: 'Auto', light: 'Claro', dark: 'Escuro' }[pref]}</span>`,
      onclick: () => {
        const next = { system: 'light', light: 'dark', dark: 'system' }[pref];
        setThemePref(next);
        renderRail();
      },
    }),
    el('button', { class: 'rail-btn', title: 'Sobre', html: `${icon('info', 20)}<span>Sobre</span>`, onclick: about }),
  );
}

// ---------------------------------------------------------------------------
// Apps em iframes (carregados sob demanda e mantidos vivos)

function loadApp(id) {
  if (state.frames[id]) return state.frames[id].ready;
  const frame = el('iframe', { class: 'app-frame view hidden', src: APPS[id].url, title: APPS[id].full });
  const entry = { frame, ready: null };
  entry.ready = new Promise((resolve) => {
    frame.addEventListener('load', () => {
      const wait = () => (frame.contentWindow.luminaApp ? resolve(frame.contentWindow.luminaApp) : setTimeout(wait, 30));
      wait();
    }, { once: true });
  });
  state.frames[id] = entry;
  $('#stage').append(frame);
  return entry.ready;
}

async function show(id) {
  state.current = id;
  $('#home').classList.toggle('hidden', id !== 'home');
  for (const [k, f] of Object.entries(state.frames)) f.frame.classList.toggle('hidden', k !== id);
  if (id === 'home') {
    renderHome();
    document.title = 'Lumina Studio';
  } else {
    const api = await loadApp(id);
    state.frames[id].frame.classList.remove('hidden');
    state.frames[id].frame.contentWindow.focus();
    document.title = state.frames[id].frame.contentDocument?.title || APPS[id].full;
    api.activated?.();
  }
  renderRail();
}

function appForFile(name) {
  const e = extOf(name);
  return Object.keys(APPS).find((id) => APPS[id].exts.includes(e)) || null;
}

async function openFiles(files) {
  const groups = {};
  for (const f of files) {
    const id = appForFile(f.name);
    if (!id) { toast(`Formato não suportado: ${f.name}`); continue; }
    (groups[id] ||= []).push(f);
  }
  let last = null;
  for (const [id, list] of Object.entries(groups)) {
    const api = await loadApp(id);
    await api.openFiles(list);
    last = id;
  }
  if (last) show(last);
}

async function newFromTemplate(appId, tpl) {
  const api = await loadApp(appId);
  await show(appId);
  await api.newDocument(tpl);
}

async function openDialog() {
  const all = Object.values(APPS).flatMap((a) => a.exts);
  const files = await pickFiles([
    { name: 'Todos os suportados', extensions: all },
    ...Object.values(APPS).map((a) => ({ name: a.full, extensions: a.exts })),
  ]);
  if (files.length) openFiles(files);
}

// ---------------------------------------------------------------------------
// Recentes

function getRecents() {
  try { return JSON.parse(localStorage.getItem(RECENTS_KEY) || '[]'); } catch { return []; }
}
function saveRecents(list) {
  try { localStorage.setItem(RECENTS_KEY, JSON.stringify(list.slice(0, 40))); } catch { /* indisponível */ }
}
function addRecent({ app, name, path }) {
  if (!name) return;
  const list = getRecents().filter((r) => !(r.path && r.path === path) && !(!r.path && r.name === name && r.app === app));
  list.unshift({ app, name, path: path || null, date: Date.now() });
  saveRecents(list);
  if (state.current === 'home') renderHome();
}

async function openRecent(r) {
  if (!r.path) {
    toast('Este arquivo foi salvo pelo navegador. Use "Abrir arquivo…" para escolhê-lo.');
    return;
  }
  try {
    openFiles([await readPath(r.path)]);
  } catch (err) {
    toast(`Não foi possível abrir: ${err.message}`);
    saveRecents(getRecents().filter((x) => x.path !== r.path));
    renderHome();
  }
}

// ---------------------------------------------------------------------------
// Página inicial

function tplThumb(appId, tpl) {
  if (appId === 'docs') {
    const lines = tpl === 'blank' ? [] : tpl === 'resume' ? ['h', '', '', 'h', '', ''] : ['h', '', '', '', '', ''];
    return el('div', { class: 'mini-page' }, ...lines.map((c) => el('i', { class: c, style: { width: c ? '' : `${70 + Math.random() * 30}%` } })));
  }
  if (appId === 'sheets') {
    return el('div', { class: 'mini-sheet' }, ...Array.from({ length: 28 }, (_, i) => el('span', {
      class: tpl === 'blank' ? '' : i < 4 ? 'hd' : Math.floor(i / 4) % 2 ? 'alt' : '',
    })));
  }
  if (appId === 'slides') {
    const [bg, fg, acc] = SLIDE_THEMES[tpl];
    return el('div', { class: 'mini-slide', style: { background: bg } },
      el('b', { style: { background: fg } }), el('s', { style: { background: acc } }));
  }
  return el('div', { class: 'app-badge', style: { background: APPS.photo.color, width: '54px', height: '54px', borderRadius: '14px' }, html: icon(tpl === 'open' ? 'open' : 'plus', 26) });
}

function renderHome() {
  const home = $('#home');
  const hour = new Date().getHours();
  const greet = hour < 12 ? 'Bom dia' : hour < 18 ? 'Boa tarde' : 'Boa noite';
  const search = el('input', { class: 'home-search', type: 'search', placeholder: 'Buscar arquivos recentes', value: state.search });
  search.addEventListener('input', debounce(() => { state.search = search.value; renderRecents(); }, 120));

  const tabs = el('div', { class: 'tabs' }, ...Object.entries(APPS).map(([id, a]) => el('button', {
    class: state.homeTab === id ? 'on' : '', onclick: () => { state.homeTab = id; renderHome(); },
  }, a.name)));

  const recentsBox = el('div', { class: 'recents' });
  const renderRecents = () => {
    const q = state.search.trim().toLowerCase();
    const list = getRecents().filter((r) => !q || r.name.toLowerCase().includes(q) || (r.path || '').toLowerCase().includes(q));
    if (!list.length) {
      recentsBox.replaceChildren(el('div', { class: 'empty' }, q ? 'Nenhum arquivo encontrado.' : 'Seus arquivos recentes aparecerão aqui.'));
      return;
    }
    recentsBox.replaceChildren(...list.map((r) => {
      const a = APPS[r.app] || APPS.docs;
      return el('div', { class: 'recent', onclick: () => openRecent(r), title: r.path || r.name },
        el('span', { class: 'r-icon', style: { background: a.color }, html: icon(a.icon, 16) }),
        el('div', {}, el('div', { class: 'r-name' }, r.name), r.path ? el('div', { class: 'r-path' }, r.path) : null),
        el('span', { class: 'r-date' }, a.full),
        el('span', { class: 'r-date' }, fmtDate(new Date(r.date))),
        el('button', {
          class: 'icon-btn', title: 'Remover da lista', html: icon('x', 14),
          onclick: (e) => { e.stopPropagation(); saveRecents(getRecents().filter((x) => x.date !== r.date)); renderRecents(); },
        }));
    }));
  };
  renderRecents();

  home.replaceChildren(el('div', { class: 'home-inner' },
    el('div', { class: 'home-top' },
      el('div', {}, el('h1', {}, `${greet}!`), el('p', {}, 'O que vamos criar hoje?')),
      el('span', { class: 'spacer' }),
      el('div', { class: 'search-wrap', html: icon('search', 16) }, search),
      el('button', { class: 'btn', onclick: openDialog, html: `${icon('open', 16)} Abrir arquivo…` })),
    el('div', { class: 'app-grid' }, ...Object.entries(APPS).map(([id, a]) => el('button', { class: 'app-card', onclick: () => show(id) },
      el('span', { class: 'badge', style: { background: a.color }, html: icon(a.icon, 24) }),
      el('h3', {}, a.full),
      el('p', {}, a.desc)))),
    el('div', { class: 'section-title' }, el('h2', {}, 'Comece com um modelo'), el('span', { class: 'spacer' }), tabs),
    el('div', { class: 'templates' }, ...TEMPLATES[state.homeTab].map(([tpl, name, sub]) => el('button', {
      class: 'tpl', onclick: () => newFromTemplate(state.homeTab, tpl),
    }, el('div', { class: 'tpl-thumb' }, tplThumb(state.homeTab, tpl)), el('div', { class: 'tpl-name' }, name), sub ? el('div', { class: 'tpl-sub' }, sub) : null))),
    el('div', { class: 'section-title' }, el('h2', {}, 'Recentes')),
    recentsBox));
}

function about() {
  dialog({
    title: 'Lumina Studio',
    cancelText: null,
    body: el('div', {},
      el('p', {}, el('b', {}, 'Lumina Studio 2.0'), ' — suíte de produtividade e criação.'),
      el('p', { class: 'muted' }, 'Foto, Texto, Planilhas e Apresentações em um só aplicativo, funcionando offline.'),
      el('p', { class: 'muted' }, isNative ? 'Aplicativo desktop.' : 'Executando no navegador.')),
  });
}

// ---------------------------------------------------------------------------
// Fechamento, arquivos externos e arrastar e soltar

async function confirmCloseAll() {
  for (const [id, f] of Object.entries(state.frames)) {
    const api = f.frame.contentWindow.luminaApp;
    if (!api?.unsaved?.().length) continue;
    await show(id);
    if (!(await api.beforeClose())) return false;
  }
  return true;
}

function setupDrop() {
  let depth = 0;
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  document.addEventListener('dragenter', (e) => { if (hasFiles(e) && state.current === 'home') { depth++; document.body.classList.add('dragging'); } });
  document.addEventListener('dragleave', () => { if (--depth <= 0) { depth = 0; document.body.classList.remove('dragging'); } });
  document.addEventListener('dragover', (e) => { if (hasFiles(e)) e.preventDefault(); });
  document.addEventListener('drop', async (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    depth = 0;
    document.body.classList.remove('dragging');
    const files = [];
    for (const f of e.dataTransfer.files) files.push({ name: f.name, path: native?.pathForFile?.(f) || null, data: new Uint8Array(await f.arrayBuffer()) });
    openFiles(files);
  });
  document.body.append(el('div', { class: 'drop-hint' }, 'Solte para abrir'));
}

function init() {
  initTheme();
  window.luminaShell = {
    setTitle: (t, win) => {
      const f = state.frames[state.current];
      if (f && (!win || f.frame.contentWindow === win)) document.title = t;
    },
    addRecent,
    openApp: show,
    openFiles,
    apps: APPS,
  };
  renderRail();
  renderHome();
  setupDrop();
  if (isNative) {
    native.onCloseRequest(async () => { if (await confirmCloseAll()) native.confirmClose(); });
    native.onOpenFiles((files) => openFiles(files));
    native.getLaunchFiles().then((files) => { if (files?.length) openFiles(files); });
  } else {
    window.addEventListener('beforeunload', (e) => {
      const dirty = Object.values(state.frames).some((f) => f.frame.contentWindow.luminaApp?.unsaved?.().length);
      if (dirty) { e.preventDefault(); e.returnValue = ''; }
    });
  }
  // Abre um app direto via #foto, #texto etc. (útil para atalhos)
  const hash = location.hash.slice(1);
  if (APPS[hash]) show(hash);
}

init();
