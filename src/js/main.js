// Ponto de entrada: monta a interface e liga tudo.
import { app, BLEND_MODES } from './state.js';
import { createDocument, documentFromImage, Layer } from './document.js';
import { Viewport } from './view.js';
import { ColorPicker, dialog, toast, field, sliderField, isDialogOpen } from './ui.js';
import { TOOLS, TOOLBAR_GROUPS, initToolOptions, saveToolOptions, transformTool, editText } from './tools.js';
import { FILTERS, FILTER_MENU, histogram } from './filters.js';
import {
  maskFull, invertMask, maskedCopy, maskBounds, alphaBounds,
} from './selection.js';
import {
  pickFiles, fileToDocument, saveProject, exportImage, writeClipboardImage, readClipboardImage,
  isNative, EXPORT_FORMATS,
} from './io.js';
import { el, createCanvas, cloneCanvas, clamp, debounceFrame } from './utils.js';

const $ = (s) => document.querySelector(s);
const native = window.lumina;
const IS_MAC = /Mac/.test(navigator.platform);

const svg = (inner, size = 18) =>
  `<svg class="icon" viewBox="0 0 24 24" width="${size}" height="${size}" aria-hidden="true">${inner}</svg>`;

const ICONS = {
  eye: '<path d="M2 12s3.5-7 10-7 10 7 10 7-3.5 7-10 7S2 12 2 12Z"/><circle cx="12" cy="12" r="3"/>',
  lock: '<rect x="4" y="11" width="16" height="10" rx="2"/><path d="M8 11V7a4 4 0 0 1 8 0v4"/>',
  plus: '<path d="M12 5v14M5 12h14"/>',
  copy: '<rect x="8" y="8" width="13" height="13" rx="2"/><path d="M4 16a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h10a2 2 0 0 1 2 2"/>',
  trash: '<path d="M3 6h18M19 6l-1 14a2 2 0 0 1-2 2H8a2 2 0 0 1-2-2L5 6M9 6V4a2 2 0 0 1 2-2h2a2 2 0 0 1 2 2v2"/>',
  merge: '<path d="M12 3v12M7 10l5 5 5-5M5 21h14"/>',
  up: '<path d="m18 15-6-6-6 6"/>',
  down: '<path d="m6 9 6 6 6-6"/>',
  fx: '<path d="M14 4h-2a3 3 0 0 0-3 3v13M6 11h7M15 13l5 6M20 13l-5 6"/>',
};

// ---------------------------------------------------------------------------
// Documentos e abas

function addDocument(doc) {
  app.docs.push(doc);
  switchDoc(doc);
}

function switchDoc(doc) {
  if (app.doc && app.doc !== doc) finishPending();
  app.doc = doc;
  $('#welcome').classList.toggle('hidden', !!doc);
  if (doc && !doc.view.fitted) app.view.fit();
  app.view.requestRender(true);
  app.emit('doc-changed', doc);
  app.emit('selection');
  app.emit('zoom');
}

async function askSave(doc) {
  let choice = 'cancel';
  const ok = await dialog({
    title: 'Salvar alterações?',
    body: el('p', {}, `Deseja salvar as alterações em "${doc.name}" antes de fechar?`),
    okText: 'Salvar',
    extraButtons: [{ label: 'Não salvar', onClick: (close) => { choice = 'discard'; close(false); } }],
  });
  if (ok) return (await saveProject(doc)) ? 'saved' : 'cancel';
  return choice;
}

async function closeDoc(doc = app.doc) {
  if (!doc) return true;
  if (doc.dirty) {
    const r = await askSave(doc);
    if (r === 'cancel') return false;
  }
  const i = app.docs.indexOf(doc);
  app.docs.splice(i, 1);
  if (app.doc === doc) switchDoc(app.docs[Math.min(i, app.docs.length - 1)] || null);
  else renderTabs();
  return true;
}

function renderTabs() {
  const tabs = $('#tabs');
  tabs.replaceChildren(...app.docs.map((d) => el('div', {
    class: `tab${d === app.doc ? ' on' : ''}`,
    title: d.filePath || d.name,
    onmousedown: (e) => { if (e.button === 1) { e.preventDefault(); closeDoc(d); } },
    onclick: () => switchDoc(d),
  },
  el('span', { class: 'name' }, `${d.name}${d.dirty ? ' •' : ''}`),
  el('span', { class: 'muted' }, `${Math.round(d.view.zoom * 100)}%`),
  el('button', { class: 'close', title: 'Fechar', onclick: (e) => { e.stopPropagation(); closeDoc(d); } }, '×'))));
  const d = app.doc;
  document.title = d ? `${d.name}${d.dirty ? ' •' : ''} — Lumina Studio` : 'Lumina Studio';
}

// ---------------------------------------------------------------------------
// Ferramentas

function setTool(id) {
  const next = app.tools[id];
  if (!next || next === app.tool) return;
  const prev = app.tool;
  if (prev?.deactivate) prev.deactivate();
  if (next.activate && next.activate() === false) {
    if (prev?.activate && prev !== next) prev.activate();
    return;
  }
  if (!next.hidden) app.prevToolId = id;
  app.tool = next;
  app.view.canvas.style.cursor = next.cursor || 'default';
  renderToolbar();
  renderOptions();
  $('#st-hint').textContent = next.hint || '';
  app.view.requestRender(true);
}

function startTransform() {
  if (!app.doc) return;
  finishPending();
  if (transformTool.begin()) {
    const prev = app.prevToolId;
    app.tool?.deactivate?.();
    app.tool = transformTool;
    app.prevToolId = prev;
    app.view.canvas.style.cursor = 'default';
    renderToolbar();
    renderOptions();
    $('#st-hint').textContent = transformTool.hint;
  }
}

app.on('transform-done', () => {
  app.tool = null;
  setTool(app.prevToolId || 'move');
});

// Conclui operações pendentes (transformação) antes de outros comandos.
function finishPending() {
  if (app.tool === transformTool && transformTool.running) transformTool.commit();
}

function cycleToolByKey(key) {
  const list = TOOLS.filter((t) => t.key === key && !t.hidden);
  if (!list.length) return false;
  const i = list.indexOf(app.tool);
  setTool(list[(i + 1) % list.length].id);
  return true;
}

function renderToolbar() {
  const bar = $('#toolbar');
  const items = [];
  TOOLBAR_GROUPS.forEach((group, gi) => {
    if (gi) items.push(el('div', { class: 'tool-sep' }));
    for (const id of group) {
      const t = app.tools[id];
      items.push(el('button', {
        class: `tool-btn${app.tool === t ? ' on' : ''}`,
        title: `${t.name}${t.key ? ` (${t.key.toUpperCase()})` : ''}`,
        html: svg(t.icon),
        onclick: () => setTool(id),
      }));
    }
  });
  const fg = el('div', { class: `swatch-fg${!app.editingBg ? ' editing' : ''}`, title: 'Cor de primeiro plano', style: { background: app.fg }, onclick: () => { app.editingBg = false; app.emit('colors'); } });
  const bg = el('div', { class: `swatch-bg${app.editingBg ? ' editing' : ''}`, title: 'Cor de fundo', style: { background: app.bg }, onclick: () => { app.editingBg = true; app.emit('colors'); } });
  items.push(el('div', { class: 'colors-widget' }, fg, bg));
  items.push(el('div', { class: 'colors-mini' },
    el('button', { title: 'Inverter cores (X)', onclick: swapColors }, '⇄'),
    el('button', { title: 'Cores padrão (D)', onclick: resetColors }, '◪')));
  bar.replaceChildren(...items);
}

function swapColors() {
  const f = app.fg;
  app.fg = app.bg;
  app.bg = f;
  app.emit('colors');
}

function resetColors() {
  app.fg = '#000000';
  app.bg = '#ffffff';
  app.emit('colors');
}

app.on('colors', () => renderToolbar());

// Barra de opções da ferramenta
function sliderToValue(o, pos) {
  if (!o.log) return Math.round(pos);
  const min = Math.max(1, o.min);
  return Math.round(min * Math.pow(o.max / min, pos / 1000));
}
function valueToSlider(o, v) {
  if (!o.log) return v;
  const min = Math.max(1, o.min);
  return (Math.log(Math.max(min, v) / min) / Math.log(o.max / min)) * 1000;
}

function renderOptions() {
  const t = app.tool;
  const bar = $('#optionsbar');
  if (!t) { bar.replaceChildren(); return; }
  const set = (k, v) => {
    t.opts[k] = v;
    saveToolOptions();
    app.view.requestRender(false);
  };
  const items = [el('div', { class: 'tool-name', html: `${svg(t.icon, 16)}<span>${t.name}</span>` })];
  for (const o of t.options || []) {
    if (o.type === 'hidden') continue;
    if (o.type === 'range') {
      const num = el('input', { type: 'number', min: o.min, max: o.max, value: t.opts[o.key] });
      const range = el('input', { type: 'range', min: o.log ? 0 : o.min, max: o.log ? 1000 : o.max, step: o.step || 1, value: valueToSlider(o, t.opts[o.key]) });
      range.addEventListener('input', () => { const v = sliderToValue(o, +range.value); num.value = v; set(o.key, v); });
      num.addEventListener('change', () => { const v = clamp(+num.value || o.min, o.min, o.max); num.value = v; range.value = valueToSlider(o, v); set(o.key, v); });
      items.push(el('label', { class: 'opt' }, o.label, range, num, o.unit ? el('span', {}, o.unit) : null));
    } else if (o.type === 'checkbox') {
      const c = el('input', { type: 'checkbox', checked: !!t.opts[o.key] });
      c.addEventListener('change', () => set(o.key, c.checked));
      items.push(el('label', { class: 'opt' }, c, o.label));
    } else if (o.type === 'select') {
      const s = el('select', {}, ...o.choices.map(([v, l]) => el('option', { value: v, selected: t.opts[o.key] === v }, l)));
      s.addEventListener('change', () => set(o.key, s.value));
      items.push(el('label', { class: 'opt' }, o.label, s));
    } else if (o.type === 'buttons') {
      const seg = el('div', { class: 'seg' }, ...o.choices.map(([v, l]) => el('button', {
        class: t.opts[o.key] === v ? 'on' : '',
        onclick: (e) => { set(o.key, v); seg.querySelectorAll('button').forEach((b) => b.classList.remove('on')); e.currentTarget.classList.add('on'); },
      }, l)));
      items.push(el('div', { class: 'opt' }, o.label, seg));
    } else if (o.type === 'color') {
      const c = el('input', { type: 'color', value: t.opts[o.key] });
      c.addEventListener('input', () => set(o.key, c.value));
      items.push(el('label', { class: 'opt' }, o.label, c));
    } else if (o.type === 'action') {
      items.push(el('button', { class: 'btn small', onclick: () => o.action(t) }, o.label));
    } else if (o.type === 'info') {
      const span = el('span', { class: 'opt-info' }, o.get(t));
      span._get = () => o.get(t);
      items.push(span);
    }
  }
  bar.replaceChildren(...items);
}

app.on('tool-info', () => {
  document.querySelectorAll('#optionsbar .opt-info').forEach((s) => { s.textContent = s._get(); });
});
app.on('tool-options', () => renderOptions());

// ---------------------------------------------------------------------------
// Painel de camadas

const thumbCache = new WeakMap();
function layerThumb(layer, d) {
  const c = el('canvas', { width: 80, height: 60 });
  const ctx = c.getContext('2d');
  const s = Math.min(80 / d.width, 60 / d.height);
  const w = d.width * s, h = d.height * s;
  ctx.imageSmoothingQuality = 'high';
  ctx.drawImage(layer.canvas, (80 - w) / 2, (60 - h) / 2, w, h);
  return c;
}

function renderLayers() {
  const d = app.doc;
  const list = $('#layers');
  const blend = $('#layer-blend');
  const op = $('#layer-opacity');
  if (!d) {
    list.replaceChildren();
    return;
  }
  const al = d.activeLayer;
  blend.value = al.blend;
  op.value = Math.round(al.opacity * 100);
  $('#layer-opacity-val').textContent = `${op.value}%`;
  const items = [];
  for (let i = d.layers.length - 1; i >= 0; i--) {
    const l = d.layers[i];
    let thumb = thumbCache.get(l.canvas);
    if (!thumb) {
      thumb = layerThumb(l, d);
      thumbCache.set(l.canvas, thumb);
    }
    const name = el('span', { class: 'lname', title: l.name }, l.name);
    const li = el('li', {
      class: `layer${i === d.active ? ' on' : ''}`,
      draggable: 'true',
      onclick: () => { if (d.active !== i) { finishPending(); d.active = i; renderLayers(); } },
      ondragstart: (e) => { e.dataTransfer.setData('text/x-layer', String(i)); e.dataTransfer.effectAllowed = 'move'; },
      ondragover: (e) => { e.preventDefault(); li.classList.add('drag-over'); },
      ondragleave: () => li.classList.remove('drag-over'),
      ondrop: (e) => {
        e.preventDefault();
        li.classList.remove('drag-over');
        const from = +e.dataTransfer.getData('text/x-layer');
        if (!Number.isNaN(from)) { finishPending(); d.moveLayer(from, i); }
      },
    },
    el('button', {
      class: `vis${l.visible ? '' : ' off'}`, title: 'Mostrar/ocultar', html: svg(ICONS.eye, 15),
      onclick: (e) => { e.stopPropagation(); l.visible = !l.visible; d.commit(l.visible ? 'Mostrar camada' : 'Ocultar camada'); app.view.requestRender(true); },
    }),
    thumb,
    name,
    l.text ? el('span', { class: 'badge', title: 'Camada de texto (duplo clique para editar)' }, 'T') : null,
    el('button', {
      class: `lock${l.locked ? ' on' : ''}`, title: 'Bloquear', html: svg(ICONS.lock, 14),
      onclick: (e) => { e.stopPropagation(); l.locked = !l.locked; d.commit(l.locked ? 'Bloquear camada' : 'Desbloquear camada'); },
    }));
    thumb.title = 'Ctrl+clique: selecionar pixels da camada';
    thumb.onclick = (e) => {
      if (e.ctrlKey || e.metaKey) {
        e.stopPropagation();
        d.setSelection(alphaBounds(l.canvas) ? cloneCanvas(l.canvas) : null);
        d.commit('Selecionar pixels da camada');
        app.view.requestRender(false);
      }
    };
    name.ondblclick = (e) => {
      e.stopPropagation();
      const inp = el('input', { type: 'text', class: 'rename', value: l.name });
      name.replaceWith(inp);
      inp.focus();
      inp.select();
      const done = (save) => {
        if (save && inp.value.trim() && inp.value !== l.name) { l.name = inp.value.trim(); d.commit('Renomear camada'); }
        else renderLayers();
      };
      inp.onkeydown = (ev) => { ev.stopPropagation(); if (ev.key === 'Enter') done(true); if (ev.key === 'Escape') done(false); };
      inp.onblur = () => done(true);
    };
    if (l.text) thumb.ondblclick = (e) => { e.stopPropagation(); d.active = i; editText(l, false); };
    items.push(li);
  }
  list.replaceChildren(...items);
}

function setupLayerPanel() {
  const blend = $('#layer-blend');
  blend.replaceChildren(...BLEND_MODES.map(([v, l]) => el('option', { value: v }, l)));
  blend.addEventListener('change', () => {
    const d = app.doc;
    if (!d) return;
    d.activeLayer.blend = blend.value;
    d.commit('Modo de mesclagem');
    app.view.requestRender(true);
  });
  const op = $('#layer-opacity');
  op.addEventListener('input', () => {
    const d = app.doc;
    if (!d) return;
    d.activeLayer.opacity = op.value / 100;
    $('#layer-opacity-val').textContent = `${op.value}%`;
    app.view.requestRender(true);
  });
  op.addEventListener('change', () => app.doc?.commit('Opacidade da camada'));
  const actions = [
    ['plus', 'Nova camada (Ctrl+Shift+N)', 'newLayer'],
    ['copy', 'Duplicar camada (Ctrl+J)', 'duplicateLayer'],
    ['fx', 'Efeitos: Sombra projetada', 'fx'],
    ['merge', 'Mesclar para baixo (Ctrl+E)', 'mergeDown'],
    ['up', 'Mover para cima (Ctrl+])', 'layerUp'],
    ['down', 'Mover para baixo (Ctrl+[)', 'layerDown'],
    ['trash', 'Excluir camada', 'deleteLayer'],
  ];
  $('#layer-actions').replaceChildren(...actions.map(([ic, title, cmd]) =>
    el('button', { class: 'icon-btn', title, html: svg(ICONS[ic], 16), onclick: () => run(cmd) })));
}

function renderHistory() {
  const d = app.doc;
  const list = $('#history');
  if (!d) { list.replaceChildren(); return; }
  list.replaceChildren(...d.history.map((h, i) => el('li', {
    class: i === d.hIndex ? 'on' : i > d.hIndex ? 'future' : '',
    onclick: () => { finishPending(); d.restore(i); },
  }, h.label)));
  list.querySelector('.on')?.scrollIntoView({ block: 'nearest' });
}

const refreshPanels = debounceFrame(() => {
  renderLayers();
  renderHistory();
  renderTabs();
  updateStatus();
});

app.on('doc-changed', () => { app.view?.requestRender(true); refreshPanels(); });
app.on('layers', () => { app.view?.requestRender(true); refreshPanels(); });
app.on('doc-meta', () => refreshPanels());
app.on('zoom', () => { updateStatus(); renderTabs(); });

// ---------------------------------------------------------------------------
// Amostras

const SWATCH_KEY = 'lumina.swatches.v1';
function renderSwatches() {
  const box = $('#swatches');
  box.replaceChildren(
    ...app.swatches.map((c, i) => el('button', {
      style: { background: c }, title: `${c} — clique: frente · Alt+clique: fundo · botão direito: remover`,
      onclick: (e) => (e.altKey || e.ctrlKey ? app.setBg(c) : app.setFg(c)),
      oncontextmenu: (e) => { e.preventDefault(); app.swatches.splice(i, 1); saveSwatches(); renderSwatches(); },
    })),
    el('button', { class: 'add', title: 'Adicionar cor de frente', onclick: () => { app.swatches.push(app.fg); saveSwatches(); renderSwatches(); } }, '+')
  );
}
function saveSwatches() {
  try { localStorage.setItem(SWATCH_KEY, JSON.stringify(app.swatches)); } catch { /* indisponível */ }
}
function loadSwatches() {
  try {
    const s = JSON.parse(localStorage.getItem(SWATCH_KEY) || 'null');
    if (Array.isArray(s)) app.swatches = s;
  } catch { /* indisponível */ }
}

// ---------------------------------------------------------------------------
// Status

function updateStatus() {
  const d = app.doc;
  $('#st-zoom').textContent = d ? `${(d.view.zoom * 100).toFixed(d.view.zoom < 0.1 ? 1 : 0)}%` : '';
  $('#zoom-indicator').textContent = d ? `${(d.view.zoom * 100).toFixed(0)}%` : '';
  $('#st-size').textContent = d ? `${d.width} × ${d.height} px · ${d.layers.length} camada(s)` : '';
}
app.on('status-info', (t) => { $('#st-info').textContent = t; });

// ---------------------------------------------------------------------------
// Diálogos de comandos

const PRESETS = [
  ['Full HD', 1920, 1080], ['4K UHD', 3840, 2160], ['HD', 1280, 720], ['Instagram', 1080, 1080],
  ['Stories', 1080, 1920], ['A4 300ppi', 2480, 3508], ['Carta 300ppi', 2550, 3300], ['Ícone', 512, 512], ['800 × 600', 800, 600],
];

async function newDocumentDialog() {
  const name = el('input', { type: 'text', value: `Sem título-${app.docs.length + 1}` });
  const w = el('input', { type: 'number', min: 1, max: 16000, value: 1920 });
  const h = el('input', { type: 'number', min: 1, max: 16000, value: 1080 });
  const bgSel = el('select', {},
    el('option', { value: '#ffffff' }, 'Branco'), el('option', { value: '#000000' }, 'Preto'),
    el('option', { value: 'transparent' }, 'Transparente'), el('option', { value: 'bg' }, 'Cor de fundo'));
  const presets = el('div', { class: 'presets' }, ...PRESETS.map(([n, pw, ph]) =>
    el('button', { class: 'btn small', onclick: () => { w.value = pw; h.value = ph; } }, `${n}`)));
  const body = el('div', {}, field('Nome', name), el('div', { class: 'grid2' }, field('Largura (px)', w), field('Altura (px)', h)),
    field('Predefinições', presets), field('Fundo', bgSel));
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  if (!(await dialog({ title: 'Novo documento', body, okText: 'Criar', width: 420 }))) return;
  const W = clamp(Math.round(+w.value) || 1, 1, 16000), H = clamp(Math.round(+h.value) || 1, 1, 16000);
  const background = bgSel.value === 'bg' ? app.bg : bgSel.value;
  addDocument(createDocument(W, H, { name: name.value || 'Sem título', background }));
}

async function imageSizeDialog() {
  const d = app.doc;
  const w = el('input', { type: 'number', min: 1, max: 16000, value: d.width });
  const h = el('input', { type: 'number', min: 1, max: 16000, value: d.height });
  const pct = el('input', { type: 'number', min: 1, max: 1000, value: 100 });
  const keep = el('input', { type: 'checkbox', checked: true });
  const ratio = d.width / d.height;
  w.addEventListener('input', () => { if (keep.checked) h.value = Math.round(+w.value / ratio); pct.value = Math.round((+w.value / d.width) * 100); });
  h.addEventListener('input', () => { if (keep.checked) w.value = Math.round(+h.value * ratio); pct.value = Math.round((+h.value / d.height) * 100); });
  pct.addEventListener('input', () => { w.value = Math.round((d.width * +pct.value) / 100); h.value = Math.round((d.height * +pct.value) / 100); });
  const body = el('div', { class: 'grid3' }, field('Largura (px)', w), field('Altura (px)', h), field('Porcentagem', pct));
  const ok = await dialog({ title: 'Tamanho da imagem', body: [body, el('label', { class: 'check' }, keep, ' Manter proporções')], width: 400 });
  if (!ok) return;
  const W = clamp(Math.round(+w.value), 1, 16000), H = clamp(Math.round(+h.value), 1, 16000);
  if (W !== d.width || H !== d.height) { d.resizeImage(W, H); app.view.fit(); }
}

async function canvasSizeDialog() {
  const d = app.doc;
  const w = el('input', { type: 'number', min: 1, max: 16000, value: d.width });
  const h = el('input', { type: 'number', min: 1, max: 16000, value: d.height });
  let anchor = [0.5, 0.5];
  const grid = el('div', { class: 'anchor-grid' });
  const draw = () => grid.replaceChildren(...[0, 0.5, 1].flatMap((ay) => [0, 0.5, 1].map((ax) => el('button', {
    class: anchor[0] === ax && anchor[1] === ay ? 'on' : '', onclick: () => { anchor = [ax, ay]; draw(); },
  }))));
  draw();
  const body = el('div', {}, el('div', { class: 'grid2' }, field('Largura (px)', w), field('Altura (px)', h)), field('Âncora', grid));
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  if (!(await dialog({ title: 'Tamanho da tela', body, width: 360 }))) return;
  d.resizeCanvas(clamp(Math.round(+w.value), 1, 16000), clamp(Math.round(+h.value), 1, 16000), anchor[0], anchor[1]);
  app.view.fit();
}

async function exportDialog() {
  const d = app.doc;
  const fmt = el('select', {}, ...Object.entries(EXPORT_FORMATS).map(([k, f]) => el('option', { value: k }, f.name)));
  const q = { label: 'Qualidade', min: 1, max: 100, value: 92 };
  let quality = 92;
  const qField = sliderField(q, 92, (v) => { quality = v; });
  const scale = el('input', { type: 'number', min: 1, max: 800, value: 100 });
  const info = el('p', { class: 'muted' }, `${d.width} × ${d.height} px`);
  scale.addEventListener('input', () => { info.textContent = `${Math.round((d.width * +scale.value) / 100)} × ${Math.round((d.height * +scale.value) / 100)} px`; });
  const sync = () => { qField.style.display = fmt.value === 'png' ? 'none' : ''; };
  fmt.addEventListener('change', sync);
  sync();
  const body = el('div', {}, el('div', { class: 'grid2' }, field('Formato', fmt), field('Escala (%)', scale)), qField, info);
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  if (!(await dialog({ title: 'Exportar como', body, okText: 'Exportar…', width: 380 }))) return;
  try {
    const path = await exportImage(d, { format: fmt.value, quality: quality / 100, scale: clamp(+scale.value || 100, 1, 800) / 100, background: '#ffffff' });
    if (path) toast(`Exportado: ${String(path).split(/[\\/]/).pop()}`);
  } catch (err) {
    toast(`Erro ao exportar: ${err.message}`);
  }
}

async function fillDialog() {
  const what = el('select', {},
    el('option', { value: 'fg' }, 'Cor de primeiro plano'), el('option', { value: 'bg' }, 'Cor de fundo'),
    el('option', { value: '#000000' }, 'Preto'), el('option', { value: '#808080' }, '50% cinza'), el('option', { value: '#ffffff' }, 'Branco'));
  const mode = el('select', {}, ...BLEND_MODES.map(([v, l]) => el('option', { value: v }, l)));
  let opacity = 100;
  const body = el('div', {}, field('Conteúdo', what), field('Modo', mode), sliderField({ label: 'Opacidade (%)', min: 1, max: 100 }, 100, (v) => { opacity = v; }));
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  if (!(await dialog({ title: 'Preencher', body, width: 340 }))) return;
  const color = what.value === 'fg' ? app.fg : what.value === 'bg' ? app.bg : what.value;
  fillArea(color, 'Preencher', mode.value, opacity / 100);
}

async function strokeDialog() {
  const d = app.doc;
  if (!d.selection) { toast('Faça uma seleção primeiro.'); return; }
  let width = 4;
  const color = el('input', { type: 'color', value: app.fg });
  const body = el('div', {}, sliderField({ label: 'Largura (px)', min: 1, max: 100 }, 4, (v) => { width = v; }), field('Cor', color));
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  if (!(await dialog({ title: 'Traçar seleção', body, width: 320 }))) return;
  const layer = d.activeLayer;
  if (!editable(layer)) return;
  d.editLayer(layer, (ctx) => {
    ctx.strokeStyle = color.value;
    ctx.lineWidth = width;
    ctx.lineJoin = 'round';
    ctx.stroke(d.selection._ants);
  });
  d.commit('Traçar');
}

function editable(layer) {
  if (!layer) return false;
  if (layer.locked) { toast('A camada está bloqueada.'); return false; }
  return true;
}

function fillArea(color, label, mode = 'source-over', alpha = 1) {
  const d = app.doc;
  const layer = d.activeLayer;
  if (!editable(layer)) return;
  const t = createCanvas(d.width, d.height);
  const tc = t.getContext('2d');
  tc.fillStyle = color;
  tc.fillRect(0, 0, d.width, d.height);
  if (d.selection) { tc.globalCompositeOperation = 'destination-in'; tc.drawImage(d.selection, 0, 0); }
  d.editLayer(layer, (ctx) => { ctx.globalAlpha = alpha; ctx.globalCompositeOperation = mode; ctx.drawImage(t, 0, 0); });
  d.commit(label);
}

// Mistura o resultado filtrado de volta respeitando a seleção (alfa linear).
function mixWithSelection(orig, filtered, mask) {
  if (!mask) return filtered;
  const out = cloneCanvas(orig);
  const ctx = out.getContext('2d');
  ctx.globalCompositeOperation = 'destination-out';
  ctx.drawImage(mask, 0, 0);
  ctx.globalCompositeOperation = 'lighter';
  ctx.drawImage(maskedCopy(filtered, mask), 0, 0);
  return out;
}

const lastFilterParams = {};
let lastFilterId = null;

async function runFilter(id, repeat = false) {
  const d = app.doc;
  if (!d) return;
  finishPending();
  const f = FILTERS[id];
  const layer = d.activeLayer;
  if (!editable(layer)) return;
  const src = layer.canvas;
  const params = {};
  for (const p of f.params || []) params[p.key] = lastFilterParams[id]?.[p.key] ?? p.value;
  const compute = () => mixWithSelection(src, f.run(src, params), d.selection);

  if (!f.params || repeat) {
    document.body.style.cursor = 'progress';
    await new Promise((r) => setTimeout(r, 10));
    try {
      d.replaceLayerCanvas(layer, compute());
      d.commit(f.name);
      lastFilterId = id;
    } finally {
      document.body.style.cursor = '';
    }
    return;
  }

  let previewOn = true, timer = 0, previewValid = false;
  const schedule = () => {
    previewValid = false;
    clearTimeout(timer);
    timer = setTimeout(() => {
      if (previewOn) {
        layer.preview = compute();
        previewValid = true;
      } else layer.preview = null;
      app.view.requestRender(true);
    }, 40);
  };
  const controls = (f.params || []).map((p) => {
    if (p.type === 'checkbox') {
      const c = el('input', { type: 'checkbox', checked: !!params[p.key] });
      c.addEventListener('change', () => { params[p.key] = c.checked; schedule(); });
      return el('label', { class: 'check' }, c, p.label);
    }
    if (p.type === 'color') {
      const c = el('input', { type: 'color', value: params[p.key] });
      c.addEventListener('input', () => { params[p.key] = c.value; schedule(); });
      return field(p.label, c);
    }
    return sliderField(p, params[p.key], (v) => { params[p.key] = v; schedule(); });
  });
  const pv = el('input', { type: 'checkbox', checked: true });
  pv.addEventListener('change', () => { previewOn = pv.checked; schedule(); });
  const extras = [];
  if (f.histogram) {
    const hc = el('canvas', { class: 'histogram', width: 256, height: 90 });
    const hist = histogram(src);
    const ctx = hc.getContext('2d');
    const max = Math.max(...hist.l) || 1;
    ctx.fillStyle = '#9aa0aa';
    for (let i = 0; i < 256; i++) {
      const v = (Math.sqrt(hist.l[i]) / Math.sqrt(max)) * 88;
      ctx.fillRect(i, 90 - v, 1, v);
    }
    extras.push(hc);
  }
  const body = el('div', {}, ...extras, ...controls, el('label', { class: 'check' }, pv, ' Visualizar'));
  body.style.cssText = 'display:flex;flex-direction:column;gap:10px';
  schedule();
  const ok = await dialog({ title: f.name, body, width: 380 });
  clearTimeout(timer);
  if (ok) {
    const result = previewValid && layer.preview ? layer.preview : compute();
    layer.preview = null;
    d.replaceLayerCanvas(layer, result);
    lastFilterParams[id] = { ...params };
    lastFilterId = id;
    d.commit(f.name);
  } else {
    layer.preview = null;
  }
  app.view.requestRender(true);
}

function shortcutsDialog() {
  const rows = [
    ['V', 'Mover'], ['M', 'Letreiro (retangular/elíptico)'], ['L', 'Laço'], ['W', 'Varinha mágica'], ['C', 'Cortar'],
    ['I', 'Conta-gotas'], ['B / N', 'Pincel / Lápis'], ['E', 'Borracha'], ['S', 'Carimbo'], ['R', 'Desfoque/Nitidez'],
    ['O', 'Subexposição/Superexposição'], ['G', 'Balde / Degradê'], ['T', 'Texto'], ['U', 'Formas'], ['H / Espaço', 'Mão'], ['Z', 'Zoom'],
    ['[ ]', 'Diminuir / aumentar pincel'], ['1 … 0', 'Opacidade 10% … 100%'], ['X / D', 'Inverter / redefinir cores'],
    ['Ctrl+T', 'Transformação livre'], ['Ctrl+Z / Ctrl+Shift+Z', 'Desfazer / Refazer'], ['Ctrl+J', 'Camada via cópia'],
    ['Ctrl+A / Ctrl+D', 'Selecionar tudo / Desmarcar'], ['Ctrl+Shift+I', 'Inverter seleção'], ['Del', 'Apagar seleção'],
    ['Alt+Backspace', 'Preencher com cor de frente'], ['Ctrl+Backspace', 'Preencher com cor de fundo'],
    ['Ctrl+0 / Ctrl+1', 'Ajustar à tela / 100%'], ['Roda do mouse', 'Zoom no cursor (Shift: rolar)'], ['Tab', 'Ocultar painéis'],
  ];
  const grid = el('div', { class: 'shortcuts' }, ...rows.flatMap(([k, v]) => [el('kbd', {}, k), el('span', {}, v)]));
  dialog({ title: 'Atalhos de teclado', body: grid, cancelText: null, width: 480 });
}

function aboutDialog() {
  dialog({
    title: 'Sobre o Lumina Studio',
    body: el('div', {},
      el('p', {}, el('b', {}, 'Lumina Studio 1.0'), ' — editor de imagens para desktop.'),
      el('p', { class: 'muted' }, 'Camadas, modos de mesclagem, seleções, transformação livre, texto editável, pincéis com pressão, carimbo, ajustes e filtros com visualização ao vivo, histórico ilimitado por documento e múltiplas abas.'),
      el('p', { class: 'muted' }, isNative ? 'Executando como aplicativo desktop.' : 'Executando no navegador.')),
    cancelText: null,
  });
}

// ---------------------------------------------------------------------------
// Seleção: modificar

function morphSelection(radius, mode) {
  const d = app.doc;
  if (!d.selection) return;
  const m = createCanvas(d.width, d.height);
  const ctx = m.getContext('2d');
  const s = mode === 'feather' ? radius : mode === 'smooth' ? radius : radius / 1.86;
  ctx.filter = `blur(${s}px)`;
  ctx.drawImage(d.selection, 0, 0);
  ctx.filter = 'none';
  if (mode !== 'feather') {
    const id = ctx.getImageData(0, 0, d.width, d.height);
    const a = id.data;
    const th = mode === 'expand' ? 8 : mode === 'contract' ? 247 : 128;
    for (let i = 3; i < a.length; i += 4) {
      const on = a[i] > th;
      a[i - 3] = a[i - 2] = a[i - 1] = 255;
      a[i] = on ? 255 : 0;
    }
    ctx.putImageData(id, 0, 0);
  }
  d.setSelection(maskBounds(m) ? m : null);
  d.commit({ expand: 'Expandir seleção', contract: 'Contrair seleção', feather: 'Difusão', smooth: 'Suavizar seleção' }[mode]);
  app.view.requestRender(false);
}

async function askNumber(title, label, value, min, max) {
  const inp = el('input', { type: 'number', min, max, value });
  if (!(await dialog({ title, body: field(label, inp), width: 300 }))) return null;
  return clamp(+inp.value || 0, min, max);
}

// ---------------------------------------------------------------------------
// Área de transferência

function selectionCrop(source) {
  const d = app.doc;
  const content = maskedCopy(source, d.selection);
  const b = (d.selection ? maskBounds(d.selection) : { x: 0, y: 0, w: d.width, h: d.height });
  if (!b) return null;
  const c = createCanvas(b.w, b.h);
  c.getContext('2d').drawImage(content, -b.x, -b.y);
  return { canvas: c, x: b.x, y: b.y };
}

async function copy(merged = false) {
  const d = app.doc;
  if (!d) return;
  finishPending();
  const clip = selectionCrop(merged ? d.flattenToCanvas() : d.activeLayer.canvas);
  if (!clip) return;
  app.clipboard = clip;
  await writeClipboardImage(clip.canvas);
  toast(merged ? 'Copiado (mesclado).' : 'Copiado.');
}

async function cut() {
  const d = app.doc;
  if (!d || !editable(d.activeLayer)) return;
  await copy();
  clearSelectionPixels('Recortar');
}

function clearSelectionPixels(label = 'Apagar') {
  const d = app.doc;
  const layer = d.activeLayer;
  if (!editable(layer)) return;
  d.editLayer(layer, (ctx) => {
    if (d.selection) { ctx.globalCompositeOperation = 'destination-out'; ctx.drawImage(d.selection, 0, 0); }
    else ctx.clearRect(0, 0, d.width, d.height);
  });
  d.commit(label);
}

function pasteImage(img, internal = null) {
  if (!app.doc) {
    addDocument(documentFromImage(img, 'Colado'));
    return;
  }
  finishPending();
  const d = app.doc;
  const c = createCanvas(d.width, d.height);
  let x, y;
  if (internal && internal.canvas.width === img.width && internal.canvas.height === img.height) {
    ({ x, y } = internal);
  } else {
    const center = app.view.toDoc(app.view.width / 2, app.view.height / 2);
    x = Math.round(clamp(center.x, 0, d.width) - img.width / 2);
    y = Math.round(clamp(center.y, 0, d.height) - img.height / 2);
  }
  c.getContext('2d').drawImage(img, x, y);
  d.addLayer('Camada colada', c);
  d.commit('Colar');
}

async function pasteFromMenu() {
  const img = await readClipboardImage();
  if (img) pasteImage(img, app.clipboard);
  else if (app.clipboard) pasteImage(app.clipboard.canvas, app.clipboard);
  else toast('A área de transferência não contém imagem.');
}

async function placeImage() {
  const files = await pickFiles();
  for (const f of files) {
    try {
      const tmp = await fileToDocument(f);
      const src = tmp.flattenToCanvas();
      if (!app.doc) { addDocument(tmp); continue; }
      const d = app.doc;
      const c = createCanvas(d.width, d.height);
      const s = Math.min(1, d.width / src.width, d.height / src.height);
      const w = src.width * s, h = src.height * s;
      c.getContext('2d').drawImage(src, (d.width - w) / 2, (d.height - h) / 2, w, h);
      d.addLayer(f.name.replace(/\.[^.]+$/, ''), c);
      d.commit('Inserir imagem');
    } catch (err) {
      toast(`Não foi possível abrir ${f.name}: ${err.message}`);
    }
  }
}

async function openFiles(files) {
  for (const f of files) {
    try {
      addDocument(await fileToDocument(f));
    } catch (err) {
      toast(`Não foi possível abrir ${f.name}: ${err.message}`);
    }
  }
}

// ---------------------------------------------------------------------------
// Comandos

const needsDoc = (fn) => (...a) => { if (app.doc) return fn(...a); return undefined; };

const commands = {
  new: newDocumentDialog,
  open: async () => openFiles(await pickFiles()),
  place: needsDoc(placeImage),
  save: needsDoc(async () => { finishPending(); if (await saveProject(app.doc)) toast('Projeto salvo.'); }),
  saveAs: needsDoc(async () => { finishPending(); if (await saveProject(app.doc, true)) toast('Projeto salvo.'); }),
  export: needsDoc(() => { finishPending(); return exportDialog(); }),
  quickPng: needsDoc(async () => { finishPending(); const p = await exportImage(app.doc, { format: 'png' }); if (p) toast('PNG exportado.'); }),
  close: needsDoc(() => closeDoc()),
  quit: () => (isNative ? window.close() : null),

  undo: needsDoc(() => {
    if (app.tool === transformTool && transformTool.running) { transformTool.cancel(); return; }
    app.tool?.cancel?.();
    app.doc.undo();
  }),
  redo: needsDoc(() => { finishPending(); app.doc.redo(); }),
  cut: needsDoc(cut),
  copy: needsDoc(() => copy(false)),
  copyMerged: needsDoc(() => copy(true)),
  paste: pasteFromMenu,
  clear: needsDoc(() => { finishPending(); clearSelectionPixels(); }),
  fill: needsDoc(() => { finishPending(); fillDialog(); }),
  fillFg: needsDoc(() => { finishPending(); fillArea(app.fg, 'Preencher'); }),
  fillBg: needsDoc(() => { finishPending(); fillArea(app.bg, 'Preencher'); }),
  stroke: needsDoc(() => { finishPending(); strokeDialog(); }),
  transform: needsDoc(startTransform),

  selectAll: needsDoc(() => { finishPending(); app.doc.setSelection(maskFull(app.doc.width, app.doc.height)); app.doc.commit('Selecionar tudo'); app.view.requestRender(false); }),
  deselect: needsDoc(() => { finishPending(); if (!app.doc.selection) return; app.doc.setSelection(null); app.doc.commit('Desmarcar'); app.view.requestRender(false); }),
  invertSel: needsDoc(() => {
    finishPending();
    const d = app.doc;
    d.setSelection(d.selection ? invertMask(d.selection, d.width, d.height) : maskFull(d.width, d.height));
    d.commit('Inverter seleção');
    app.view.requestRender(false);
  }),
  selectLayer: needsDoc(() => {
    const d = app.doc;
    const l = d.activeLayer;
    d.setSelection(alphaBounds(l.canvas) ? cloneCanvas(l.canvas) : null);
    d.commit('Selecionar pixels da camada');
    app.view.requestRender(false);
  }),
  expandSel: needsDoc(async () => { const r = await askNumber('Expandir seleção', 'Expandir em (px)', 5, 1, 200); if (r) morphSelection(r, 'expand'); }),
  contractSel: needsDoc(async () => { const r = await askNumber('Contrair seleção', 'Contrair em (px)', 5, 1, 200); if (r) morphSelection(r, 'contract'); }),
  featherSel: needsDoc(async () => { const r = await askNumber('Difusão', 'Raio de difusão (px)', 8, 0.5, 250); if (r) morphSelection(r, 'feather'); }),
  smoothSel: needsDoc(async () => { const r = await askNumber('Suavizar seleção', 'Raio (px)', 4, 1, 100); if (r) morphSelection(r, 'smooth'); }),

  imageSize: needsDoc(() => { finishPending(); imageSizeDialog(); }),
  canvasSize: needsDoc(() => { finishPending(); canvasSizeDialog(); }),
  rotateCW: needsDoc(() => { finishPending(); app.doc.rotate(90); app.view.fit(); }),
  rotateCCW: needsDoc(() => { finishPending(); app.doc.rotate(270); app.view.fit(); }),
  rotate180: needsDoc(() => { finishPending(); app.doc.rotate(180); }),
  flipH: needsDoc(() => { finishPending(); app.doc.flip(true); }),
  flipV: needsDoc(() => { finishPending(); app.doc.flip(false); }),
  cropToSel: needsDoc(() => {
    finishPending();
    const d = app.doc;
    const b = d.selection && maskBounds(d.selection);
    if (!b) { toast('Faça uma seleção primeiro.'); return; }
    d.crop(b.x, b.y, b.w, b.h);
    app.view.fit();
  }),
  trim: needsDoc(() => {
    finishPending();
    const d = app.doc;
    const b = alphaBounds(d.flattenToCanvas());
    if (!b) { toast('A imagem está vazia.'); return; }
    if (b.w === d.width && b.h === d.height) { toast('Nada a aparar (sem bordas transparentes).'); return; }
    d.crop(b.x, b.y, b.w, b.h);
    app.view.fit();
  }),

  newLayer: needsDoc(() => { finishPending(); app.doc.addLayer(`Camada ${app.doc.layers.length}`); app.doc.commit('Nova camada'); }),
  duplicateLayer: needsDoc(() => {
    finishPending();
    const d = app.doc;
    if (d.selection) commands.layerViaCopy();
    else d.duplicateLayer();
  }),
  layerViaCopy: needsDoc(() => {
    finishPending();
    const d = app.doc;
    if (!d.selection) { d.duplicateLayer(); return; }
    d.addLayer('Camada 1', maskedCopy(d.activeLayer.canvas, d.selection));
    d.commit('Camada via cópia');
  }),
  layerViaCut: needsDoc(() => {
    finishPending();
    const d = app.doc;
    const src = d.activeLayer;
    if (!d.selection || !editable(src)) return;
    const piece = maskedCopy(src.canvas, d.selection);
    d.editLayer(src, (ctx) => { ctx.globalCompositeOperation = 'destination-out'; ctx.drawImage(d.selection, 0, 0); });
    d.addLayer('Camada recortada', piece);
    d.commit('Camada via recorte');
  }),
  deleteLayer: needsDoc(() => { finishPending(); if (!app.doc.deleteLayer()) toast('O documento precisa de pelo menos uma camada.'); }),
  mergeDown: needsDoc(() => { finishPending(); if (!app.doc.mergeDown()) toast('Não há camada abaixo para mesclar.'); }),
  mergeVisible: needsDoc(() => { finishPending(); app.doc.mergeVisible(); }),
  flatten: needsDoc(() => { finishPending(); app.doc.flatten(); }),
  layerUp: needsDoc(() => { finishPending(); app.doc.moveLayer(app.doc.active, app.doc.active + 1); }),
  layerDown: needsDoc(() => { finishPending(); app.doc.moveLayer(app.doc.active, app.doc.active - 1); }),
  layerTop: needsDoc(() => { finishPending(); app.doc.moveLayer(app.doc.active, app.doc.layers.length - 1); }),
  layerBottom: needsDoc(() => { finishPending(); app.doc.moveLayer(app.doc.active, 0); }),
  editTextLayer: needsDoc(() => {
    const l = app.doc.activeLayer;
    if (l.text) editText(l, false);
    else toast('A camada ativa não é uma camada de texto.');
  }),
  rasterize: needsDoc(() => { const l = app.doc.activeLayer; if (l.text) { l.text = null; app.doc.commit('Rasterizar texto'); } }),
  flipLayerH: needsDoc(() => flipLayer(true)),
  flipLayerV: needsDoc(() => flipLayer(false)),
  fx: needsDoc(() => runFilter('dropShadow')),

  repeatFilter: needsDoc(() => { if (lastFilterId) runFilter(lastFilterId, true); }),

  zoomIn: needsDoc(() => app.view.zoomStep(1)),
  zoomOut: needsDoc(() => app.view.zoomStep(-1)),
  fit: needsDoc(() => app.view.fit()),
  actual: needsDoc(() => { app.view.setZoom(1); app.view.center(); }),
  grid: () => { app.showGrid = !app.showGrid; app.view.requestRender(false); },
  panels: () => { document.body.classList.toggle('hide-panels'); },
  fullscreen: () => {
    if (native?.toggleFullscreen) native.toggleFullscreen();
    else if (document.fullscreenElement) document.exitFullscreen();
    else document.documentElement.requestFullscreen?.();
  },
  shortcuts: shortcutsDialog,
  about: aboutDialog,
};

function flipLayer(h) {
  finishPending();
  const d = app.doc;
  const l = d.activeLayer;
  if (!editable(l)) return;
  const c = createCanvas(d.width, d.height);
  const ctx = c.getContext('2d');
  if (h) { ctx.translate(d.width, 0); ctx.scale(-1, 1); } else { ctx.translate(0, d.height); ctx.scale(1, -1); }
  ctx.drawImage(l.canvas, 0, 0);
  d.replaceLayerCanvas(l, c);
  d.commit(h ? 'Espelhar camada (H)' : 'Espelhar camada (V)');
}

for (const id of Object.keys(FILTERS)) commands[`filter:${id}`] = needsDoc(() => runFilter(id));

function run(cmd) {
  if (isDialogOpen()) return;
  const fn = commands[cmd];
  if (fn) Promise.resolve(fn()).catch((err) => { console.error(err); toast(`Erro: ${err.message}`); });
}

// ---------------------------------------------------------------------------
// Menus

const filterItems = (ids) => ids.map((id) => (id === '-' ? '-' : { label: `${FILTERS[id].name}${FILTERS[id].params ? '…' : ''}`, cmd: `filter:${id}`, key: FILTER_KEYS[id] }));
const FILTER_KEYS = {
  levels: 'Ctrl+L', curves: 'Ctrl+M', hueSaturation: 'Ctrl+U', colorBalance: 'Ctrl+B', invert: 'Ctrl+I',
  desaturate: 'Ctrl+Shift+U', autoTone: 'Ctrl+Shift+L', blackWhite: 'Ctrl+Alt+Shift+B',
};

const MENUS = [
  ['Arquivo', [
    { label: 'Novo…', cmd: 'new', key: 'Ctrl+N' },
    { label: 'Abrir…', cmd: 'open', key: 'Ctrl+O' },
    { label: 'Inserir imagem como camada…', cmd: 'place', key: 'Ctrl+Shift+P' },
    '-',
    { label: 'Salvar projeto', cmd: 'save', key: 'Ctrl+S' },
    { label: 'Salvar projeto como…', cmd: 'saveAs', key: 'Ctrl+Shift+S' },
    { label: 'Exportar como…', cmd: 'export', key: 'Ctrl+Alt+Shift+W' },
    { label: 'Exportação rápida como PNG', cmd: 'quickPng', key: 'Ctrl+Shift+E' },
    '-',
    { label: 'Fechar', cmd: 'close', key: 'Ctrl+W' },
    isNative ? { label: 'Sair', cmd: 'quit', key: 'Ctrl+Q' } : null,
  ]],
  ['Editar', [
    { label: 'Desfazer', cmd: 'undo', key: 'Ctrl+Z' },
    { label: 'Refazer', cmd: 'redo', key: 'Ctrl+Shift+Z', alt: ['Ctrl+Y'] },
    '-',
    { label: 'Recortar', cmd: 'cut', key: 'Ctrl+X' },
    { label: 'Copiar', cmd: 'copy', key: 'Ctrl+C' },
    { label: 'Copiar mesclado', cmd: 'copyMerged', key: 'Ctrl+Shift+C' },
    { label: 'Colar', cmd: 'paste', key: 'Ctrl+V', noKey: true },
    { label: 'Apagar', cmd: 'clear', key: 'Delete', alt: ['Backspace'] },
    '-',
    { label: 'Preencher…', cmd: 'fill', key: 'Shift+F5' },
    { label: 'Preencher com cor de frente', cmd: 'fillFg', key: 'Alt+Backspace' },
    { label: 'Preencher com cor de fundo', cmd: 'fillBg', key: 'Ctrl+Backspace' },
    { label: 'Traçar…', cmd: 'stroke' },
    '-',
    { label: 'Transformação livre', cmd: 'transform', key: 'Ctrl+T' },
  ]],
  ['Imagem', [
    { label: 'Ajustes', sub: filterItems(FILTER_MENU.adjust) },
    '-',
    { label: 'Tamanho da imagem…', cmd: 'imageSize', key: 'Ctrl+Alt+I' },
    { label: 'Tamanho da tela…', cmd: 'canvasSize', key: 'Ctrl+Alt+C' },
    { label: 'Rotação da imagem', sub: [
      { label: '90° horário', cmd: 'rotateCW' },
      { label: '90° anti-horário', cmd: 'rotateCCW' },
      { label: '180°', cmd: 'rotate180' },
      '-',
      { label: 'Espelhar tela horizontalmente', cmd: 'flipH' },
      { label: 'Espelhar tela verticalmente', cmd: 'flipV' },
    ] },
    '-',
    { label: 'Cortar para seleção', cmd: 'cropToSel' },
    { label: 'Aparar transparência', cmd: 'trim' },
  ]],
  ['Camada', [
    { label: 'Nova camada', cmd: 'newLayer', key: 'Ctrl+Shift+N' },
    { label: 'Duplicar camada', cmd: 'duplicateLayer' },
    { label: 'Camada via cópia', cmd: 'layerViaCopy', key: 'Ctrl+J' },
    { label: 'Camada via recorte', cmd: 'layerViaCut', key: 'Ctrl+Shift+J' },
    { label: 'Excluir camada', cmd: 'deleteLayer' },
    '-',
    { label: 'Organizar', sub: [
      { label: 'Trazer para a frente', cmd: 'layerTop', key: 'Ctrl+Shift+]' },
      { label: 'Avançar', cmd: 'layerUp', key: 'Ctrl+]' },
      { label: 'Recuar', cmd: 'layerDown', key: 'Ctrl+[' },
      { label: 'Enviar para trás', cmd: 'layerBottom', key: 'Ctrl+Shift+[' },
    ] },
    { label: 'Espelhar camada horizontalmente', cmd: 'flipLayerH' },
    { label: 'Espelhar camada verticalmente', cmd: 'flipLayerV' },
    '-',
    { label: 'Editar texto…', cmd: 'editTextLayer' },
    { label: 'Rasterizar texto', cmd: 'rasterize' },
    '-',
    { label: 'Mesclar para baixo', cmd: 'mergeDown', key: 'Ctrl+E' },
    { label: 'Mesclar visíveis', cmd: 'mergeVisible', key: 'Ctrl+Shift+Alt+E' },
    { label: 'Achatar imagem', cmd: 'flatten' },
  ]],
  ['Selecionar', [
    { label: 'Tudo', cmd: 'selectAll', key: 'Ctrl+A' },
    { label: 'Desmarcar', cmd: 'deselect', key: 'Ctrl+D' },
    { label: 'Inverter', cmd: 'invertSel', key: 'Ctrl+Shift+I' },
    { label: 'Pixels da camada', cmd: 'selectLayer' },
    '-',
    { label: 'Modificar', sub: [
      { label: 'Expandir…', cmd: 'expandSel' },
      { label: 'Contrair…', cmd: 'contractSel' },
      { label: 'Suavizar…', cmd: 'smoothSel' },
      { label: 'Difusão…', cmd: 'featherSel', key: 'Shift+F6' },
    ] },
  ]],
  ['Filtro', [
    { label: 'Repetir último filtro', cmd: 'repeatFilter', key: 'Ctrl+F' },
    '-',
    ...FILTER_MENU.filters.map(([label, ids]) => ({ label, sub: filterItems(ids) })),
  ]],
  ['Exibir', [
    { label: 'Ampliar', cmd: 'zoomIn', key: 'Ctrl+=', alt: ['Ctrl++'] },
    { label: 'Reduzir', cmd: 'zoomOut', key: 'Ctrl+-' },
    { label: 'Ajustar à tela', cmd: 'fit', key: 'Ctrl+0' },
    { label: '100% (pixels reais)', cmd: 'actual', key: 'Ctrl+1' },
    '-',
    { label: 'Grade', cmd: 'grid', key: "Ctrl+'", check: () => app.showGrid },
    { label: 'Mostrar/ocultar painéis', cmd: 'panels', key: 'Tab' },
    { label: 'Tela cheia', cmd: 'fullscreen', key: 'F11' },
  ]],
  ['Ajuda', [
    { label: 'Atalhos de teclado', cmd: 'shortcuts', key: 'F1' },
    { label: 'Sobre o Lumina Studio', cmd: 'about' },
  ]],
];

const shortcutMap = new Map();
const displayKey = (k) => (IS_MAC ? k.replace(/Ctrl\+/g, '⌘').replace(/Alt\+/g, '⌥').replace(/Shift\+/g, '⇧') : k);

function buildMenuList(items) {
  const list = el('div', { class: 'menu-list' });
  for (const it of items) {
    if (!it) continue;
    if (it === '-') { list.append(el('div', { class: 'menu-sep' })); continue; }
    if (it.sub) {
      const row = el('div', { class: 'menu-item sub' }, el('span', {}, it.label));
      row.append(buildMenuList(it.sub));
      list.append(row);
      continue;
    }
    if (it.key && !it.noKey) {
      shortcutMap.set(normalizeCombo(it.key), it.cmd);
      for (const a of it.alt || []) shortcutMap.set(normalizeCombo(a), it.cmd);
    }
    const row = el('div', {
      class: `menu-item${it.check ? ' check' : ''}`,
      onclick: () => { closeMenus(); run(it.cmd); },
    }, el('span', {}, it.label), it.key ? el('span', { class: 'kbd' }, displayKey(it.key)) : null);
    if (it.check) row._check = it.check;
    list.append(row);
  }
  return list;
}

let menuOpen = false;
function closeMenus() {
  document.querySelectorAll('.menu.open').forEach((m) => m.classList.remove('open'));
  menuOpen = false;
}

function setupMenus() {
  const nav = $('#menus');
  for (const [label, items] of MENUS) {
    const m = el('div', { class: 'menu' });
    const btn = el('button', {}, label);
    btn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      const was = m.classList.contains('open');
      closeMenus();
      if (!was) { openMenu(m); }
    });
    btn.addEventListener('mouseenter', () => { if (menuOpen && !m.classList.contains('open')) { closeMenus(); openMenu(m); } });
    m.append(btn, buildMenuList(items));
    nav.append(m);
  }
  document.addEventListener('mousedown', (e) => { if (!e.target.closest('.menu')) closeMenus(); });
}

function openMenu(m) {
  m.classList.add('open');
  menuOpen = true;
  m.querySelectorAll('.menu-item.check').forEach((r) => r.classList.toggle('on', !!r._check()));
}

// ---------------------------------------------------------------------------
// Teclado

let spaceDown = false;

function comboOf(e) {
  let key;
  const code = e.code || '';
  if (code.startsWith('Key')) key = code.slice(3);
  else if (code.startsWith('Digit')) key = code.slice(5);
  else if (code.startsWith('Numpad') && /\d$/.test(code)) key = code.slice(6);
  else key = CODE_KEYS[code];
  if (!key) key = e.key.length === 1 ? e.key.toUpperCase() : e.key;
  return { combo: makeCombo(e.ctrlKey || e.metaKey, e.altKey, e.shiftKey, key), key };
}

const CODE_KEYS = { Equal: '=', Minus: '-', NumpadAdd: '+', NumpadSubtract: '-', BracketLeft: '[', BracketRight: ']', Quote: "'", Backquote: "'" };

function makeCombo(ctrl, alt, shift, key) {
  return [ctrl && 'Ctrl', alt && 'Alt', shift && 'Shift', key].filter(Boolean).join('+');
}

// Normaliza "Ctrl+Shift+Alt+E" → "Ctrl+Alt+Shift+E" (e trata "Ctrl++").
function normalizeCombo(str) {
  const plus = str.endsWith('++');
  const parts = (plus ? str.slice(0, -2) : str).split('+').filter(Boolean);
  const key = plus ? '+' : parts.pop();
  return makeCombo(parts.includes('Ctrl'), parts.includes('Alt'), parts.includes('Shift'), key.length === 1 ? key.toUpperCase() : key);
}

function setupKeyboard() {
  document.addEventListener('keydown', (e) => {
    if (isDialogOpen()) return;
    const t = e.target;
    if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement) {
      if (e.key === 'Escape') t.blur();
      return;
    }
    if (e.key === 'Alt') e.preventDefault();
    if (e.code === 'Space' && !spaceDown) {
      spaceDown = true;
      app.view.canvas.style.cursor = 'grab';
      e.preventDefault();
      return;
    }
    if (menuOpen && e.key === 'Escape') { closeMenus(); return; }
    if (app.tool?.keydown && app.doc && !e.ctrlKey && !e.metaKey && app.tool.keydown(e)) { e.preventDefault(); return; }

    const { combo, key } = comboOf(e);
    const cmd = shortcutMap.get(combo);
    if (cmd) { e.preventDefault(); run(cmd); return; }
    if (e.ctrlKey || e.metaKey || e.altKey) return;

    if (key === 'Tab') { e.preventDefault(); run('panels'); return; }
    if (key === 'X') { swapColors(); return; }
    if (key === 'D') { resetColors(); return; }
    if (key === '[' || key === ']') {
      const tool = app.tool;
      if (tool?.opts?.size != null) {
        const s = tool.opts.size;
        const step = s < 10 ? 1 : s < 50 ? 5 : s < 200 ? 10 : 25;
        tool.opts.size = clamp(s + (key === ']' ? step : -step), 1, 1000);
        saveToolOptions();
        renderOptions();
        app.view.requestRender(false);
      }
      return;
    }
    if (/^[0-9]$/.test(key) && app.tool?.opts?.opacity != null && app.tool.options.some((o) => o.key === 'opacity' && o.type === 'range')) {
      app.tool.opts.opacity = key === '0' ? 100 : +key * 10;
      saveToolOptions();
      renderOptions();
      return;
    }
    if (key.length === 1 && cycleToolByKey(key.toLowerCase())) e.preventDefault();
  });
  document.addEventListener('keyup', (e) => {
    if (e.code === 'Space') {
      spaceDown = false;
      if (app.tool) app.view.canvas.style.cursor = app.tool.cursor || 'default';
    }
  });
  window.addEventListener('blur', () => { spaceDown = false; });
}

// ---------------------------------------------------------------------------
// Ponteiro sobre a tela

function setupPointer() {
  const canvas = app.view.canvas;
  let mode = null; // 'tool' | 'pan'
  let last = null;
  const sp = (e) => {
    const r = canvas.getBoundingClientRect();
    return { x: e.clientX - r.left, y: e.clientY - r.top };
  };
  canvas.addEventListener('contextmenu', (e) => e.preventDefault());
  canvas.addEventListener('pointerdown', (e) => {
    if (!app.doc || isDialogOpen()) return;
    closeMenus();
    document.activeElement?.blur?.();
    const s = sp(e);
    canvas.setPointerCapture(e.pointerId);
    if (e.button === 1 || (e.button === 0 && spaceDown)) {
      mode = 'pan';
      app.view.interacting = true;
      last = s;
      canvas.style.cursor = 'grabbing';
      e.preventDefault();
      return;
    }
    if (e.button !== 0) return;
    mode = 'tool';
    app.view.interacting = true;
    app.tool?.down?.(e, app.view.toDoc(s.x, s.y), s);
  });
  canvas.addEventListener('pointermove', (e) => {
    if (!app.doc) return;
    const s = sp(e);
    const p = app.view.toDoc(s.x, s.y);
    app.view.pointer = p;
    app.view.pointerScreen = s;
    const inside = p.x >= 0 && p.y >= 0 && p.x < app.doc.width && p.y < app.doc.height;
    $('#st-pos').textContent = inside ? `X: ${Math.floor(p.x)}  Y: ${Math.floor(p.y)}` : '';
    if (mode === 'pan') {
      app.view.pan(s.x - last.x, s.y - last.y);
      last = s;
      return;
    }
    if (mode === 'tool') {
      const evs = e.getCoalescedEvents?.() || [e];
      for (const ev of evs.length ? evs : [e]) {
        const es = sp(ev);
        app.tool?.move?.(ev, app.view.toDoc(es.x, es.y), es);
      }
    } else {
      app.tool?.hoverMove?.(e, p, s);
    }
    if (app.tool?.drawOverlay) app.view.requestRender(false);
  });
  const end = (e) => {
    if (mode === 'pan') {
      canvas.style.cursor = spaceDown ? 'grab' : app.tool?.cursor || 'default';
    } else if (mode === 'tool') {
      const s = sp(e);
      app.tool?.up?.(e, app.view.toDoc(s.x, s.y), s);
      app.emit('status-info', '');
    }
    mode = null;
    app.view.interacting = false;
    app.view.requestRender(false);
  };
  canvas.addEventListener('pointerup', end);
  canvas.addEventListener('pointercancel', end);
  canvas.addEventListener('pointerleave', () => {
    app.view.pointerScreen = null;
    $('#st-pos').textContent = '';
    app.view.requestRender(false);
  });
  canvas.addEventListener('wheel', (e) => {
    if (!app.doc) return;
    e.preventDefault();
    const s = sp(e);
    if (e.shiftKey && !e.ctrlKey) { app.view.pan(-e.deltaY, 0); return; }
    const factor = Math.exp(-e.deltaY * (e.ctrlKey ? 0.01 : 0.0015));
    app.view.setZoom(app.doc.view.zoom * factor, s.x, s.y);
  }, { passive: false });
}

// ---------------------------------------------------------------------------
// Arrastar e soltar / colar

function setupDropAndPaste() {
  const wrap = $('#canvas-wrap');
  const hasFiles = (e) => [...(e.dataTransfer?.types || [])].includes('Files');
  wrap.addEventListener('dragover', (e) => { if (hasFiles(e)) { e.preventDefault(); wrap.classList.add('drop'); } });
  wrap.addEventListener('dragleave', () => wrap.classList.remove('drop'));
  wrap.addEventListener('drop', async (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    wrap.classList.remove('drop');
    const files = [];
    for (const f of e.dataTransfer.files) files.push({ name: f.name, path: native?.pathForFile?.(f) || null, data: new Uint8Array(await f.arrayBuffer()) });
    await openFiles(files);
  });
  document.addEventListener('paste', async (e) => {
    if (isDialogOpen()) return;
    const t = e.target;
    if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement) return;
    const item = [...(e.clipboardData?.items || [])].find((i) => i.type.startsWith('image/'));
    e.preventDefault();
    if (item) {
      const blob = item.getAsFile();
      const url = URL.createObjectURL(blob);
      const img = new Image();
      img.src = url;
      await img.decode();
      URL.revokeObjectURL(url);
      pasteImage(img, app.clipboard);
    } else if (app.clipboard) {
      pasteImage(app.clipboard.canvas, app.clipboard);
    }
  });
}

// ---------------------------------------------------------------------------
// Integração com o aplicativo desktop

function setupNative() {
  window.addEventListener('beforeunload', (e) => {
    if (!isNative && app.docs.some((d) => d.dirty)) { e.preventDefault(); e.returnValue = ''; }
  });
  if (!isNative) return;
  native.onCloseRequest(async () => {
    for (const d of [...app.docs]) {
      if (d.dirty) {
        switchDoc(d);
        const r = await askSave(d);
        if (r === 'cancel') return;
      }
    }
    native.confirmClose();
  });
  native.onOpenFiles((files) => openFiles(files));
  native.getLaunchFiles().then((files) => { if (files?.length) openFiles(files); });
}

// ---------------------------------------------------------------------------
// Inicialização

function setupPanels() {
  document.querySelectorAll('.panel > h3').forEach((h) => h.addEventListener('click', () => h.parentElement.classList.toggle('collapsed')));
  new ColorPicker($('#color-picker'));
  loadSwatches();
  renderSwatches();
  setupLayerPanel();
}

function init() {
  initToolOptions();
  app.view = new Viewport($('#view'));
  setupMenus();
  setupPanels();
  setupKeyboard();
  setupPointer();
  setupDropAndPaste();
  setupNative();
  document.querySelectorAll('[data-cmd]').forEach((b) => b.addEventListener('click', () => run(b.dataset.cmd)));
  setTool('brush');
  renderToolbar();
  renderTabs();
  updateStatus();
  window.__lumina = { app, commands, run, setTool, Layer };
}

init();
