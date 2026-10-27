// Lumina Apresentações — aplicação principal.
import {
  el, $, $$, toast, dialog, prompt, field, select, menubar, toolButton, sep, popover, closePopovers, colorPopover, dropdown,
  contextMenu, clamp, debounce, comboOf, isDialogOpen, PALETTE,
} from '../../shared/ui.js';
import { icon } from '../../shared/icons.js';
import {
  pickFiles, saveFile, exportPdf, isNative, native, registerApp, standaloneBoot, setTitle, addRecent, extOf, fileName,
  initTheme, pickImageDataUrl, loadVendor,
} from '../../shared/native.js';
import {
  THEMES, LAYOUTS, W, H, uid, newSlide, newPresentation, themedPresentation, layoutElements,
} from './themes.js';
import { renderSlide, SHAPES, SLIDE_CSS, themeOf, resolveColor, bgCss } from './render.js';
import * as io from './io.js';

const S = { decks: [], deck: null, seq: 1, sel: new Set(), editing: null, clipboard: null, scale: 1, focus: 'stage', chartImages: {}, charts: new Map() };
const deck = () => S.deck;
const pres = () => S.deck.pres;
const slide = () => S.deck.pres.slides[S.deck.current];
const holder = $('#holder');
const overlay = $('#overlay');
const stage = $('#stage');

// Estilos dos slides no documento
document.head.append(el('style', {}, SLIDE_CSS));

// ---------------------------------------------------------------------------
// Apresentações abertas

function makeDeck(p, { filePath = null, ext = 'lumslides' } = {}) {
  const d = { id: S.seq++, pres: p, filePath, ext, dirty: false, current: 0, hist: { stack: [], index: -1 } };
  pushHistory(d);
  return d;
}

function isPristine(d) {
  return d && !d.dirty && !d.filePath && d.pres.slides.length === 1 && d.pres.slides[0].elements.every((e) => !e.text?.html && !(e.image?.asset)) && d.hist.stack.length === 1;
}

function addDeck(d) {
  finishEditing();
  if (S.decks.length === 1 && isPristine(S.decks[0])) S.decks = [];
  S.decks.push(d);
  activate(d);
}

function activate(d) {
  finishEditing();
  S.deck = d;
  S.sel.clear();
  renderAll();
}

async function closeDeck(d = S.deck) {
  if (d.dirty && !(await askSave(d))) return false;
  const i = S.decks.indexOf(d);
  S.decks.splice(i, 1);
  if (!S.decks.length) S.decks.push(makeDeck(newPresentation()));
  activate(S.decks[Math.min(i, S.decks.length - 1)]);
  return true;
}

async function askSave(d) {
  if (d !== S.deck) activate(d);
  let choice = 'cancel';
  const ok = await dialog({
    title: 'Salvar alterações?', okText: 'Salvar',
    body: el('p', {}, `A apresentação "${d.pres.title}" tem alterações não salvas.`),
    extraButtons: [{ label: 'Não salvar', onClick: (close) => { choice = 'discard'; close(false); } }],
  });
  if (ok) return save(d);
  return choice === 'discard';
}

function newDeck(tpl = 'blank') {
  const p = tpl === 'blank' ? newPresentation('claro') : themedPresentation(tpl);
  if (tpl !== 'blank' && tpl !== 'pitch') p.title = `Apresentação ${THEMES[tpl]?.name || ''}`.trim();
  addDeck(makeDeck(p));
}

async function openFiles(files) {
  for (const f of files) {
    const ext = extOf(f.name);
    try {
      if (ext === 'pptx') {
        toast('Importar .pptx ainda não é suportado. Exporte para PDF ou recrie no Lumina — a exportação para PowerPoint funciona.');
        continue;
      }
      const p = io.parse(f.data);
      addDeck(makeDeck(p, { filePath: f.path, ext }));
      addRecent({ app: 'slides', name: f.name, path: f.path });
    } catch (err) {
      toast(`Não foi possível abrir ${f.name}: ${err.message}`);
    }
  }
}

async function openDialog() {
  const files = await pickFiles([{ name: 'Apresentação Lumina', extensions: ['lumslides'] }]);
  if (files.length) openFiles(files);
}

async function save(d = S.deck, saveAs = false) {
  finishEditing();
  try {
    const path = await saveFile(io.serialize(d.pres), {
      name: `${d.pres.title}.lumslides`, filters: [{ name: 'Apresentação Lumina', extensions: ['lumslides'] }],
      path: saveAs ? null : d.filePath,
    });
    if (!path) return false;
    if (isNative) d.filePath = path;
    d.dirty = false;
    addRecent({ app: 'slides', name: fileName(path), path: isNative ? path : null });
    toast(isNative ? `Salvo: ${fileName(path)}` : 'Apresentação baixada.');
    renderTitle();
    return true;
  } catch (err) {
    toast(`Erro ao salvar: ${err.message}`);
    return false;
  }
}

async function exportAs(fmt) {
  finishEditing();
  const p = pres();
  try {
    await ensureChartImages(p);
    if (fmt === 'pptx') {
      const bytes = await io.exportPptx(p, S.chartImages);
      const out = await saveFile(bytes, { name: `${p.title}.pptx`, filters: [{ name: 'PowerPoint', extensions: ['pptx'] }] });
      if (out) toast(`Exportado: ${fileName(out)}`);
    } else if (fmt === 'pdf') {
      const out = await exportPdf(io.presentationHtml(p, S.chartImages), { name: `${p.title}.pdf`, options: { pageSize: { width: 13.333, height: 7.5 } } });
      if (out) toast(isNative ? `PDF exportado: ${fileName(out)}` : 'Use "Salvar como PDF" na janela de impressão.');
    } else if (fmt === 'png') {
      const url = await io.slideToPng(slide(), p, { chartImages: S.chartImages });
      const bytes = Uint8Array.from(atob(url.split(',')[1]), (c) => c.charCodeAt(0));
      const out = await saveFile(bytes, { name: `${p.title} - slide ${deck().current + 1}.png`, filters: [{ name: 'PNG', extensions: ['png'] }] });
      if (out) toast('Imagem exportada.');
    }
  } catch (err) {
    console.error(err);
    toast(`Erro ao exportar: ${err.message}`);
  }
}

// ---------------------------------------------------------------------------
// Histórico

function snapshot(p) {
  const { assets, ...rest } = p;
  void assets;
  return JSON.stringify(rest);
}

function pushHistory(d = S.deck) {
  const h = d.hist;
  const snap = snapshot(d.pres);
  if (h.stack[h.index] === snap) return false;
  h.stack.splice(h.index + 1);
  h.stack.push(snap);
  if (h.stack.length > 150) h.stack.shift();
  h.index = h.stack.length - 1;
  return true;
}

// Registra uma alteração do usuário.
function commit() {
  const d = S.deck;
  if (!pushHistory(d)) return;
  if (!d.dirty) { d.dirty = true; renderTitle(); }
  renderThumbs();
}

function restore(i) {
  const d = S.deck;
  const assets = d.pres.assets;
  d.pres = { ...JSON.parse(d.hist.stack[i]), assets };
  d.hist.index = i;
  d.current = clamp(d.current, 0, d.pres.slides.length - 1);
  S.sel = new Set([...S.sel].filter((id) => slide().elements.some((e) => e.id === id)));
  d.dirty = true;
  renderAll();
}

function undo() {
  if (S.editing) { finishEditing(); }
  const h = S.deck.hist;
  if (h.index > 0) restore(h.index - 1);
}

function redo() {
  const h = S.deck.hist;
  if (h.index < h.stack.length - 1) restore(h.index + 1);
}

// ---------------------------------------------------------------------------
// Renderização: palco, miniaturas e painel

function renderAll() {
  renderTitle();
  renderDeckTabs();
  renderThumbs();
  renderStage();
  renderProps();
  renderStatus();
  $('#notes').value = slide().notes || '';
}

function fitScale() {
  const r = stage.getBoundingClientRect();
  S.scale = Math.max(0.1, Math.min((r.width - 64) / W, (r.height - 64) / H));
  const w = W * S.scale, h = H * S.scale;
  const left = (r.width - w) / 2, top = (r.height - h) / 2;
  Object.assign(holder.style, { left: `${left}px`, top: `${top}px`, width: `${W}px`, height: `${H}px`, transform: `scale(${S.scale})` });
  Object.assign(overlay.style, { left: `${left}px`, top: `${top}px`, width: `${w}px`, height: `${h}px` });
}

function renderStage() {
  for (const c of S.charts.values()) c.destroy();
  S.charts.clear();
  const node = renderSlide(slide(), pres(), { mode: 'edit', editingId: S.editing });
  holder.replaceChildren(node);
  fitScale();
  initCharts(node);
  if (S.editing) {
    const inner = holder.querySelector(`[data-id="${S.editing}"] .sl-text`);
    if (inner) { inner.contentEditable = 'true'; }
  }
  renderOverlay();
}

async function initCharts(root) {
  const canvases = root.querySelectorAll('.sl-chart-canvas');
  if (!canvases.length) return;
  const Chart = await loadVendor('chart.umd.min.js', 'Chart');
  for (const c of canvases) {
    const id = c.closest('.sl-el').dataset.id;
    const e = slide().elements.find((x) => x.id === id);
    if (!e) continue;
    c.width = e.w;
    c.height = e.h;
    const chart = new Chart(c, chartConfig(e, false));
    S.charts.set(id, chart);
    S.chartImages[id] = chart.toBase64Image();
  }
  renderThumbs();
}

function chartConfig(e, animate) {
  const t = themeOf(pres());
  const ch = e.chart;
  const pal = io.chartPalette(t);
  const pie = ch.kind === 'pie' || ch.kind === 'doughnut';
  return {
    type: { column: 'bar', bar: 'bar', line: 'line', area: 'line', pie: 'pie', doughnut: 'doughnut' }[ch.kind] || 'bar',
    data: {
      labels: ch.labels,
      datasets: pie
        ? [{ label: ch.series[0]?.name, data: ch.series[0]?.values || [], backgroundColor: ch.labels.map((_, i) => pal[i % pal.length]), borderWidth: 0 }]
        : ch.series.map((s, i) => ({ label: s.name, data: s.values, backgroundColor: ch.kind === 'area' ? `${pal[i % pal.length]}66` : pal[i % pal.length], borderColor: pal[i % pal.length], borderWidth: ch.kind === 'line' || ch.kind === 'area' ? 6 : 0, borderRadius: 10, fill: ch.kind === 'area', tension: 0.35, pointRadius: ch.kind === 'line' ? 8 : 0 })),
    },
    options: {
      responsive: false, animation: animate ? { duration: 900 } : false, devicePixelRatio: 1,
      indexAxis: ch.kind === 'bar' ? 'y' : 'x',
      plugins: {
        title: { display: !!ch.title, text: ch.title, color: t.title, font: { size: 40, weight: '700', family: t.head } },
        legend: { display: pie || ch.series.length > 1, position: 'bottom', labels: { color: t.text, font: { size: 26, family: t.body }, padding: 24 } },
      },
      scales: pie ? {} : {
        x: { ticks: { color: t.text, font: { size: 24, family: t.body } }, grid: { display: false } },
        y: { ticks: { color: t.text, font: { size: 24, family: t.body } }, grid: { color: `${t.text}22` }, beginAtZero: true },
      },
    },
  };
}

// Gera imagens dos gráficos de todos os slides (para miniaturas, apresentação e exportação).
async function ensureChartImages(p) {
  const todo = p.slides.flatMap((s) => s.elements.filter((e) => e.type === 'chart'));
  if (!todo.length) return;
  const Chart = await loadVendor('chart.umd.min.js', 'Chart');
  const prev = S.deck.pres;
  for (const e of todo) {
    const c = document.createElement('canvas');
    c.width = e.w;
    c.height = e.h;
    S.deck.pres = p;
    const chart = new Chart(c, chartConfig(e, false));
    S.chartImages[e.id] = chart.toBase64Image();
    chart.destroy();
  }
  S.deck.pres = prev;
}

const THUMB_SCALE = 182 / W;
function renderThumbs() {
  const box = $('#thumbs');
  const p = pres();
  box.replaceChildren(...p.slides.map((s, i) => {
    const node = renderSlide(s, p, { mode: 'thumb', chartImages: S.chartImages });
    node.style.transform = `scale(${THUMB_SCALE})`;
    const badges = el('div', { class: 'badges' },
      s.transition && s.transition !== 'none' ? el('span', { title: 'Transição' }, '⇢') : null,
      s.elements.some((e) => e.anim && e.anim.type !== 'none') ? el('span', { title: 'Animações' }, '★') : null,
      s.notes ? el('span', { title: 'Anotações' }, '✎') : null);
    const t = el('div', {
      class: `thumb${i === deck().current ? ' on' : ''}`, draggable: 'true',
      onclick: () => goTo(i),
      oncontextmenu: (e) => { e.preventDefault(); goTo(i); slideMenu(e); },
      ondragstart: (e) => { e.dataTransfer.setData('text/x-slide', String(i)); },
      ondragover: (e) => { e.preventDefault(); t.classList.add('drop-before'); },
      ondragleave: () => t.classList.remove('drop-before'),
      ondrop: (e) => { e.preventDefault(); t.classList.remove('drop-before'); const from = +e.dataTransfer.getData('text/x-slide'); if (!Number.isNaN(from)) moveSlide(from, i); },
    }, el('span', { class: 'num' }, i + 1), el('div', { class: 'frame' }, node, badges));
    return t;
  }), el('button', { class: 'btn small block', style: { marginTop: '4px', marginLeft: '26px', width: '182px' }, onclick: () => addSlide(), html: `${icon('plus', 14)} Novo slide` }));
  box.querySelector('.thumb.on')?.scrollIntoView({ block: 'nearest' });
}

const refreshCurrentThumb = debounce(() => {
  const t = $$('#thumbs .thumb')[deck().current];
  if (!t) return;
  const node = renderSlide(slide(), pres(), { mode: 'thumb', chartImages: S.chartImages });
  node.style.transform = `scale(${THUMB_SCALE})`;
  t.querySelector('.slide').replaceWith(node);
}, 250);

function renderTitle() {
  const d = deck();
  const t = $('#doc-title');
  if (document.activeElement !== t) t.value = d.pres.title;
  $('#save-state').textContent = d.dirty ? 'Alterações não salvas' : d.filePath ? `Salvo · ${fileName(d.filePath)}` : 'Não salvo ainda';
  setTitle(`${d.pres.title}${d.dirty ? ' •' : ''} — Lumina Apresentações`);
  renderDeckTabs();
}

function renderDeckTabs() {
  const box = $('#deck-tabs');
  box.classList.toggle('hidden', S.decks.length < 2);
  box.replaceChildren(...S.decks.map((d) => el('div', { class: `doc-tab${d === S.deck ? ' on' : ''}`, onclick: () => d !== S.deck && activate(d) },
    el('span', {}, `${d.pres.title}${d.dirty ? ' •' : ''}`),
    el('button', { class: 'x', html: icon('x', 12), onclick: (e) => { e.stopPropagation(); closeDeck(d); } }))));
}

function renderStatus() {
  const n = pres().slides.length;
  $('#status').replaceChildren(
    el('span', {}, `Slide ${deck().current + 1} de ${n}`),
    el('span', {}, `Tema: ${themeOf(pres()).name}`),
    el('span', { class: 'spacer' }),
    el('button', { class: 'btn small', onclick: () => $('#notes-wrap').classList.toggle('hidden'), html: `${icon('notes', 14)} Anotações` }),
    el('button', { class: 'btn small', onclick: () => present(deck().current), html: `${icon('play', 14)} Apresentar daqui` }));
}

// ---------------------------------------------------------------------------
// Slides

function goTo(i) {
  finishEditing();
  const d = deck();
  d.current = clamp(i, 0, pres().slides.length - 1);
  S.sel.clear();
  renderStage();
  renderProps();
  renderStatus();
  $$('#thumbs .thumb').forEach((t, k) => t.classList.toggle('on', k === d.current));
  $$('#thumbs .thumb')[d.current]?.scrollIntoView({ block: 'nearest' });
  $('#notes').value = slide().notes || '';
}

function addSlide(layout = null, at = deck().current + 1) {
  finishEditing();
  const cur = slide();
  const l = layout || (cur?.layout === 'title' ? 'titleContent' : cur?.layout || 'titleContent');
  const s = newSlide(l);
  s.transition = cur?.transition || 'fade';
  pres().slides.splice(at, 0, s);
  deck().current = at;
  S.sel.clear();
  commit();
  renderAll();
}

function duplicateSlide() {
  const s = JSON.parse(JSON.stringify(slide()));
  s.id = uid('s');
  for (const e of s.elements) e.id = uid();
  pres().slides.splice(deck().current + 1, 0, s);
  deck().current++;
  commit();
  renderAll();
}

function deleteSlide() {
  const p = pres();
  if (p.slides.length <= 1) { toast('A apresentação precisa de pelo menos um slide.'); return; }
  p.slides.splice(deck().current, 1);
  deck().current = Math.min(deck().current, p.slides.length - 1);
  S.sel.clear();
  commit();
  renderAll();
}

function moveSlide(from, to) {
  const p = pres();
  if (from === to) return;
  const [s] = p.slides.splice(from, 1);
  p.slides.splice(to > from ? to - 1 : to, 0, s);
  deck().current = p.slides.indexOf(s);
  commit();
  renderAll();
}

function slideMenu(e) {
  contextMenu([
    { label: 'Novo slide', icon: 'plus', key: 'Ctrl+M', action: () => addSlide() },
    { label: 'Duplicar slide', icon: 'copy', key: 'Ctrl+Shift+D', action: duplicateSlide },
    { label: 'Excluir slide', icon: 'trash', action: deleteSlide },
    '-',
    { label: 'Mover para cima', disabled: deck().current === 0, action: () => moveSlide(deck().current, deck().current - 1) },
    { label: 'Mover para baixo', disabled: deck().current >= pres().slides.length - 1, action: () => moveSlide(deck().current, deck().current + 2) },
    '-',
    { label: 'Apresentar a partir deste', icon: 'play', key: 'Shift+F5', action: () => present(deck().current) },
  ], e.clientX, e.clientY);
}

// ---------------------------------------------------------------------------
// Elementos: inserção

function addElement(e, { select: sel = true } = {}) {
  finishEditing();
  slide().elements.push(e);
  if (sel) S.sel = new Set([e.id]);
  commit();
  renderStage();
  renderProps();
}

function insertText() {
  addElement({ id: uid(), type: 'text', role: 'body', x: 560, y: 440, w: 800, h: 160, rot: 0, opacity: 1, text: { html: '<p>Texto</p>', size: 48, color: null, align: 'left', valign: 'top', bold: false, font: null, line: 1.25 } });
  startEditing(slide().elements[slide().elements.length - 1].id, true);
}

function insertShape(kind) {
  const isLine = kind === 'line';
  addElement({
    id: uid(), type: 'shape', x: 760, y: isLine ? 535 : 340, w: 400, h: isLine ? 10 : kind === 'arrow' || kind === 'chevron' ? 240 : 400, rot: 0, opacity: 1,
    shape: { kind, fill: isLine ? null : 'accent', stroke: isLine ? 'accent' : null, strokeW: isLine ? 8 : 0, radius: 32 },
    text: { html: '', size: 40, color: '#ffffff', align: 'center', valign: 'middle', bold: true, font: null, line: 1.2 },
  });
}

async function addImageFromDataUrl(src, at = null) {
  const img = new Image();
  img.src = src;
  await img.decode().catch(() => {});
  const iw = img.naturalWidth || 800, ih = img.naturalHeight || 600;
  const s = Math.min(1, 1100 / iw, 760 / ih);
  const w = Math.round(iw * s), h = Math.round(ih * s);
  const asset = uid('a');
  pres().assets[asset] = src;
  addElement({ id: uid(), type: 'image', x: at ? at.x - w / 2 : (W - w) / 2, y: at ? at.y - h / 2 : (H - h) / 2, w, h, rot: 0, opacity: 1, image: { asset, fit: 'cover', radius: 0 } });
}

async function insertImage(target = null) {
  const src = await pickImageDataUrl();
  if (!src) return;
  if (target) {
    const asset = uid('a');
    pres().assets[asset] = src;
    target.image.asset = asset;
    commit();
    renderStage();
    renderProps();
    return;
  }
  addImageFromDataUrl(src);
}

function insertTable(rows = 4, cols = 3) {
  const data = Array.from({ length: rows }, (_, r) => Array.from({ length: cols }, (_, c) => (r === 0 ? `Coluna ${c + 1}` : '')));
  addElement({ id: uid(), type: 'table', x: 260, y: 300, w: 1400, h: rows * 96, rot: 0, opacity: 1, table: { rows: data, header: true, size: 32 } });
}

function insertChart(kind = 'column') {
  addElement({
    id: uid(), type: 'chart', x: 360, y: 260, w: 1200, h: 680, rot: 0, opacity: 1,
    chart: { kind, title: '', labels: ['T1', 'T2', 'T3', 'T4'], series: [{ name: '2025', values: [32, 45, 51, 68] }, { name: '2026', values: [40, 58, 66, 90] }] },
  });
}

// ---------------------------------------------------------------------------
// Seleção, arrastar, redimensionar, girar

const elById = (id) => slide().elements.find((e) => e.id === id);
const selected = () => slide().elements.filter((e) => S.sel.has(e.id));

function clientToSlide(x, y) {
  const r = holder.getBoundingClientRect();
  return { x: (x - r.left) / S.scale, y: (y - r.top) / S.scale };
}

function bounds(list) {
  const xs = [], ys = [];
  for (const e of list) {
    for (const [px, py] of corners(e)) { xs.push(px); ys.push(py); }
  }
  return { x: Math.min(...xs), y: Math.min(...ys), w: Math.max(...xs) - Math.min(...xs), h: Math.max(...ys) - Math.min(...ys) };
}

function corners(e) {
  const cx = e.x + e.w / 2, cy = e.y + e.h / 2, a = ((e.rot || 0) * Math.PI) / 180;
  return [[-1, -1], [1, -1], [1, 1], [-1, 1]].map(([sx, sy]) => {
    const lx = (sx * e.w) / 2, ly = (sy * e.h) / 2;
    return [cx + lx * Math.cos(a) - ly * Math.sin(a), cy + lx * Math.sin(a) + ly * Math.cos(a)];
  });
}

function renderOverlay(guides = [], marquee = null, tip = null) {
  const sc = S.scale;
  const items = [];
  const list = selected();
  if (list.length === 1 && !S.editing) {
    const e = list[0];
    const box = el('div', { class: 'sel-box', style: { left: `${e.x * sc}px`, top: `${e.y * sc}px`, width: `${e.w * sc}px`, height: `${e.h * sc}px`, transform: `rotate(${e.rot || 0}deg)` } });
    const hs = e.type === 'shape' && e.shape.kind === 'line' ? ['w', 'e'] : ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'];
    const pos = { nw: [0, 0], n: [50, 0], ne: [100, 0], e: [100, 50], se: [100, 100], s: [50, 100], sw: [0, 100], w: [0, 50] };
    const cursor = { nw: 'nwse-resize', se: 'nwse-resize', ne: 'nesw-resize', sw: 'nesw-resize', n: 'ns-resize', s: 'ns-resize', e: 'ew-resize', w: 'ew-resize' };
    for (const h of hs) box.append(el('div', { class: 'handle', 'data-h': h, style: { left: `${pos[h][0]}%`, top: `${pos[h][1]}%`, cursor: cursor[h] } }));
    box.append(el('div', { class: 'handle rot', 'data-h': 'rot', style: { left: '50%', top: '-30px' } }));
    items.push(box);
  } else if (list.length > 1) {
    for (const e of list) items.push(el('div', { class: 'sel-box', style: { left: `${e.x * sc}px`, top: `${e.y * sc}px`, width: `${e.w * sc}px`, height: `${e.h * sc}px`, transform: `rotate(${e.rot || 0}deg)` } }));
    const b = bounds(list);
    items.push(el('div', { class: 'sel-box multi', style: { left: `${b.x * sc - 4}px`, top: `${b.y * sc - 4}px`, width: `${b.w * sc + 8}px`, height: `${b.h * sc + 8}px` } }));
  }
  for (const g of guides) items.push(el('div', { class: `guide ${g.axis}`, style: g.axis === 'v' ? { left: `${g.pos * sc}px` } : { top: `${g.pos * sc}px` } }));
  if (marquee) items.push(el('div', { class: 'marquee', style: { left: `${marquee.x * sc}px`, top: `${marquee.y * sc}px`, width: `${marquee.w * sc}px`, height: `${marquee.h * sc}px` } }));
  if (tip) items.push(el('div', { class: 'size-tip', style: { left: `${tip.x * sc + 8}px`, top: `${tip.y * sc + 8}px` } }, tip.text));
  overlay.replaceChildren(...items);
}

// Encaixe magnético nas bordas/centros do slide e dos outros elementos.
function snap(box, movingIds, threshold = 8 / S.scale) {
  const xs = [0, W / 2, W], ys = [0, H / 2, H];
  for (const e of slide().elements) {
    if (movingIds.has(e.id)) continue;
    const b = bounds([e]);
    xs.push(b.x, b.x + b.w / 2, b.x + b.w);
    ys.push(b.y, b.y + b.h / 2, b.y + b.h);
  }
  const best = (vals, edges) => {
    let out = null;
    for (const edge of edges) for (const v of vals) {
      const d = v - edge;
      if (Math.abs(d) <= threshold && (!out || Math.abs(d) < Math.abs(out.d))) out = { d, pos: v };
    }
    return out;
  };
  const sx = best(xs, [box.x, box.x + box.w / 2, box.x + box.w]);
  const sy = best(ys, [box.y, box.y + box.h / 2, box.y + box.h]);
  return { dx: sx?.d || 0, dy: sy?.d || 0, guides: [sx && { axis: 'v', pos: sx.pos }, sy && { axis: 'h', pos: sy.pos }].filter(Boolean) };
}

function onStageDown(e) {
  if (e.button !== 0) return;
  const handle = e.target.closest('.handle');
  if (handle) { e.preventDefault(); startHandleDrag(handle.dataset.h); return; }
  const node = e.target.closest('.sl-el');
  if (S.editing) {
    if (node && node.dataset.id === S.editing) return; // clique dentro do texto em edição
    finishEditing();
  }
  S.focus = 'stage';
  if (!node) {
    if (!e.shiftKey) S.sel.clear();
    startMarquee(e);
    renderProps();
    return;
  }
  e.preventDefault();
  const id = node.dataset.id;
  if (e.shiftKey) {
    if (S.sel.has(id)) S.sel.delete(id); else S.sel.add(id);
    renderOverlay();
    renderProps();
    return;
  }
  if (!S.sel.has(id)) { S.sel = new Set([id]); renderProps(); }
  startMove(e);
}

function startMove(e) {
  const start = clientToSlide(e.clientX, e.clientY);
  const list = selected();
  const orig = list.map((x) => ({ x: x.x, y: x.y }));
  const ids = new Set(list.map((x) => x.id));
  let moved = false;
  // Alt+arrastar duplica
  let dup = e.altKey;
  const move = (ev) => {
    const p = clientToSlide(ev.clientX, ev.clientY);
    let dx = p.x - start.x, dy = p.y - start.y;
    if (!moved && Math.hypot(dx, dy) * S.scale < 3) return;
    if (!moved && dup) {
      const copies = list.map((x) => ({ ...JSON.parse(JSON.stringify(x)), id: uid() }));
      slide().elements.push(...copies);
      S.sel = new Set(copies.map((c) => c.id));
      list.splice(0, list.length, ...copies);
      ids.clear();
      copies.forEach((c) => ids.add(c.id));
      dup = false;
      renderStage();
    }
    moved = true;
    if (ev.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    list.forEach((x, i) => { x.x = orig[i].x + dx; x.y = orig[i].y + dy; });
    const b = bounds(list);
    const sn = ev.altKey ? { dx: 0, dy: 0, guides: [] } : snap(b, ids);
    list.forEach((x) => { x.x = Math.round(x.x + sn.dx); x.y = Math.round(x.y + sn.dy); });
    for (const x of list) {
      const n = holder.querySelector(`[data-id="${x.id}"]`);
      if (n) { n.style.left = `${x.x}px`; n.style.top = `${x.y}px`; }
    }
    renderOverlay(sn.guides);
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    renderOverlay();
    if (moved) { commit(); renderProps(); }
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startMarquee(e) {
  const start = clientToSlide(e.clientX, e.clientY);
  const base = new Set(S.sel);
  const move = (ev) => {
    const p = clientToSlide(ev.clientX, ev.clientY);
    const m = { x: Math.min(start.x, p.x), y: Math.min(start.y, p.y), w: Math.abs(p.x - start.x), h: Math.abs(p.y - start.y) };
    S.sel = new Set(base);
    for (const x of slide().elements) {
      const b = bounds([x]);
      if (b.x < m.x + m.w && b.x + b.w > m.x && b.y < m.y + m.h && b.y + b.h > m.y) S.sel.add(x.id);
    }
    renderOverlay([], m);
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    renderOverlay();
    renderProps();
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

function startHandleDrag(h) {
  const el0 = selected()[0];
  if (!el0) return;
  const o = { x: el0.x, y: el0.y, w: el0.w, h: el0.h, rot: el0.rot || 0 };
  const cx = o.x + o.w / 2, cy = o.y + o.h / 2;
  const a = (o.rot * Math.PI) / 180;
  const node = holder.querySelector(`[data-id="${el0.id}"]`);
  const move = (ev) => {
    const p = clientToSlide(ev.clientX, ev.clientY);
    if (h === 'rot') {
      let ang = (Math.atan2(p.y - cy, p.x - cx) * 180) / Math.PI + 90;
      if (ev.shiftKey) ang = Math.round(ang / 15) * 15;
      else { const near = Math.round(ang / 90) * 90; if (Math.abs(ang - near) < 3) ang = near; }
      el0.rot = Math.round(((ang % 360) + 360) % 360);
      if (el0.rot > 180) el0.rot -= 360;
      node.style.transform = `rotate(${el0.rot}deg)`;
      renderOverlay([], null, { x: p.x, y: p.y, text: `${el0.rot}°` });
      return;
    }
    // Coordenadas locais (sem rotação) relativas ao centro original
    const dx = p.x - cx, dy = p.y - cy;
    const lx = dx * Math.cos(-a) - dy * Math.sin(-a), ly = dx * Math.sin(-a) + dy * Math.cos(-a);
    const hx = h.includes('e') ? 1 : h.includes('w') ? -1 : 0;
    const hy = h.includes('s') ? 1 : h.includes('n') ? -1 : 0;
    let left = -o.w / 2, right = o.w / 2, top = -o.h / 2, bottom = o.h / 2;
    if (hx > 0) right = lx; if (hx < 0) left = lx;
    if (hy > 0) bottom = ly; if (hy < 0) top = ly;
    let w = Math.max(20, right - left), hh = Math.max(el0.type === 'shape' && el0.shape.kind === 'line' ? 4 : 20, bottom - top);
    const keep = (el0.type === 'image' && hx && hy) ? !ev.shiftKey : ev.shiftKey && hx && hy;
    if (keep) {
      const ratio = o.w / o.h;
      if (w / hh > ratio) w = hh * ratio; else hh = w / ratio;
      if (hx < 0) left = right - w; else right = left + w;
      if (hy < 0) top = bottom - hh; else bottom = top + hh;
    }
    if (hx < 0) left = right - w; else right = left + w;
    if (hy < 0) top = bottom - hh; else bottom = top + hh;
    const lcx = (left + right) / 2, lcy = (top + bottom) / 2;
    const ncx = cx + lcx * Math.cos(a) - lcy * Math.sin(a), ncy = cy + lcx * Math.sin(a) + lcy * Math.cos(a);
    el0.w = Math.round(w);
    el0.h = Math.round(hh);
    el0.x = Math.round(ncx - w / 2);
    el0.y = Math.round(ncy - hh / 2);
    Object.assign(node.style, { left: `${el0.x}px`, top: `${el0.y}px`, width: `${el0.w}px`, height: `${el0.h}px` });
    const ch = S.charts.get(el0.id);
    if (ch) { ch.canvas.width = el0.w; ch.canvas.height = el0.h; ch.resize(el0.w, el0.h); }
    renderOverlay([], null, { x: p.x, y: p.y, text: `${el0.w} × ${el0.h}` });
  };
  const up = () => {
    document.removeEventListener('mousemove', move);
    document.removeEventListener('mouseup', up);
    if (el0.x !== o.x || el0.y !== o.y || el0.w !== o.w || el0.h !== o.h || (el0.rot || 0) !== o.rot) { commit(); renderStage(); }
    renderOverlay();
    renderProps();
  };
  document.addEventListener('mousemove', move);
  document.addEventListener('mouseup', up);
}

// ---------------------------------------------------------------------------
// Edição de texto

function startEditing(id, selectAll = false) {
  const e = elById(id);
  if (!e || !e.text) return;
  S.editing = id;
  S.sel = new Set([id]);
  const node = holder.querySelector(`[data-id="${id}"]`);
  let inner = node?.querySelector('.sl-text');
  if (!inner) { renderStage(); inner = holder.querySelector(`[data-id="${id}"] .sl-text`); }
  if (!inner) return;
  node.classList.add('editing');
  inner.classList.remove('is-empty');
  if (!e.text.html && e.list) inner.innerHTML = '<ul><li><br></li></ul>';
  else if (!e.text.html) inner.innerHTML = '<p><br></p>';
  inner.contentEditable = 'true';
  inner.focus();
  const r = document.createRange();
  r.selectNodeContents(inner);
  if (!selectAll) r.collapse(false);
  const s = getSelection();
  s.removeAllRanges();
  s.addRange(r);
  renderOverlay();
  renderProps();
  inner.addEventListener('input', () => {
    e.text.html = inner.innerHTML;
    refreshCurrentThumb();
  });
}

function finishEditing() {
  if (!S.editing) return;
  const id = S.editing;
  S.editing = null;
  const e = elById(id);
  const inner = holder.querySelector(`[data-id="${id}"] .sl-text`);
  if (e && inner) {
    const html = inner.innerHTML.replace(/<p><br><\/p>|<ul><li><br><\/li><\/ul>/g, '').trim();
    e.text.html = inner.textContent.trim() || inner.querySelector('img') ? html : '';
    inner.contentEditable = 'false';
  }
  commit();
  renderStage();
  renderProps();
}

function onStageDblClick(e) {
  const node = e.target.closest('.sl-el');
  if (!node) return;
  const x = elById(node.dataset.id);
  if (!x) return;
  if (x.type === 'image') insertImage(x);
  else if (x.type === 'chart') chartDialog(x);
  else if (x.type === 'table') tableCellEdit(x, e);
  else if (x.text) startEditing(x.id);
}

function tableCellEdit(x, e) {
  const td = e.target.closest('td');
  if (!td) return;
  const r = +td.dataset.r, c = +td.dataset.c;
  td.contentEditable = 'true';
  td.focus();
  const sel = getSelection();
  const rg = document.createRange();
  rg.selectNodeContents(td);
  sel.removeAllRanges();
  sel.addRange(rg);
  const done = () => {
    td.removeEventListener('blur', done);
    x.table.rows[r][c] = td.innerHTML.replace(/<br>$/, '');
    commit();
    renderStage();
  };
  td.addEventListener('blur', done);
  td.addEventListener('keydown', (ev) => {
    ev.stopPropagation();
    if (ev.key === 'Tab') {
      ev.preventDefault();
      td.blur();
      const nc = c + 1 < x.table.rows[0].length ? c + 1 : 0;
      const nr = c + 1 < x.table.rows[0].length ? r : r + 1;
      if (nr < x.table.rows.length) setTimeout(() => {
        const next = holder.querySelector(`[data-id="${x.id}"] td[data-r="${nr}"][data-c="${nc}"]`);
        next?.dispatchEvent(new MouseEvent('dblclick', { bubbles: true }));
      }, 0);
    }
    if (ev.key === 'Escape') td.blur();
  });
}

// ---------------------------------------------------------------------------
// Operações em elementos

function mutate(fn, { stage: rerender = true } = {}) {
  const list = selected();
  if (!list.length) return;
  list.forEach(fn);
  commit();
  if (rerender) renderStage();
  renderProps();
}

function deleteSelection() {
  const s = slide();
  if (!S.sel.size) return;
  s.elements = s.elements.filter((e) => !S.sel.has(e.id));
  S.sel.clear();
  commit();
  renderStage();
  renderProps();
}

function duplicateSelection() {
  const copies = selected().map((e) => ({ ...JSON.parse(JSON.stringify(e)), id: uid(), x: e.x + 40, y: e.y + 40 }));
  if (!copies.length) return;
  slide().elements.push(...copies);
  S.sel = new Set(copies.map((c) => c.id));
  commit();
  renderStage();
  renderProps();
}

function arrange(dir) {
  const s = slide();
  const list = selected();
  if (!list.length) return;
  const rest = s.elements.filter((e) => !S.sel.has(e.id));
  if (dir === 'front') s.elements = [...rest, ...list];
  else if (dir === 'back') s.elements = [...list, ...rest];
  else {
    for (const e of dir === 'forward' ? [...list].reverse() : list) {
      const i = s.elements.indexOf(e);
      const j = dir === 'forward' ? Math.min(s.elements.length - 1, i + 1) : Math.max(0, i - 1);
      s.elements.splice(i, 1);
      s.elements.splice(j, 0, e);
    }
  }
  commit();
  renderStage();
}

function align(kind) {
  const list = selected();
  if (!list.length) return;
  const ref = list.length > 1 ? bounds(list) : { x: 0, y: 0, w: W, h: H };
  for (const e of list) {
    const b = bounds([e]);
    if (kind === 'left') e.x += ref.x - b.x;
    if (kind === 'hcenter') e.x += ref.x + ref.w / 2 - (b.x + b.w / 2);
    if (kind === 'right') e.x += ref.x + ref.w - (b.x + b.w);
    if (kind === 'top') e.y += ref.y - b.y;
    if (kind === 'vcenter') e.y += ref.y + ref.h / 2 - (b.y + b.h / 2);
    if (kind === 'bottom') e.y += ref.y + ref.h - (b.y + b.h);
    e.x = Math.round(e.x); e.y = Math.round(e.y);
  }
  if ((kind === 'hdist' || kind === 'vdist') && list.length > 2) {
    const key = kind === 'hdist' ? 'x' : 'y', size = kind === 'hdist' ? 'w' : 'h';
    const sorted = [...list].sort((a, b) => a[key] - b[key]);
    const total = sorted.reduce((s, e) => s + e[size], 0);
    const span = sorted[sorted.length - 1][key] + sorted[sorted.length - 1][size] - sorted[0][key];
    const gap = (span - total) / (sorted.length - 1);
    let pos = sorted[0][key];
    for (const e of sorted) { e[key] = Math.round(pos); pos += e[size] + gap; }
  }
  commit();
  renderStage();
  renderProps();
}

function nudge(dx, dy) {
  mutate((e) => { e.x += dx; e.y += dy; });
}

// Área de transferência de elementos
function copySelection(cut = false) {
  const list = selected();
  if (!list.length) return false;
  const assets = {};
  for (const e of list) if (e.type === 'image' && e.image.asset) assets[e.image.asset] = pres().assets[e.image.asset];
  S.clipboard = { elements: JSON.parse(JSON.stringify(list)), assets };
  if (cut) deleteSelection();
  return true;
}

function pasteElements() {
  if (!S.clipboard) return false;
  Object.assign(pres().assets, S.clipboard.assets);
  const same = S.clipboard.elements.every((c) => slide().elements.some((e) => e.x === c.x && e.y === c.y));
  const copies = S.clipboard.elements.map((e) => ({ ...JSON.parse(JSON.stringify(e)), id: uid(), x: e.x + (same ? 40 : 0), y: e.y + (same ? 40 : 0) }));
  slide().elements.push(...copies);
  S.clipboard.elements = copies.map((c) => JSON.parse(JSON.stringify(c)));
  S.sel = new Set(copies.map((c) => c.id));
  commit();
  renderStage();
  renderProps();
  return true;
}

// ---------------------------------------------------------------------------
// Painel de propriedades

const COLOR_TOKENS = [['accent', 'Destaque do tema'], ['title', 'Cor do título'], ['text', 'Cor do texto'], ['accent2', 'Destaque suave']];

function colorRow(current, onPick, { allowNone = false } = {}) {
  const t = themeOf(pres());
  const sw = el('div', { class: 'swatches' });
  if (allowNone) sw.append(el('button', { class: `none${!current ? ' on' : ''}`, title: 'Nenhuma', onclick: () => onPick(null) }));
  for (const [tok, label] of COLOR_TOKENS) sw.append(el('button', { class: current === tok ? 'on' : '', title: label, style: { background: resolveColor(tok, t) }, onclick: () => onPick(tok) }));
  for (const c of ['#ffffff', '#000000', '#e03131', '#f08c00', '#2f9e44', '#1971c2', '#7048e8']) sw.append(el('button', { class: current === c ? 'on' : '', title: c, style: { background: c }, onclick: () => onPick(c) }));
  const more = el('button', { title: 'Mais cores', style: { background: 'conic-gradient(red, yellow, lime, cyan, blue, magenta, red)' }, 'data-popover-anchor': '' });
  more.addEventListener('click', () => colorPopover(more, (c) => c && onPick(c)));
  sw.append(more);
  return sw;
}

function seg(options, value, onPick) {
  return el('div', { class: 'seg2' }, ...options.map(([v, ic, title]) => el('button', { class: value === v ? 'on' : '', title, html: ic.startsWith('<') ? ic : icon(ic, 16), onclick: () => onPick(v) })));
}

function numInput(value, onChange, { step = 1, min = -99999, max = 99999 } = {}) {
  const i = el('input', { type: 'number', value: Math.round(value * 100) / 100, step, min, max });
  i.addEventListener('change', () => onChange(clamp(parseFloat(i.value) || 0, min, max)));
  return i;
}

function renderProps() {
  const box = $('#props');
  const list = selected();
  const secs = [];
  const sec = (title, ...children) => el('div', { class: 'p-sec' }, el('div', { class: 'p-title' }, title), ...children);
  if (!list.length) {
    const s = slide();
    const t = themeOf(pres());
    secs.push(sec('Tema da apresentação', el('div', { class: 'theme-grid' }, ...Object.entries(THEMES).map(([id, th]) => el('button', {
      class: `theme-card${pres().theme === id ? ' on' : ''}`, onclick: () => { pres().theme = id; commit(); renderAll(); },
    }, el('div', { class: 'tp', style: { background: bgCss(th.bg, th) } }, el('b', { style: { color: th.title, fontFamily: th.head } }, 'Aa Título'), el('i', { style: { background: th.accent } })), el('span', {}, th.name))))));
    const bgType = s.bg?.type || 'theme';
    secs.push(sec('Fundo do slide',
      select([['theme', 'Do tema'], ['color', 'Cor sólida'], ['gradient', 'Degradê'], ['image', 'Imagem']], bgType, { onchange: async (ev) => {
        const v = ev.target.value;
        if (v === 'theme') s.bg = null;
        else if (v === 'color') s.bg = { type: 'color', color: t.bg.color || t.bg.from || '#ffffff' };
        else if (v === 'gradient') s.bg = { type: 'gradient', from: t.bg.from || t.accent, to: t.bg.to || t.title, angle: 135 };
        else {
          const src = await pickImageDataUrl();
          if (!src) { renderProps(); return; }
          const asset = uid('a');
          pres().assets[asset] = src;
          s.bg = { type: 'image', asset };
        }
        commit(); renderStage(); renderProps();
      } }),
      s.bg?.type === 'color' ? el('input', { type: 'color', value: s.bg.color, oninput: (ev) => { s.bg.color = ev.target.value; renderStage(); }, onchange: () => commit() }) : null,
      s.bg?.type === 'gradient' ? el('div', { class: 'p-row' },
        el('input', { type: 'color', value: s.bg.from, oninput: (ev) => { s.bg.from = ev.target.value; renderStage(); }, onchange: () => commit() }),
        el('input', { type: 'color', value: s.bg.to, oninput: (ev) => { s.bg.to = ev.target.value; renderStage(); }, onchange: () => commit() }),
        numInput(s.bg.angle ?? 135, (v) => { s.bg.angle = v; commit(); renderStage(); }, { min: 0, max: 360 })) : null,
      el('button', { class: 'btn small', onclick: () => { for (const x of pres().slides) x.bg = s.bg ? JSON.parse(JSON.stringify(s.bg)) : null; commit(); renderAll(); toast('Fundo aplicado a todos os slides.'); } }, 'Aplicar a todos')));
    secs.push(sec('Transição',
      select([['none', 'Nenhuma'], ['fade', 'Esmaecer'], ['push', 'Empurrar'], ['zoom', 'Zoom'], ['cover', 'Cobrir'], ['flip', 'Girar']], s.transition || 'none', { onchange: (ev) => { s.transition = ev.target.value; commit(); } }),
      el('button', { class: 'btn small', onclick: () => { for (const x of pres().slides) x.transition = s.transition; commit(); toast('Transição aplicada a todos os slides.'); } }, 'Aplicar a todos')));
    secs.push(sec('Layout', el('div', { class: 'p-row' },
      select(Object.entries(LAYOUTS), s.layout || 'blank', { onchange: (ev) => resetLayout(ev.target.value) }))));
    box.replaceChildren(...secs);
    return;
  }
  const e = list[0];
  const single = list.length === 1;
  if (single) {
    secs.push(sec('Posição e tamanho',
      el('div', { class: 'p-row' }, el('label', {}, 'X', numInput(e.x, (v) => mutate((x) => { x.x = v; }))), el('label', {}, 'Y', numInput(e.y, (v) => mutate((x) => { x.y = v; })))),
      el('div', { class: 'p-row' }, el('label', {}, 'L', numInput(e.w, (v) => mutate((x) => { x.w = Math.max(10, v); }))), el('label', {}, 'A', numInput(e.h, (v) => mutate((x) => { x.h = Math.max(4, v); })))),
      el('div', { class: 'p-row' }, el('label', {}, '↻', numInput(e.rot || 0, (v) => mutate((x) => { x.rot = v; }), { min: -360, max: 360 })),
        el('label', {}, 'Opac.', el('input', { type: 'range', min: 0, max: 100, value: Math.round((e.opacity ?? 1) * 100), oninput: (ev) => { e.opacity = ev.target.value / 100; holder.querySelector(`[data-id="${e.id}"]`).style.opacity = e.opacity; }, onchange: () => commit() })))));
  }
  secs.push(sec(single ? 'Organizar' : `${list.length} elementos selecionados`,
    seg([['left', 'hAlignLeft', 'Alinhar à esquerda'], ['hcenter', 'hAlignCenter', 'Centralizar na horizontal'], ['right', 'hAlignRight', 'Alinhar à direita'], ['top', 'alignTop', 'Alinhar em cima'], ['vcenter', 'alignMiddle', 'Centralizar na vertical'], ['bottom', 'alignBottom', 'Alinhar embaixo']], null, align),
    !single ? seg([['hdist', 'more', 'Distribuir na horizontal'], ['vdist', 'more', 'Distribuir na vertical']], null, align) : null,
    el('div', { class: 'p-row' },
      el('button', { class: 'btn small', html: `${icon('bringFront', 14)} Frente`, onclick: () => arrange('front') }),
      el('button', { class: 'btn small', html: `${icon('sendBack', 14)} Trás`, onclick: () => arrange('back') })),
    el('label', { class: 'check' }, el('input', { type: 'checkbox', checked: !!e.shadow, onchange: (ev) => mutate((x) => { x.shadow = ev.target.checked; }) }), 'Sombra suave')));
  if (single && (e.type === 'text' || e.type === 'shape') && e.text) {
    const t = e.text;
    const fontSel = select([['', 'Do tema (automática)'], ['head', 'Títulos do tema'], ['body', 'Corpo do tema'], ['Georgia, serif', 'Georgia'], ['"Times New Roman", serif', 'Times New Roman'], ['Arial, sans-serif', 'Arial'], ['Verdana, sans-serif', 'Verdana'], ['"Trebuchet MS", sans-serif', 'Trebuchet'], ['Impact, sans-serif', 'Impact'], ['"Courier New", monospace', 'Courier New'], ['"Comic Sans MS", cursive', 'Comic Sans']], t.font || '', { onchange: (ev) => mutate((x) => { x.text.font = ev.target.value || null; }) });
    secs.push(sec('Texto',
      fontSel,
      el('div', { class: 'p-row' }, el('label', {}, 'Tam.', numInput(t.size, (v) => mutate((x) => { x.text.size = clamp(v, 8, 400); }), { min: 8, max: 400 })),
        el('label', {}, 'Entrelinha', numInput(t.line || 1.25, (v) => mutate((x) => { x.text.line = v; }), { step: 0.05, min: 0.8, max: 3 }))),
      el('div', { class: 'p-row' },
        seg([['b', 'bold', 'Negrito'], ['i', 'italic', 'Itálico']], null, (k) => mutate((x) => { if (k === 'b') x.text.bold = !x.text.bold; else x.text.italic = !x.text.italic; })),
        seg([['left', 'alignLeft', 'Esquerda'], ['center', 'alignCenter', 'Centro'], ['right', 'alignRight', 'Direita']], t.align || 'left', (v) => mutate((x) => { x.text.align = v; }))),
      seg([['top', 'alignTop', 'Topo'], ['middle', 'alignMiddle', 'Meio'], ['bottom', 'alignBottom', 'Base']], t.valign || 'top', (v) => mutate((x) => { x.text.valign = v; })),
      el('div', { class: 'p-title' }, 'Cor do texto'),
      colorRow(t.color, (c) => mutate((x) => { x.text.color = c; }), { allowNone: true })));
  }
  if (single && e.type === 'shape') {
    const s = e.shape;
    secs.push(sec('Forma',
      select(Object.entries(SHAPES), s.kind, { onchange: (ev) => mutate((x) => { x.shape.kind = ev.target.value; }) }),
      el('div', { class: 'p-title' }, 'Preenchimento'),
      colorRow(s.fill, (c) => mutate((x) => { x.shape.fill = c; }), { allowNone: true }),
      el('div', { class: 'p-title' }, 'Contorno'),
      colorRow(s.stroke, (c) => mutate((x) => { x.shape.stroke = c; if (c && !x.shape.strokeW) x.shape.strokeW = 4; }), { allowNone: true }),
      el('div', { class: 'p-row' }, el('label', {}, 'Espessura', numInput(s.strokeW || 0, (v) => mutate((x) => { x.shape.strokeW = v; }), { min: 0, max: 60 }))),
      s.kind === 'round' ? el('label', {}, 'Cantos', el('input', { type: 'range', min: 0, max: 300, value: s.radius ?? 24, oninput: (ev) => { s.radius = +ev.target.value; renderStage(); }, onchange: () => commit() })) : null));
  }
  if (single && e.type === 'image') {
    secs.push(sec('Imagem',
      el('button', { class: 'btn small', html: `${icon('image', 14)} Trocar imagem`, onclick: () => insertImage(e) }),
      select([['cover', 'Preencher (cortar)'], ['contain', 'Ajustar (inteira)']], e.image.fit || 'cover', { onchange: (ev) => mutate((x) => { x.image.fit = ev.target.value; }) }),
      el('label', {}, 'Cantos arredondados', el('input', { type: 'range', min: 0, max: 400, value: e.image.radius || 0, oninput: (ev) => { e.image.radius = +ev.target.value; renderStage(); }, onchange: () => commit() })),
      select([['', 'Sem filtro'], ['grayscale(1)', 'Preto e branco'], ['sepia(.8)', 'Sépia'], ['saturate(1.6)', 'Vívido'], ['brightness(.7)', 'Escurecer'], ['blur(6px)', 'Desfocar'], ['contrast(1.3)', 'Contraste']], e.image.filter || '', { onchange: (ev) => mutate((x) => { x.image.filter = ev.target.value || null; }) })));
  }
  if (single && e.type === 'table') {
    const t = e.table;
    secs.push(sec('Tabela',
      el('div', { class: 'p-row' },
        el('button', { class: 'btn small', onclick: () => mutate((x) => { x.table.rows.push(x.table.rows[0].map(() => '')); x.h += 90; }) }, '+ Linha'),
        el('button', { class: 'btn small', onclick: () => mutate((x) => { if (x.table.rows.length > 1) { x.table.rows.pop(); x.h = Math.max(90, x.h - 90); } }) }, '− Linha')),
      el('div', { class: 'p-row' },
        el('button', { class: 'btn small', onclick: () => mutate((x) => { x.table.rows.forEach((r, i) => r.push(i === 0 && x.table.header ? `Coluna ${r.length + 1}` : '')); }) }, '+ Coluna'),
        el('button', { class: 'btn small', onclick: () => mutate((x) => { if (x.table.rows[0].length > 1) x.table.rows.forEach((r) => r.pop()); }) }, '− Coluna')),
      el('label', { class: 'check' }, el('input', { type: 'checkbox', checked: t.header, onchange: (ev) => mutate((x) => { x.table.header = ev.target.checked; }) }), 'Linha de cabeçalho'),
      el('label', {}, 'Fonte', numInput(t.size || 30, (v) => mutate((x) => { x.table.size = v; }), { min: 10, max: 120 })),
      el('small', { class: 'muted' }, 'Clique duas vezes numa célula para editar; Tab vai para a próxima.')));
  }
  if (single && e.type === 'chart') {
    secs.push(sec('Gráfico',
      select([['column', 'Colunas'], ['bar', 'Barras'], ['line', 'Linhas'], ['area', 'Área'], ['pie', 'Pizza'], ['doughnut', 'Rosca']], e.chart.kind, { onchange: (ev) => mutate((x) => { x.chart.kind = ev.target.value; }) }),
      el('button', { class: 'btn small', html: `${icon('table', 14)} Editar dados…`, onclick: () => chartDialog(e) })));
  }
  if (single) {
    const anim = e.anim || { type: 'none' };
    secs.push(sec('Animação de entrada',
      select([['none', 'Nenhuma'], ['fade', 'Surgir'], ['up', 'Subir'], ['left', 'Deslizar da esquerda'], ['zoom', 'Zoom'], ['pop', 'Saltar'], ['blur', 'Desfoque']], anim.type, { onchange: (ev) => mutate((x) => { x.anim = { type: ev.target.value }; }, { stage: false }) }),
      el('small', { class: 'muted' }, 'Na apresentação, cada clique revela o próximo elemento animado, na ordem das camadas.')));
  }
  secs.push(el('div', { class: 'p-sec' }, el('div', { class: 'p-row' },
    el('button', { class: 'btn small', html: `${icon('copy', 14)} Duplicar`, onclick: duplicateSelection }),
    el('button', { class: 'btn small danger', html: `${icon('trash', 14)} Excluir`, onclick: deleteSelection }))));
  box.replaceChildren(...secs);
}

function resetLayout(layout) {
  const s = slide();
  const texts = s.elements.filter((e) => e.text?.html).map((e) => e.text.html);
  s.layout = layout;
  const fresh = layoutElements(layout);
  // Mantém textos já escritos, na ordem
  fresh.filter((e) => e.text).forEach((e, i) => { if (texts[i]) e.text.html = texts[i]; });
  const keep = s.elements.filter((e) => e.type === 'image' && e.image.asset || e.type === 'chart' || e.type === 'table');
  s.elements = [...fresh, ...keep];
  S.sel.clear();
  commit();
  renderStage();
  renderProps();
}

// ---------------------------------------------------------------------------
// Diálogo de dados do gráfico

async function chartDialog(e) {
  const ch = e.chart;
  const lines = [['', ...ch.series.map((s) => s.name)].join('\t'), ...ch.labels.map((l, i) => [l, ...ch.series.map((s) => s.values[i] ?? '')].join('\t'))].join('\n');
  const ta = el('textarea', { rows: 10, style: { width: '100%', fontFamily: 'var(--mono)' } });
  ta.value = lines;
  const title = el('input', { type: 'text', value: ch.title || '', placeholder: 'Título (opcional)' });
  const ok = await dialog({
    title: 'Dados do gráfico', width: 560,
    body: [field('Título', title), el('p', { class: 'muted' }, 'Primeira linha: nomes das séries. Primeira coluna: categorias. Separe por Tab, ";" ou "," — dá para colar direto do Lumina Planilhas ou do Excel.'), ta],
  });
  if (!ok) return;
  const rows = ta.value.trim().split(/\r?\n/).map((l) => l.split(/\t|;|,(?!\d)/).map((x) => x.trim()));
  if (rows.length < 2) return;
  const head = rows[0].slice(1);
  ch.title = title.value;
  ch.labels = rows.slice(1).map((r) => r[0]);
  ch.series = head.map((name, j) => ({ name: name || `Série ${j + 1}`, values: rows.slice(1).map((r) => parseFloat(String(r[j + 1] ?? '').replace(/\./g, '').replace(',', '.')) || 0) }));
  commit();
  renderStage();
}

// ---------------------------------------------------------------------------
// Modo apresentação

const P = { on: false, idx: 0, step: 0, steps: [], layer: null, start: 0, timer: 0, typed: '' };

async function present(from = 0) {
  finishEditing();
  await ensureChartImages(pres());
  const root = $('#present');
  root.classList.remove('hidden');
  root.replaceChildren(el('div', { class: 'p-black' }), el('div', { class: 'laser' }), el('div', { class: 'p-notes' }), el('div', { class: 'p-hud' }));
  P.on = true;
  P.start = Date.now();
  P.layer = null;
  try { if (native?.setFullscreen) await native.setFullscreen(true); else await document.documentElement.requestFullscreen?.(); } catch { /* sem tela cheia */ }
  showSlide(from, 1, false);
  clearInterval(P.timer);
  P.timer = setInterval(updateNotes, 1000);
  root.focus();
}

function presentScale() {
  return Math.min(innerWidth / W, innerHeight / H);
}

function showSlide(i, dir = 1, animate = true, revealAll = false) {
  const p = pres();
  const root = $('#present');
  if (i >= p.slides.length) { endScreen(); return; }
  if (i < 0) return;
  root.querySelector('.p-end')?.remove();
  P.idx = i;
  const s = p.slides[i];
  const node = renderSlide(s, p, { mode: 'present', chartImages: S.chartImages });
  const sc = presentScale();
  const layer = el('div', { class: 'p-layer', style: { left: `${(innerWidth - W * sc) / 2}px`, top: `${(innerHeight - H * sc) / 2}px`, width: `${W}px`, height: `${H}px`, transform: `scale(${sc})` } }, node);
  P.steps = [...node.querySelectorAll('[data-anim]')];
  P.step = revealAll ? P.steps.length : 0;
  for (const n of P.steps) {
    n.style.setProperty('--o', n.style.opacity || 1);
    if (!revealAll) n.classList.add('pre-anim');
  }
  const old = P.layer;
  root.prepend(layer);
  P.layer = layer;
  const tr = animate ? s.transition || 'none' : 'none';
  if (tr !== 'none') {
    const inCls = `tr-${tr}-in`;
    layer.classList.add(inCls);
    if (dir < 0) layer.classList.add('back');
    if (tr === 'push' && old) { old.classList.add('tr-push-out'); if (dir < 0) old.classList.add('back'); }
    root.append(layer);
    setTimeout(() => { old?.remove(); layer.classList.remove(inCls, 'back'); }, 700);
  } else {
    old?.remove();
  }
  updateNotes();
  root.querySelector('.p-hud').textContent = `${i + 1} / ${p.slides.length}`;
}

function endScreen() {
  const root = $('#present');
  P.idx = pres().slides.length;
  P.layer?.remove();
  P.layer = null;
  root.append(el('div', { class: 'p-end' }, 'Fim da apresentação. Clique ou pressione Esc para sair.'));
}

function nextStep() {
  if (P.idx >= pres().slides.length) { exitPresent(); return; }
  if (P.step < P.steps.length) {
    const n = P.steps[P.step++];
    n.classList.remove('pre-anim');
    n.classList.add(`anim-${n.dataset.anim}`);
    return;
  }
  showSlide(P.idx + 1, 1);
}

function prevStep() {
  if (P.idx >= pres().slides.length) { showSlide(pres().slides.length - 1, -1, false, true); return; }
  if (P.step > 0) {
    const n = P.steps[--P.step];
    n.classList.remove(`anim-${n.dataset.anim}`);
    n.classList.add('pre-anim');
    return;
  }
  if (P.idx > 0) showSlide(P.idx - 1, -1, true, true);
}

function updateNotes() {
  if (!P.on) return;
  const box = $('#present .p-notes');
  if (!box?.classList.contains('on')) return;
  const p = pres();
  const s = p.slides[P.idx];
  const secs = Math.floor((Date.now() - P.start) / 1000);
  const next = p.slides[P.idx + 1];
  const nextBox = el('div', { class: 'n-next' });
  if (next) {
    const n = renderSlide(next, p, { mode: 'export', chartImages: S.chartImages });
    n.style.transform = `scale(${240 / W})`;
    n.style.transformOrigin = '0 0';
    nextBox.append(n);
  }
  box.replaceChildren(
    el('div', { class: 'n-text' }, s?.notes || '(Sem anotações neste slide)'),
    el('div', { class: 'n-side' },
      el('div', { class: 'n-timer' }, `${String(Math.floor(secs / 60)).padStart(2, '0')}:${String(secs % 60).padStart(2, '0')}`),
      el('div', {}, `Slide ${Math.min(P.idx + 1, p.slides.length)} de ${p.slides.length}`),
      el('div', { style: { color: '#999', fontSize: '12px' } }, 'Próximo:'), nextBox));
}

async function exitPresent() {
  if (!P.on) return;
  P.on = false;
  clearInterval(P.timer);
  $('#present').classList.add('hidden');
  $('#present').replaceChildren();
  try { if (native?.setFullscreen) await native.setFullscreen(false); else if (document.fullscreenElement) await document.exitFullscreen(); } catch { /* ignorar */ }
  if (P.idx < pres().slides.length) goTo(P.idx);
}

function setupPresentEvents() {
  const root = $('#present');
  let hideTimer = 0;
  root.addEventListener('mousemove', (e) => {
    root.classList.add('cursor');
    clearTimeout(hideTimer);
    hideTimer = setTimeout(() => root.classList.remove('cursor'), 2000);
    const laser = root.querySelector('.laser');
    if (laser) { laser.style.left = `${e.clientX}px`; laser.style.top = `${e.clientY}px`; }
  });
  root.addEventListener('click', (e) => { if (!e.target.closest('.p-notes')) nextStep(); });
  root.addEventListener('contextmenu', (e) => { e.preventDefault(); prevStep(); });
  root.addEventListener('wheel', debounce((e) => (e.deltaY > 0 ? nextStep() : prevStep()), 60), { passive: true });
  document.addEventListener('keydown', (e) => {
    if (!P.on) return;
    e.preventDefault();
    e.stopPropagation();
    const k = e.key;
    if (/^\d$/.test(k)) { P.typed += k; return; }
    if (k === 'Enter' && P.typed) { showSlide(clamp(+P.typed - 1, 0, pres().slides.length - 1), 1, false); P.typed = ''; return; }
    P.typed = '';
    if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter', 'n'].includes(k)) nextStep();
    else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p'].includes(k)) prevStep();
    else if (k === 'Home') showSlide(0, -1, false);
    else if (k === 'End') showSlide(pres().slides.length - 1, 1, false, true);
    else if (k === 'Escape') exitPresent();
    else if (k === 'b' || k === 'B' || k === '.') { const b = root.querySelector('.p-black'); b.classList.remove('white'); b.classList.toggle('on'); }
    else if (k === 'w' || k === 'W' || k === ',') { const b = root.querySelector('.p-black'); b.classList.add('white'); b.classList.toggle('on'); }
    else if (k === 'l' || k === 'L') root.classList.toggle('laser-on');
    else if (k === 'N' || k === 'm' || k === 'M' || k === 's' || k === 'S') { root.querySelector('.p-notes').classList.toggle('on'); updateNotes(); }
  }, true);
  window.addEventListener('resize', () => {
    if (!P.on || !P.layer) return;
    const sc = presentScale();
    Object.assign(P.layer.style, { left: `${(innerWidth - W * sc) / 2}px`, top: `${(innerHeight - H * sc) / 2}px`, transform: `scale(${sc})` });
  });
}

// ---------------------------------------------------------------------------
// Cabeçalho, menus e barra de ferramentas

function exportItems() {
  return [
    { label: 'PowerPoint (.pptx)', icon: 'slides', action: () => exportAs('pptx') },
    { label: 'PDF (.pdf)', icon: 'pdf', key: 'Ctrl+P', action: () => exportAs('pdf') },
    { label: 'Imagem do slide (.png)', icon: 'image', action: () => exportAs('png') },
  ];
}

function setupHeader() {
  $('#badge').innerHTML = icon('slides', 20);
  const t = $('#doc-title');
  t.addEventListener('input', () => { pres().title = t.value || 'Apresentação sem título'; deck().dirty = true; $('#save-state').textContent = 'Alterações não salvas'; renderDeckTabs(); });
  t.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); t.blur(); } });
  t.addEventListener('blur', () => { commit(); renderTitle(); });
  const exp = el('button', { class: 'btn', html: `${icon('download', 16)} Exportar`, 'data-popover-anchor': '' });
  exp.addEventListener('click', () => dropdown(exp, exportItems()));
  $('#header-actions').append(
    exp,
    el('button', { class: 'btn', html: `${icon('save', 16)} Salvar`, onclick: () => save() }),
    el('button', { class: 'btn primary', html: `${icon('play', 16)} Apresentar`, title: 'Apresentar do início (F5)', onclick: () => present(0) }));
}

function setupMenus() {
  menubar($('#menus'), [
    ['Arquivo', () => [
      { label: 'Nova apresentação', icon: 'plus', key: 'Ctrl+N', action: () => newDeck() },
      { label: 'Nova a partir de tema', icon: 'palette', sub: [['pitch', 'Pitch de startup'], ...Object.entries(THEMES).map(([k, v]) => [k, v.name])].map(([k, l]) => ({ label: l, action: () => newDeck(k) })) },
      { label: 'Abrir…', icon: 'open', key: 'Ctrl+O', action: openDialog },
      '-',
      { label: 'Salvar', icon: 'save', key: 'Ctrl+S', action: () => save() },
      { label: 'Salvar como…', key: 'Ctrl+Shift+S', action: () => save(deck(), true) },
      { label: 'Exportar', icon: 'download', sub: exportItems() },
      '-',
      { label: 'Fechar', action: () => closeDeck() },
    ]],
    ['Editar', () => [
      { label: 'Desfazer', icon: 'undo', key: 'Ctrl+Z', action: undo },
      { label: 'Refazer', icon: 'redo', key: 'Ctrl+Y', action: redo },
      '-',
      { label: 'Recortar', key: 'Ctrl+X', action: () => copySelection(true) },
      { label: 'Copiar', key: 'Ctrl+C', action: () => copySelection() },
      { label: 'Colar', key: 'Ctrl+V', action: pasteElements },
      { label: 'Duplicar', key: 'Ctrl+D', action: duplicateSelection },
      { label: 'Excluir', key: 'Delete', action: deleteSelection },
      '-',
      { label: 'Selecionar tudo', key: 'Ctrl+A', action: () => { S.sel = new Set(slide().elements.map((e) => e.id)); renderOverlay(); renderProps(); } },
    ]],
    ['Inserir', () => [
      { label: 'Novo slide', icon: 'plus', key: 'Ctrl+M', action: () => addSlide() },
      { label: 'Novo slide com layout', icon: 'layout', sub: Object.entries(LAYOUTS).map(([k, l]) => ({ label: l, action: () => addSlide(k) })) },
      { label: 'Duplicar slide', icon: 'copy', key: 'Ctrl+Shift+D', action: duplicateSlide },
      '-',
      { label: 'Caixa de texto', icon: 'text', key: 'T', action: insertText },
      { label: 'Imagem…', icon: 'image', action: () => insertImage() },
      { label: 'Forma', icon: 'shapes', sub: Object.entries(SHAPES).map(([k, l]) => ({ label: l, action: () => insertShape(k) })) },
      { label: 'Tabela', icon: 'table', action: () => insertTable() },
      { label: 'Gráfico', icon: 'chart', sub: [['column', 'Colunas'], ['bar', 'Barras'], ['line', 'Linhas'], ['area', 'Área'], ['pie', 'Pizza'], ['doughnut', 'Rosca']].map(([k, l]) => ({ label: l, action: () => insertChart(k) })) },
    ]],
    ['Organizar', () => [
      { label: 'Trazer para a frente', icon: 'bringFront', key: 'Ctrl+Shift+]', action: () => arrange('front') },
      { label: 'Avançar', key: 'Ctrl+]', action: () => arrange('forward') },
      { label: 'Recuar', key: 'Ctrl+[', action: () => arrange('backward') },
      { label: 'Enviar para trás', icon: 'sendBack', key: 'Ctrl+Shift+[', action: () => arrange('back') },
      '-',
      { label: 'Alinhar', sub: [['left', 'À esquerda'], ['hcenter', 'Ao centro'], ['right', 'À direita'], ['top', 'Em cima'], ['vcenter', 'Ao meio'], ['bottom', 'Embaixo'], ['hdist', 'Distribuir na horizontal'], ['vdist', 'Distribuir na vertical']].map(([k, l]) => ({ label: l, action: () => align(k) })) },
    ]],
    ['Apresentação', () => [
      { label: 'Apresentar do início', icon: 'play', key: 'F5', action: () => present(0) },
      { label: 'Apresentar deste slide', key: 'Shift+F5', action: () => present(deck().current) },
      '-',
      { label: 'Anotações do apresentador', icon: 'notes', action: () => $('#notes-wrap').classList.toggle('hidden') },
    ]],
    ['Ajuda', () => [{ label: 'Atalhos e recursos', icon: 'info', key: 'F1', action: helpDialog }]],
  ]);
}

function setupToolbar() {
  const bar = $('#toolbar');
  const newBtn = toolButton('plus', 'Novo slide (Ctrl+M)', () => addSlide(), { label: 'Novo slide' });
  const layoutBtn = toolButton('chevronDown', 'Escolher layout', () => layoutPicker(layoutBtn));
  layoutBtn.setAttribute('data-popover-anchor', '');
  layoutBtn.style.minWidth = '20px';
  layoutBtn.style.padding = '0 2px';
  const shapesBtn = toolButton('shapes', 'Formas', () => shapePicker(shapesBtn), { label: 'Formas' });
  shapesBtn.setAttribute('data-popover-anchor', '');
  const chartBtn = toolButton('chart', 'Gráfico', () => dropdown(chartBtn, [['column', 'Colunas'], ['bar', 'Barras'], ['line', 'Linhas'], ['area', 'Área'], ['pie', 'Pizza'], ['doughnut', 'Rosca']].map(([k, l]) => ({ label: l, action: () => insertChart(k) }))), { label: 'Gráfico' });
  chartBtn.setAttribute('data-popover-anchor', '');
  const themeBtn = toolButton('palette', 'Temas', () => { S.sel.clear(); finishEditing(); renderOverlay(); renderProps(); $('#props').scrollTop = 0; }, { label: 'Temas' });
  const fmt = (cmd, ic, title) => toolButton(ic, title, () => { if (S.editing) { document.execCommand(cmd); const e = elById(S.editing); const inner = holder.querySelector(`[data-id="${S.editing}"] .sl-text`); if (e && inner) e.text.html = inner.innerHTML; } else mutateTextFlag(cmd); });
  const colorBtn = toolButton('textColor', 'Cor do texto selecionado', () => colorPopover(colorBtn, (c) => {
    if (!c) return;
    if (S.editing) { document.execCommand('styleWithCSS', false, true); document.execCommand('foreColor', false, c); const e = elById(S.editing); e.text.html = holder.querySelector(`[data-id="${S.editing}"] .sl-text`).innerHTML; }
    else mutate((x) => { if (x.text) x.text.color = c; });
  }));
  colorBtn.setAttribute('data-popover-anchor', '');
  bar.append(
    newBtn, layoutBtn,
    toolButton('undo', 'Desfazer', undo, { key: 'Ctrl+Z' }), toolButton('redo', 'Refazer', redo, { key: 'Ctrl+Y' }),
    sep(),
    toolButton('text', 'Caixa de texto', insertText, { label: 'Texto' }),
    toolButton('image', 'Imagem', () => insertImage(), { label: 'Imagem' }),
    shapesBtn,
    toolButton('table', 'Tabela', () => insertTable(), { label: 'Tabela' }),
    chartBtn,
    sep(),
    fmt('bold', 'bold', 'Negrito'), fmt('italic', 'italic', 'Itálico'), fmt('underline', 'underline', 'Sublinhado'), colorBtn,
    fmt('insertUnorderedList', 'list', 'Lista com marcadores'), fmt('insertOrderedList', 'listOrdered', 'Lista numerada'),
    sep(),
    themeBtn,
    el('span', { class: 'spacer' }),
    toolButton('play', 'Apresentar (F5)', () => present(0), { label: 'Apresentar' }),
  );
}

function mutateTextFlag(cmd) {
  if (cmd === 'bold') mutate((x) => { if (x.text) x.text.bold = !x.text.bold; });
  else if (cmd === 'italic') mutate((x) => { if (x.text) x.text.italic = !x.text.italic; });
  else if (cmd === 'insertUnorderedList' || cmd === 'insertOrderedList') {
    const tag = cmd === 'insertOrderedList' ? 'ol' : 'ul';
    mutate((x) => {
      if (!x.text) return;
      const d = document.createElement('div');
      d.innerHTML = x.text.html || '';
      const items = [...d.querySelectorAll('p, li')].map((n) => n.innerHTML).filter(Boolean);
      x.text.html = d.querySelector(tag) ? items.map((i) => `<p>${i}</p>`).join('') : `<${tag}>${(items.length ? items : ['']).map((i) => `<li>${i}</li>`).join('')}</${tag}>`;
    });
  }
}

function layoutPicker(anchor) {
  const t = themeOf(pres());
  const grid = el('div', { class: 'layout-grid' }, ...Object.entries(LAYOUTS).map(([k, l]) => {
    const preview = renderSlide({ id: 'p', elements: layoutElements(k).map((e) => (e.text ? { ...e, text: { ...e.text, html: e.text.html || `<p>${e.ph || ''}</p>` } } : e)), bg: null }, { ...pres(), assets: {} }, { mode: 'export' });
    preview.style.transform = `scale(${150 / W})`;
    preview.style.transformOrigin = '0 0';
    return el('button', { class: 'layout-card', onclick: () => { closePopovers(); addSlide(k); } }, el('div', { class: 'lp', style: { background: bgCss(null, t) } }, preview), l);
  }));
  popover(grid, anchor);
}

function shapePicker(anchor) {
  const svgFor = (k) => {
    const prev = renderSlide({ id: 'x', elements: [{ id: 's', type: 'shape', x: 0, y: 0, w: 100, h: k === 'line' ? 4 : k === 'arrow' || k === 'chevron' ? 60 : 100, rot: 0, opacity: 1, shape: { kind: k, fill: k === 'line' ? null : '#5b6476', stroke: k === 'line' ? '#5b6476' : null, strokeW: k === 'line' ? 3 : 0, radius: 22 }, text: { html: '' } }] }, { theme: 'claro', assets: {} }, { mode: 'export' });
    prev.style.cssText = 'width:100px;height:100px;position:relative;transform:scale(.26);transform-origin:0 0;background:transparent';
    const wrap = el('div', { style: { width: '26px', height: '26px', overflow: 'hidden', display: 'flex', alignItems: 'center' } }, prev);
    return wrap;
  };
  popover(el('div', { class: 'shape-grid' }, ...Object.entries(SHAPES).map(([k, l]) => el('button', { title: l, onclick: () => { closePopovers(); insertShape(k); } }, svgFor(k)))), anchor);
}

function helpDialog() {
  const rows = [
    ['Clique', 'Selecionar (Shift: vários; arraste no vazio para selecionar por área)'], ['Duplo clique', 'Editar texto, trocar imagem, editar dados do gráfico ou célula da tabela'],
    ['Arrastar', 'Mover com guias magnéticas (Alt: sem encaixe; Alt+arrastar: duplicar)'], ['Alças', 'Redimensionar (Shift: proporcional) e girar (Shift: de 15 em 15°)'],
    ['Setas', 'Mover 1 px (Shift: 10 px). Sem seleção: trocar de slide'], ['Ctrl+D', 'Duplicar elemento'], ['Ctrl+M', 'Novo slide'], ['Ctrl+Shift+D', 'Duplicar slide'],
    ['Ctrl+] / Ctrl+[', 'Avançar / recuar camada'], ['Ctrl+C / V', 'Copiar e colar elementos (cole imagens direto do sistema)'],
    ['F5 / Shift+F5', 'Apresentar do início / deste slide'],
    ['Na apresentação', '→ ou clique: avançar · ←: voltar · B/W: tela preta/branca · L: laser · S: anotações e cronômetro · número+Enter: ir ao slide · Esc: sair'],
  ];
  dialog({ title: 'Atalhos e recursos', cancelText: null, wide: true, body: el('div', { style: { display: 'grid', gridTemplateColumns: 'auto 1fr', gap: '6px 18px' } }, ...rows.flatMap(([k, v]) => [el('kbd', {}, k), el('span', {}, v)])) });
}

// ---------------------------------------------------------------------------
// Teclado, colar e arrastar arquivos

function setupKeyboard() {
  document.addEventListener('keydown', (e) => {
    if (P.on || isDialogOpen()) return;
    const t = e.target;
    const inField = t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement || t.isContentEditable;
    const combo = comboOf(e);
    const global = {
      'Ctrl+S': () => save(), 'Ctrl+Shift+S': () => save(deck(), true), 'Ctrl+O': openDialog, 'Ctrl+N': () => newDeck(), 'Ctrl+P': () => exportAs('pdf'),
      F5: () => present(0), 'Shift+F5': () => present(deck().current), 'Ctrl+M': () => addSlide(), F1: helpDialog,
    };
    if (global[combo]) { e.preventDefault(); global[combo](); return; }
    if (S.editing) {
      if (e.key === 'Escape') { e.preventDefault(); const id = S.editing; finishEditing(); S.sel = new Set([id]); renderOverlay(); }
      return;
    }
    if (inField) return;
    const map = {
      'Ctrl+Z': undo, 'Ctrl+Y': redo, 'Ctrl+Shift+Z': redo,
      'Ctrl+C': () => copySelection(), 'Ctrl+X': () => copySelection(true),
      'Ctrl+D': duplicateSelection, 'Ctrl+Shift+D': duplicateSlide,
      'Ctrl+A': () => { S.sel = new Set(slide().elements.map((x) => x.id)); renderOverlay(); renderProps(); },
      'Ctrl+]': () => arrange('forward'), 'Ctrl+[': () => arrange('backward'), 'Ctrl+Shift+]': () => arrange('front'), 'Ctrl+Shift+[': () => arrange('back'),
      Delete: () => (S.sel.size ? deleteSelection() : S.focus === 'thumbs' && deleteSlide()), Backspace: () => (S.sel.size ? deleteSelection() : S.focus === 'thumbs' && deleteSlide()),
      Escape: () => { S.sel.clear(); renderOverlay(); renderProps(); },
      Enter: () => { const l = selected(); if (l.length === 1 && l[0].text) startEditing(l[0].id, true); },
      F2: () => { const l = selected(); if (l.length === 1 && l[0].text) startEditing(l[0].id); },
      T: () => insertText(),
    };
    if (map[combo]) { e.preventDefault(); map[combo](); return; }
    const arrows = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] }[e.key];
    if (arrows && !e.ctrlKey && !e.altKey) {
      e.preventDefault();
      if (S.sel.size) nudge(arrows[0] * (e.shiftKey ? 10 : 1), arrows[1] * (e.shiftKey ? 10 : 1));
      else goTo(deck().current + (arrows[0] || arrows[1]));
    }
  });
  document.addEventListener('paste', async (e) => {
    if (P.on || isDialogOpen() || S.editing) return;
    const t = e.target;
    if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t.isContentEditable) return;
    const item = [...(e.clipboardData?.items || [])].find((i) => i.type.startsWith('image/'));
    if (item) {
      e.preventDefault();
      const reader = new FileReader();
      reader.onload = () => addImageFromDataUrl(reader.result);
      reader.readAsDataURL(item.getAsFile());
      return;
    }
    const text = e.clipboardData?.getData('text/plain');
    if (pasteElements()) { e.preventDefault(); return; }
    if (text) {
      e.preventDefault();
      addElement({ id: uid(), type: 'text', role: 'body', x: 360, y: 400, w: 1200, h: 280, rot: 0, opacity: 1, text: { html: text.split(/\r?\n/).map((l) => `<p>${l.replace(/[&<>]/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;' })[c]) || '<br>'}</p>`).join(''), size: 40, color: null, align: 'left', valign: 'top', bold: false, font: null, line: 1.25 } });
    }
  });
  stage.addEventListener('dragover', (e) => { if ([...e.dataTransfer.types].includes('Files')) e.preventDefault(); });
  stage.addEventListener('drop', async (e) => {
    const files = [...e.dataTransfer.files].filter((f) => f.type.startsWith('image/'));
    if (!files.length) return;
    e.preventDefault();
    const at = clientToSlide(e.clientX, e.clientY);
    for (const f of files) {
      const url = await new Promise((r) => { const fr = new FileReader(); fr.onload = () => r(fr.result); fr.readAsDataURL(f); });
      await addImageFromDataUrl(url, at);
    }
  });
}

// ---------------------------------------------------------------------------
// Inicialização

registerApp({
  openFiles,
  newDocument: (tpl) => newDeck(tpl),
  unsaved: () => S.decks.filter((d) => d.dirty).map((d) => d.pres.title),
  beforeClose: async () => {
    for (const d of [...S.decks]) if (d.dirty && !(await askSave(d))) return false;
    return true;
  },
  activated: () => fitScale(),
});

initTheme();
setupHeader();
setupMenus();
setupToolbar();
setupKeyboard();
setupPresentEvents();
holder.addEventListener('mousedown', onStageDown);
overlay.addEventListener('mousedown', onStageDown);
stage.addEventListener('mousedown', (e) => { if (e.target === stage) onStageDown(e); });
holder.addEventListener('dblclick', onStageDblClick);
holder.addEventListener('contextmenu', (e) => {
  e.preventDefault();
  const node = e.target.closest('.sl-el');
  if (node && !S.sel.has(node.dataset.id)) { S.sel = new Set([node.dataset.id]); renderOverlay(); renderProps(); }
  contextMenu(S.sel.size ? [
    { label: 'Recortar', key: 'Ctrl+X', action: () => copySelection(true) }, { label: 'Copiar', key: 'Ctrl+C', action: () => copySelection() },
    { label: 'Colar', key: 'Ctrl+V', action: pasteElements, disabled: !S.clipboard }, { label: 'Duplicar', key: 'Ctrl+D', action: duplicateSelection },
    '-', { label: 'Trazer para a frente', icon: 'bringFront', action: () => arrange('front') }, { label: 'Enviar para trás', icon: 'sendBack', action: () => arrange('back') },
    '-', { label: 'Excluir', icon: 'trash', key: 'Delete', action: deleteSelection },
  ] : [
    { label: 'Colar', key: 'Ctrl+V', action: pasteElements, disabled: !S.clipboard }, { label: 'Novo slide', icon: 'plus', action: () => addSlide() },
    { label: 'Caixa de texto', icon: 'text', action: insertText }, { label: 'Imagem…', icon: 'image', action: () => insertImage() },
  ], e.clientX, e.clientY);
});
$('#thumbs').addEventListener('focus', () => { S.focus = 'thumbs'; });
$('#thumbs').addEventListener('keydown', (e) => {
  if (e.key === 'ArrowDown' || e.key === 'ArrowUp') { e.preventDefault(); e.stopPropagation(); goTo(deck().current + (e.key === 'ArrowDown' ? 1 : -1)); }
});
const notes = $('#notes');
notes.addEventListener('input', debounce(() => { slide().notes = notes.value; commit(); }, 500));
new ResizeObserver(() => { fitScale(); renderOverlay(); }).observe(stage);
S.decks.push(makeDeck(newPresentation()));
activate(S.decks[0]);
standaloneBoot({ openFiles });
window.__slides = { S, present, exportAs, newDeck, addSlide, insertShape, insertChart, insertTable, io, pres, slide };
void PALETTE;
void prompt;
