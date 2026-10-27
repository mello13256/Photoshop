// Componentes de interface compartilhados pelos apps da suíte.
import { icon } from './icons.js';

export const IS_MAC = /Mac/.test(navigator.platform);

export function el(tag, attrs = {}, ...children) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v == null || v === false) continue;
    if (k === 'class') e.className = v;
    else if (k === 'style' && typeof v === 'object') Object.assign(e.style, v);
    else if (k.startsWith('on') && typeof v === 'function') e.addEventListener(k.slice(2), v);
    else if (k === 'html') e.innerHTML = v;
    else if (k === 'value') e.value = v;
    else if (k === 'checked' || k === 'selected' || k === 'disabled') e[k] = !!v;
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const c of children.flat(Infinity)) {
    if (c == null || c === false) continue;
    e.append(c instanceof Node ? c : document.createTextNode(String(c)));
  }
  return e;
}

export const $ = (s, root = document) => root.querySelector(s);
export const $$ = (s, root = document) => [...root.querySelectorAll(s)];
export const clamp = (v, a, b) => (v < a ? a : v > b ? b : v);

export function debounce(fn, ms) {
  let t = 0;
  const d = (...a) => { clearTimeout(t); t = setTimeout(() => fn(...a), ms); };
  d.flush = (...a) => { clearTimeout(t); fn(...a); };
  d.cancel = () => clearTimeout(t);
  return d;
}

export function uid(prefix = '') {
  return prefix + Math.random().toString(36).slice(2, 10) + Date.now().toString(36).slice(-4);
}

// ---------------------------------------------------------------------------
// Notificações

export function toast(msg, ms = 2400) {
  let host = $('#toasts');
  if (!host) { host = el('div', { id: 'toasts' }); document.body.append(host); }
  const t = el('div', { class: 'toast' }, msg);
  host.append(t);
  setTimeout(() => t.classList.add('out'), ms);
  setTimeout(() => t.remove(), ms + 400);
}

// ---------------------------------------------------------------------------
// Diálogos

let dialogCount = 0;
export const isDialogOpen = () => dialogCount > 0;

export function dialog({ title, body, okText = 'OK', cancelText = 'Cancelar', width, onOpen, extraButtons = [], wide = false }) {
  return new Promise((resolve) => {
    dialogCount++;
    const overlay = el('div', { class: 'dialog-overlay' });
    const close = (v) => {
      dialogCount--;
      overlay.remove();
      document.removeEventListener('keydown', onKey, true);
      resolve(v);
    };
    const onKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(false); }
      else if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement) && !e.target.isContentEditable && !(e.target instanceof HTMLButtonElement)) {
        e.preventDefault(); e.stopPropagation(); close(true);
      } else e.stopPropagation();
    };
    const box = el('div', { class: `dialog${wide ? ' wide' : ''}`, style: width ? { width: `${width}px` } : null, role: 'dialog', 'aria-label': title },
      el('div', { class: 'dialog-title' }, title, el('button', { class: 'icon-btn dialog-x', title: 'Fechar', html: icon('x', 16), onclick: () => close(false) })),
      el('div', { class: 'dialog-body' }, body),
      el('div', { class: 'dialog-buttons' },
        ...extraButtons.map((b) => el('button', { class: 'btn', onclick: () => b.onClick(close) }, b.label)),
        el('span', { class: 'spacer' }),
        cancelText ? el('button', { class: 'btn', onclick: () => close(false) }, cancelText) : null,
        okText ? el('button', { class: 'btn primary', onclick: () => close(true) }, okText) : null));
    overlay.append(box);
    overlay.addEventListener('mousedown', (e) => { if (e.target === overlay) close(false); });
    document.body.append(overlay);
    document.addEventListener('keydown', onKey, true);
    const first = box.querySelector('.dialog-body input, .dialog-body textarea, .dialog-body select');
    if (first) setTimeout(() => { first.focus(); first.select?.(); }, 0);
    onOpen?.(box, close);
  });
}

export async function prompt(title, label, value = '', { placeholder = '', okText = 'OK' } = {}) {
  const inp = el('input', { type: 'text', value, placeholder });
  const ok = await dialog({ title, body: field(label, inp), okText, width: 400 });
  return ok ? inp.value : null;
}

export async function confirm(title, msg, okText = 'OK') {
  return dialog({ title, body: el('p', {}, msg), okText, width: 420 });
}

export function field(label, input, hint) {
  return el('label', { class: 'field' }, el('span', { class: 'field-label' }, label), input, hint ? el('small', { class: 'muted' }, hint) : null);
}

export function select(options, value, attrs = {}) {
  return el('select', attrs, ...options.map((o) => {
    const [v, l] = Array.isArray(o) ? o : [o, o];
    return el('option', { value: v, selected: String(v) === String(value) }, l);
  }));
}

// ---------------------------------------------------------------------------
// Menus suspensos e de contexto

let openPopover = null;
export function closePopovers() {
  if (openPopover) { openPopover.remove(); openPopover = null; }
  $$('.menubar .menu.open').forEach((m) => m.classList.remove('open'));
}
document.addEventListener('mousedown', (e) => {
  if (openPopover && !openPopover.contains(e.target) && !e.target.closest('[data-popover-anchor]')) closePopovers();
  if (!e.target.closest('.menubar')) $$('.menubar .menu.open').forEach((m) => m.classList.remove('open'));
});
document.addEventListener('keydown', (e) => { if (e.key === 'Escape' && openPopover) closePopovers(); });
window.addEventListener('blur', () => closePopovers());

// Abre um popover posicionado junto a um elemento ou ponto.
export function popover(content, anchor, { align = 'start', className = '' } = {}) {
  closePopovers();
  const pop = el('div', { class: `popover ${className}` }, content);
  document.body.append(pop);
  const r = anchor instanceof Element ? anchor.getBoundingClientRect() : { left: anchor.x, right: anchor.x, top: anchor.y, bottom: anchor.y };
  const pw = pop.offsetWidth, ph = pop.offsetHeight;
  let x = align === 'end' ? r.right - pw : r.left;
  let y = r.bottom + 4;
  if (x + pw > innerWidth - 8) x = innerWidth - pw - 8;
  if (y + ph > innerHeight - 8) y = Math.max(8, r.top - ph - 4);
  pop.style.left = `${Math.max(8, x)}px`;
  pop.style.top = `${y}px`;
  openPopover = pop;
  return pop;
}

// items: [{ label, icon, key, action, disabled, checked, sub: [...] } | '-']
export function menuList(items, onPick) {
  const list = el('div', { class: 'menu-list' });
  for (const it of items) {
    if (!it) continue;
    if (it === '-') { list.append(el('div', { class: 'menu-sep' })); continue; }
    if (it.header) { list.append(el('div', { class: 'menu-header' }, it.header)); continue; }
    const row = el('div', {
      class: `menu-item${it.disabled ? ' disabled' : ''}${it.sub ? ' sub' : ''}${it.checked ? ' checked' : ''}`,
      onclick: (e) => {
        if (it.disabled || it.sub) return;
        e.stopPropagation();
        closePopovers();
        onPick?.();
        it.action?.();
      },
    },
    el('span', { class: 'mi-icon', html: it.checked ? icon('check', 15) : it.icon ? icon(it.icon, 15) : '' }),
    el('span', { class: 'mi-label' }, it.label),
    it.key ? el('span', { class: 'kbd' }, displayKey(it.key)) : null,
    it.sub ? el('span', { class: 'mi-arrow', html: icon('chevronRight', 14) }) : null);
    if (it.sub) row.append(menuList(typeof it.sub === 'function' ? it.sub() : it.sub, onPick));
    list.append(row);
  }
  return list;
}

export function contextMenu(items, x, y) {
  return popover(menuList(items), { x, y }, { className: 'context' });
}

export function dropdown(anchor, items) {
  return popover(menuList(items), anchor, { className: 'context' });
}

// Barra de menus: [[rótulo, itens | () => itens], ...]
export function menubar(root, menus) {
  root.classList.add('menubar');
  let active = false;
  for (const [label, items] of menus) {
    const m = el('div', { class: 'menu' });
    const btn = el('button', { class: 'menu-btn' }, label);
    const open = () => {
      $$('.menubar .menu.open').forEach((x) => { x.classList.remove('open'); x.querySelector('.menu-list')?.remove(); });
      m.append(menuList(typeof items === 'function' ? items() : items, () => { m.classList.remove('open'); active = false; }));
      m.classList.add('open');
      active = true;
    };
    btn.addEventListener('mousedown', (e) => {
      e.preventDefault();
      if (m.classList.contains('open')) { m.classList.remove('open'); active = false; } else open();
    });
    btn.addEventListener('mouseenter', () => { if (active && !m.classList.contains('open')) open(); });
    m.append(btn);
    root.append(m);
  }
}

// ---------------------------------------------------------------------------
// Atalhos de teclado

export const displayKey = (k) => (IS_MAC ? k.replace(/Ctrl\+/g, '⌘').replace(/Alt\+/g, '⌥').replace(/Shift\+/g, '⇧') : k);

const CODE_KEYS = { Equal: '=', Minus: '-', NumpadAdd: '+', NumpadSubtract: '-', BracketLeft: '[', BracketRight: ']', Quote: "'", Backquote: "'", Period: '.', Comma: ',', Slash: '/', Backslash: '\\', Semicolon: ';' };

function makeCombo(ctrl, alt, shift, key) {
  return [ctrl && 'Ctrl', alt && 'Alt', shift && 'Shift', key].filter(Boolean).join('+');
}

export function normalizeCombo(str) {
  const plus = str.endsWith('++');
  const parts = (plus ? str.slice(0, -2) : str).split('+').filter(Boolean);
  const key = plus ? '+' : parts.pop();
  return makeCombo(parts.includes('Ctrl'), parts.includes('Alt'), parts.includes('Shift'), key.length === 1 ? key.toUpperCase() : key);
}

export function comboOf(e) {
  const code = e.code || '';
  let key;
  if (code.startsWith('Key')) key = code.slice(3);
  else if (code.startsWith('Digit')) key = code.slice(5);
  else key = CODE_KEYS[code];
  if (!key) key = e.key.length === 1 ? e.key.toUpperCase() : e.key;
  return makeCombo(e.ctrlKey || e.metaKey, e.altKey, e.shiftKey, key);
}

// Registra atalhos { 'Ctrl+S': fn }. Ignora quando um diálogo está aberto.
export function shortcuts(map, { allowInInputs = [] } = {}) {
  const table = new Map(Object.entries(map).map(([k, v]) => [normalizeCombo(k), v]));
  const allow = new Set(allowInInputs.map(normalizeCombo));
  document.addEventListener('keydown', (e) => {
    if (isDialogOpen()) return;
    const combo = comboOf(e);
    const fn = table.get(combo);
    if (!fn) return;
    const t = e.target;
    const inField = t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement || t instanceof HTMLSelectElement;
    if (inField && !allow.has(combo)) return;
    if (fn(e) !== false) e.preventDefault();
  });
}

// ---------------------------------------------------------------------------
// Seletor de cor em popover

export const PALETTE = [
  '#000000', '#434343', '#666666', '#999999', '#b7b7b7', '#cccccc', '#d9d9d9', '#efefef', '#f3f3f3', '#ffffff',
  '#980000', '#ff0000', '#ff9900', '#ffff00', '#00ff00', '#00ffff', '#4a86e8', '#0000ff', '#9900ff', '#ff00ff',
  '#e6b8af', '#f4cccc', '#fce5cd', '#fff2cc', '#d9ead3', '#d0e0e3', '#c9daf8', '#cfe2f3', '#d9d2e9', '#ead1dc',
  '#dd7e6b', '#ea9999', '#f9cb9c', '#ffe599', '#b6d7a8', '#a2c4c9', '#a4c2f4', '#9fc5e8', '#b4a7d6', '#d5a6bd',
  '#cc4125', '#e06666', '#f6b26b', '#ffd966', '#93c47d', '#76a5af', '#6d9eeb', '#6fa8dc', '#8e7cc3', '#c27ba0',
  '#a61c00', '#cc0000', '#e69138', '#f1c232', '#6aa84f', '#45818e', '#3c78d8', '#3d85c6', '#674ea7', '#a64d79',
  '#5b0f00', '#660000', '#783f04', '#7f6000', '#274e13', '#0c343d', '#1c4587', '#073763', '#20124d', '#4c1130',
];

export function colorPopover(anchor, onPick, { allowNone = false, noneLabel = 'Nenhuma' } = {}) {
  const custom = el('input', { type: 'color', value: '#3366ff', title: 'Cor personalizada' });
  const content = el('div', { class: 'color-pop' },
    allowNone ? el('button', { class: 'btn small block', onclick: () => { closePopovers(); onPick(null); } }, noneLabel) : null,
    el('div', { class: 'swatch-grid' }, ...PALETTE.map((c) => el('button', {
      class: 'swatch', style: { background: c }, title: c, onclick: () => { closePopovers(); onPick(c); },
    }))),
    el('div', { class: 'row' }, el('span', { class: 'muted' }, 'Personalizada'), custom,
      el('button', { class: 'btn small', onclick: () => { closePopovers(); onPick(custom.value); } }, 'Aplicar')));
  return popover(content, anchor);
}

// Botão de ferramenta com ícone.
export function toolButton(iconName, title, onClick, { key, cls = '', label } = {}) {
  return el('button', {
    class: `tb-btn ${cls}`,
    title: key ? `${title} (${displayKey(key)})` : title,
    'aria-label': title,
    html: icon(iconName, 18) + (label ? `<span class="tb-label">${label}</span>` : ''),
    onmousedown: (e) => e.preventDefault(), // não rouba o foco do editor
    onclick: onClick,
  });
}

export function sep() {
  return el('span', { class: 'tb-sep' });
}

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
}
