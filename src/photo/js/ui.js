// Componentes de interface: diálogos, notificações, seletor de cores.
import { el, hexToRgb, rgbToHex, rgbToHsv, hsvToRgb, clamp } from './utils.js';
import { app } from './state.js';

export function toast(msg, ms = 2200) {
  const t = el('div', { class: 'toast' }, msg);
  document.getElementById('toasts').append(t);
  setTimeout(() => t.classList.add('out'), ms);
  setTimeout(() => t.remove(), ms + 400);
}

let dialogOpen = 0;
export const isDialogOpen = () => dialogOpen > 0;

// Abre um diálogo modal. Retorna uma Promise com true (OK) ou false (Cancelar).
export function dialog({ title, body, okText = 'OK', cancelText = 'Cancelar', width, onOpen, extraButtons = [] }) {
  return new Promise((resolve) => {
    dialogOpen++;
    const overlay = el('div', { class: 'dialog-overlay' });
    const close = (v) => {
      dialogOpen--;
      overlay.remove();
      document.removeEventListener('keydown', onKey, true);
      resolve(v);
    };
    const onKey = (e) => {
      if (e.key === 'Escape') { e.preventDefault(); e.stopPropagation(); close(false); }
      else if (e.key === 'Enter' && !(e.target instanceof HTMLTextAreaElement)) {
        e.preventDefault(); e.stopPropagation(); close(true);
      } else e.stopPropagation();
    };
    const box = el('div', { class: 'dialog', style: width ? { width: `${width}px` } : null },
      el('div', { class: 'dialog-title' }, title),
      el('div', { class: 'dialog-body' }, body),
      el('div', { class: 'dialog-buttons' },
        ...extraButtons.map((b) => el('button', { class: 'btn', onclick: () => b.onClick(close) }, b.label)),
        el('span', { class: 'spacer' }),
        cancelText ? el('button', { class: 'btn', onclick: () => close(false) }, cancelText) : null,
        el('button', { class: 'btn primary', onclick: () => close(true) }, okText)
      )
    );
    overlay.append(box);
    document.body.append(overlay);
    document.addEventListener('keydown', onKey, true);
    const first = box.querySelector('input, textarea, select');
    if (first) setTimeout(() => { first.focus(); first.select?.(); }, 0);
    onOpen?.(box, close);
  });
}

export async function confirmDialog(title, msg, okText = 'OK') {
  return dialog({ title, body: el('p', {}, msg), okText });
}

export function field(label, input) {
  return el('label', { class: 'field' }, el('span', {}, label), input);
}

export function numberInput(value, attrs = {}) {
  return el('input', { type: 'number', value, ...attrs });
}

// Slider + caixa numérica sincronizados.
export function sliderField(p, value, onInput) {
  const step = p.step ?? 1;
  const range = el('input', { type: 'range', min: p.min, max: p.max, step, value });
  const num = el('input', { type: 'number', min: p.min, max: p.max, step, value, class: 'num' });
  range.addEventListener('input', () => { num.value = range.value; onInput(+range.value); });
  num.addEventListener('input', () => {
    const v = clamp(+num.value, p.min, p.max);
    range.value = v;
    onInput(v);
  });
  return el('div', { class: 'slider-field' }, el('div', { class: 'slider-label' }, el('span', {}, p.label), num), range);
}

// ---------------------------------------------------------------------------
// Seletor de cores HSV do painel "Cor".
export class ColorPicker {
  constructor(root) {
    this.root = root;
    this.hsv = { h: 0, s: 0, v: 0 };
    this.sv = el('canvas', { class: 'cp-sv', width: 220, height: 140 });
    this.hue = el('canvas', { class: 'cp-hue', width: 220, height: 12 });
    this.hex = el('input', { class: 'cp-hex', maxlength: 7, spellcheck: 'false' });
    this.rgbInputs = ['R', 'G', 'B'].map(() => el('input', { type: 'number', min: 0, max: 255, class: 'cp-num' }));
    root.append(
      this.sv, this.hue,
      el('div', { class: 'cp-row' },
        el('span', { class: 'cp-label' }, '#'), this.hex,
        ...this.rgbInputs.flatMap((inp, i) => [el('span', { class: 'cp-label' }, 'RGB'[i]), inp]))
    );
    this.bindDrag(this.sv, (x, y) => {
      this.hsv.s = clamp(x / this.sv.clientWidth, 0, 1);
      this.hsv.v = 1 - clamp(y / this.sv.clientHeight, 0, 1);
      this.commit();
    });
    this.bindDrag(this.hue, (x) => {
      this.hsv.h = clamp(x / this.hue.clientWidth, 0, 0.9999) * 360;
      this.commit();
    });
    this.hex.addEventListener('change', () => {
      const v = this.hex.value.trim();
      if (/^#?[0-9a-f]{3}([0-9a-f]{3})?$/i.test(v)) this.setColor(v.startsWith('#') ? v : `#${v}`, true);
    });
    this.rgbInputs.forEach((inp) => inp.addEventListener('change', () => {
      const [r, g, b] = this.rgbInputs.map((i) => clamp(+i.value || 0, 0, 255));
      this.setColor(rgbToHex(r, g, b), true);
    }));
    app.on('colors', () => this.sync());
    this.drawHue();
    this.sync();
  }

  bindDrag(canvas, fn) {
    canvas.addEventListener('pointerdown', (e) => {
      canvas.setPointerCapture(e.pointerId);
      const r = canvas.getBoundingClientRect();
      fn(e.clientX - r.left, e.clientY - r.top);
      const move = (ev) => fn(ev.clientX - r.left, ev.clientY - r.top);
      const up = () => { canvas.removeEventListener('pointermove', move); canvas.removeEventListener('pointerup', up); };
      canvas.addEventListener('pointermove', move);
      canvas.addEventListener('pointerup', up);
    });
  }

  current() {
    return app.editingBg ? app.bg : app.fg;
  }

  sync() {
    if (this._self) return;
    const { r, g, b } = hexToRgb(this.current());
    const hsv = rgbToHsv(r, g, b);
    if (hsv.s === 0 || hsv.v === 0) hsv.h = this.hsv.h; // preserva matiz em cinzas
    this.hsv = hsv;
    this.updateFields();
  }

  setColor(hex, emit) {
    const { r, g, b } = hexToRgb(hex);
    this.hsv = rgbToHsv(r, g, b);
    if (emit) {
      if (app.editingBg) app.setBg(hex.toLowerCase());
      else app.setFg(hex.toLowerCase());
    }
    this.updateFields();
  }

  commit() {
    const { r, g, b } = hsvToRgb(this.hsv.h, this.hsv.s, this.hsv.v);
    const hex = rgbToHex(r, g, b);
    this._self = true;
    if (app.editingBg) app.setBg(hex);
    else app.setFg(hex);
    this._self = false;
    this.updateFields();
  }

  updateFields() {
    const { r, g, b } = hsvToRgb(this.hsv.h, this.hsv.s, this.hsv.v);
    const hex = rgbToHex(r, g, b);
    if (document.activeElement !== this.hex) this.hex.value = hex;
    [r, g, b].forEach((v, i) => { if (document.activeElement !== this.rgbInputs[i]) this.rgbInputs[i].value = Math.round(v); });
    this.drawSV();
  }

  drawHue() {
    const ctx = this.hue.getContext('2d');
    const g = ctx.createLinearGradient(0, 0, this.hue.width, 0);
    for (let i = 0; i <= 6; i++) g.addColorStop(i / 6, `hsl(${i * 60},100%,50%)`);
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, this.hue.width, this.hue.height);
    this.hueBase = ctx.getImageData(0, 0, this.hue.width, this.hue.height);
  }

  drawSV() {
    const { sv } = this;
    const ctx = sv.getContext('2d');
    const w = sv.width, h = sv.height;
    ctx.fillStyle = `hsl(${this.hsv.h},100%,50%)`;
    ctx.fillRect(0, 0, w, h);
    let g = ctx.createLinearGradient(0, 0, w, 0);
    g.addColorStop(0, '#fff');
    g.addColorStop(1, 'rgba(255,255,255,0)');
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, w, h);
    g = ctx.createLinearGradient(0, 0, 0, h);
    g.addColorStop(0, 'rgba(0,0,0,0)');
    g.addColorStop(1, '#000');
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, w, h);
    const x = this.hsv.s * w, y = (1 - this.hsv.v) * h;
    ctx.lineWidth = 2;
    ctx.strokeStyle = this.hsv.v > 0.5 ? '#000' : '#fff';
    ctx.beginPath();
    ctx.arc(x, y, 6, 0, Math.PI * 2);
    ctx.stroke();
    // marcador de matiz
    const hctx = this.hue.getContext('2d');
    hctx.putImageData(this.hueBase, 0, 0);
    const hx = (this.hsv.h / 360) * this.hue.width;
    hctx.fillStyle = '#fff';
    hctx.strokeStyle = '#000';
    hctx.lineWidth = 1;
    hctx.fillRect(hx - 2, 0, 4, this.hue.height);
    hctx.strokeRect(hx - 2.5, 0.5, 5, this.hue.height - 1);
  }
}
