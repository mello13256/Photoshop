// Estado global da aplicação e barramento de eventos simples.

export const app = {
  docs: [],
  doc: null,
  tool: null,
  tools: {},
  prevToolId: null,
  fg: '#000000',
  bg: '#ffffff',
  editingBg: false,
  view: null,
  clipboard: null,
  showGrid: false,
  swatches: [
    '#000000', '#ffffff', '#7f7f7f', '#c3c3c3', '#ff0000', '#ff7f00', '#ffff00', '#7fff00',
    '#00ff00', '#00ff7f', '#00ffff', '#007fff', '#0000ff', '#7f00ff', '#ff00ff', '#ff007f',
    '#880015', '#b97a57', '#ffaec9', '#ffc90e', '#efe4b0', '#b5e61d', '#99d9ea', '#7092be',
  ],
  _listeners: {},
  on(ev, fn) {
    (this._listeners[ev] ||= []).push(fn);
  },
  emit(ev, ...args) {
    for (const fn of this._listeners[ev] || []) fn(...args);
  },
  setFg(c) {
    this.fg = c;
    this.emit('colors');
  },
  setBg(c) {
    this.bg = c;
    this.emit('colors');
  },
};

export const BLEND_MODES = [
  ['source-over', 'Normal'],
  ['multiply', 'Multiplicação'],
  ['screen', 'Divisão (Tela)'],
  ['overlay', 'Sobrepor'],
  ['darken', 'Escurecer'],
  ['lighten', 'Clarear'],
  ['color-dodge', 'Subexposição de cores'],
  ['color-burn', 'Superexposição de cores'],
  ['hard-light', 'Luz direta'],
  ['soft-light', 'Luz suave'],
  ['difference', 'Diferença'],
  ['exclusion', 'Exclusão'],
  ['hue', 'Matiz'],
  ['saturation', 'Saturação'],
  ['color', 'Cor'],
  ['luminosity', 'Luminosidade'],
];
