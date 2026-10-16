// Documento, camadas e histórico (copy-on-write: canvases nunca são alterados depois de entrarem no histórico).
import { createCanvas, cloneCanvas } from './utils.js';
import { buildAntsPath, translateMask } from './selection.js';
import { app } from './state.js';

const MAX_HISTORY = 60;
const HISTORY_BUDGET = 1.5e9; // bytes aproximados de pixels mantidos pelo histórico
let nextLayerId = 1;
let nextDocId = 1;

export class Layer {
  constructor(name, w, h, canvas = null) {
    this.id = nextLayerId++;
    this.name = name;
    this.canvas = canvas || createCanvas(w, h);
    this.visible = true;
    this.opacity = 1;
    this.blend = 'source-over';
    this.locked = false;
    this.text = null; // metadados de camada de texto editável
    this.preview = null; // canvas temporário usado durante edições ao vivo
    this.version = 0;
  }
}

export class Doc {
  constructor(width, height, name = 'Sem título') {
    this.id = nextDocId++;
    this.width = width;
    this.height = height;
    this.name = name;
    this.layers = [];
    this.active = 0;
    this.selection = null;
    this.history = [];
    this.hIndex = -1;
    this.savedIndex = 0;
    this.filePath = null;
    this.view = { zoom: 1, panX: 0, panY: 0, fitted: false };
  }

  get activeLayer() {
    return this.layers[this.active];
  }

  get dirty() {
    return this.hIndex !== this.savedIndex;
  }

  markSaved() {
    this.savedIndex = this.hIndex;
    app.emit('doc-meta');
  }

  addLayer(name, canvas = null, index = this.active + 1) {
    const l = new Layer(name || `Camada ${this.layers.length}`, this.width, this.height, canvas);
    index = Math.min(Math.max(0, index), this.layers.length);
    this.layers.splice(index, 0, l);
    this.active = index;
    return l;
  }

  // Executa fn em uma cópia do canvas da camada e substitui (mantém o histórico intacto).
  editLayer(layer, fn, keepText = false) {
    const c = cloneCanvas(layer.canvas);
    const ctx = c.getContext('2d');
    fn(ctx, c);
    layer.canvas = c;
    layer.version++;
    if (!keepText) layer.text = null;
  }

  replaceLayerCanvas(layer, canvas, keepText = false) {
    layer.canvas = canvas;
    layer.version++;
    if (!keepText) layer.text = null;
  }

  setSelection(mask) {
    this.selection = mask || null;
    if (mask) buildAntsPath(mask);
    app.emit('selection');
  }

  snapshot(label) {
    return {
      label,
      width: this.width,
      height: this.height,
      active: this.active,
      selection: this.selection,
      layers: this.layers.map((l) => ({
        id: l.id, name: l.name, canvas: l.canvas, visible: l.visible, opacity: l.opacity,
        blend: l.blend, locked: l.locked, text: l.text ? { ...l.text } : null, version: l.version,
      })),
    };
  }

  commit(label) {
    this.history.splice(this.hIndex + 1);
    this.history.push(this.snapshot(label));
    // Imagens grandes guardam menos passos para não esgotar a memória.
    const limit = Math.max(10, Math.min(MAX_HISTORY, Math.floor(HISTORY_BUDGET / (this.width * this.height * 4))));
    if (this.history.length > limit) {
      const drop = this.history.length - limit;
      this.history.splice(0, drop);
      this.savedIndex -= drop;
    }
    this.hIndex = this.history.length - 1;
    app.emit('doc-changed', this);
  }

  restore(index) {
    const s = this.history[index];
    if (!s) return;
    this.hIndex = index;
    this.width = s.width;
    this.height = s.height;
    this.layers = s.layers.map((d) => {
      const l = new Layer(d.name, s.width, s.height, d.canvas);
      Object.assign(l, d, { text: d.text ? { ...d.text } : null, preview: null });
      return l;
    });
    this.active = Math.min(s.active, this.layers.length - 1);
    this.selection = s.selection;
    if (s.selection) buildAntsPath(s.selection);
    app.emit('doc-changed', this);
    app.emit('selection');
  }

  undo() {
    if (this.hIndex > 0) this.restore(this.hIndex - 1);
  }

  redo() {
    if (this.hIndex < this.history.length - 1) this.restore(this.hIndex + 1);
  }

  composite(ctx, { hidePreview = false, rect = null } = {}) {
    const r = rect || { x: 0, y: 0, w: this.width, h: this.height };
    ctx.save();
    ctx.setTransform(1, 0, 0, 1, 0, 0);
    ctx.clearRect(r.x, r.y, r.w, r.h);
    for (const l of this.layers) {
      if (!l.visible || l.opacity <= 0) continue;
      ctx.globalAlpha = l.opacity;
      ctx.globalCompositeOperation = l.blend;
      const src = !hidePreview && l.preview ? l.preview : l.canvas;
      if (rect) ctx.drawImage(src, r.x, r.y, r.w, r.h, r.x, r.y, r.w, r.h);
      else ctx.drawImage(src, 0, 0);
    }
    ctx.restore();
  }

  flattenToCanvas() {
    const c = createCanvas(this.width, this.height);
    this.composite(c.getContext('2d'), { hidePreview: true });
    return c;
  }

  // ---- Operações de imagem -----------------------------------------------

  transformAll(newW, newH, draw) {
    for (const l of this.layers) {
      const c = createCanvas(newW, newH);
      const ctx = c.getContext('2d');
      ctx.imageSmoothingQuality = 'high';
      draw(ctx, l.canvas);
      l.canvas = c;
      l.version++;
      l.text = null;
    }
    if (this.selection) {
      const m = createCanvas(newW, newH);
      draw(m.getContext('2d'), this.selection);
      this.selection = m;
      buildAntsPath(m);
    }
    this.width = newW;
    this.height = newH;
  }

  resizeImage(w, h) {
    this.transformAll(w, h, (ctx, src) => {
      ctx.imageSmoothingEnabled = true;
      // Redução em etapas para melhor qualidade.
      let cur = src;
      while (cur.width / 2 > w && cur.height / 2 > h) {
        const half = createCanvas(Math.ceil(cur.width / 2), Math.ceil(cur.height / 2));
        const hctx = half.getContext('2d');
        hctx.imageSmoothingQuality = 'high';
        hctx.drawImage(cur, 0, 0, half.width, half.height);
        cur = half;
      }
      ctx.drawImage(cur, 0, 0, w, h);
    });
    this.commit('Tamanho da imagem');
  }

  resizeCanvas(w, h, anchorX = 0.5, anchorY = 0.5) {
    const ox = Math.round((w - this.width) * anchorX);
    const oy = Math.round((h - this.height) * anchorY);
    this.transformAll(w, h, (ctx, src) => ctx.drawImage(src, ox, oy));
    this.commit('Tamanho da tela');
  }

  crop(x, y, w, h) {
    x = Math.round(x); y = Math.round(y); w = Math.round(w); h = Math.round(h);
    if (w < 1 || h < 1) return;
    this.transformAll(w, h, (ctx, src) => ctx.drawImage(src, -x, -y));
    this.selection = null;
    this.commit('Cortar');
    app.emit('selection');
  }

  rotate(deg) {
    const swap = deg % 180 !== 0;
    const w = swap ? this.height : this.width;
    const h = swap ? this.width : this.height;
    const ow = this.width, oh = this.height;
    this.transformAll(w, h, (ctx, src) => {
      ctx.translate(w / 2, h / 2);
      ctx.rotate((deg * Math.PI) / 180);
      ctx.drawImage(src, -ow / 2, -oh / 2);
    });
    this.commit(`Girar ${deg}°`);
  }

  flip(horizontal) {
    const { width: w, height: h } = this;
    this.transformAll(w, h, (ctx, src) => {
      if (horizontal) { ctx.translate(w, 0); ctx.scale(-1, 1); }
      else { ctx.translate(0, h); ctx.scale(1, -1); }
      ctx.drawImage(src, 0, 0);
    });
    this.commit(horizontal ? 'Espelhar horizontal' : 'Espelhar vertical');
  }

  // ---- Operações de camada ----------------------------------------------

  duplicateLayer() {
    const src = this.activeLayer;
    const l = this.addLayer(`${src.name} cópia`, cloneCanvas(src.canvas));
    Object.assign(l, { opacity: src.opacity, blend: src.blend, visible: src.visible, text: src.text ? { ...src.text } : null });
    this.commit('Duplicar camada');
  }

  deleteLayer() {
    if (this.layers.length <= 1) return false;
    this.layers.splice(this.active, 1);
    this.active = Math.max(0, this.active - 1);
    this.commit('Excluir camada');
    return true;
  }

  moveLayer(from, to) {
    if (to < 0 || to >= this.layers.length || from === to) return;
    const [l] = this.layers.splice(from, 1);
    this.layers.splice(to, 0, l);
    this.active = to;
    this.commit('Reordenar camadas');
  }

  mergeDown() {
    if (this.active === 0) return false;
    const top = this.layers[this.active];
    const below = this.layers[this.active - 1];
    this.editLayer(below, (ctx) => {
      ctx.globalAlpha = top.visible ? top.opacity : 0;
      ctx.globalCompositeOperation = top.blend;
      ctx.drawImage(top.canvas, 0, 0);
    });
    this.layers.splice(this.active, 1);
    this.active--;
    this.commit('Mesclar para baixo');
    return true;
  }

  mergeVisible() {
    const visible = this.layers.filter((l) => l.visible);
    if (visible.length < 2) return;
    const c = createCanvas(this.width, this.height);
    this.composite(c.getContext('2d'), { hidePreview: true });
    const idx = this.layers.indexOf(visible[0]);
    const merged = new Layer('Mesclada', this.width, this.height, c);
    this.layers = this.layers.filter((l) => !l.visible);
    this.layers.splice(Math.min(idx, this.layers.length), 0, merged);
    this.active = this.layers.indexOf(merged);
    this.commit('Mesclar visíveis');
  }

  flatten() {
    const c = createCanvas(this.width, this.height);
    const ctx = c.getContext('2d');
    ctx.fillStyle = '#ffffff';
    ctx.fillRect(0, 0, this.width, this.height);
    ctx.drawImage(this.flattenToCanvas(), 0, 0);
    this.layers = [new Layer('Fundo', this.width, this.height, c)];
    this.active = 0;
    this.commit('Achatar imagem');
  }

  moveSelection(dx, dy) {
    if (this.selection) this.setSelection(translateMask(this.selection, dx, dy));
  }
}

export function createDocument(w, h, { name = 'Sem título', background = '#ffffff' } = {}) {
  const doc = new Doc(w, h, name);
  const bgLayer = doc.addLayer(background === 'transparent' ? 'Camada 1' : 'Fundo', null, 0);
  if (background !== 'transparent') {
    const ctx = bgLayer.canvas.getContext('2d');
    ctx.fillStyle = background;
    ctx.fillRect(0, 0, w, h);
  }
  doc.commit('Novo documento');
  doc.savedIndex = doc.hIndex;
  return doc;
}

export function documentFromImage(img, name) {
  const w = img.naturalWidth || img.width;
  const h = img.naturalHeight || img.height;
  const doc = new Doc(w, h, name);
  const c = createCanvas(w, h);
  c.getContext('2d').drawImage(img, 0, 0);
  doc.addLayer('Fundo', c, 0);
  doc.commit('Abrir');
  doc.savedIndex = doc.hIndex;
  return doc;
}
