// Área de trabalho: zoom, panorâmica, renderização e sobreposições.
import { app } from './state.js';
import { createCanvas, clamp } from './utils.js';

export function unionRect(a, b) {
  if (!a) return { ...b };
  if (!b) return { ...a };
  const x = Math.min(a.x, b.x), y = Math.min(a.y, b.y);
  return { x, y, w: Math.max(a.x + a.w, b.x + b.w) - x, h: Math.max(a.y + a.h, b.y + b.h) - y };
}

// Converte para inteiros dentro do documento. null = documento inteiro; false = vazio.
function clipRect(r, doc) {
  if (!r) return null;
  const x0 = Math.max(0, Math.floor(r.x)), y0 = Math.max(0, Math.floor(r.y));
  const x1 = Math.min(doc.width, Math.ceil(r.x + r.w)), y1 = Math.min(doc.height, Math.ceil(r.y + r.h));
  if (x1 <= x0 || y1 <= y0) return false;
  return { x: x0, y: y0, w: x1 - x0, h: y1 - y0 };
}

export class Viewport {
  constructor(canvas) {
    this.canvas = canvas;
    this.ctx = canvas.getContext('2d');
    this.comp = createCanvas(1, 1);
    this.compDirty = true;
    this.needsRender = true;
    this.antsOffset = 0;
    this.pointer = null; // posição do ponteiro em coordenadas do documento
    this.pointerScreen = null;
    this.checker = this.makeChecker();
    new ResizeObserver(() => this.resize()).observe(canvas.parentElement);
    this.resize();
    const loop = () => {
      if (this.needsRender) this.render();
      requestAnimationFrame(loop);
    };
    requestAnimationFrame(loop);
    setInterval(() => {
      if (app.doc?.selection || app.tool?.animatesAnts) {
        this.antsOffset = (this.antsOffset + 1) % 16;
        this.needsRender = true;
      }
    }, 90);
  }

  makeChecker() {
    const c = createCanvas(16, 16);
    const x = c.getContext('2d');
    x.fillStyle = '#ffffff';
    x.fillRect(0, 0, 16, 16);
    x.fillStyle = '#d6d6d6';
    x.fillRect(0, 0, 8, 8);
    x.fillRect(8, 8, 8, 8);
    return this.ctx.createPattern(c, 'repeat');
  }

  resize() {
    const r = this.canvas.parentElement.getBoundingClientRect();
    const dpr = window.devicePixelRatio || 1;
    this.width = r.width;
    this.height = r.height;
    this.canvas.width = Math.max(1, Math.round(r.width * dpr));
    this.canvas.height = Math.max(1, Math.round(r.height * dpr));
    this.canvas.style.width = `${r.width}px`;
    this.canvas.style.height = `${r.height}px`;
    if (app.doc && !app.doc.view.fitted) this.fit();
    this.requestRender();
  }

  // rect (opcional): recompõe apenas essa região do documento.
  requestRender(recomposite = false, rect = null) {
    if (recomposite) {
      if (rect && (!this.compDirty || this.compRect)) this.compRect = unionRect(this.compRect, rect);
      else this.compRect = null;
      this.compDirty = true;
    }
    this.needsRender = true;
  }

  // ---- Transformações de visualização ------------------------------------

  toDoc(sx, sy) {
    const v = app.doc.view;
    return { x: (sx - v.panX) / v.zoom, y: (sy - v.panY) / v.zoom };
  }

  toScreen(x, y) {
    const v = app.doc.view;
    return { x: x * v.zoom + v.panX, y: y * v.zoom + v.panY };
  }

  fit() {
    const doc = app.doc;
    if (!doc) return;
    const pad = 40;
    const z = Math.min((this.width - pad) / doc.width, (this.height - pad) / doc.height, 1);
    this.setZoom(z > 0 ? z : 1);
    this.center();
    doc.view.fitted = true;
  }

  center() {
    const doc = app.doc;
    doc.view.panX = Math.round((this.width - doc.width * doc.view.zoom) / 2);
    doc.view.panY = Math.round((this.height - doc.height * doc.view.zoom) / 2);
    this.requestRender();
  }

  setZoom(z, sx = this.width / 2, sy = this.height / 2) {
    const v = app.doc.view;
    z = clamp(z, 0.02, 64);
    const p = this.toDoc(sx, sy);
    v.zoom = z;
    v.panX = sx - p.x * z;
    v.panY = sy - p.y * z;
    v.fitted = true;
    app.emit('zoom');
    this.requestRender();
  }

  zoomStep(dir, sx, sy) {
    const steps = [0.02, 0.05, 0.0833, 0.125, 0.1667, 0.25, 0.333, 0.5, 0.667, 1, 1.5, 2, 3, 4, 5, 6, 8, 12, 16, 24, 32, 48, 64];
    const z = app.doc.view.zoom;
    let next;
    if (dir > 0) next = steps.find((s) => s > z * 1.001) ?? 64;
    else next = [...steps].reverse().find((s) => s < z * 0.999) ?? 0.02;
    this.setZoom(next, sx, sy);
  }

  pan(dx, dy) {
    app.doc.view.panX += dx;
    app.doc.view.panY += dy;
    this.requestRender();
  }

  // ---- Renderização -------------------------------------------------------

  render() {
    this.needsRender = false;
    const { ctx } = this;
    const dpr = window.devicePixelRatio || 1;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    this.bgColor ||= getComputedStyle(document.body).getPropertyValue('--canvas-bg').trim() || '#16171a';
    ctx.fillStyle = this.bgColor;
    ctx.fillRect(0, 0, this.width, this.height);
    const doc = app.doc;
    if (!doc) return;
    const { zoom, panX, panY } = doc.view;

    if (this.comp.width !== doc.width || this.comp.height !== doc.height) {
      this.comp.width = doc.width;
      this.comp.height = doc.height;
      this.compDirty = true;
      this.compRect = null;
    }
    if (this.compDirty) {
      const rect = clipRect(this.compRect, doc);
      this.compRect = null;
      this.compDirty = false;
      if (rect !== false) {
        app.tool?.prepareRender?.(rect);
        doc.composite(this.comp.getContext('2d'), { rect });
      }
    }

    // Sombra do documento
    ctx.save();
    ctx.shadowColor = 'rgba(0,0,0,0.5)';
    ctx.shadowBlur = this.interacting ? 0 : 18;
    ctx.fillStyle = '#fff';
    ctx.fillRect(panX, panY, doc.width * zoom, doc.height * zoom);
    ctx.restore();

    // Xadrez de transparência (tamanho constante na tela)
    ctx.save();
    this.checker.setTransform(new DOMMatrix().translate(panX, panY));
    ctx.fillStyle = this.checker;
    ctx.fillRect(panX, panY, doc.width * zoom, doc.height * zoom);
    ctx.restore();

    ctx.save();
    ctx.translate(panX, panY);
    ctx.scale(zoom, zoom);
    ctx.imageSmoothingEnabled = zoom < 1;
    // Durante arrastos usa filtragem rápida; ao soltar, redesenha em alta qualidade.
    ctx.imageSmoothingQuality = this.interacting ? 'low' : 'high';
    ctx.drawImage(this.comp, 0, 0);
    ctx.restore();

    if (zoom >= 10) this.drawPixelGrid(doc);
    if (app.showGrid) this.drawGrid(doc);

    if (doc.selection?._ants && !app.tool?.hideAnts) this.drawAnts(doc.selection._ants);
    app.tool?.drawOverlay?.(ctx, this);
  }

  drawAnts(path) {
    const { ctx } = this;
    const { zoom, panX, panY } = app.doc.view;
    ctx.save();
    ctx.translate(panX, panY);
    ctx.scale(zoom, zoom);
    ctx.lineWidth = 1 / zoom;
    ctx.setLineDash([4 / zoom, 4 / zoom]);
    ctx.strokeStyle = '#000';
    ctx.lineDashOffset = -this.antsOffset / zoom;
    ctx.stroke(path);
    ctx.strokeStyle = '#fff';
    ctx.lineDashOffset = (-this.antsOffset + 4) / zoom;
    ctx.stroke(path);
    ctx.restore();
  }

  drawPixelGrid(doc) {
    const { ctx } = this;
    const { zoom, panX, panY } = doc.view;
    const x0 = Math.max(0, Math.floor(-panX / zoom)), y0 = Math.max(0, Math.floor(-panY / zoom));
    const x1 = Math.min(doc.width, Math.ceil((this.width - panX) / zoom));
    const y1 = Math.min(doc.height, Math.ceil((this.height - panY) / zoom));
    ctx.save();
    ctx.strokeStyle = 'rgba(128,128,128,0.35)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let x = x0; x <= x1; x++) {
      const sx = Math.round(x * zoom + panX) + 0.5;
      ctx.moveTo(sx, y0 * zoom + panY);
      ctx.lineTo(sx, y1 * zoom + panY);
    }
    for (let y = y0; y <= y1; y++) {
      const sy = Math.round(y * zoom + panY) + 0.5;
      ctx.moveTo(x0 * zoom + panX, sy);
      ctx.lineTo(x1 * zoom + panX, sy);
    }
    ctx.stroke();
    ctx.restore();
  }

  drawGrid(doc) {
    const { ctx } = this;
    const { zoom, panX, panY } = doc.view;
    let step = 50;
    while (step * zoom < 24) step *= 2;
    ctx.save();
    ctx.strokeStyle = 'rgba(0,160,255,0.45)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let x = 0; x <= doc.width; x += step) {
      const sx = Math.round(x * zoom + panX) + 0.5;
      ctx.moveTo(sx, panY);
      ctx.lineTo(sx, panY + doc.height * zoom);
    }
    for (let y = 0; y <= doc.height; y += step) {
      const sy = Math.round(y * zoom + panY) + 0.5;
      ctx.moveTo(panX, sy);
      ctx.lineTo(panX + doc.width * zoom, sy);
    }
    ctx.stroke();
    ctx.restore();
  }

  // Lê um pixel da composição atual.
  samplePixel(x, y, size = 1) {
    const doc = app.doc;
    x = Math.floor(x); y = Math.floor(y);
    const r = Math.floor(size / 2);
    const x0 = clamp(x - r, 0, doc.width - 1), y0 = clamp(y - r, 0, doc.height - 1);
    const x1 = clamp(x + r, 0, doc.width - 1), y1 = clamp(y + r, 0, doc.height - 1);
    const flat = this.comp;
    const d = flat.getContext('2d').getImageData(x0, y0, x1 - x0 + 1, y1 - y0 + 1).data;
    let R = 0, G = 0, B = 0, A = 0, n = 0;
    for (let i = 0; i < d.length; i += 4) {
      const a = d[i + 3] / 255;
      R += d[i] * a; G += d[i + 1] * a; B += d[i + 2] * a; A += a; n++;
    }
    if (A === 0) return null;
    return { r: R / A, g: G / A, b: B / A, a: A / n };
  }
}
