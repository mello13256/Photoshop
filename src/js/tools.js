// Ferramentas da caixa de ferramentas.
import { app } from './state.js';
import {
  createCanvas, cloneCanvas, scratch, clamp, rgba, rgbToHex, el,
} from './utils.js';
import {
  maskFromPath, combineMasks, translateMask, floodRegion, maskFromBits, maskedCopy, maskedHole,
  pointInMask, alphaBounds, buildAntsPath,
} from './selection.js';
import { toast, dialog, field } from './ui.js';
import { FILTERS } from './filters.js';
import { BLEND_MODES } from './state.js';
import { unionRect } from './view.js';

const doc = () => app.doc;
const render = (recomp = true, rect = null) => app.view?.requestRender(recomp, rect);

// Desenha src → ctx apenas na região r, aplicando a máscara de seleção (se houver).
function drawMaskedRegion(ctx, src, mask, r) {
  if (!mask) {
    ctx.drawImage(src, r.x, r.y, r.w, r.h, r.x, r.y, r.w, r.h);
    return;
  }
  const tmp = scratch('region', r.w, r.h);
  const t = tmp.getContext('2d');
  t.drawImage(src, r.x, r.y, r.w, r.h, 0, 0, r.w, r.h);
  t.globalCompositeOperation = 'destination-in';
  t.drawImage(mask, r.x, r.y, r.w, r.h, 0, 0, r.w, r.h);
  t.globalCompositeOperation = 'source-over';
  ctx.drawImage(tmp, r.x, r.y);
}

function ensureEditable(layer) {
  if (!layer) return false;
  if (layer.locked) { toast('A camada está bloqueada.'); return false; }
  if (!layer.visible) { toast('A camada está oculta. Torne-a visível para editar.'); return false; }
  return true;
}

function selMode(e, base) {
  if (e.shiftKey && e.altKey) return 'int';
  if (e.shiftKey) return 'add';
  if (e.altKey) return 'sub';
  return base;
}

function pick(p, toBg = false, size = 1) {
  const c = app.view.samplePixel(p.x, p.y, size);
  if (!c) return;
  const hex = rgbToHex(c.r, c.g, c.b);
  if (toBg) app.setBg(hex);
  else app.setFg(hex);
}

function maskWithSelection(canvas, key = 'masked') {
  const sel = doc().selection;
  if (!sel) return canvas;
  const m = scratch(key, canvas.width, canvas.height);
  const ctx = m.getContext('2d');
  ctx.drawImage(canvas, 0, 0);
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(sel, 0, 0);
  ctx.globalCompositeOperation = 'source-over';
  return m;
}

const blendChoices = BLEND_MODES.map(([v, l]) => [v, l]);

// ---------------------------------------------------------------------------
// Pontas de pincel
const tipCache = new Map();
function brushTip(size, hardness, color) {
  const key = `${size}|${hardness}|${color}`;
  let tip = tipCache.get(key);
  if (tip) return tip;
  if (tipCache.size > 64) tipCache.clear();
  const d = Math.max(2, Math.ceil(size) + 2);
  tip = createCanvas(d, d);
  const ctx = tip.getContext('2d');
  const r = Math.max(0.5, size / 2), c = d / 2;
  if (hardness >= 0.99) {
    ctx.fillStyle = color;
    ctx.beginPath();
    ctx.arc(c, c, r, 0, Math.PI * 2);
    ctx.fill();
  } else {
    const g = ctx.createRadialGradient(c, c, 0, c, c, r);
    const h = clamp(hardness, 0, 0.99);
    g.addColorStop(0, rgba(color, 1));
    g.addColorStop(h, rgba(color, 1));
    for (let i = 1; i <= 6; i++) {
      const t = i / 6;
      const a = 1 - t * t * (3 - 2 * t); // smoothstep invertido
      g.addColorStop(h + (1 - h) * t, rgba(color, a));
    }
    ctx.fillStyle = g;
    ctx.fillRect(0, 0, d, d);
  }
  tipCache.set(key, tip);
  return tip;
}

// ---------------------------------------------------------------------------
// Base para ferramentas de pintura por carimbos (pincel, borracha, carimbo, etc.)
class PaintTool {
  constructor(cfg) {
    Object.assign(this, { cursor: 'crosshair', altPicks: true, op: 'source-over', spacing: 0.12 }, cfg);
  }

  get size() { return this.opts.size; }

  pressureOf(e) {
    return e.pointerType === 'pen' && this.opts.pressure !== false ? clamp(e.pressure || 0.5, 0.05, 1) : 1;
  }

  down(e, p) {
    const d = doc();
    if (e.altKey && this.altPicks) { this.picking = true; pick(p); return; }
    const layer = d.activeLayer;
    if (!ensureEditable(layer)) return;
    if (this.beforeStroke && this.beforeStroke(e, p) === false) return;
    this.layer = layer;
    this.stroke = scratch('stroke', d.width, d.height);
    this.sctx = this.stroke.getContext('2d');
    this.pv = scratch('preview', d.width, d.height);
    this.pv.getContext('2d').drawImage(layer.canvas, 0, 0);
    layer.preview = this.pv;
    this.dirty = null;
    const pr = this.pressureOf(e);
    this.active = true;
    if (e.shiftKey && this.lastPoint && this.lastDoc === d) {
      this.last = this.lastPoint;
      this.carry = 0;
      this.dab(this.last, pr);
      this.lineTo(p, pr);
    } else {
      this.last = p;
      this.carry = 0;
      this.dab(p, pr);
    }
    this.flush();
  }

  move(e, p) {
    if (this.picking) { pick(p); return; }
    if (!this.active) return;
    this.lineTo(p, this.pressureOf(e));
    this.flush();
  }

  markDirty(x, y, w) {
    const r = { x: x - 2, y: y - 2, w: w + 4, h: w + 4 };
    this.dirty = this.dirty ? unionRect(this.dirty, r) : r;
  }

  flush() {
    if (this.dirty) render(true, this.dirty);
    this.dirty = null;
  }

  up() {
    if (this.picking) { this.picking = false; return; }
    if (!this.active) return;
    this.active = false;
    this.lastPoint = this.last;
    this.lastDoc = doc();
    const d = doc(), layer = this.layer;
    const stroke = maskWithSelection(this.stroke);
    d.editLayer(layer, (ctx) => {
      ctx.globalAlpha = this.opts.opacity / 100;
      ctx.globalCompositeOperation = this.blendOp();
      ctx.drawImage(stroke, 0, 0);
    });
    layer.preview = null;
    this.source = null;
    d.commit(this.name);
    render();
  }

  cancel() {
    if (!this.active) return;
    this.active = false;
    this.layer.preview = null;
    render();
  }

  blendOp() {
    return this.opts.mode || this.op;
  }

  lineTo(p, pr) {
    const dx = p.x - this.last.x, dy = p.y - this.last.y;
    const len = Math.hypot(dx, dy);
    if (len === 0) return;
    const step = Math.max(0.5, this.size * pr * this.spacing);
    let t = step - this.carry;
    while (t <= len) {
      this.dab({ x: this.last.x + (dx * t) / len, y: this.last.y + (dy * t) / len }, pr);
      t += step;
    }
    this.carry = len - (t - step);
    this.last = p;
  }

  dab(p, pr) {
    const s = this.size * pr;
    const tip = brushTip(this.size, this.opts.hardness / 100, this.tipColor());
    const scale = s / this.size;
    const w = tip.width * scale;
    this.sctx.globalAlpha = (this.opts.flow ?? 100) / 100;
    this.sctx.drawImage(tip, p.x - w / 2, p.y - w / 2, w, w);
    this.markDirty(p.x - w / 2, p.y - w / 2, w);
  }

  tipColor() { return app.fg; }

  prepareRender(rect) {
    if (!this.active || !this.layer) return;
    const d = doc();
    const r = rect || { x: 0, y: 0, w: d.width, h: d.height };
    const ctx = this.pv.getContext('2d');
    ctx.save();
    ctx.clearRect(r.x, r.y, r.w, r.h);
    ctx.drawImage(this.layer.canvas, r.x, r.y, r.w, r.h, r.x, r.y, r.w, r.h);
    ctx.globalAlpha = this.opts.opacity / 100;
    ctx.globalCompositeOperation = this.blendOp();
    drawMaskedRegion(ctx, this.stroke, d.selection, r);
    ctx.restore();
    this.layer.preview = this.pv;
  }

  drawOverlay(ctx, view) {
    const s = view.pointerScreen;
    if (!s || !app.doc) return;
    const r = (this.size / 2) * app.doc.view.zoom;
    ctx.save();
    ctx.lineWidth = 1;
    ctx.strokeStyle = 'rgba(0,0,0,0.8)';
    ctx.beginPath();
    ctx.arc(s.x, s.y, Math.max(1.5, r), 0, Math.PI * 2);
    ctx.stroke();
    ctx.strokeStyle = 'rgba(255,255,255,0.8)';
    ctx.beginPath();
    ctx.arc(s.x, s.y, Math.max(1.5, r) + 1, 0, Math.PI * 2);
    ctx.stroke();
    ctx.restore();
  }
}

const brushOptions = (extra = []) => [
  { key: 'size', label: 'Tamanho', type: 'range', min: 1, max: 1000, value: 20, unit: 'px', log: true },
  { key: 'hardness', label: 'Dureza', type: 'range', min: 0, max: 100, value: 80, unit: '%' },
  ...extra,
  { key: 'opacity', label: 'Opacidade', type: 'range', min: 1, max: 100, value: 100, unit: '%' },
  { key: 'flow', label: 'Fluxo', type: 'range', min: 1, max: 100, value: 100, unit: '%' },
  { key: 'pressure', label: 'Pressão', type: 'checkbox', value: true },
];

// Ferramentas que pintam a partir de uma "fonte" (carimbo, desfoque, subexposição…)
class SourceTool extends PaintTool {
  beforeStroke() {
    const d = doc();
    this.source = this.makeSource(d.activeLayer);
    if (!this.source) return false;
    this.dabCanvas = null;
    return true;
  }

  dab(p, pr) {
    const s = Math.max(1, this.size * pr);
    const D = Math.max(2, Math.ceil(s) + 2);
    if (!this.dabCanvas || this.dabCanvas.width !== D) this.dabCanvas = createCanvas(D, D);
    const dc = this.dabCanvas.getContext('2d');
    dc.globalCompositeOperation = 'source-over';
    dc.clearRect(0, 0, D, D);
    const o = this.offset || { x: 0, y: 0 };
    dc.drawImage(this.source, p.x + o.x - D / 2, p.y + o.y - D / 2, D, D, 0, 0, D, D);
    dc.globalCompositeOperation = 'destination-in';
    const tip = brushTip(this.size, this.opts.hardness / 100, '#000');
    dc.drawImage(tip, 0, 0, D, D);
    this.sctx.globalAlpha = (this.opts.flow ?? 100) / 100;
    this.sctx.drawImage(this.dabCanvas, p.x - D / 2, p.y - D / 2);
    this.markDirty(p.x - D / 2, p.y - D / 2, D);
  }
}

class CloneTool extends SourceTool {
  down(e, p) {
    if (e.altKey) {
      this.src = { x: p.x, y: p.y };
      this.offset = null;
      toast('Origem do carimbo definida.');
      return;
    }
    if (!this.src) { toast('Alt+clique para definir a origem da clonagem.'); return; }
    if (!this.offset || !this.opts.aligned) this.offset = { x: this.src.x - p.x, y: this.src.y - p.y };
    super.down({ altKey: false, shiftKey: e.shiftKey, pointerType: e.pointerType, pressure: e.pressure }, p);
  }

  makeSource(layer) {
    return this.opts.sampleAll ? cloneCanvas(app.view.comp) : cloneCanvas(layer.canvas);
  }

  drawOverlay(ctx, view) {
    super.drawOverlay(ctx, view);
    const pt = view.pointer;
    let srcPt = null;
    if (this.offset && pt) srcPt = { x: pt.x + this.offset.x, y: pt.y + this.offset.y };
    else if (this.src) srcPt = this.src;
    if (!srcPt) return;
    const s = view.toScreen(srcPt.x, srcPt.y);
    ctx.save();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 1;
    ctx.beginPath();
    ctx.moveTo(s.x - 8, s.y); ctx.lineTo(s.x + 8, s.y);
    ctx.moveTo(s.x, s.y - 8); ctx.lineTo(s.x, s.y + 8);
    ctx.stroke();
    ctx.strokeStyle = '#000';
    ctx.setLineDash([2, 2]);
    ctx.stroke();
    ctx.restore();
  }
}

class FocusTool extends SourceTool {
  makeSource(layer) {
    const st = this.opts.strength / 100;
    if (this.opts.kind === 'sharpen') return FILTERS.sharpen.run(layer.canvas, { amount: 60 + st * 200, radius: 1.2, threshold: 0 });
    return FILTERS.gaussianBlur.run(layer.canvas, { radius: 1 + st * 6 });
  }
  blendOp() { return 'source-atop'; }
}

class ToneTool extends SourceTool {
  makeSource(layer) {
    const st = this.opts.strength / 100;
    switch (this.opts.kind) {
      case 'burn': return FILTERS.exposure.run(layer.canvas, { exposure: -0.3 - st * 1.2, offset: 0, gamma: 1 });
      case 'sponge-': return FILTERS.hueSaturation.run(layer.canvas, { hue: 0, saturation: -30 - st * 70, lightness: 0 });
      case 'sponge+': return FILTERS.vibrance.run(layer.canvas, { vibrance: 0, saturation: 20 + st * 80 });
      default: return FILTERS.exposure.run(layer.canvas, { exposure: 0.3 + st * 1.2, offset: 0, gamma: 1 });
    }
  }
  blendOp() { return 'source-atop'; }
}

// ---------------------------------------------------------------------------
// Seleções

function shapePath(a, b, kind, { square = false, fromCenter = false, radius = 0 } = {}) {
  let x0 = a.x, y0 = a.y, x1 = b.x, y1 = b.y;
  if (square) {
    const s = Math.max(Math.abs(x1 - x0), Math.abs(y1 - y0));
    x1 = x0 + Math.sign(x1 - x0 || 1) * s;
    y1 = y0 + Math.sign(y1 - y0 || 1) * s;
  }
  if (fromCenter) {
    const dx = x1 - x0, dy = y1 - y0;
    x0 -= dx; y0 -= dy;
  }
  const x = Math.min(x0, x1), y = Math.min(y0, y1), w = Math.abs(x1 - x0), h = Math.abs(y1 - y0);
  const p = new Path2D();
  if (kind === 'ellipse') p.ellipse(x + w / 2, y + h / 2, w / 2, h / 2, 0, 0, Math.PI * 2);
  else if (kind === 'rounded') p.roundRect(x, y, w, h, Math.min(radius, w / 2, h / 2));
  else p.rect(x, y, w, h);
  return { path: p, x, y, w, h };
}

class MarqueeTool {
  constructor(cfg) {
    Object.assign(this, { cursor: 'crosshair' }, cfg);
  }

  down(e, p) {
    const d = doc();
    this.mode = selMode(e, this.opts.mode);
    this.shiftAtStart = e.shiftKey;
    if (this.mode === 'new' && d.selection && pointInMask(d.selection, p.x, p.y)) {
      this.movingSel = { start: p, orig: d.selection, dx: 0, dy: 0 };
      this.hideAnts = true;
      return;
    }
    const snap = (v) => (this.kind === 'rect' && !this.opts.feather ? Math.round(v) : v);
    this.a = { x: snap(p.x), y: snap(p.y) };
    this.b = this.a;
    this.dragging = true;
    this.snap = snap;
  }

  move(e, p) {
    if (this.movingSel) {
      this.movingSel.dx = Math.round(p.x - this.movingSel.start.x);
      this.movingSel.dy = Math.round(p.y - this.movingSel.start.y);
      render(false);
      return;
    }
    if (!this.dragging) return;
    this.b = { x: this.snap(p.x), y: this.snap(p.y) };
    this.square = e.shiftKey && !this.shiftAtStart;
    this.fromCenter = e.altKey && this.mode !== 'sub';
    const r = this.rect();
    app.emit('status-info', `${Math.round(r.w)} × ${Math.round(r.h)} px`);
    render(false);
  }

  rect() {
    return shapePath(this.a, this.b, this.kind, { square: this.square, fromCenter: this.fromCenter });
  }

  up() {
    const d = doc();
    if (this.movingSel) {
      const { orig, dx, dy } = this.movingSel;
      this.movingSel = null;
      this.hideAnts = false;
      if (dx || dy) {
        d.setSelection(translateMask(orig, dx, dy));
        d.commit('Mover seleção');
      } else {
        d.setSelection(null);
        d.commit('Desmarcar');
      }
      render(false);
      return;
    }
    if (!this.dragging) return;
    this.dragging = false;
    const r = this.rect();
    if (r.w < 1 || r.h < 1) {
      if (this.mode === 'new' && d.selection) {
        d.setSelection(null);
        d.commit('Desmarcar');
      }
      render(false);
      return;
    }
    const mask = maskFromPath(d.width, d.height, r.path, this.opts.feather || 0);
    d.setSelection(combineMasks(d.selection, mask, this.mode));
    d.commit(this.kind === 'ellipse' ? 'Letreiro elíptico' : 'Letreiro retangular');
    render(false);
  }

  drawOverlay(ctx, view) {
    const { zoom, panX, panY } = doc().view;
    if (this.movingSel) {
      const path = buildAntsPath(this.movingSel.orig);
      ctx.save();
      ctx.translate(panX + this.movingSel.dx * zoom, panY + this.movingSel.dy * zoom);
      ctx.scale(zoom, zoom);
      ctx.lineWidth = 1 / zoom;
      ctx.setLineDash([4 / zoom, 4 / zoom]);
      ctx.strokeStyle = '#000';
      ctx.stroke(path);
      ctx.strokeStyle = '#fff';
      ctx.lineDashOffset = 4 / zoom;
      ctx.stroke(path);
      ctx.restore();
      return;
    }
    if (!this.dragging) return;
    const r = this.rect();
    ctx.save();
    ctx.translate(panX, panY);
    ctx.scale(zoom, zoom);
    ctx.lineWidth = 1 / zoom;
    ctx.setLineDash([4 / zoom, 4 / zoom]);
    ctx.lineDashOffset = -view.antsOffset / zoom;
    ctx.strokeStyle = '#000';
    ctx.stroke(r.path);
    ctx.strokeStyle = '#fff';
    ctx.lineDashOffset = (-view.antsOffset + 4) / zoom;
    ctx.stroke(r.path);
    ctx.restore();
  }
}

const selectionOptions = (extra = []) => [
  { key: 'mode', label: 'Modo', type: 'buttons', value: 'new', choices: [['new', 'Nova'], ['add', 'Adicionar'], ['sub', 'Subtrair'], ['int', 'Interseção']] },
  ...extra,
];

class LassoTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }

  down(e, p) {
    if (this.opts.poly) {
      const now = performance.now();
      if (!this.points) {
        this.points = [p];
        this.mode = selMode(e, this.opts.mode);
      } else {
        const first = app.view.toScreen(this.points[0].x, this.points[0].y);
        const prev = this.points[this.points.length - 1];
        const lastS = app.view.toScreen(prev.x, prev.y);
        const sp = app.view.toScreen(p.x, p.y);
        const dbl = now - this.lastClick < 350 && Math.hypot(lastS.x - sp.x, lastS.y - sp.y) < 6;
        if ((this.points.length > 2 && Math.hypot(first.x - sp.x, first.y - sp.y) < 8) || dbl) {
          this.finish();
          return;
        }
        this.points.push(p);
      }
      this.lastClick = now;
      this.hover = p;
      render(false);
      return;
    }
    this.mode = selMode(e, this.opts.mode);
    this.points = [p];
    this.dragging = true;
  }

  hoverMove(e, p) {
    if (this.opts.poly && this.points) { this.hover = p; render(false); }
  }

  move(e, p) {
    if (this.opts.poly) { this.hoverMove(e, p); return; }
    if (!this.dragging) return;
    const last = this.points[this.points.length - 1];
    if (Math.hypot(p.x - last.x, p.y - last.y) * doc().view.zoom >= 1.5) this.points.push(p);
    render(false);
  }

  up() {
    if (this.opts.poly) return;
    if (!this.dragging) return;
    this.dragging = false;
    this.finish();
  }

  finish() {
    const d = doc();
    const pts = this.points;
    this.points = null;
    this.hover = null;
    if (!pts || pts.length < 3) {
      if (this.mode === 'new' && d.selection) { d.setSelection(null); d.commit('Desmarcar'); }
      render(false);
      return;
    }
    const path = new Path2D();
    path.moveTo(pts[0].x, pts[0].y);
    for (const q of pts.slice(1)) path.lineTo(q.x, q.y);
    path.closePath();
    const mask = maskFromPath(d.width, d.height, path, this.opts.feather || 0);
    d.setSelection(combineMasks(d.selection, mask, this.mode));
    d.commit(this.opts.poly ? 'Laço poligonal' : 'Laço');
    render(false);
  }

  keydown(e) {
    if (!this.points) return false;
    if (e.key === 'Enter') { this.finish(); return true; }
    if (e.key === 'Escape') { this.points = null; render(false); return true; }
    if (e.key === 'Backspace' && this.opts.poly && this.points.length > 1) { this.points.pop(); render(false); return true; }
    return false;
  }

  cancel() { this.points = null; this.dragging = false; render(false); }

  drawOverlay(ctx, view) {
    if (!this.points) return;
    ctx.save();
    ctx.beginPath();
    this.points.forEach((q, i) => {
      const s = view.toScreen(q.x, q.y);
      if (i) ctx.lineTo(s.x, s.y); else ctx.moveTo(s.x, s.y);
    });
    if (this.hover) { const s = view.toScreen(this.hover.x, this.hover.y); ctx.lineTo(s.x, s.y); }
    ctx.lineWidth = 1;
    ctx.strokeStyle = '#000';
    ctx.setLineDash([4, 4]);
    ctx.stroke();
    ctx.strokeStyle = '#fff';
    ctx.lineDashOffset = 4;
    ctx.stroke();
    ctx.restore();
  }
}

class WandTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }
  down(e, p) {
    const d = doc();
    const x = Math.floor(p.x), y = Math.floor(p.y);
    if (x < 0 || y < 0 || x >= d.width || y >= d.height) return;
    const src = this.opts.sampleAll ? app.view.comp : d.activeLayer.canvas;
    const id = src.getContext('2d').getImageData(0, 0, d.width, d.height);
    const bits = floodRegion(id, x, y, this.opts.tolerance, this.opts.contiguous);
    let mask = maskFromBits(bits, d.width, d.height);
    if (this.opts.antialias) {
      const m = createCanvas(d.width, d.height);
      const c = m.getContext('2d');
      c.filter = 'blur(0.6px)';
      c.drawImage(mask, 0, 0);
      mask = m;
    }
    d.setSelection(combineMasks(d.selection, mask, selMode(e, this.opts.mode)));
    d.commit('Varinha mágica');
    render(false);
  }
}

// ---------------------------------------------------------------------------
// Mover

class MoveTool {
  constructor(cfg) { Object.assign(this, { cursor: 'move' }, cfg); }

  lift(layer) {
    const d = doc();
    this.layer = layer;
    this.sel = d.selection;
    this.float = maskedCopy(layer.canvas, this.sel);
    this.hole = maskedHole(layer.canvas, this.sel);
    this.dx = this.dy = 0;
  }

  down(e, p) {
    const d = doc();
    if (this.opts.autoSelect && !d.selection) {
      const x = Math.floor(p.x), y = Math.floor(p.y);
      for (let i = d.layers.length - 1; i >= 0; i--) {
        const l = d.layers[i];
        if (!l.visible || x < 0 || y < 0 || x >= d.width || y >= d.height) continue;
        if (l.canvas.getContext('2d').getImageData(x, y, 1, 1).data[3] > 10) {
          if (d.active !== i) { d.active = i; app.emit('layers'); }
          break;
        }
      }
    }
    const layer = d.activeLayer;
    if (!ensureEditable(layer)) return;
    this.lift(layer);
    this.start = p;
    this.active = true;
    this.hideAnts = !!this.sel;
  }

  move(e, p) {
    if (!this.active) return;
    let dx = Math.round(p.x - this.start.x), dy = Math.round(p.y - this.start.y);
    if (e.shiftKey) { if (Math.abs(dx) > Math.abs(dy)) dy = 0; else dx = 0; }
    this.dx = dx; this.dy = dy;
    app.emit('status-info', `Δ ${dx}, ${dy} px`);
    render();
  }

  up() {
    if (!this.active) return;
    this.active = false;
    this.hideAnts = false;
    this.finish('Mover');
  }

  finish(label) {
    const d = doc(), layer = this.layer;
    if (this.dx || this.dy) {
      this.prepareRender(true);
      const out = cloneCanvas(layer.preview);
      const text = layer.text && !this.sel ? { ...layer.text, x: layer.text.x + this.dx, y: layer.text.y + this.dy } : null;
      d.replaceLayerCanvas(layer, out);
      if (text) layer.text = text;
      if (this.sel) d.setSelection(translateMask(this.sel, this.dx, this.dy));
      layer.preview = null;
      d.commit(label);
    }
    layer.preview = null;
    render();
  }

  prepareRender(force) {
    if ((!this.active && !force) || !this.layer) return;
    const d = doc();
    const pv = scratch('preview', d.width, d.height);
    const ctx = pv.getContext('2d');
    ctx.drawImage(this.hole, 0, 0);
    ctx.drawImage(this.float, this.dx, this.dy);
    this.layer.preview = pv;
  }

  keydown(e) {
    const dirs = { ArrowLeft: [-1, 0], ArrowRight: [1, 0], ArrowUp: [0, -1], ArrowDown: [0, 1] };
    const dir = dirs[e.key];
    if (!dir || this.active) return false;
    const layer = doc().activeLayer;
    if (!ensureEditable(layer)) return true;
    const step = e.shiftKey ? 10 : 1;
    this.lift(layer);
    this.dx = dir[0] * step;
    this.dy = dir[1] * step;
    this.finish('Deslocar');
    return true;
  }

  drawOverlay(ctx) {
    if (!this.active || !this.sel) return;
    const { zoom, panX, panY } = doc().view;
    ctx.save();
    ctx.translate(panX + this.dx * zoom, panY + this.dy * zoom);
    ctx.scale(zoom, zoom);
    ctx.lineWidth = 1 / zoom;
    ctx.setLineDash([4 / zoom, 4 / zoom]);
    ctx.strokeStyle = '#fff';
    ctx.stroke(buildAntsPath(this.sel));
    ctx.restore();
  }
}

// ---------------------------------------------------------------------------
// Transformação livre (Ctrl+T)

const HANDLES = [[-1, -1], [0, -1], [1, -1], [1, 0], [1, 1], [0, 1], [-1, 1], [-1, 0]];

class TransformTool {
  constructor(cfg) { Object.assign(this, { cursor: 'default', hideAnts: true }, cfg); }

  begin() {
    const d = doc();
    if (!d) return false;
    const layer = d.activeLayer;
    if (!ensureEditable(layer)) return false;
    const sel = d.selection;
    const content = maskedCopy(layer.canvas, sel);
    const b = alphaBounds(content);
    if (!b) { toast('Não há pixels para transformar nesta camada.'); return false; }
    this.layer = layer;
    this.sel = sel;
    this.src = createCanvas(b.w, b.h);
    this.src.getContext('2d').drawImage(content, -b.x, -b.y);
    if (sel) {
      this.maskSrc = createCanvas(b.w, b.h);
      this.maskSrc.getContext('2d').drawImage(sel, -b.x, -b.y);
    } else this.maskSrc = null;
    this.hole = maskedHole(layer.canvas, sel);
    this.box = { cx: b.x + b.w / 2, cy: b.y + b.h / 2, w: b.w, h: b.h, angle: 0 };
    this.orig = { ...this.box };
    this.running = true;
    render();
    app.emit('tool-options');
    return true;
  }

  activate() {
    if (!this.running && !this.begin()) return false;
    return true;
  }

  deactivate() {
    if (this.running) this.commit();
  }

  toLocal(p, box = this.box) {
    const dx = p.x - box.cx, dy = p.y - box.cy;
    const c = Math.cos(-box.angle), s = Math.sin(-box.angle);
    return { x: dx * c - dy * s, y: dx * s + dy * c };
  }

  toWorld(lx, ly, box = this.box) {
    const c = Math.cos(box.angle), s = Math.sin(box.angle);
    return { x: box.cx + lx * c - ly * s, y: box.cy + lx * s + ly * c };
  }

  handlePoints() {
    const { w, h } = this.box;
    return HANDLES.map(([hx, hy]) => ({ hx, hy, ...this.toWorld((hx * w) / 2, (hy * h) / 2) }));
  }

  hit(s) {
    for (const hp of this.handlePoints()) {
      const sp = app.view.toScreen(hp.x, hp.y);
      if (Math.hypot(sp.x - s.x, sp.y - s.y) <= 8) return { type: 'scale', hx: hp.hx, hy: hp.hy };
    }
    const l = this.toLocal(app.view.toDoc(s.x, s.y));
    if (Math.abs(l.x) <= Math.abs(this.box.w) / 2 && Math.abs(l.y) <= Math.abs(this.box.h) / 2) return { type: 'move' };
    return { type: 'rotate' };
  }

  hoverMove(e, p, s) {
    if (!this.running) return;
    const h = this.hit(s);
    const c = app.view.canvas;
    if (h.type === 'move') c.style.cursor = 'move';
    else if (h.type === 'rotate') c.style.cursor = 'alias';
    else {
      const ang = Math.atan2(h.hy * this.box.h, h.hx * this.box.w) + this.box.angle;
      const deg = ((ang * 180) / Math.PI + 360) % 180;
      c.style.cursor = deg < 22.5 || deg >= 157.5 ? 'ew-resize' : deg < 67.5 ? 'nwse-resize' : deg < 112.5 ? 'ns-resize' : 'nesw-resize';
    }
  }

  down(e, p, s) {
    if (!this.running) return;
    this.drag = { ...this.hit(s), start: p, box: { ...this.box } };
  }

  move(e, p) {
    if (!this.drag) return;
    const { type, start, box: b0 } = this.drag;
    const box = this.box;
    if (type === 'move') {
      box.cx = b0.cx + p.x - start.x;
      box.cy = b0.cy + p.y - start.y;
    } else if (type === 'rotate') {
      const a0 = Math.atan2(start.y - b0.cy, start.x - b0.cx);
      const a1 = Math.atan2(p.y - b0.cy, p.x - b0.cx);
      let a = b0.angle + a1 - a0;
      if (e.shiftKey) a = Math.round(a / (Math.PI / 12)) * (Math.PI / 12);
      box.angle = a;
    } else {
      const { hx, hy } = this.drag;
      const lp = this.toLocal(p, b0);
      const fromCenter = e.altKey;
      let w = b0.w, h = b0.h, cxl = 0, cyl = 0;
      const fx = fromCenter ? 0 : (-hx * b0.w) / 2;
      const fy = fromCenter ? 0 : (-hy * b0.h) / 2;
      if (hx) w = (lp.x - fx) * hx * (fromCenter ? 2 : 1);
      if (hy) h = (lp.y - fy) * hy * (fromCenter ? 2 : 1);
      if (e.shiftKey && hx && hy) {
        const sx = w / b0.w, sy = h / b0.h;
        const sc = Math.abs(sx) > Math.abs(sy) ? sx : sy;
        w = b0.w * sc; h = b0.h * sc;
      }
      if (!fromCenter) {
        if (hx) cxl = fx + (hx * w) / 2;
        if (hy) cyl = fy + (hy * h) / 2;
      }
      const c = { x: b0.cx + cxl * Math.cos(b0.angle) - cyl * Math.sin(b0.angle), y: b0.cy + cxl * Math.sin(b0.angle) + cyl * Math.cos(b0.angle) };
      Object.assign(box, { w, h, cx: c.x, cy: c.y });
    }
    app.emit('tool-info');
    render();
  }

  up() { this.drag = null; }

  drawTransformed(ctx, img) {
    const { cx, cy, w, h, angle } = this.box;
    ctx.save();
    ctx.translate(cx, cy);
    ctx.rotate(angle);
    ctx.scale(Math.sign(w) || 1, Math.sign(h) || 1);
    ctx.imageSmoothingEnabled = true;
    ctx.imageSmoothingQuality = 'high';
    ctx.drawImage(img, -Math.abs(w) / 2, -Math.abs(h) / 2, Math.abs(w), Math.abs(h));
    ctx.restore();
  }

  prepareRender() {
    if (!this.running) return;
    const d = doc();
    const pv = scratch('preview', d.width, d.height);
    const ctx = pv.getContext('2d');
    ctx.drawImage(this.hole, 0, 0);
    this.drawTransformed(ctx, this.src);
    this.layer.preview = pv;
  }

  flip(horizontal) {
    if (horizontal) this.box.w = -this.box.w;
    else this.box.h = -this.box.h;
    render();
  }

  commit() {
    if (!this.running) return;
    const d = doc();
    this.running = false;
    const out = cloneCanvas(this.hole);
    this.drawTransformed(out.getContext('2d'), this.src);
    this.layer.preview = null;
    d.replaceLayerCanvas(this.layer, out);
    if (this.maskSrc) {
      const m = createCanvas(d.width, d.height);
      this.drawTransformed(m.getContext('2d'), this.maskSrc);
      d.setSelection(m);
    }
    d.commit('Transformação livre');
    render();
    app.emit('transform-done');
  }

  cancel() {
    if (!this.running) return;
    this.running = false;
    this.layer.preview = null;
    render();
    app.emit('transform-done');
  }

  keydown(e) {
    if (e.key === 'Enter') { this.commit(); return true; }
    if (e.key === 'Escape') { this.cancel(); return true; }
    return false;
  }

  info() {
    const { w, h, angle } = this.box;
    return `L: ${((w / this.orig.w) * 100).toFixed(1)}%  A: ${((h / this.orig.h) * 100).toFixed(1)}%  Ângulo: ${((angle * 180) / Math.PI).toFixed(1)}°`;
  }

  drawOverlay(ctx, view) {
    if (!this.running) return;
    const pts = this.handlePoints().map((hp) => view.toScreen(hp.x, hp.y));
    ctx.save();
    ctx.strokeStyle = '#2d8cff';
    ctx.lineWidth = 1;
    ctx.beginPath();
    [0, 2, 4, 6].forEach((i, k) => (k ? ctx.lineTo(pts[i].x, pts[i].y) : ctx.moveTo(pts[i].x, pts[i].y)));
    ctx.closePath();
    ctx.stroke();
    ctx.fillStyle = '#fff';
    for (const s of pts) {
      ctx.fillRect(s.x - 4, s.y - 4, 8, 8);
      ctx.strokeRect(s.x - 4 + 0.5, s.y - 4 + 0.5, 7, 7);
    }
    const c = view.toScreen(this.box.cx, this.box.cy);
    ctx.beginPath();
    ctx.arc(c.x, c.y, 4, 0, Math.PI * 2);
    ctx.stroke();
    ctx.restore();
  }
}

// ---------------------------------------------------------------------------
// Cortar

class CropTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }

  activate() { this.rect = null; return true; }
  deactivate() { this.rect = null; render(false); }

  ratio() {
    const r = this.opts.ratio;
    if (r === 'free') return 0;
    if (r === 'orig') return doc().width / doc().height;
    const [a, b] = r.split(':').map(Number);
    return a / b;
  }

  norm(r) {
    return { x: Math.min(r.x0, r.x1), y: Math.min(r.y0, r.y1), w: Math.abs(r.x1 - r.x0), h: Math.abs(r.y1 - r.y0) };
  }

  hit(s) {
    if (!this.rect) return { type: 'new' };
    const r = this.norm(this.rect);
    for (const [hx, hy] of HANDLES) {
      const sp = app.view.toScreen(r.x + ((hx + 1) / 2) * r.w, r.y + ((hy + 1) / 2) * r.h);
      if (Math.hypot(sp.x - s.x, sp.y - s.y) <= 8) return { type: 'resize', hx, hy };
    }
    const p = app.view.toDoc(s.x, s.y);
    if (p.x >= r.x && p.x <= r.x + r.w && p.y >= r.y && p.y <= r.y + r.h) return { type: 'move' };
    return { type: 'new' };
  }

  hoverMove(e, p, s) {
    const h = this.hit(s);
    app.view.canvas.style.cursor = h.type === 'move' ? 'move' : h.type === 'resize'
      ? (h.hx && h.hy ? (h.hx === h.hy ? 'nwse-resize' : 'nesw-resize') : h.hx ? 'ew-resize' : 'ns-resize')
      : 'crosshair';
  }

  down(e, p, s) {
    const h = this.hit(s);
    const now = performance.now();
    if (h.type === 'move' && now - (this.lastDown || 0) < 300) { this.apply(); return; }
    this.lastDown = now;
    const r = this.rect ? this.norm(this.rect) : null;
    this.drag = { ...h, start: p, r0: r };
    if (h.type === 'new') this.rect = { x0: Math.round(p.x), y0: Math.round(p.y), x1: Math.round(p.x), y1: Math.round(p.y) };
    else this.rect = { x0: r.x, y0: r.y, x1: r.x + r.w, y1: r.y + r.h };
  }

  move(e, p) {
    if (!this.drag) return;
    const ratio = this.ratio() || (e.shiftKey ? 1 : 0);
    const R = this.rect, { type, start, r0 } = this.drag;
    const px = Math.round(p.x), py = Math.round(p.y);
    if (type === 'new') {
      R.x1 = px; R.y1 = py;
      if (ratio) R.y1 = R.y0 + Math.sign(py - R.y0 || 1) * Math.abs(R.x1 - R.x0) / ratio;
    } else if (type === 'move') {
      const dx = Math.round(p.x - start.x), dy = Math.round(p.y - start.y);
      Object.assign(R, { x0: r0.x + dx, y0: r0.y + dy, x1: r0.x + r0.w + dx, y1: r0.y + r0.h + dy });
    } else {
      const { hx, hy } = this.drag;
      if (hx === -1) R.x0 = px; else if (hx === 1) R.x1 = px;
      if (hy === -1) R.y0 = py; else if (hy === 1) R.y1 = py;
      if (ratio) {
        if (hx) {
          const hgt = Math.abs(R.x1 - R.x0) / ratio;
          if (hy === -1) R.y0 = R.y1 - hgt; else R.y1 = R.y0 + hgt;
        } else {
          R.x1 = R.x0 + Math.abs(R.y1 - R.y0) * ratio;
        }
      }
    }
    const n = this.norm(R);
    app.emit('status-info', `${Math.round(n.w)} × ${Math.round(n.h)} px`);
    render(false);
  }

  up() {
    this.drag = null;
    if (this.rect) {
      const n = this.norm(this.rect);
      if (n.w < 2 || n.h < 2) this.rect = null;
    }
    render(false);
  }

  apply() {
    if (!this.rect) { toast('Arraste para definir a área de corte.'); return; }
    const n = this.norm(this.rect);
    this.rect = null;
    doc().crop(n.x, n.y, n.w, n.h);
    app.view.fit();
    render();
  }

  keydown(e) {
    if (e.key === 'Enter') { this.apply(); return true; }
    if (e.key === 'Escape') { this.rect = null; render(false); return true; }
    return false;
  }

  drawOverlay(ctx, view) {
    if (!this.rect) return;
    const n = this.norm(this.rect);
    const a = view.toScreen(n.x, n.y), b = view.toScreen(n.x + n.w, n.y + n.h);
    ctx.save();
    ctx.fillStyle = 'rgba(0,0,0,0.55)';
    ctx.beginPath();
    ctx.rect(0, 0, view.width, view.height);
    ctx.rect(a.x, a.y, b.x - a.x, b.y - a.y);
    ctx.fill('evenodd');
    ctx.strokeStyle = 'rgba(255,255,255,0.5)';
    ctx.lineWidth = 1;
    ctx.beginPath();
    for (let i = 1; i < 3; i++) {
      const x = a.x + ((b.x - a.x) * i) / 3, y = a.y + ((b.y - a.y) * i) / 3;
      ctx.moveTo(x, a.y); ctx.lineTo(x, b.y);
      ctx.moveTo(a.x, y); ctx.lineTo(b.x, y);
    }
    ctx.stroke();
    ctx.strokeStyle = '#fff';
    ctx.strokeRect(a.x + 0.5, a.y + 0.5, b.x - a.x, b.y - a.y);
    ctx.fillStyle = '#fff';
    for (const [hx, hy] of HANDLES) {
      const x = a.x + ((hx + 1) / 2) * (b.x - a.x), y = a.y + ((hy + 1) / 2) * (b.y - a.y);
      ctx.fillRect(x - 4, y - 4, 8, 8);
    }
    ctx.restore();
  }
}

// ---------------------------------------------------------------------------
// Balde e degradê

class BucketTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }
  down(e, p) {
    const d = doc(), layer = d.activeLayer;
    if (e.altKey) { pick(p); return; }
    if (!ensureEditable(layer)) return;
    const x = Math.floor(p.x), y = Math.floor(p.y);
    if (x < 0 || y < 0 || x >= d.width || y >= d.height) return;
    const src = this.opts.sampleAll ? app.view.comp : layer.canvas;
    const id = src.getContext('2d').getImageData(0, 0, d.width, d.height);
    const bits = floodRegion(id, x, y, this.opts.tolerance, this.opts.contiguous);
    const mask = maskFromBits(bits, d.width, d.height);
    const fill = createCanvas(d.width, d.height);
    const f = fill.getContext('2d');
    f.fillStyle = app.fg;
    f.fillRect(0, 0, d.width, d.height);
    f.globalCompositeOperation = 'destination-in';
    if (this.opts.antialias) f.filter = 'blur(0.5px)';
    f.drawImage(mask, 0, 0);
    f.filter = 'none';
    if (d.selection) f.drawImage(d.selection, 0, 0);
    d.editLayer(layer, (ctx) => {
      ctx.globalAlpha = this.opts.opacity / 100;
      ctx.globalCompositeOperation = this.opts.mode;
      ctx.drawImage(fill, 0, 0);
    });
    d.commit('Balde de tinta');
    render();
  }
}

function makeGradient(ctx, a, b, kind, c0, c1) {
  let g;
  const len = Math.hypot(b.x - a.x, b.y - a.y) || 1;
  if (kind === 'radial') g = ctx.createRadialGradient(a.x, a.y, 0, a.x, a.y, len);
  else if (kind === 'angle') g = ctx.createConicGradient(Math.atan2(b.y - a.y, b.x - a.x), a.x, a.y);
  else if (kind === 'reflected') g = ctx.createLinearGradient(2 * a.x - b.x, 2 * a.y - b.y, b.x, b.y);
  else g = ctx.createLinearGradient(a.x, a.y, b.x, b.y);
  if (kind === 'reflected') {
    g.addColorStop(0, c1); g.addColorStop(0.5, c0); g.addColorStop(1, c1);
  } else {
    g.addColorStop(0, c0); g.addColorStop(1, c1);
  }
  return g;
}

class GradientTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }
  colors() {
    let c0 = app.fg, c1 = app.bg;
    if (this.opts.colors === 'fg-transparent') c1 = rgba(app.fg, 0);
    else if (this.opts.colors === 'black-white') { c0 = '#000000'; c1 = '#ffffff'; }
    if (this.opts.reverse) [c0, c1] = [c1, c0];
    return [c0, c1];
  }
  down(e, p) {
    if (!ensureEditable(doc().activeLayer)) return;
    this.layer = doc().activeLayer;
    this.a = p; this.b = p; this.active = true;
  }
  move(e, p) {
    if (!this.active) return;
    if (e.shiftKey) {
      const ang = Math.round(Math.atan2(p.y - this.a.y, p.x - this.a.x) / (Math.PI / 4)) * (Math.PI / 4);
      const len = Math.hypot(p.x - this.a.x, p.y - this.a.y);
      p = { x: this.a.x + Math.cos(ang) * len, y: this.a.y + Math.sin(ang) * len };
    }
    this.b = p;
    render();
  }
  gradientCanvas() {
    const d = doc();
    const g = scratch('gradient', d.width, d.height);
    const ctx = g.getContext('2d');
    const [c0, c1] = this.colors();
    ctx.fillStyle = makeGradient(ctx, this.a, this.b, this.opts.kind, c0, c1);
    ctx.fillRect(0, 0, d.width, d.height);
    return maskWithSelection(g, 'gradient-masked');
  }
  prepareRender() {
    if (!this.active) return;
    const d = doc();
    const pv = scratch('preview', d.width, d.height);
    const ctx = pv.getContext('2d');
    ctx.drawImage(this.layer.canvas, 0, 0);
    if (Math.hypot(this.b.x - this.a.x, this.b.y - this.a.y) > 0.5) {
      ctx.globalAlpha = this.opts.opacity / 100;
      ctx.globalCompositeOperation = this.opts.mode;
      ctx.drawImage(this.gradientCanvas(), 0, 0);
    }
    this.layer.preview = pv;
  }
  up() {
    if (!this.active) return;
    this.active = false;
    const d = doc();
    this.layer.preview = null;
    if (Math.hypot(this.b.x - this.a.x, this.b.y - this.a.y) > 0.5) {
      const g = this.gradientCanvas();
      d.editLayer(this.layer, (ctx) => {
        ctx.globalAlpha = this.opts.opacity / 100;
        ctx.globalCompositeOperation = this.opts.mode;
        ctx.drawImage(g, 0, 0);
      });
      d.commit('Degradê');
    }
    render();
  }
  drawOverlay(ctx, view) {
    if (!this.active) return;
    const a = view.toScreen(this.a.x, this.a.y), b = view.toScreen(this.b.x, this.b.y);
    ctx.save();
    ctx.strokeStyle = '#fff';
    ctx.lineWidth = 2;
    ctx.beginPath(); ctx.moveTo(a.x, a.y); ctx.lineTo(b.x, b.y); ctx.stroke();
    ctx.strokeStyle = '#000';
    ctx.lineWidth = 1;
    ctx.stroke();
    for (const s of [a, b]) { ctx.beginPath(); ctx.arc(s.x, s.y, 4, 0, Math.PI * 2); ctx.fillStyle = '#fff'; ctx.fill(); ctx.stroke(); }
    ctx.restore();
  }
}

// ---------------------------------------------------------------------------
// Formas

function starPath(cx, cy, r, points, inner) {
  const p = new Path2D();
  for (let i = 0; i < points * 2; i++) {
    const rad = i % 2 ? r * inner : r;
    const a = (i * Math.PI) / points - Math.PI / 2;
    const x = cx + Math.cos(a) * rad, y = cy + Math.sin(a) * rad;
    if (i) p.lineTo(x, y); else p.moveTo(x, y);
  }
  p.closePath();
  return p;
}

class ShapeTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }

  down(e, p) {
    const d = doc();
    if (this.opts.newLayer) {
      this.layer = d.addLayer(`Forma ${d.layers.length}`);
      this.created = true;
      app.emit('layers');
    } else {
      if (!ensureEditable(d.activeLayer)) return;
      this.layer = d.activeLayer;
      this.created = false;
    }
    this.a = p; this.b = p; this.active = true;
  }

  move(e, p) {
    if (!this.active) return;
    this.b = p;
    this.square = e.shiftKey;
    this.fromCenter = e.altKey;
    render();
  }

  draw(ctx) {
    const o = this.opts;
    const { a } = this;
    let { b } = this;
    ctx.lineJoin = 'round';
    ctx.lineCap = 'round';
    ctx.fillStyle = app.fg;
    ctx.strokeStyle = o.strokeColor;
    ctx.lineWidth = o.strokeWidth;
    if (o.shape === 'line' || o.shape === 'arrow') {
      if (this.square) {
        const ang = Math.round(Math.atan2(b.y - a.y, b.x - a.x) / (Math.PI / 4)) * (Math.PI / 4);
        const len = Math.hypot(b.x - a.x, b.y - a.y);
        b = { x: a.x + Math.cos(ang) * len, y: a.y + Math.sin(ang) * len };
      }
      ctx.strokeStyle = app.fg;
      ctx.lineWidth = Math.max(1, o.strokeWidth);
      ctx.beginPath(); ctx.moveTo(a.x, a.y); ctx.lineTo(b.x, b.y); ctx.stroke();
      if (o.shape === 'arrow') {
        const ang = Math.atan2(b.y - a.y, b.x - a.x), hl = Math.max(10, ctx.lineWidth * 4);
        ctx.beginPath();
        ctx.moveTo(b.x, b.y);
        ctx.lineTo(b.x - hl * Math.cos(ang - 0.45), b.y - hl * Math.sin(ang - 0.45));
        ctx.lineTo(b.x - hl * Math.cos(ang + 0.45), b.y - hl * Math.sin(ang + 0.45));
        ctx.closePath();
        ctx.fill();
      }
      return;
    }
    let path;
    if (o.shape === 'star') {
      path = starPath(a.x, a.y, Math.hypot(b.x - a.x, b.y - a.y), o.sides, 0.45);
    } else if (o.shape === 'polygon') {
      const r = Math.hypot(b.x - a.x, b.y - a.y);
      path = new Path2D();
      for (let i = 0; i < o.sides; i++) {
        const ang = (i * 2 * Math.PI) / o.sides - Math.PI / 2;
        const x = a.x + Math.cos(ang) * r, y = a.y + Math.sin(ang) * r;
        if (i) path.lineTo(x, y); else path.moveTo(x, y);
      }
      path.closePath();
    } else {
      path = shapePath(a, b, o.shape, { square: this.square, fromCenter: this.fromCenter, radius: o.radius }).path;
    }
    if (o.fill) ctx.fill(path);
    if (o.stroke && o.strokeWidth > 0) ctx.stroke(path);
  }

  prepareRender() {
    if (!this.active) return;
    const d = doc();
    const pv = scratch('preview', d.width, d.height);
    const ctx = pv.getContext('2d');
    ctx.drawImage(this.layer.canvas, 0, 0);
    const sh = scratch('shape', d.width, d.height);
    this.draw(sh.getContext('2d'));
    ctx.drawImage(maskWithSelection(sh), 0, 0);
    this.layer.preview = pv;
  }

  up() {
    if (!this.active) return;
    this.active = false;
    const d = doc();
    this.layer.preview = null;
    if (Math.hypot(this.b.x - this.a.x, this.b.y - this.a.y) < 1) {
      if (this.created) {
        d.layers.splice(d.layers.indexOf(this.layer), 1);
        d.active = Math.max(0, d.active - 1);
        app.emit('layers');
      }
      render();
      return;
    }
    const sh = scratch('shape', d.width, d.height);
    this.draw(sh.getContext('2d'));
    const masked = maskWithSelection(sh);
    d.editLayer(this.layer, (ctx) => ctx.drawImage(masked, 0, 0));
    d.commit('Forma');
    render();
  }
}

// ---------------------------------------------------------------------------
// Texto

export const FONTS = ['Arial', 'Helvetica', 'Segoe UI', 'Roboto', 'Verdana', 'Tahoma', 'Trebuchet MS', 'Georgia', 'Times New Roman', 'Garamond', 'Courier New', 'Consolas', 'Impact', 'Comic Sans MS', 'Brush Script MT', 'system-ui', 'serif', 'sans-serif', 'monospace', 'cursive'];

export function renderText(meta, w, h) {
  const c = createCanvas(w, h);
  const ctx = c.getContext('2d');
  ctx.font = `${meta.italic ? 'italic ' : ''}${meta.bold ? 'bold ' : ''}${meta.size}px "${meta.font}"`;
  if ('letterSpacing' in ctx) ctx.letterSpacing = `${meta.tracking || 0}px`;
  ctx.textBaseline = 'top';
  ctx.textAlign = meta.align;
  ctx.fillStyle = meta.color;
  const lines = meta.text.split('\n');
  const lh = meta.size * meta.lineHeight;
  let maxW = 0;
  lines.forEach((line, i) => {
    ctx.fillText(line, meta.x, meta.y + i * lh);
    maxW = Math.max(maxW, ctx.measureText(line).width);
  });
  if (meta.underline) {
    ctx.fillRect(meta.align === 'center' ? meta.x - maxW / 2 : meta.align === 'right' ? meta.x - maxW : meta.x, meta.y + lines.length * lh - lh * 0.12, maxW, Math.max(1, meta.size / 16));
  }
  const bx = meta.align === 'center' ? meta.x - maxW / 2 : meta.align === 'right' ? meta.x - maxW : meta.x;
  meta.bounds = { x: bx, y: meta.y, w: maxW, h: lines.length * lh };
  return c;
}

export async function editText(layer, isNew) {
  const d = doc();
  const meta = { ...layer.text };
  const ta = el('textarea', { rows: 4, class: 'text-input' });
  ta.value = meta.text;
  const font = el('select', {}, ...FONTS.map((f) => el('option', { value: f, selected: f === meta.font }, f)));
  const size = el('input', { type: 'number', min: 1, max: 2000, value: meta.size });
  const color = el('input', { type: 'color', value: meta.color });
  const bold = el('input', { type: 'checkbox', checked: meta.bold });
  const italic = el('input', { type: 'checkbox', checked: meta.italic });
  const underline = el('input', { type: 'checkbox', checked: meta.underline });
  const align = el('select', {}, ...[['left', 'Esquerda'], ['center', 'Centro'], ['right', 'Direita']].map(([v, l]) => el('option', { value: v, selected: v === meta.align }, l)));
  const lh = el('input', { type: 'number', min: 0.5, max: 4, step: 0.05, value: meta.lineHeight });
  const tracking = el('input', { type: 'number', min: -50, max: 200, step: 1, value: meta.tracking || 0 });
  const update = () => {
    Object.assign(meta, {
      text: ta.value, font: font.value, size: +size.value || 12, color: color.value, bold: bold.checked,
      italic: italic.checked, underline: underline.checked, align: align.value, lineHeight: +lh.value || 1.2, tracking: +tracking.value || 0,
    });
    layer.preview = renderText(meta, d.width, d.height);
    render();
  };
  [ta, font, size, color, bold, italic, underline, align, lh, tracking].forEach((i) => i.addEventListener('input', update));
  update();
  const body = el('div', { class: 'text-dialog' },
    ta,
    el('div', { class: 'grid2' },
      field('Fonte', font), field('Tamanho (px)', size),
      field('Cor', color), field('Alinhamento', align),
      field('Entrelinha', lh), field('Espaçamento', tracking)),
    el('div', { class: 'row' },
      el('label', { class: 'check' }, bold, ' Negrito'),
      el('label', { class: 'check' }, italic, ' Itálico'),
      el('label', { class: 'check' }, underline, ' Sublinhado')));
  const ok = await dialog({ title: isNew ? 'Novo texto' : 'Editar texto', body, width: 460 });
  layer.preview = null;
  if (ok && meta.text.trim()) {
    d.replaceLayerCanvas(layer, renderText(meta, d.width, d.height), true);
    layer.text = meta;
    layer.name = meta.text.split('\n')[0].slice(0, 32) || 'Texto';
    d.commit(isNew ? 'Texto' : 'Editar texto');
  } else if (isNew) {
    d.layers.splice(d.layers.indexOf(layer), 1);
    d.active = clamp(d.active - 1, 0, d.layers.length - 1);
  }
  app.emit('layers');
  render();
}

class TextTool {
  constructor(cfg) { Object.assign(this, { cursor: 'text' }, cfg); }
  down(e, p) {
    const d = doc();
    const al = d.activeLayer;
    if (al?.text?.bounds) {
      const b = al.text.bounds;
      if (p.x >= b.x - 4 && p.x <= b.x + b.w + 4 && p.y >= b.y - 4 && p.y <= b.y + b.h + 4) {
        editText(al, false);
        return;
      }
    }
    const o = this.opts;
    const layer = d.addLayer('Texto');
    layer.text = {
      text: 'Seu texto aqui', x: Math.round(p.x), y: Math.round(p.y), font: o.font, size: o.size, color: app.fg,
      bold: o.bold, italic: o.italic, underline: false, align: o.align, lineHeight: 1.2, tracking: 0,
    };
    app.emit('layers');
    editText(layer, true);
  }
}

// ---------------------------------------------------------------------------
// Navegação

class HandTool {
  constructor(cfg) { Object.assign(this, { cursor: 'grab' }, cfg); }
  down(e, p, s) { this.last = s; app.view.canvas.style.cursor = 'grabbing'; }
  move(e, p, s) {
    if (!this.last) return;
    app.view.pan(s.x - this.last.x, s.y - this.last.y);
    this.last = s;
  }
  up() { this.last = null; app.view.canvas.style.cursor = this.cursor; }
}

class ZoomTool {
  constructor(cfg) { Object.assign(this, { cursor: 'zoom-in' }, cfg); }
  down(e, p, s) { app.view.zoomStep(e.altKey || this.opts.out ? -1 : 1, s.x, s.y); }
}

class EyedropperTool {
  constructor(cfg) { Object.assign(this, { cursor: 'crosshair' }, cfg); }
  down(e, p) { this.active = true; this.toBg = e.altKey; pick(p, this.toBg, +this.opts.sample); }
  move(e, p) { if (this.active) pick(p, this.toBg, +this.opts.sample); }
  up() { this.active = false; }
}

// ---------------------------------------------------------------------------
// Registro das ferramentas

const I = {
  move: '<path d="M12 2v20M2 12h20M9 5l3-3 3 3M9 19l3 3 3-3M5 9l-3 3 3 3M19 9l3 3-3 3"/>',
  marquee: '<rect x="4" y="5" width="16" height="14" stroke-dasharray="3 2.2"/>',
  ellipse: '<ellipse cx="12" cy="12" rx="9" ry="7" stroke-dasharray="3 2.2"/>',
  lasso: '<path d="M7 17.5C4.5 16.5 3 14.6 3 12.5 3 8.4 7 5 12 5s9 3.4 9 7.5S17 20 12 20c-1.3 0-2.6-.2-3.7-.6"/><circle cx="7" cy="18" r="1.6"/><path d="M7 19.6c0 1.3.8 2.4 2 2.4"/>',
  wand: '<path d="m3 21 12-12"/><path d="M15 3v2.5M15 12.5V15M21 9h-2.5M11.5 9H9M19.2 4.8l-1.8 1.8M19.2 13.2l-1.8-1.8M10.8 4.8l1.8 1.8"/>',
  crop: '<path d="M6 2v14a2 2 0 0 0 2 2h14"/><path d="M18 22V8a2 2 0 0 0-2-2H2"/>',
  eyedropper: '<path d="m2 22 1-1h3l9-9"/><path d="M3 21v-3l9-9"/><path d="m15 6 3.4-3.4a2.1 2.1 0 1 1 3 3L18 9l.4.4a2.1 2.1 0 1 1-3 3l-3.8-3.8a2.1 2.1 0 1 1 3-3l.4.4Z"/>',
  brush: '<path d="m9.06 11.9 8.07-8.06a2.85 2.85 0 1 1 4.03 4.03l-8.06 8.08"/><path d="M7.07 14.94c-1.66 0-3 1.35-3 3.02 0 1.33-2.5 1.52-2 2.02 1.08 1.1 2.49 2.02 4 2.02 2.2 0 4-1.8 4-4.04a3.01 3.01 0 0 0-3-3.02z"/>',
  pencil: '<path d="M17 3a2.85 2.83 0 1 1 4 4L7.5 20.5 2 22l1.5-5.5Z"/><path d="m15 5 4 4"/>',
  eraser: '<path d="m7 21-4.3-4.3c-1-1-1-2.5 0-3.4l9.6-9.6c1-1 2.5-1 3.4 0l5.6 5.6c1 1 1 2.5 0 3.4L13 21"/><path d="M22 21H7"/><path d="m5 11 9 9"/>',
  clone: '<path d="M5 22h14"/><path d="M19.27 13.73A2.5 2.5 0 0 0 17.5 13h-11A2.5 2.5 0 0 0 4 15.5V17a1 1 0 0 0 1 1h14a1 1 0 0 0 1-1v-1.5c0-.66-.26-1.3-.73-1.77Z"/><path d="M14 13V8.5C14 7 15 7 15 5a3 3 0 0 0-3-3c-1.69 0-3 1-3 3 0 2 1 2 1 3.5V13"/>',
  bucket: '<path d="m19 11-8-8-8.6 8.6a2 2 0 0 0 0 2.8l5.2 5.2c.8.8 2 .8 2.8 0L19 11Z"/><path d="m5 2 5 5"/><path d="M2 13h15"/><path d="M22 20a2 2 0 1 1-4 0c0-1.6 1.7-2.4 2-4 .3 1.6 2 2.4 2 4Z"/>',
  gradient: '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M7 3v18" opacity=".9"/><path d="M11 3v18" opacity=".6"/><path d="M15 3v18" opacity=".35"/>',
  focus: '<path d="M12 22a7 7 0 0 0 7-7c0-2-1-3.9-3-5.5s-3.5-4-4-6.5c-.5 2.5-2 4.9-4 6.5C6 11.1 5 13 5 15a7 7 0 0 0 7 7z"/>',
  tone: '<circle cx="12" cy="9" r="6"/><path d="M12 15v7"/>',
  text: '<path d="M4 7V4h16v3M9 20h6M12 4v16"/>',
  shape: '<rect x="3" y="3" width="11" height="11" rx="1"/><circle cx="15.5" cy="15.5" r="5.5"/>',
  hand: '<path d="M18 11V6a2 2 0 0 0-4 0M14 10V4a2 2 0 0 0-4 0v2M10 10.5V6a2 2 0 0 0-4 0v8"/><path d="M18 8a2 2 0 1 1 4 0v6a8 8 0 0 1-8 8h-2c-2.8 0-4.5-.86-5.99-2.34l-3.6-3.6a2 2 0 0 1 2.83-2.82L7 15"/>',
  zoom: '<circle cx="11" cy="11" r="8"/><path d="m21 21-4.3-4.3M11 8v6M8 11h6"/>',
  transform: '<path d="M4 4h4M4 4v4M20 4h-4M20 4v4M4 20h4M4 20v-4M20 20h-4M20 20v-4"/><rect x="8" y="8" width="8" height="8"/>',
};

const defs = [
  new MoveTool({ id: 'move', name: 'Mover', key: 'v', icon: I.move, options: [
    { key: 'autoSelect', label: 'Selecionar camada automaticamente', type: 'checkbox', value: false },
  ], hint: 'Arraste para mover a camada ou a seleção. Setas deslocam 1px (Shift: 10px).' }),
  new MarqueeTool({ id: 'marquee', name: 'Letreiro retangular', key: 'm', kind: 'rect', icon: I.marquee, options: selectionOptions([
    { key: 'feather', label: 'Difusão', type: 'range', min: 0, max: 200, value: 0, unit: 'px' },
  ]), hint: 'Shift: adicionar/quadrado · Alt: subtrair/do centro · arraste dentro para mover a seleção.' }),
  new MarqueeTool({ id: 'ellipse', name: 'Letreiro elíptico', key: 'm', kind: 'ellipse', icon: I.ellipse, options: selectionOptions([
    { key: 'feather', label: 'Difusão', type: 'range', min: 0, max: 200, value: 0, unit: 'px' },
  ]), hint: 'Shift: adicionar/círculo · Alt: subtrair/do centro.' }),
  new LassoTool({ id: 'lasso', name: 'Laço', key: 'l', icon: I.lasso, options: selectionOptions([
    { key: 'poly', label: 'Poligonal', type: 'checkbox', value: false },
    { key: 'feather', label: 'Difusão', type: 'range', min: 0, max: 200, value: 0, unit: 'px' },
  ]), hint: 'Desenhe à mão livre. No modo poligonal: clique para pontos, duplo clique ou Enter para fechar.' }),
  new WandTool({ id: 'wand', name: 'Varinha mágica', key: 'w', icon: I.wand, options: selectionOptions([
    { key: 'tolerance', label: 'Tolerância', type: 'range', min: 0, max: 100, value: 32 },
    { key: 'contiguous', label: 'Contíguo', type: 'checkbox', value: true },
    { key: 'antialias', label: 'Suavização', type: 'checkbox', value: true },
    { key: 'sampleAll', label: 'Todas as camadas', type: 'checkbox', value: false },
  ]), hint: 'Clique para selecionar pixels de cor semelhante.' }),
  new CropTool({ id: 'crop', name: 'Cortar', key: 'c', icon: I.crop, options: [
    { key: 'ratio', label: 'Proporção', type: 'select', value: 'free', choices: [['free', 'Livre'], ['orig', 'Original'], ['1:1', '1:1 (quadrado)'], ['4:3', '4:3'], ['3:2', '3:2'], ['16:9', '16:9'], ['9:16', '9:16 (stories)'], ['4:5', '4:5 (retrato)']] },
    { label: 'Aplicar', type: 'action', action: (t) => t.apply() },
  ], hint: 'Arraste para definir a área. Enter ou duplo clique aplica, Esc cancela.' }),
  new EyedropperTool({ id: 'eyedropper', name: 'Conta-gotas', key: 'i', icon: I.eyedropper, options: [
    { key: 'sample', label: 'Amostra', type: 'select', value: '1', choices: [['1', 'Ponto'], ['3', 'Média 3×3'], ['5', 'Média 5×5'], ['11', 'Média 11×11']] },
  ], hint: 'Clique para cor de primeiro plano. Alt+clique para cor de fundo.' }),
  new PaintTool({ id: 'brush', name: 'Pincel', key: 'b', icon: I.brush, options: brushOptions([
    { key: 'mode', label: 'Modo', type: 'select', value: 'source-over', choices: blendChoices },
  ]), hint: 'Shift+clique traça linha reta · Alt: conta-gotas · [ ] tamanho · 1–0 opacidade.' }),
  new PaintTool({ id: 'pencil', name: 'Lápis', key: 'n', icon: I.pencil, spacing: 0.05, options: [
    { key: 'size', label: 'Tamanho', type: 'range', min: 1, max: 200, value: 1, unit: 'px' },
    { key: 'hardness', type: 'hidden', value: 100 },
    { key: 'opacity', label: 'Opacidade', type: 'range', min: 1, max: 100, value: 100, unit: '%' },
    { key: 'mode', label: 'Modo', type: 'select', value: 'source-over', choices: blendChoices },
  ], hint: 'Traço duro, ideal para pixel art.' }),
  new PaintTool({ id: 'eraser', name: 'Borracha', key: 'e', icon: I.eraser, op: 'destination-out', altPicks: false, options: brushOptions(), hint: 'Apaga pixels da camada ativa.' }),
  new CloneTool({ id: 'clone', name: 'Carimbo', key: 's', icon: I.clone, options: brushOptions([
    { key: 'aligned', label: 'Alinhado', type: 'checkbox', value: true },
    { key: 'sampleAll', label: 'Todas as camadas', type: 'checkbox', value: false },
  ]), hint: 'Alt+clique define a origem; depois pinte para clonar.' }),
  new FocusTool({ id: 'focus', name: 'Desfoque / Nitidez', key: 'r', icon: I.focus, altPicks: false, options: [
    { key: 'kind', label: 'Tipo', type: 'buttons', value: 'blur', choices: [['blur', 'Desfocar'], ['sharpen', 'Nitidez']] },
    { key: 'size', label: 'Tamanho', type: 'range', min: 1, max: 1000, value: 50, unit: 'px', log: true },
    { key: 'hardness', label: 'Dureza', type: 'range', min: 0, max: 100, value: 0, unit: '%' },
    { key: 'strength', label: 'Intensidade', type: 'range', min: 1, max: 100, value: 50, unit: '%' },
    { key: 'opacity', type: 'hidden', value: 100 },
  ], hint: 'Pinte para desfocar ou dar nitidez a áreas.' }),
  new ToneTool({ id: 'tone', name: 'Subexposição / Superexposição / Esponja', key: 'o', icon: I.tone, altPicks: false, options: [
    { key: 'kind', label: 'Tipo', type: 'buttons', value: 'dodge', choices: [['dodge', 'Clarear'], ['burn', 'Escurecer'], ['sponge-', 'Dessaturar'], ['sponge+', 'Saturar']] },
    { key: 'size', label: 'Tamanho', type: 'range', min: 1, max: 1000, value: 60, unit: 'px', log: true },
    { key: 'hardness', label: 'Dureza', type: 'range', min: 0, max: 100, value: 0, unit: '%' },
    { key: 'strength', label: 'Exposição', type: 'range', min: 1, max: 100, value: 40, unit: '%' },
    { key: 'opacity', label: 'Opacidade', type: 'range', min: 1, max: 100, value: 60, unit: '%' },
  ], hint: 'Clareie, escureça ou ajuste a saturação pintando.' }),
  new BucketTool({ id: 'bucket', name: 'Balde de tinta', key: 'g', icon: I.bucket, options: [
    { key: 'tolerance', label: 'Tolerância', type: 'range', min: 0, max: 100, value: 32 },
    { key: 'opacity', label: 'Opacidade', type: 'range', min: 1, max: 100, value: 100, unit: '%' },
    { key: 'mode', label: 'Modo', type: 'select', value: 'source-over', choices: blendChoices },
    { key: 'contiguous', label: 'Contíguo', type: 'checkbox', value: true },
    { key: 'antialias', label: 'Suavização', type: 'checkbox', value: true },
    { key: 'sampleAll', label: 'Todas as camadas', type: 'checkbox', value: false },
  ], hint: 'Preenche áreas de cor semelhante com a cor de primeiro plano.' }),
  new GradientTool({ id: 'gradient', name: 'Degradê', key: 'g', icon: I.gradient, options: [
    { key: 'kind', label: 'Tipo', type: 'buttons', value: 'linear', choices: [['linear', 'Linear'], ['radial', 'Radial'], ['angle', 'Angular'], ['reflected', 'Refletido']] },
    { key: 'colors', label: 'Cores', type: 'select', value: 'fg-bg', choices: [['fg-bg', 'Frente → Fundo'], ['fg-transparent', 'Frente → Transparente'], ['black-white', 'Preto → Branco']] },
    { key: 'reverse', label: 'Inverter', type: 'checkbox', value: false },
    { key: 'opacity', label: 'Opacidade', type: 'range', min: 1, max: 100, value: 100, unit: '%' },
    { key: 'mode', label: 'Modo', type: 'select', value: 'source-over', choices: blendChoices },
  ], hint: 'Arraste para criar o degradê. Shift restringe a 45°.' }),
  new TextTool({ id: 'text', name: 'Texto', key: 't', icon: I.text, options: [
    { key: 'font', label: 'Fonte', type: 'select', value: 'Arial', choices: FONTS.map((f) => [f, f]) },
    { key: 'size', label: 'Tamanho', type: 'range', min: 6, max: 400, value: 48, unit: 'px', log: true },
    { key: 'bold', label: 'Negrito', type: 'checkbox', value: false },
    { key: 'italic', label: 'Itálico', type: 'checkbox', value: false },
    { key: 'align', label: 'Alinhar', type: 'buttons', value: 'left', choices: [['left', 'Esq.'], ['center', 'Centro'], ['right', 'Dir.']] },
  ], hint: 'Clique para criar texto. Clique num texto existente (camada ativa) para editá-lo.' }),
  new ShapeTool({ id: 'shape', name: 'Formas', key: 'u', icon: I.shape, options: [
    { key: 'shape', label: 'Forma', type: 'select', value: 'rect', choices: [['rect', 'Retângulo'], ['rounded', 'Retângulo arredondado'], ['ellipse', 'Elipse'], ['polygon', 'Polígono'], ['star', 'Estrela'], ['line', 'Linha'], ['arrow', 'Seta']] },
    { key: 'fill', label: 'Preencher', type: 'checkbox', value: true },
    { key: 'stroke', label: 'Contorno', type: 'checkbox', value: false },
    { key: 'strokeColor', label: 'Cor do contorno', type: 'color', value: '#000000' },
    { key: 'strokeWidth', label: 'Espessura', type: 'range', min: 0, max: 100, value: 4, unit: 'px' },
    { key: 'radius', label: 'Raio', type: 'range', min: 0, max: 200, value: 16, unit: 'px' },
    { key: 'sides', label: 'Lados/Pontas', type: 'range', min: 3, max: 24, value: 5 },
    { key: 'newLayer', label: 'Nova camada', type: 'checkbox', value: true },
  ], hint: 'Arraste para desenhar. Shift: proporção · Alt: a partir do centro. Preenchimento usa a cor de frente.' }),
  new HandTool({ id: 'hand', name: 'Mão', key: 'h', icon: I.hand, options: [
    { label: 'Ajustar à tela', type: 'action', action: () => app.view.fit() },
    { label: '100%', type: 'action', action: () => { app.view.setZoom(1); app.view.center(); } },
  ], hint: 'Arraste para navegar. Segure Espaço com qualquer ferramenta.' }),
  new ZoomTool({ id: 'zoom', name: 'Zoom', key: 'z', icon: I.zoom, options: [
    { key: 'out', label: 'Reduzir', type: 'checkbox', value: false },
    { label: 'Ajustar à tela', type: 'action', action: () => app.view.fit() },
    { label: '100%', type: 'action', action: () => { app.view.setZoom(1); app.view.center(); } },
  ], hint: 'Clique para ampliar · Alt+clique para reduzir · Roda do mouse também dá zoom.' }),
];

export const transformTool = new TransformTool({ id: 'transform', name: 'Transformação livre', key: null, icon: I.transform, hidden: true, options: [
  { type: 'info', get: (t) => (t.running ? t.info() : '') },
  { label: 'Espelhar H', type: 'action', action: (t) => t.flip(true) },
  { label: 'Espelhar V', type: 'action', action: (t) => t.flip(false) },
  { label: '✓ Aplicar', type: 'action', action: (t) => t.commit() },
  { label: '✕ Cancelar', type: 'action', action: (t) => t.cancel() },
], hint: 'Arraste alças para escalar (Shift: proporcional, Alt: do centro), fora da caixa para girar. Enter aplica.' });

export const TOOLS = [...defs, transformTool];
export const TOOLBAR_GROUPS = [
  ['move'], ['marquee', 'ellipse', 'lasso', 'wand'], ['crop', 'eyedropper'],
  ['brush', 'pencil', 'eraser', 'clone'], ['focus', 'tone'], ['bucket', 'gradient'], ['text', 'shape'], ['hand', 'zoom'],
];

const STORAGE_KEY = 'lumina.toolOptions.v1';
export function initToolOptions() {
  let saved = {};
  try { saved = JSON.parse(localStorage.getItem(STORAGE_KEY) || '{}'); } catch { saved = {}; }
  for (const t of TOOLS) {
    t.opts = {};
    for (const o of t.options || []) if (o.key) t.opts[o.key] = saved[t.id]?.[o.key] ?? o.value;
    app.tools[t.id] = t;
  }
}

export function saveToolOptions() {
  try {
    const out = {};
    for (const t of TOOLS) out[t.id] = t.opts;
    localStorage.setItem(STORAGE_KEY, JSON.stringify(out));
  } catch { /* armazenamento indisponível */ }
}

