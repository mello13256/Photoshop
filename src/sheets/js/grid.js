// Grade da planilha desenhada em canvas, com seleção, edição na célula e interação por mouse.
import { KEY, RC } from './model.js';
import { display, editText } from './format.js';
import { isErr, isMatrix, colName } from './values.js';
import { formulaRefs } from './formula.js';
import { FUNCS } from './functions.js';

export const RH_W = 46; // largura do cabeçalho de linhas
export const CH_H = 24; // altura do cabeçalho de colunas
const DEFAULT_FONT = 'Calibri, Carlito, "Segoe UI", Arial, sans-serif';
export const REF_COLORS = ['#2f6bff', '#e8590c', '#2b8a3e', '#ae3ec9', '#d6336c', '#0c8599', '#f08c00', '#5c7cfa'];

export class Grid {
  constructor(root, app) {
    this.app = app;
    this.root = root;
    this.zoom = 1;
    this.sel = { r: 0, c: 0, r1: 0, c1: 0, r2: 0, c2: 0 };
    this.copyRange = null;
    this.refHighlights = [];
    this.scroller = document.createElement('div');
    this.scroller.className = 'grid-scroller';
    this.spacer = document.createElement('div');
    this.spacer.className = 'grid-spacer';
    this.canvas = document.createElement('canvas');
    this.canvas.className = 'grid-canvas';
    this.overlay = document.createElement('div');
    this.overlay.className = 'grid-overlay';
    this.scroller.append(this.spacer);
    root.append(this.canvas, this.scroller, this.overlay);
    this.ctx = this.canvas.getContext('2d');
    this.dirty = true;
    this.antsOffset = 0;
    new ResizeObserver(() => this.resize()).observe(root);
    this.scroller.addEventListener('scroll', () => { this.extendIfNeeded(); this.invalidate(); this.app.onScroll?.(); }, { passive: true });
    this.scroller.addEventListener('mousedown', (e) => this.onMouseDown(e));
    this.scroller.addEventListener('mousemove', (e) => this.onHover(e));
    this.scroller.addEventListener('dblclick', (e) => this.onDblClick(e));
    this.scroller.addEventListener('contextmenu', (e) => { e.preventDefault(); this.app.onContextMenu?.(e, this.hit(...this.local(e))); });
    this.scroller.addEventListener('wheel', (e) => {
      if (!e.ctrlKey) return;
      e.preventDefault();
      this.setZoom(this.zoom * (e.deltaY < 0 ? 1.1 : 1 / 1.1));
    }, { passive: false });
    const loop = () => {
      if (this.dirty) this.render();
      requestAnimationFrame(loop);
    };
    requestAnimationFrame(loop);
    setInterval(() => { if (this.copyRange) { this.antsOffset = (this.antsOffset + 1) % 8; this.invalidate(); } }, 100);
    this.readTheme();
  }

  get sheet() { return this.app.wb.sheet; }
  get wb() { return this.app.wb; }

  readTheme() {
    const cs = getComputedStyle(document.documentElement);
    const v = (n, d) => cs.getPropertyValue(n).trim() || d;
    this.colors = {
      bg: v('--grid-bg', '#ffffff'), line: v('--grid-line', '#e2e5ea'), head: v('--grid-head', '#f6f7f9'), headText: v('--grid-head-text', '#5d6472'),
      headSel: v('--grid-head-sel', '#dbe6ff'), headSelText: v('--accent', '#3b6cf6'), text: v('--grid-text', '#1d2129'), accent: v('--accent', '#3b6cf6'),
      selFill: v('--grid-sel-fill', 'rgba(59,108,246,.10)'), frozen: v('--grid-frozen', '#b9c0cc'),
    };
    this.invalidate();
  }

  invalidate() { this.dirty = true; }

  setZoom(z) {
    this.zoom = Math.min(3, Math.max(0.4, z));
    this.resize();
    this.app.onZoom?.(this.zoom);
  }

  resize() {
    const r = this.root.getBoundingClientRect();
    const dpr = devicePixelRatio || 1;
    this.w = r.width;
    this.h = r.height;
    this.canvas.width = Math.round(r.width * dpr);
    this.canvas.height = Math.round(r.height * dpr);
    this.canvas.style.width = `${r.width}px`;
    this.canvas.style.height = `${r.height}px`;
    this.updateSpacer();
    this.invalidate();
  }

  updateSpacer() {
    const s = this.sheet;
    this.spacer.style.width = `${RH_W + s.cols.total() * this.zoom + 40}px`;
    this.spacer.style.height = `${CH_H + s.rows.total() * this.zoom + 40}px`;
  }

  // Aumenta a grade ao rolar perto do fim.
  extendIfNeeded() {
    const s = this.sheet;
    const sc = this.scroller;
    let changed = false;
    if (sc.scrollTop + sc.clientHeight > sc.scrollHeight - 400 && s.rows.count < 1048576) { s.rows.count = Math.min(1048576, s.rows.count + 500); changed = true; }
    if (sc.scrollLeft + sc.clientWidth > sc.scrollWidth - 300 && s.cols.count < 16384) { s.cols.count = Math.min(16384, s.cols.count + 10); changed = true; }
    if (changed) { s.rows.dirty = s.cols.dirty = true; this.updateSpacer(); }
  }

  // ---- Geometria -----------------------------------------------------------------

  frozenW() { return this.sheet.cols.pos(this.sheet.freeze.c) * this.zoom; }
  frozenH() { return this.sheet.rows.pos(this.sheet.freeze.r) * this.zoom; }

  colX(c) {
    const s = this.sheet, z = this.zoom, fc = s.freeze.c;
    if (c < fc) return RH_W + s.cols.pos(c) * z;
    return RH_W + this.frozenW() + (s.cols.pos(c) - s.cols.pos(fc)) * z - this.scroller.scrollLeft;
  }

  rowY(r) {
    const s = this.sheet, z = this.zoom, fr = s.freeze.r;
    if (r < fr) return CH_H + s.rows.pos(r) * z;
    return CH_H + this.frozenH() + (s.rows.pos(r) - s.rows.pos(fr)) * z - this.scroller.scrollTop;
  }

  cellRect(r, c) {
    const s = this.sheet;
    const m = s.mergeAt(r, c);
    if (m) {
      const x = this.colX(m.c1), y = this.rowY(m.r1);
      return { x, y, w: this.colX(m.c2 + 1) - x, h: this.rowY(m.r2 + 1) - y, m };
    }
    const x = this.colX(c), y = this.rowY(r);
    return { x, y, w: s.cols.size(c) * this.zoom, h: s.rows.size(r) * this.zoom };
  }

  rangeRect(rg) {
    const x = this.colX(rg.c1), y = this.rowY(rg.r1);
    return { x, y, w: this.colX(rg.c2 + 1) - x, h: this.rowY(rg.r2 + 1) - y };
  }

  local(e) {
    const r = this.root.getBoundingClientRect();
    return [e.clientX - r.left, e.clientY - r.top];
  }

  colAt(x) {
    const s = this.sheet, z = this.zoom, fw = this.frozenW();
    if (x - RH_W < fw) return s.cols.indexAt((x - RH_W) / z);
    return s.cols.indexAt(s.cols.pos(s.freeze.c) + (x - RH_W - fw + this.scroller.scrollLeft) / z);
  }

  rowAt(y) {
    const s = this.sheet, z = this.zoom, fh = this.frozenH();
    if (y - CH_H < fh) return s.rows.indexAt((y - CH_H) / z);
    return s.rows.indexAt(s.rows.pos(s.freeze.r) + (y - CH_H - fh + this.scroller.scrollTop) / z);
  }

  hit(x, y) {
    const out = { x, y };
    if (x < RH_W && y < CH_H) return { ...out, area: 'corner' };
    const c = this.colAt(Math.max(x, RH_W));
    const r = this.rowAt(Math.max(y, CH_H));
    if (y < CH_H) {
      // borda para redimensionar
      const right = this.colX(c + 1), left = this.colX(c);
      if (Math.abs(x - right) <= 4) return { ...out, area: 'colresize', c };
      if (Math.abs(x - left) <= 4 && c > 0) return { ...out, area: 'colresize', c: c - 1 };
      return { ...out, area: 'colhead', c };
    }
    if (x < RH_W) {
      const bottom = this.rowY(r + 1), top = this.rowY(r);
      if (Math.abs(y - bottom) <= 3) return { ...out, area: 'rowresize', r };
      if (Math.abs(y - top) <= 3 && r > 0) return { ...out, area: 'rowresize', r: r - 1 };
      return { ...out, area: 'rowhead', r };
    }
    const rr = this.rangeRect(this.sel);
    if (Math.abs(x - (rr.x + rr.w)) <= 5 && Math.abs(y - (rr.y + rr.h)) <= 5) return { ...out, area: 'fill', r, c };
    return { ...out, area: 'cell', r, c };
  }

  scrollIntoView(r, c) {
    const s = this.sheet, z = this.zoom, sc = this.scroller;
    if (c >= s.freeze.c) {
      const x0 = (s.cols.pos(c) - s.cols.pos(s.freeze.c)) * z, x1 = x0 + s.cols.size(c) * z;
      const vw = this.w - RH_W - this.frozenW() - 14;
      if (x0 < sc.scrollLeft) sc.scrollLeft = x0;
      else if (x1 > sc.scrollLeft + vw) sc.scrollLeft = x1 - vw;
    }
    if (r >= s.freeze.r) {
      const y0 = (s.rows.pos(r) - s.rows.pos(s.freeze.r)) * z, y1 = y0 + s.rows.size(r) * z;
      const vh = this.h - CH_H - this.frozenH() - 14;
      if (y0 < sc.scrollTop) sc.scrollTop = y0;
      else if (y1 > sc.scrollTop + vh) sc.scrollTop = y1 - vh;
    }
    this.invalidate();
  }

  // ---- Seleção -----------------------------------------------------------------

  select(r, c, extend = false) {
    const s = this.sheet;
    r = Math.max(0, Math.min(r, s.rows.count - 1));
    c = Math.max(0, Math.min(c, s.cols.count - 1));
    if (extend) {
      this.sel = { ...this.sel, ...norm(this.sel.r, this.sel.c, r, c) };
      this.expandMerges();
      this.selEnd = { r, c };
    } else {
      const m = s.mergeAt(r, c);
      this.sel = m ? { r: m.r1, c: m.c1, r1: m.r1, c1: m.c1, r2: m.r2, c2: m.c2 } : { r, c, r1: r, c1: c, r2: r, c2: c };
      this.selEnd = { r, c };
    }
    this.invalidate();
    this.app.onSelect?.();
  }

  setRange(rg, active = { r: rg.r1, c: rg.c1 }) {
    this.sel = { r: active.r, c: active.c, r1: rg.r1, c1: rg.c1, r2: rg.r2, c2: rg.c2 };
    this.selEnd = { r: rg.r2, c: rg.c2 };
    this.invalidate();
    this.app.onSelect?.();
  }

  expandMerges() {
    const s = this.sheet;
    let changed = true;
    while (changed) {
      changed = false;
      for (const m of s.merges) {
        const sel = this.sel;
        if (m.r2 < sel.r1 || m.r1 > sel.r2 || m.c2 < sel.c1 || m.c1 > sel.c2) continue;
        if (m.r1 < sel.r1 || m.r2 > sel.r2 || m.c1 < sel.c1 || m.c2 > sel.c2) {
          Object.assign(sel, { r1: Math.min(sel.r1, m.r1), r2: Math.max(sel.r2, m.r2), c1: Math.min(sel.c1, m.c1), c2: Math.max(sel.c2, m.c2) });
          changed = true;
        }
      }
    }
  }

  // ---- Mouse -----------------------------------------------------------------

  onHover(e) {
    if (this.dragging) return;
    const h = this.hit(...this.local(e));
    const cur = h.area === 'colresize' ? 'col-resize' : h.area === 'rowresize' ? 'row-resize' : h.area === 'fill' ? 'crosshair'
      : h.area === 'colhead' ? 's-resize' : h.area === 'rowhead' ? 'e-resize' : 'cell';
    this.scroller.style.cursor = cur;
  }

  onMouseDown(e) {
    if (e.button !== 0) {
      if (e.button === 2) {
        const h = this.hit(...this.local(e));
        if (h.area === 'cell' && !inRange(this.sel, h.r, h.c)) this.select(h.r, h.c);
      }
      return;
    }
    const [x, y] = this.local(e);
    // Clique na barra de rolagem nativa
    if (x > this.scroller.clientWidth || y > this.scroller.clientHeight) return;
    const h = this.hit(x, y);
    if (this.app.editor?.active && this.app.editor.pointMode() && (h.area === 'cell' || h.area === 'fill')) {
      e.preventDefault();
      this.app.editor.insertRef(h.r, h.c, h.r, h.c, true);
      this.drag({ kind: 'ref', r0: h.r, c0: h.c });
      return;
    }
    if (this.app.editor?.active) this.app.editor.commit();
    e.preventDefault();
    this.app.focusGrid?.();
    switch (h.area) {
      case 'corner':
        this.setRange({ r1: 0, c1: 0, r2: this.sheet.rows.count - 1, c2: this.sheet.cols.count - 1 }, { r: this.sel.r, c: this.sel.c });
        break;
      case 'colhead':
        if (e.shiftKey) this.setRange({ r1: 0, r2: this.sheet.rows.count - 1, c1: Math.min(this.sel.c, h.c), c2: Math.max(this.sel.c, h.c) }, { r: this.sel.r, c: this.sel.c });
        else this.setRange({ r1: 0, r2: this.sheet.rows.count - 1, c1: h.c, c2: h.c }, { r: Math.max(0, this.rowAt(CH_H + this.frozenH() + 1)), c: h.c });
        this.drag({ kind: 'cols', c0: h.c });
        break;
      case 'rowhead':
        if (e.shiftKey) this.setRange({ c1: 0, c2: this.sheet.cols.count - 1, r1: Math.min(this.sel.r, h.r), r2: Math.max(this.sel.r, h.r) }, { r: this.sel.r, c: this.sel.c });
        else this.setRange({ c1: 0, c2: this.sheet.cols.count - 1, r1: h.r, r2: h.r }, { r: h.r, c: 0 });
        this.drag({ kind: 'rows', r0: h.r });
        break;
      case 'colresize':
        this.drag({ kind: 'colresize', c: h.c, start: x, size: this.sheet.cols.size(h.c) });
        break;
      case 'rowresize':
        this.drag({ kind: 'rowresize', r: h.r, start: y, size: this.sheet.rows.size(h.r) });
        break;
      case 'fill':
        this.drag({ kind: 'fill', from: { ...this.sel } });
        break;
      default:
        this.select(h.r, h.c, e.shiftKey);
        this.app.onCellClick?.(h.r, h.c, e);
        this.drag({ kind: 'cells' });
    }
  }

  drag(state) {
    this.dragging = state;
    let lastEvent = null;
    let raf = 0;
    const autoScroll = () => {
      if (!this.dragging || !lastEvent) return;
      const [x, y] = this.local(lastEvent);
      const sc = this.scroller;
      let moved = false;
      if (y > this.h - 20) { sc.scrollTop += 20; moved = true; } else if (y < CH_H + this.frozenH() && y < CH_H + 4 && sc.scrollTop > 0) { sc.scrollTop -= 20; moved = true; }
      if (x > this.w - 20) { sc.scrollLeft += 20; moved = true; } else if (x < RH_W + 4 && sc.scrollLeft > 0) { sc.scrollLeft -= 20; moved = true; }
      if (moved) move(lastEvent);
      raf = requestAnimationFrame(autoScroll);
    };
    const move = (ev) => {
      lastEvent = ev;
      const [x, y] = this.local(ev);
      const st = this.dragging;
      const s = this.sheet;
      const r = this.rowAt(Math.max(CH_H + 1, Math.min(y, this.h - 2)));
      const c = this.colAt(Math.max(RH_W + 1, Math.min(x, this.w - 2)));
      if (st.kind === 'cells') this.select(r, c, true);
      else if (st.kind === 'cols') this.setRange({ r1: 0, r2: s.rows.count - 1, c1: Math.min(st.c0, c), c2: Math.max(st.c0, c) }, { r: this.sel.r, c: st.c0 });
      else if (st.kind === 'rows') this.setRange({ c1: 0, c2: s.cols.count - 1, r1: Math.min(st.r0, r), r2: Math.max(st.r0, r) }, { r: st.r0, c: this.sel.c });
      else if (st.kind === 'colresize') { st.now = Math.max(8, st.size + (x - st.start) / this.zoom); this.resizePreview = { kind: 'col', i: st.c, size: st.now }; this.invalidate(); }
      else if (st.kind === 'rowresize') { st.now = Math.max(8, st.size + (y - st.start) / this.zoom); this.resizePreview = { kind: 'row', i: st.r, size: st.now }; this.invalidate(); }
      else if (st.kind === 'fill') {
        const f = st.from;
        let target;
        const dr = r > f.r2 ? r - f.r2 : r < f.r1 ? r - f.r1 : 0;
        const dc = c > f.c2 ? c - f.c2 : c < f.c1 ? c - f.c1 : 0;
        if (Math.abs(dr) >= Math.abs(dc)) target = dr > 0 ? { ...f, r2: r } : dr < 0 ? { ...f, r1: r } : { ...f };
        else target = dc > 0 ? { ...f, c2: c } : { ...f, c1: c };
        this.fillPreview = target;
        this.invalidate();
      } else if (st.kind === 'ref') {
        this.app.editor.insertRef(Math.min(st.r0, r), Math.min(st.c0, c), Math.max(st.r0, r), Math.max(st.c0, c), false);
      }
    };
    const up = () => {
      document.removeEventListener('mousemove', move);
      document.removeEventListener('mouseup', up);
      cancelAnimationFrame(raf);
      const st = this.dragging;
      this.dragging = null;
      this.resizePreview = null;
      if (st.kind === 'colresize' && st.now) this.app.resizeAxis('col', st.c, st.now);
      if (st.kind === 'rowresize' && st.now) this.app.resizeAxis('row', st.r, st.now);
      if (st.kind === 'fill' && this.fillPreview) {
        const t = this.fillPreview;
        this.fillPreview = null;
        if (t.r1 !== st.from.r1 || t.r2 !== st.from.r2 || t.c1 !== st.from.c1 || t.c2 !== st.from.c2) this.app.fill(st.from, t);
      }
      if (st.kind === 'ref') this.app.editor.focus();
      this.invalidate();
    };
    document.addEventListener('mousemove', move);
    document.addEventListener('mouseup', up);
    raf = requestAnimationFrame(autoScroll);
  }

  onDblClick(e) {
    const h = this.hit(...this.local(e));
    if (h.area === 'colresize') this.app.autofit('col', h.c);
    else if (h.area === 'rowresize') this.app.autofit('row', h.r);
    else if (h.area === 'fill') this.app.fillDownToNeighbors();
    else if (h.area === 'cell') this.app.startEdit?.({ keepValue: true });
  }

  // ---- Desenho ---------------------------------------------------------------

  render() {
    this.dirty = false;
    const { ctx } = this;
    const dpr = devicePixelRatio || 1;
    const s = this.sheet;
    const C = this.colors;
    const z = this.zoom;
    ctx.setTransform(dpr, 0, 0, dpr, 0, 0);
    ctx.fillStyle = C.bg;
    ctx.fillRect(0, 0, this.w, this.h);
    const fw = this.frozenW(), fh = this.frozenH();
    const fr = s.freeze.r, fc = s.freeze.c;
    const sizeOf = (axis, i, kind) => (this.resizePreview?.kind === kind && this.resizePreview.i === i ? this.resizePreview.size : axis.size(i)) * z;

    const visibleCols = (from, x0, x1) => {
      const out = [];
      for (let c = from, x = x0; c < s.cols.count && x < x1; c++) {
        const w = sizeOf(s.cols, c, 'col');
        if (w > 0) out.push({ c, x, w });
        x += w;
      }
      return out;
    };
    const visibleRows = (from, y0, y1) => {
      const out = [];
      for (let r = from, y = y0; r < s.rows.count && y < y1; r++) {
        const h = sizeOf(s.rows, r, 'row');
        if (h > 0) out.push({ r, y, h });
        y += h;
      }
      return out;
    };
    const firstCol = fc + (this.scroller.scrollLeft ? s.cols.indexAt(s.cols.pos(fc) + this.scroller.scrollLeft / z) - fc : 0);
    const firstRow = fr + (this.scroller.scrollTop ? s.rows.indexAt(s.rows.pos(fr) + this.scroller.scrollTop / z) - fr : 0);
    const mainX0 = this.colX(firstCol), mainY0 = this.rowY(firstRow);
    const frozenCols = visibleCols(0, RH_W, RH_W + fw).filter((x) => x.c < fc);
    const frozenRows = visibleRows(0, CH_H, CH_H + fh).filter((x) => x.r < fr);
    const mainCols = visibleCols(firstCol, mainX0, this.w);
    const mainRows = visibleRows(firstRow, mainY0, this.h);
    this.visible = { cols: [...frozenCols, ...mainCols], rows: [...frozenRows, ...mainRows] };

    const panes = [
      { cols: mainCols, rows: mainRows, clip: [RH_W + fw, CH_H + fh, this.w, this.h] },
      { cols: frozenCols, rows: mainRows, clip: [RH_W, CH_H + fh, RH_W + fw, this.h] },
      { cols: mainCols, rows: frozenRows, clip: [RH_W + fw, CH_H, this.w, CH_H + fh] },
      { cols: frozenCols, rows: frozenRows, clip: [RH_W, CH_H, RH_W + fw, CH_H + fh] },
    ];
    for (const p of panes) if (p.cols.length && p.rows.length) this.drawPane(p);

    this.drawSelection();
    this.drawHeaders(frozenCols, mainCols, frozenRows, mainRows, fw, fh);
    if (fc || fr) {
      ctx.strokeStyle = C.frozen;
      ctx.lineWidth = 2;
      ctx.beginPath();
      if (fc) { ctx.moveTo(RH_W + fw, 0); ctx.lineTo(RH_W + fw, this.h); }
      if (fr) { ctx.moveTo(0, CH_H + fh); ctx.lineTo(this.w, CH_H + fh); }
      ctx.stroke();
    }
    this.app.afterRender?.();
  }

  drawPane({ cols, rows, clip }) {
    const { ctx } = this;
    const s = this.sheet;
    const C = this.colors;
    const z = this.zoom;
    const wb = this.wb;
    ctx.save();
    ctx.beginPath();
    ctx.rect(clip[0], clip[1], clip[2] - clip[0], clip[3] - clip[1]);
    ctx.clip();

    // Grade
    if (s.showGrid) {
      ctx.strokeStyle = C.line;
      ctx.lineWidth = 1;
      ctx.beginPath();
      const right = cols[cols.length - 1].x + cols[cols.length - 1].w, bottom = rows[rows.length - 1].y + rows[rows.length - 1].h;
      for (const col of cols) { const x = Math.round(col.x + col.w) - 0.5; ctx.moveTo(x, rows[0].y); ctx.lineTo(x, bottom); }
      for (const row of rows) { const y = Math.round(row.y + row.h) - 0.5; ctx.moveTo(cols[0].x, y); ctx.lineTo(right, y); }
      ctx.stroke();
    }

    const cfRules = s.cf.filter((rule) => rule.range.r2 >= rows[0].r && rule.range.r1 <= rows[rows.length - 1].r);
    const drawn = new Set();
    const overflow = [];
    const filter = s.filter;
    for (const row of rows) {
      for (let ci = 0; ci < cols.length; ci++) {
        const col = cols[ci];
        const r = row.r, c = col.c;
        let rect = { x: col.x, y: row.y, w: col.w, h: row.h };
        const m = s.merges.length ? s.mergeAt(r, c) : null;
        if (m) {
          const key = `${m.r1}:${m.c1}`;
          if (drawn.has(key)) continue;
          drawn.add(key);
          rect = this.cellRect(m.r1, m.c1);
        }
        const ar = m ? m.r1 : r, ac = m ? m.c1 : c;
        const cell = s.cells.get(KEY(ar, ac));
        const value = cell?.f || !cell ? wb.value(s, ar, ac) : cell.v;
        const style = cell?.s || {};
        let bg = style.bg, color = style.color, bold = style.b;
        let bar = null;
        for (const rule of cfRules) {
          if (r < rule.range.r1 || r > rule.range.r2 || c < rule.range.c1 || c > rule.range.c2) continue;
          const res = applyCf(rule, value);
          if (!res) continue;
          if (res.bg) bg = res.bg;
          if (res.color) color = res.color;
          if (res.bold) bold = true;
          if (res.bar) bar = res.bar;
        }
        if (bg || m) {
          ctx.fillStyle = bg || C.bg;
          ctx.fillRect(rect.x, rect.y, rect.w - (bg ? 0 : 1), rect.h - (bg ? 0 : 1));
          if (bg && s.showGrid && !m) {
            ctx.strokeStyle = 'rgba(0,0,0,.06)';
            ctx.strokeRect(rect.x + 0.5, rect.y + 0.5, rect.w - 1, rect.h - 1);
          }
        }
        if (bar) {
          ctx.fillStyle = bar.color;
          ctx.globalAlpha = 0.55;
          ctx.fillRect(rect.x + 2, rect.y + 3, Math.max(0, (rect.w - 4) * bar.frac), rect.h - 6);
          ctx.globalAlpha = 1;
        }
        if (style.bd) this.drawBorders(rect, style.bd);
        if (value === null || value === undefined || value === '') continue;
        const fmtStyle = cell?.autoFmt && !style.fmt ? { ...style, fmt: cell.autoFmt } : style;
        let text = display(value, fmtStyle);
        if (cell?.parseError) text = '#NOME?';
        const isNum = typeof value === 'number';
        const errV = isErr(value);
        let ha = style.ha || (isNum ? 'right' : typeof value === 'boolean' || errV ? 'center' : 'left');
        const fontPx = (style.size || 11) * (4 / 3) * z;
        ctx.font = `${style.i ? 'italic ' : ''}${bold ? '700 ' : ''}${fontPx}px ${style.font || DEFAULT_FONT}`;
        ctx.fillStyle = errV ? '#d6336c' : color || C.text;
        const pad = 4 * z;
        if (style.wrap) {
          this.drawWrapped(text, rect, ha, style.va || 'bottom', fontPx, pad, style);
          continue;
        }
        let tw = ctx.measureText(text).width;
        let clipW = rect.w;
        // Números que não cabem viram ###
        if (isNum && tw > rect.w - pad * 2 && !fmtStyle.fmt?.startsWith('general') && fmtStyle.fmt && fmtStyle.fmt !== 'general') {
          text = '#'.repeat(Math.max(1, Math.floor((rect.w - pad * 2) / ctx.measureText('#').width)));
          tw = ctx.measureText(text).width;
        }
        // Texto que transborda para células vazias à direita
        if (!isNum && ha === 'left' && tw > rect.w - pad && !m) {
          let extra = 0;
          for (let k = ci + 1; k < cols.length; k++) {
            const nc = cols[k].c;
            if (wb.value(s, r, nc) !== null || s.cells.get(KEY(r, nc))?.v != null) break;
            extra += cols[k].w;
            if (tw <= rect.w + extra - pad) break;
          }
          clipW += extra;
        }
        const va = style.va || 'bottom';
        const ty = va === 'top' ? rect.y + pad + fontPx * 0.8 : va === 'middle' ? rect.y + rect.h / 2 + fontPx * 0.32 : rect.y + rect.h - 5 * z;
        let tx = ha === 'right' ? rect.x + rect.w - pad - tw : ha === 'center' ? rect.x + (rect.w - tw) / 2 : rect.x + pad + (style.indent || 0) * 10 * z;
        if (clipW > rect.w) overflow.push({ text, x: tx, y: ty, clip: [rect.x, rect.y, clipW, rect.h], font: ctx.font, fill: ctx.fillStyle, bg: bg || C.bg });
        else {
          ctx.save();
          ctx.beginPath();
          ctx.rect(rect.x + 1, rect.y, rect.w - 2, rect.h);
          ctx.clip();
          ctx.fillText(text, tx, ty);
          if (style.u || style.st) {
            ctx.fillRect(tx, style.u ? ty + 2 * z : ty - fontPx * 0.3, tw, Math.max(1, z));
          }
          ctx.restore();
        }
        if (filter && r === filter.r1 && c >= filter.c1 && c <= filter.c2) this.drawFilterButton(rect, filter.criteria?.[c]);
      }
    }
    for (const o of overflow) {
      ctx.save();
      ctx.beginPath();
      ctx.rect(o.clip[0] + 1, o.clip[1], o.clip[2] - 2, o.clip[3]);
      ctx.clip();
      ctx.font = o.font;
      ctx.fillStyle = o.fill;
      ctx.fillText(o.text, o.x, o.y);
      ctx.restore();
    }
    // Botões de filtro em células vazias do cabeçalho
    if (filter) {
      for (const col of cols) {
        if (col.c < filter.c1 || col.c > filter.c2) continue;
        const rowInfo = rows.find((x) => x.r === filter.r1);
        if (!rowInfo) continue;
        if (wb.value(s, filter.r1, col.c) === null) this.drawFilterButton({ x: col.x, y: rowInfo.y, w: col.w, h: rowInfo.h }, filter.criteria?.[col.c]);
      }
    }
    ctx.restore();
  }

  drawWrapped(text, rect, ha, va, fontPx, pad, style) {
    const { ctx } = this;
    const lines = [];
    for (const para of String(text).split('\n')) {
      let line = '';
      for (const word of para.split(/(\s+)/)) {
        const test = line + word;
        if (ctx.measureText(test).width > rect.w - pad * 2 && line) { lines.push(line.trimEnd()); line = word.trimStart(); }
        else line = test;
      }
      lines.push(line);
    }
    const lh = fontPx * 1.22;
    const total = lines.length * lh;
    let y = va === 'top' ? rect.y + pad + fontPx * 0.85 : va === 'middle' ? rect.y + (rect.h - total) / 2 + fontPx * 0.85 : rect.y + rect.h - total + fontPx * 0.85 - 3;
    ctx.save();
    ctx.beginPath();
    ctx.rect(rect.x + 1, rect.y, rect.w - 2, rect.h);
    ctx.clip();
    for (const l of lines) {
      const tw = ctx.measureText(l).width;
      const x = ha === 'right' ? rect.x + rect.w - pad - tw : ha === 'center' ? rect.x + (rect.w - tw) / 2 : rect.x + pad;
      ctx.fillText(l, x, y);
      if (style.u) ctx.fillRect(x, y + 2, tw, 1);
      y += lh;
    }
    ctx.restore();
  }

  drawBorders(rect, bd) {
    const { ctx } = this;
    const side = (b, x1, y1, x2, y2) => {
      if (!b) return;
      ctx.save();
      ctx.strokeStyle = b.color || '#1d2129';
      ctx.lineWidth = b.w === 'thick' ? 3 : b.w === 'medium' ? 2 : 1;
      if (b.w === 'dashed') ctx.setLineDash([3, 2]);
      ctx.beginPath();
      ctx.moveTo(x1, y1);
      ctx.lineTo(x2, y2);
      ctx.stroke();
      ctx.restore();
    };
    const x = Math.round(rect.x) - 0.5, y = Math.round(rect.y) - 0.5, X = Math.round(rect.x + rect.w) - 0.5, Y = Math.round(rect.y + rect.h) - 0.5;
    side(bd.t, x, y, X, y);
    side(bd.b, x, Y, X, Y);
    side(bd.l, x, y, x, Y);
    side(bd.r, X, y, X, Y);
  }

  drawFilterButton(rect, active) {
    const { ctx } = this;
    const s = 16, x = rect.x + rect.w - s - 3, y = rect.y + (rect.h - s) / 2;
    ctx.fillStyle = active ? this.colors.accent : this.colors.head;
    ctx.strokeStyle = this.colors.line;
    ctx.beginPath();
    ctx.roundRect(x, y, s, s, 3);
    ctx.fill();
    ctx.stroke();
    ctx.fillStyle = active ? '#fff' : this.colors.headText;
    ctx.beginPath();
    ctx.moveTo(x + 4, y + 6);
    ctx.lineTo(x + 12, y + 6);
    ctx.lineTo(x + 8, y + 11);
    ctx.closePath();
    ctx.fill();
  }

  drawSelection() {
    const { ctx } = this;
    const C = this.colors;
    ctx.save();
    ctx.beginPath();
    ctx.rect(RH_W, CH_H, this.w - RH_W, this.h - CH_H);
    ctx.clip();
    // Referências da fórmula em edição
    for (const h of this.refHighlights) {
      const rr = this.rangeRect(h);
      ctx.fillStyle = `${h.color}14`;
      ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
      ctx.strokeStyle = h.color;
      ctx.lineWidth = 2;
      ctx.setLineDash(h.dashed ? [4, 3] : []);
      ctx.strokeRect(rr.x + 1, rr.y + 1, rr.w - 2, rr.h - 2);
      ctx.setLineDash([]);
    }
    // Área de despejo de matriz dinâmica
    const anchor = this.spillOf(this.sel.r, this.sel.c);
    if (anchor) {
      const rr = this.rangeRect(anchor);
      ctx.strokeStyle = C.accent;
      ctx.lineWidth = 1;
      ctx.setLineDash([4, 3]);
      ctx.strokeRect(rr.x + 0.5, rr.y + 0.5, rr.w - 1, rr.h - 1);
      ctx.setLineDash([]);
    }
    const sel = this.sel;
    const rr = this.rangeRect(sel);
    const multi = sel.r1 !== sel.r2 || sel.c1 !== sel.c2;
    if (multi) {
      ctx.fillStyle = C.selFill;
      ctx.fillRect(rr.x, rr.y, rr.w, rr.h);
      const ac = this.cellRect(sel.r, sel.c);
      ctx.strokeStyle = C.accent;
      ctx.lineWidth = 1;
      ctx.strokeRect(ac.x + 0.5, ac.y + 0.5, ac.w - 1, ac.h - 1);
    }
    ctx.strokeStyle = C.accent;
    ctx.lineWidth = 2;
    ctx.strokeRect(rr.x, rr.y, rr.w - 1, rr.h - 1);
    if (!this.app.editor?.active) {
      ctx.fillStyle = C.accent;
      ctx.fillRect(rr.x + rr.w - 4, rr.y + rr.h - 4, 7, 7);
      ctx.strokeStyle = C.bg;
      ctx.lineWidth = 1;
      ctx.strokeRect(rr.x + rr.w - 4.5, rr.y + rr.h - 4.5, 8, 8);
    }
    if (this.fillPreview) {
      const fr = this.rangeRect(this.fillPreview);
      ctx.strokeStyle = C.accent;
      ctx.setLineDash([3, 3]);
      ctx.lineWidth = 1;
      ctx.strokeRect(fr.x + 0.5, fr.y + 0.5, fr.w - 1, fr.h - 1);
      ctx.setLineDash([]);
    }
    if (this.copyRange && this.copyRange.sheet === this.sheet.id) {
      const cr = this.rangeRect(this.copyRange);
      ctx.strokeStyle = C.accent;
      ctx.lineWidth = 2;
      ctx.setLineDash([5, 3]);
      ctx.lineDashOffset = -this.antsOffset;
      ctx.strokeRect(cr.x + 1, cr.y + 1, cr.w - 2, cr.h - 2);
      ctx.setLineDash([]);
    }
    ctx.restore();
  }

  spillOf(r, c) {
    const s = this.sheet;
    let ar = r, ac = c;
    const anchor = this.wb.spillAnchor(s, r, c);
    if (anchor) [ar, ac] = anchor;
    const cell = s.get(ar, ac);
    if (!cell?.spillRows) return null;
    return { r1: ar, c1: ac, r2: ar + cell.spillRows - 1, c2: ac + cell.spillCols - 1 };
  }

  drawHeaders(frozenCols, mainCols, frozenRows, mainRows, fw, fh) {
    const { ctx } = this;
    const C = this.colors;
    const sel = this.sel;
    const z = this.zoom;
    ctx.font = `${12 * Math.min(1.2, Math.max(0.85, z))}px "Segoe UI", system-ui, sans-serif`;
    ctx.textBaseline = 'middle';
    const fullCols = sel.r1 === 0 && sel.r2 >= this.sheet.rows.count - 1;
    const fullRows = sel.c1 === 0 && sel.c2 >= this.sheet.cols.count - 1;
    // Colunas
    ctx.fillStyle = C.head;
    ctx.fillRect(0, 0, this.w, CH_H);
    const drawCols = (cols, clip) => {
      ctx.save();
      ctx.beginPath();
      ctx.rect(clip[0], 0, clip[1] - clip[0], CH_H);
      ctx.clip();
      for (const col of cols) {
        const on = col.c >= sel.c1 && col.c <= sel.c2;
        if (on) { ctx.fillStyle = fullCols ? C.accent : C.headSel; ctx.fillRect(col.x, 0, col.w, CH_H); }
        ctx.fillStyle = on ? (fullCols ? '#fff' : C.headSelText) : C.headText;
        ctx.textAlign = 'center';
        ctx.fillText(colName(col.c), col.x + col.w / 2, CH_H / 2 + 1);
        ctx.fillStyle = C.line;
        ctx.fillRect(Math.round(col.x + col.w) - 1, 0, 1, CH_H);
      }
      ctx.restore();
    };
    drawCols(frozenCols, [RH_W, RH_W + fw]);
    drawCols(mainCols, [RH_W + fw, this.w]);
    // Linhas
    ctx.fillStyle = C.head;
    ctx.fillRect(0, CH_H, RH_W, this.h);
    const drawRows = (rows, clip) => {
      ctx.save();
      ctx.beginPath();
      ctx.rect(0, clip[0], RH_W, clip[1] - clip[0]);
      ctx.clip();
      for (const row of rows) {
        const on = row.r >= sel.r1 && row.r <= sel.r2;
        if (on) { ctx.fillStyle = fullRows ? C.accent : C.headSel; ctx.fillRect(0, row.y, RH_W, row.h); }
        ctx.fillStyle = on ? (fullRows ? '#fff' : C.headSelText) : C.headText;
        ctx.textAlign = 'center';
        if (row.h > 9) ctx.fillText(String(row.r + 1), RH_W / 2, row.y + row.h / 2 + 1);
        ctx.fillStyle = C.line;
        ctx.fillRect(0, Math.round(row.y + row.h) - 1, RH_W, 1);
      }
      ctx.restore();
    };
    drawRows(frozenRows, [CH_H, CH_H + fh]);
    drawRows(mainRows, [CH_H + fh, this.h]);
    ctx.fillStyle = C.line;
    ctx.fillRect(RH_W - 1, 0, 1, this.h);
    ctx.fillRect(0, CH_H - 1, this.w, 1);
    ctx.fillStyle = C.head;
    ctx.fillRect(0, 0, RH_W - 1, CH_H - 1);
    // Triângulo do canto (selecionar tudo)
    ctx.fillStyle = C.line;
    ctx.beginPath();
    ctx.moveTo(RH_W - 6, CH_H - 14);
    ctx.lineTo(RH_W - 6, CH_H - 6);
    ctx.lineTo(RH_W - 14, CH_H - 6);
    ctx.closePath();
    ctx.fill();
    ctx.textBaseline = 'alphabetic';
    ctx.textAlign = 'left';
  }
}

// ---------------------------------------------------------------------------
// Formatação condicional

const lerpColor = (a, b, t) => {
  const pa = hex(a), pb = hex(b);
  return `rgb(${pa.map((x, i) => Math.round(x + (pb[i] - x) * t)).join(',')})`;
};
const hex = (h) => {
  const n = parseInt(h.slice(1), 16);
  return [(n >> 16) & 255, (n >> 8) & 255, n & 255];
};

export function applyCf(rule, v) {
  const st = rule._stats;
  if (!st) return null;
  const p = rule.params || {};
  const style = rule.style || { bg: '#ffc9c9', color: '#9c0006' };
  const n = typeof v === 'number' ? v : null;
  switch (rule.type) {
    case 'scale2':
    case 'scale3': {
      if (n === null) return null;
      const lo = p.min ?? '#f8696b', mid = p.mid ?? '#ffeb84', hi = p.max ?? '#63be7b';
      if (st.max === st.min) return { bg: rule.type === 'scale3' ? mid : hi };
      const t = (n - st.min) / (st.max - st.min);
      if (rule.type === 'scale2') return { bg: lerpColor(lo, hi, t) };
      const tm = (st.mid - st.min) / (st.max - st.min) || 0.5;
      return { bg: t < tm ? lerpColor(lo, mid, t / tm) : lerpColor(mid, hi, (t - tm) / (1 - tm || 1)) };
    }
    case 'bar':
      if (n === null) return null;
      return { bar: { color: p.color || '#5b8cff', frac: st.max <= 0 ? 0 : Math.max(0, Math.min(1, (n - Math.min(0, st.min)) / (st.max - Math.min(0, st.min) || 1))) } };
    case 'gt': return n !== null && n > p.a ? style : null;
    case 'lt': return n !== null && n < p.a ? style : null;
    case 'between': return n !== null && n >= p.a && n <= p.b ? style : null;
    case 'eq': return (n !== null ? n === p.a : String(v ?? '').toLowerCase() === String(p.a).toLowerCase()) ? style : null;
    case 'text': return typeof v === 'string' && v.toLowerCase().includes(String(p.a).toLowerCase()) ? style : null;
    case 'dup': {
      if (v === null || v === '') return null;
      const k = typeof v === 'string' ? v.toLowerCase() : String(v);
      return st.counts.get(k) > 1 ? style : null;
    }
    case 'unique': {
      if (v === null || v === '') return null;
      const k = typeof v === 'string' ? v.toLowerCase() : String(v);
      return st.counts.get(k) === 1 ? style : null;
    }
    case 'top': return n !== null && n >= st.top ? style : null;
    case 'bottom': return n !== null && n <= st.bottom ? style : null;
    case 'aboveAvg': return n !== null && n > st.avg ? style : null;
    case 'belowAvg': return n !== null && n < st.avg ? style : null;
    case 'blank': return v === null || v === '' ? style : null;
    case 'error': return isErr(v) ? style : null;
    default: return null;
  }
}

// ---------------------------------------------------------------------------
// Utilidades

export function norm(r1, c1, r2, c2) {
  return { r1: Math.min(r1, r2), c1: Math.min(c1, c2), r2: Math.max(r1, r2), c2: Math.max(c1, c2) };
}

export function inRange(rg, r, c) {
  return r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2;
}

// ---------------------------------------------------------------------------
// Editor de célula (sobreposto) com autocompletar de funções

const FN_LIST = Object.values(FUNCS).map((f) => ({ ...f, label: f.pt.split('|')[0], search: f.pt.split('|')[0].normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase() }))
  .sort((a, b) => a.label.localeCompare(b.label, 'pt-BR'));

export class CellEditor {
  constructor(app, grid, formulaBar) {
    this.app = app;
    this.grid = grid;
    this.bar = formulaBar;
    this.active = false;
    this.input = document.createElement('textarea');
    this.input.className = 'cell-editor';
    this.input.spellcheck = false;
    this.input.wrap = 'off';
    grid.root.append(this.input);
    this.ac = document.createElement('div');
    this.ac.className = 'fn-ac hidden';
    this.tip = document.createElement('div');
    this.tip.className = 'fn-tip hidden';
    document.body.append(this.ac, this.tip);
    this.acIndex = 0;
    this.source = null; // 'cell' | 'bar'
    this.input.addEventListener('input', () => { this.bar.value = this.input.value; this.update(); });
    this.input.addEventListener('keydown', (e) => this.onKey(e, this.input));
    this.input.addEventListener('keyup', () => this.update());
    this.input.addEventListener('click', () => this.update());
    this.bar.addEventListener('focus', () => {
      if (!this.active) this.begin({ keepValue: true, from: 'bar' });
      this.source = 'bar';
    });
    this.bar.addEventListener('input', () => { this.input.value = this.bar.value; this.update(); });
    this.bar.addEventListener('keydown', (e) => this.onKey(e, this.bar));
    this.bar.addEventListener('keyup', () => this.update());
    this.bar.addEventListener('click', () => this.update());
    this.input.addEventListener('blur', () => setTimeout(() => this.maybeCommitOnBlur(), 0));
    this.bar.addEventListener('blur', () => setTimeout(() => this.maybeCommitOnBlur(), 0));
  }

  get el() { return this.source === 'bar' ? this.bar : this.input; }

  maybeCommitOnBlur() {
    if (!this.active) return;
    const a = document.activeElement;
    if (a === this.input || a === this.bar || this.grid.dragging || a?.closest?.('.fn-ac')) return;
    this.commit();
  }

  begin({ text = null, keepValue = false, from = 'cell' } = {}) {
    const s = this.grid.sheet;
    const { r, c } = this.grid.sel;
    const cell = s.get(r, c);
    if (this.app.isLocked?.(r, c)) return;
    this.active = true;
    this.typed = !keepValue;
    this.source = from;
    this.r = r;
    this.c = c;
    this.sheetId = s.id;
    this.original = editText(cell);
    const v = text !== null ? text : keepValue ? this.original : '';
    this.input.value = v;
    this.bar.value = v;
    this.position();
    this.input.classList.add('on');
    const style = cell?.s || {};
    this.input.style.font = `${style.i ? 'italic ' : ''}${style.b ? '700 ' : ''}${(style.size || 11) * (4 / 3) * this.grid.zoom}px ${style.font || DEFAULT_FONT}`;
    this.input.style.textAlign = style.ha || 'left';
    this.input.style.color = style.color || '';
    this.input.style.background = style.bg || '';
    if (from === 'cell') {
      this.input.focus();
      this.input.setSelectionRange(v.length, v.length);
    }
    this.update();
    this.grid.invalidate();
  }

  position() {
    const rect = this.grid.cellRect(this.r, this.c);
    const s = this.input.style;
    s.left = `${rect.x}px`;
    s.top = `${rect.y}px`;
    s.minWidth = `${rect.w}px`;
    s.height = `${rect.h}px`;
    s.minHeight = `${rect.h}px`;
    this.autoSize();
  }

  autoSize() {
    const s = this.input.style;
    s.width = '0px';
    const rect = this.grid.cellRect(this.r, this.c);
    s.width = `${Math.min(Math.max(rect.w, this.input.scrollWidth + 8), this.grid.w - rect.x - 4)}px`;
    s.height = `${Math.max(rect.h, this.input.scrollHeight)}px`;
  }

  focus() {
    this.el.focus();
  }

  isFormula() { return this.input.value.startsWith('='); }

  // Pode inserir referência clicando na grade? (após operador, "(" ou separador)
  pointMode() {
    if (!this.isFormula()) return false;
    const el = this.el;
    const before = el.value.slice(0, el.selectionStart);
    if (this.lastRef && this.lastRef.end === el.selectionStart && el.selectionStart === el.selectionEnd) return true;
    return /[=(;,+\-*/^&<>:\s]$/.test(before);
  }

  // Insere (ou substitui a última inserida) uma referência na posição do cursor.
  insertRef(r1, c1, r2, c2, fresh) {
    const el = this.el;
    const sameSheet = this.sheetId === this.grid.sheet.id;
    const prefix = sameSheet ? '' : `${quoteSheetName(this.grid.sheet.name)}!`;
    const ref = prefix + (r1 === r2 && c1 === c2 ? `${colName(c1)}${r1 + 1}` : `${colName(c1)}${r1 + 1}:${colName(c2)}${r2 + 1}`);
    let start = el.selectionStart, end = el.selectionEnd;
    if (!fresh && this.lastRef) { start = this.lastRef.start; end = this.lastRef.end; }
    else if (this.lastRef && this.lastRef.end === start && start === end) start = this.lastRef.start;
    el.value = el.value.slice(0, start) + ref + el.value.slice(end);
    const pos = start + ref.length;
    el.setSelectionRange(pos, pos);
    this.lastRef = { start, end: pos };
    if (el === this.bar) this.input.value = el.value; else this.bar.value = el.value;
    this.update(true);
  }

  update(keepRef = false) {
    if (!this.active) return;
    if (!keepRef) {
      const el = this.el;
      if (this.lastRef && (el.selectionStart !== this.lastRef.end)) this.lastRef = null;
    }
    this.autoSize();
    // Destaques coloridos das referências
    const v = this.input.value;
    const hl = [];
    if (v.startsWith('=') && this.sheetId === this.grid.sheet.id) {
      const sheetName = this.grid.sheet.name;
      formulaRefs(v.slice(1)).forEach((ref, i) => {
        if (ref.sheet && ref.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
        if (ref.kind === 'cell') hl.push({ r1: ref.a.row, c1: ref.a.col, r2: ref.a.row, c2: ref.a.col, color: REF_COLORS[i % REF_COLORS.length] });
        else if (ref.kind === 'range') hl.push({ ...norm(ref.a.row, ref.a.col, ref.b.row, ref.b.col), color: REF_COLORS[i % REF_COLORS.length] });
      });
    }
    this.grid.refHighlights = hl;
    this.grid.invalidate();
    this.updateAutocomplete();
  }

  updateAutocomplete() {
    const el = this.el;
    const v = el.value;
    if (!v.startsWith('=') || document.activeElement !== el) { this.hideAc(); this.hideTip(); return; }
    const pos = el.selectionStart;
    const before = v.slice(0, pos);
    const m = /([A-Za-zÀ-ÿ][\wÀ-ÿ.]*)$/.exec(before);
    const inString = (before.match(/"/g) || []).length % 2 === 1;
    if (m && !inString && !/[A-Za-z]{1,3}\d+$/.test(m[1])) {
      const q = m[1].normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
      const pt = FN_LIST.filter((f) => f.search.startsWith(q));
      const en = FN_LIST.filter((f) => !f.search.startsWith(q) && f.name.startsWith(q));
      this.acItems = [...pt, ...en].slice(0, 12);
      const suppressed = this.acSuppressed === `${pos}:${m[1]}`;
      if (this.acItems.length && !suppressed && !(this.acItems.length === 1 && this.acItems[0].search === q && v[pos] === '(')) {
        this.acPrefix = m[1];
        this.acIndex = Math.min(this.acIndex, this.acItems.length - 1);
        this.showAc();
      } else this.hideAc();
    } else this.hideAc();
    // Dica de sintaxe da função atual
    const call = currentCall(before);
    if (call && !inString) {
      const f = FUNCS[resolveFn(call.name)];
      if (f) { this.showTip(f, call.arg); return; }
    }
    this.hideTip();
  }

  showAc() {
    const r = this.el.getBoundingClientRect();
    this.ac.replaceChildren(...this.acItems.map((f, i) => {
      const d = document.createElement('div');
      d.className = `fn-item${i === this.acIndex ? ' on' : ''}`;
      d.innerHTML = `<b>${f.label}</b><span>${f.d}</span>`;
      d.addEventListener('mousedown', (e) => { e.preventDefault(); this.acIndex = i; this.acceptAc(); });
      return d;
    }));
    this.ac.classList.remove('hidden');
    this.ac.style.left = `${r.left}px`;
    this.ac.style.top = `${r.bottom + 2}px`;
    this.ac.querySelector('.on')?.scrollIntoView({ block: 'nearest' });
  }

  hideAc() { this.ac.classList.add('hidden'); this.acItems = []; }

  acceptAc() {
    const f = this.acItems[this.acIndex];
    if (!f) return;
    const el = this.el;
    const pos = el.selectionStart;
    const start = pos - this.acPrefix.length;
    const insert = `${f.label}(`;
    const after = el.value.slice(pos);
    el.value = el.value.slice(0, start) + insert + (after.startsWith('(') ? after.slice(1) : after);
    const p = start + insert.length;
    el.setSelectionRange(p, p);
    if (el === this.bar) this.input.value = el.value; else this.bar.value = el.value;
    this.hideAc();
    this.update();
  }

  showTip(f, argIndex) {
    const parts = f.s.replace(/^[^(]*\(|\)$/g, '').split(';').map((x) => x.trim()).filter(Boolean);
    const name = f.pt.split('|')[0];
    const last = parts.length - 1;
    const idx = parts[last] === '…' ? Math.min(argIndex, Math.max(0, last - 1)) : Math.min(argIndex, last);
    this.tip.innerHTML = `<div class="sig">${name}(${parts.map((p, i) => (i === idx ? `<b>${p}</b>` : p)).join('; ')})</div><div class="desc">${f.d}</div>`;
    const r = this.el.getBoundingClientRect();
    this.tip.classList.remove('hidden');
    this.tip.style.left = `${r.left}px`;
    this.tip.style.top = `${Math.max(4, r.top - this.tip.offsetHeight - 4)}px`;
  }

  hideTip() { this.tip.classList.add('hidden'); }

  onKey(e, el) {
    if (!this.active) {
      if (el === this.bar && e.key === 'Escape') { this.app.focusGrid(); }
      return;
    }
    const acOpen = !this.ac.classList.contains('hidden') && this.acItems?.length;
    if (acOpen && (e.key === 'ArrowDown' || e.key === 'ArrowUp')) {
      e.preventDefault();
      this.acIndex = (this.acIndex + (e.key === 'ArrowDown' ? 1 : -1) + this.acItems.length) % this.acItems.length;
      this.showAc();
      return;
    }
    if (acOpen && (e.key === 'Tab' || (e.key === 'Enter' && !e.altKey))) {
      e.preventDefault();
      this.acceptAc();
      return;
    }
    if (e.key === 'Escape') {
      e.preventDefault();
      if (acOpen) {
        this.acSuppressed = `${el.selectionStart}:${this.acPrefix}`;
        this.hideAc();
        return;
      }
      this.cancel();
      return;
    }
    if (e.key === 'Enter' && e.altKey) {
      e.preventDefault();
      const p = el.selectionStart;
      el.value = `${el.value.slice(0, p)}\n${el.value.slice(el.selectionEnd)}`;
      el.setSelectionRange(p + 1, p + 1);
      this.input.value = el.value;
      this.bar.value = el.value;
      this.update();
      return;
    }
    if (e.key === 'Enter') {
      e.preventDefault();
      this.commit(e.ctrlKey ? { fillSelection: true } : { move: e.shiftKey ? [-1, 0] : [1, 0] });
      return;
    }
    if (e.key === 'Tab') {
      e.preventDefault();
      this.commit({ move: e.shiftKey ? [0, -1] : [0, 1] });
      return;
    }
    if (this.typed && el === this.input && !this.isFormula() && ['ArrowUp', 'ArrowDown', 'ArrowLeft', 'ArrowRight'].includes(e.key)) {
      e.preventDefault();
      const d = { ArrowUp: [-1, 0], ArrowDown: [1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1] }[e.key];
      this.commit({ move: d });
    }
  }

  commit(opts = {}) {
    if (!this.active) return;
    const value = this.input.value;
    this.end();
    this.app.returnToSheet?.(this.sheetId);
    if (value !== this.original || opts.fillSelection) this.app.setCellInput(this.r, this.c, value, opts);
    if (opts.move) this.app.moveAfterEdit(opts.move);
    this.app.focusGrid();
  }

  cancel() {
    if (!this.active) return;
    this.end();
    this.app.returnToSheet?.(this.sheetId);
    this.bar.value = this.original;
    this.app.focusGrid();
  }

  end() {
    this.active = false;
    this.acSuppressed = null;
    this.lastRef = null;
    this.input.classList.remove('on');
    this.input.value = '';
    this.grid.refHighlights = [];
    this.hideAc();
    this.hideTip();
    this.grid.invalidate();
  }
}

function quoteSheetName(n) {
  return /^[A-Za-z_À-ɏ][\wÀ-ɏ.]*$/.test(n) && !/^[A-Za-z]{1,3}\d+$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
}

// Descobre a função cujo parêntese contém o cursor e o índice do argumento.
function currentCall(before) {
  let depth = 0, arg = 0, inStr = false;
  for (let i = before.length - 1; i >= 0; i--) {
    const ch = before[i];
    if (ch === '"') { inStr = !inStr; continue; }
    if (inStr) continue;
    if (ch === ')') depth++;
    else if (ch === '(') {
      if (depth === 0) {
        const m = /([A-Za-zÀ-ÿ][\wÀ-ÿ.]*)\s*$/.exec(before.slice(0, i));
        return m ? { name: m[1], arg } : null;
      }
      depth--;
    } else if ((ch === ';' || ch === ',') && depth === 0) arg++;
  }
  return null;
}

function resolveFn(name) {
  const strip = (s) => s.normalize('NFD').replace(/[̀-ͯ]/g, '').toUpperCase();
  const u = name.toUpperCase();
  if (FUNCS[u]) return u;
  const hit = FN_LIST.find((f) => f.pt.split('|').some((p) => strip(p) === strip(name)));
  return hit?.name || u;
}

export { isMatrix, RC };
