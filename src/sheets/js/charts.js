// Gráficos flutuantes sobre a grade (Chart.js carregado sob demanda).
import { loadVendor } from '../../shared/native.js';
import { RH_W, CH_H } from './grid.js';
import { display } from './format.js';
import { isErr } from './values.js';

export const CHART_TYPES = [
  ['column', 'Colunas'], ['bar', 'Barras'], ['line', 'Linhas'], ['area', 'Área'], ['pie', 'Pizza'], ['doughnut', 'Rosca'], ['scatter', 'Dispersão'],
];
export const PALETTE = ['#3b6cf6', '#f0603f', '#2f9e6e', '#f5b400', '#9b5de5', '#00b4d8', '#ef476f', '#6c757d', '#06d6a0', '#ff8fab'];

export class ChartLayer {
  constructor(app, grid) {
    this.app = app;
    this.grid = grid;
    this.layer = grid.overlay;
    this.boxes = new Map();
    this.selected = null;
    this.Chart = null;
    this.sheetId = null;
  }

  async ensureLib() {
    if (!this.Chart) this.Chart = await loadVendor('chart.umd.min.js', 'Chart');
    return this.Chart;
  }

  // Extrai categorias e séries do intervalo.
  data(ch) {
    const wb = this.app.wb;
    const sheet = wb.byId(ch.sheetId) || wb.sheet;
    const { r1, c1, r2, c2 } = ch.range;
    const m = [];
    for (let r = r1; r <= r2; r++) {
      const row = [];
      for (let c = c1; c <= c2; c++) row.push(wb.value(sheet, r, c));
      m.push(row);
    }
    if (ch.pick) for (let i = 0; i < m.length; i++) m[i] = ch.pick.map((k) => m[i][k]);
    const byRows = ch.seriesIn === 'rows';
    const grid = byRows ? m : m[0].map((_, c) => m.map((row) => row[c]));
    // grid: cada item é uma série (coluna) com valores por categoria
    const isText = (v) => typeof v === 'string' && v !== '';
    const header = ch.header ?? grid.some((s) => isText(s[0]));
    const hasLabels = grid.length > 1 && grid[0].slice(header ? 1 : 0).some(isText);
    const labelsSeries = hasLabels ? grid[0] : null;
    const series = (hasLabels ? grid.slice(1) : grid).map((s, i) => ({
      name: header ? String(s[0] ?? `Série ${i + 1}`) : `Série ${i + 1}`,
      values: s.slice(header ? 1 : 0).map((v) => (typeof v === 'number' ? v : isErr(v) ? null : Number.isFinite(+v) && v !== '' && v !== null ? +v : null)),
    }));
    const n = series[0]?.values.length || 0;
    const labels = labelsSeries
      ? labelsSeries.slice(header ? 1 : 0).map((v) => (typeof v === 'number' ? display(v, sheet.get?.(0, 0)?.s || {}) : String(v ?? '')))
      : Array.from({ length: n }, (_, i) => String(i + 1));
    return { labels, series, rawX: labelsSeries ? labelsSeries.slice(header ? 1 : 0) : null };
  }

  config(ch) {
    const { labels, series, rawX } = this.data(ch);
    const dark = document.documentElement.dataset.theme === 'dark';
    const text = dark ? '#c9ccd6' : '#4a5061';
    const gridColor = dark ? 'rgba(255,255,255,.08)' : 'rgba(0,0,0,.07)';
    const t = ch.type;
    const colors = ch.colors?.length ? ch.colors : PALETTE;
    const pie = t === 'pie' || t === 'doughnut';
    const datasets = pie
      ? [{ label: series[0]?.name, data: series[0]?.values || [], backgroundColor: labels.map((_, i) => colors[i % colors.length]), borderColor: dark ? '#1d1f26' : '#fff', borderWidth: 2 }]
      : series.map((s, i) => {
        const c = colors[i % colors.length];
        const base = { label: s.name, borderColor: c, backgroundColor: t === 'area' ? `${c}55` : c, borderWidth: t === 'line' || t === 'area' ? 2.5 : 0, borderRadius: t === 'column' || t === 'bar' ? 4 : 0 };
        if (t === 'scatter') return { ...base, data: s.values.map((y, k) => ({ x: typeof rawX?.[k] === 'number' ? rawX[k] : k + 1, y })), pointRadius: 4 };
        return { ...base, data: s.values, tension: 0.3, fill: t === 'area', pointRadius: t === 'line' ? 3 : 0 };
      });
    return {
      type: { column: 'bar', bar: 'bar', line: 'line', area: 'line', pie: 'pie', doughnut: 'doughnut', scatter: 'scatter' }[t] || 'bar',
      data: { labels, datasets },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        animation: { duration: 300 },
        indexAxis: t === 'bar' ? 'y' : 'x',
        plugins: {
          title: { display: !!ch.title, text: ch.title, color: dark ? '#e8e9ee' : '#1d2129', font: { size: 15, weight: '600' } },
          legend: { display: ch.legend !== 'none' && (pie || series.length > 1), position: ch.legend === 'top' ? 'top' : ch.legend === 'right' ? 'right' : 'bottom', labels: { color: text, usePointStyle: true } },
          tooltip: { callbacks: { label: (ctx) => `${ctx.dataset.label ?? ''}: ${new Intl.NumberFormat('pt-BR').format(ctx.parsed.y ?? ctx.parsed)}` } },
        },
        scales: pie ? {} : {
          x: { stacked: !!ch.stacked, ticks: { color: text }, grid: { color: gridColor, display: t === 'bar' || t === 'scatter' } },
          y: { stacked: !!ch.stacked, ticks: { color: text, callback: (v) => new Intl.NumberFormat('pt-BR', { notation: 'compact' }).format(v) }, grid: { color: gridColor, display: t !== 'bar' }, beginAtZero: true },
        },
      },
    };
  }

  async sync() {
    const sheet = this.app.wb.sheet;
    const charts = sheet.charts;
    if (this.sheetId !== sheet.id) {
      for (const b of this.boxes.values()) { b.chart?.destroy(); b.el.remove(); }
      this.boxes.clear();
      this.sheetId = sheet.id;
    }
    if (charts.some((c) => c.type !== 'image')) await this.ensureLib();
    const ids = new Set(charts.map((c) => c.id));
    for (const [id, b] of this.boxes) if (!ids.has(id)) { b.chart?.destroy(); b.el.remove(); this.boxes.delete(id); }
    for (const ch of charts) {
      let b = this.boxes.get(ch.id);
      if (ch.type === 'image') {
        // Imagem flutuante (logotipos, fotos)
        if (!b) {
          const el = document.createElement('div');
          el.className = 'chart-box image-box';
          const img = document.createElement('img');
          img.draggable = false;
          const handle = document.createElement('div');
          handle.className = 'chart-resize';
          el.append(img, handle);
          this.layer.append(el);
          b = { el, img, chart: null, sig: '' };
          this.boxes.set(ch.id, b);
          this.bind(ch.id, el, handle);
        }
        if (b.img.src !== ch.src) b.img.src = ch.src;
        b.el.classList.toggle('selected', this.selected === ch.id);
        continue;
      }
      if (!b) {
        const el = document.createElement('div');
        el.className = 'chart-box';
        const canvas = document.createElement('canvas');
        const handle = document.createElement('div');
        handle.className = 'chart-resize';
        el.append(canvas, handle);
        this.layer.append(el);
        b = { el, canvas, chart: null, sig: '' };
        this.boxes.set(ch.id, b);
        this.bind(ch.id, el, handle);
      }
      const cfg = this.config(ch);
      const sig = JSON.stringify([ch.type, ch.stacked, ch.legend, ch.title, cfg.data, document.documentElement.dataset.theme]);
      if (!b.chart || b.type !== cfg.type) {
        b.chart?.destroy();
        b.chart = new this.Chart(b.canvas, cfg);
        b.type = cfg.type;
      } else if (sig !== b.sig) {
        b.chart.data = cfg.data;
        b.chart.options = cfg.options;
        b.chart.update('none');
      }
      b.sig = sig;
      b.el.classList.toggle('selected', this.selected === ch.id);
    }
    this.position();
  }

  position() {
    const z = this.grid.zoom;
    const sc = this.grid.scroller;
    const sheet = this.app.wb.sheet;
    for (const ch of sheet.charts) {
      const b = this.boxes.get(ch.id);
      if (!b) continue;
      Object.assign(b.el.style, {
        left: `${ch.x * z - sc.scrollLeft}px`, top: `${ch.y * z - sc.scrollTop}px`, width: `${ch.w * z}px`, height: `${ch.h * z}px`,
      });
    }
  }

  bind(id, el, handle) {
    const find = () => this.app.wb.sheet.charts.find((c) => c.id === id);
    el.addEventListener('mousedown', (e) => {
      if (e.button !== 0) return;
      e.preventDefault();
      e.stopPropagation();
      this.select(id);
      const ch = find();
      if (!ch) return;
      const resize = e.target === handle;
      const sx = e.clientX, sy = e.clientY, o = { x: ch.x, y: ch.y, w: ch.w, h: ch.h };
      const z = this.grid.zoom;
      const move = (ev) => {
        const dx = (ev.clientX - sx) / z, dy = (ev.clientY - sy) / z;
        if (resize) {
          const min = ch.type === 'image' ? 24 : 200;
          ch.w = Math.max(min, o.w + dx);
          ch.h = ch.type === 'image' && !ev.shiftKey ? Math.max(min, (ch.w * o.h) / o.w) : Math.max(ch.type === 'image' ? 24 : 140, o.h + dy);
        }
        else { ch.x = Math.max(0, o.x + dx); ch.y = Math.max(0, o.y + dy); }
        this.position();
      };
      const up = () => {
        document.removeEventListener('mousemove', move);
        document.removeEventListener('mouseup', up);
        if (ch.x !== o.x || ch.y !== o.y || ch.w !== o.w || ch.h !== o.h) {
          const final = { ...ch };
          Object.assign(ch, o);
          this.app.updateChart(id, final, resize ? 'Redimensionar gráfico' : 'Mover gráfico');
        }
      };
      document.addEventListener('mousemove', move);
      document.addEventListener('mouseup', up);
    });
    el.addEventListener('dblclick', (e) => { e.stopPropagation(); this.app.editChart(id); });
    el.addEventListener('contextmenu', (e) => { e.preventDefault(); e.stopPropagation(); this.select(id); this.app.chartMenu(id, e); });
  }

  select(id) {
    this.selected = id;
    for (const [k, b] of this.boxes) b.el.classList.toggle('selected', k === id);
    this.app.onChartSelect?.(id);
  }

  // Imagem PNG do gráfico (para exportação)
  image(id) {
    const b = this.boxes.get(id);
    if (b?.img) return b.img.src;
    return b?.chart?.toBase64Image() || null;
  }
}

export function defaultChartPlacement(grid, rg) {
  const s = grid.sheet;
  return {
    x: s.cols.pos(rg.c2 + 1) + 16,
    y: s.rows.pos(rg.r1),
    w: 520,
    h: 320,
  };
}

export { RH_W, CH_H };
