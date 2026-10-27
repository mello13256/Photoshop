// Ajustes e filtros. Cada filtro recebe um canvas e retorna um novo canvas do mesmo tamanho.
import { createCanvas, cloneCanvas, clamp, rgbToHsl, hslToRgb } from './utils.js';

function pixelOp(src, fn) {
  const out = cloneCanvas(src);
  const ctx = out.getContext('2d');
  const id = ctx.getImageData(0, 0, out.width, out.height);
  fn(id.data, out.width, out.height);
  ctx.putImageData(id, 0, 0);
  return out;
}

function lutOp(src, lutR, lutG = lutR, lutB = lutR) {
  return pixelOp(src, (d) => {
    for (let i = 0; i < d.length; i += 4) {
      d[i] = lutR[d[i]];
      d[i + 1] = lutG[d[i + 1]];
      d[i + 2] = lutB[d[i + 2]];
    }
  });
}

function makeLut(fn) {
  const lut = new Uint8ClampedArray(256);
  for (let i = 0; i < 256; i++) lut[i] = fn(i);
  return lut;
}

function blurCanvas(src, radius) {
  if (radius <= 0) return cloneCanvas(src);
  const pad = Math.ceil(radius * 3);
  // Estende as bordas para evitar halo transparente.
  const big = createCanvas(src.width + pad * 2, src.height + pad * 2);
  const b = big.getContext('2d');
  b.drawImage(src, pad, pad);
  b.drawImage(src, 0, 0, 1, src.height, 0, pad, pad, src.height);
  b.drawImage(src, src.width - 1, 0, 1, src.height, pad + src.width, pad, pad, src.height);
  b.drawImage(big, 0, pad, big.width, 1, 0, 0, big.width, pad);
  b.drawImage(big, 0, pad + src.height - 1, big.width, 1, 0, pad + src.height, big.width, pad);
  const out = createCanvas(src.width, src.height);
  const ctx = out.getContext('2d');
  ctx.filter = `blur(${radius}px)`;
  ctx.drawImage(big, -pad, -pad);
  // Preserva áreas totalmente transparentes da origem como desfocadas também (comportamento natural).
  return out;
}

function convolve(src, kernel, { divisor, offset = 0, keepAlpha = true } = {}) {
  const size = Math.round(Math.sqrt(kernel.length));
  const half = size >> 1;
  const div = divisor ?? (kernel.reduce((a, b) => a + b, 0) || 1);
  const w = src.width, h = src.height;
  const sctx = src.getContext('2d');
  const sd = sctx.getImageData(0, 0, w, h).data;
  const out = createCanvas(w, h);
  const octx = out.getContext('2d');
  const od = octx.createImageData(w, h);
  const o = od.data;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      let r = 0, g = 0, b = 0;
      for (let ky = 0; ky < size; ky++) {
        const yy = clamp(y + ky - half, 0, h - 1);
        for (let kx = 0; kx < size; kx++) {
          const xx = clamp(x + kx - half, 0, w - 1);
          const k = kernel[ky * size + kx];
          if (!k) continue;
          const i = (yy * w + xx) * 4;
          r += sd[i] * k; g += sd[i + 1] * k; b += sd[i + 2] * k;
        }
      }
      const i = (y * w + x) * 4;
      o[i] = r / div + offset;
      o[i + 1] = g / div + offset;
      o[i + 2] = b / div + offset;
      o[i + 3] = keepAlpha ? sd[i + 3] : 255;
    }
  }
  octx.putImageData(od, 0, 0);
  return out;
}

export function histogram(canvas) {
  const d = canvas.getContext('2d').getImageData(0, 0, canvas.width, canvas.height).data;
  const r = new Uint32Array(256), g = new Uint32Array(256), b = new Uint32Array(256), l = new Uint32Array(256);
  for (let i = 0; i < d.length; i += 4) {
    if (d[i + 3] === 0) continue;
    r[d[i]]++; g[d[i + 1]]++; b[d[i + 2]]++;
    l[Math.round(0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2])]++;
  }
  return { r, g, b, l };
}

export const FILTERS = {
  // ---------- Ajustes ----------
  brightnessContrast: {
    name: 'Brilho/Contraste',
    group: 'adjust',
    params: [
      { key: 'brightness', label: 'Brilho', min: -150, max: 150, value: 0 },
      { key: 'contrast', label: 'Contraste', min: -100, max: 100, value: 0 },
    ],
    run(src, { brightness, contrast }) {
      const c = contrast * 2.55;
      const f = (259 * (c + 255)) / (255 * (259 - c));
      return lutOp(src, makeLut((v) => f * (v + brightness - 128) + 128));
    },
  },
  levels: {
    name: 'Níveis',
    group: 'adjust',
    histogram: true,
    params: [
      { key: 'inBlack', label: 'Entrada: preto', min: 0, max: 253, value: 0 },
      { key: 'gamma', label: 'Gama (meios-tons)', min: 0.1, max: 9.99, step: 0.01, value: 1 },
      { key: 'inWhite', label: 'Entrada: branco', min: 2, max: 255, value: 255 },
      { key: 'outBlack', label: 'Saída: preto', min: 0, max: 255, value: 0 },
      { key: 'outWhite', label: 'Saída: branco', min: 0, max: 255, value: 255 },
    ],
    run(src, p) {
      const range = Math.max(1, p.inWhite - p.inBlack);
      return lutOp(src, makeLut((v) => {
        let t = clamp((v - p.inBlack) / range, 0, 1);
        t = Math.pow(t, 1 / p.gamma);
        return p.outBlack + t * (p.outWhite - p.outBlack);
      }));
    },
  },
  curves: {
    name: 'Curvas',
    group: 'adjust',
    params: [
      { key: 'shadows', label: 'Sombras', min: -100, max: 100, value: 0 },
      { key: 'midtones', label: 'Meios-tons', min: -100, max: 100, value: 0 },
      { key: 'highlights', label: 'Realces', min: -100, max: 100, value: 0 },
    ],
    run(src, { shadows, midtones, highlights }) {
      // Curva suave por pontos de controle em 25%, 50% e 75%.
      const pts = [[0, 0], [64, 64 + shadows * 0.6], [128, 128 + midtones * 0.6], [192, 192 + highlights * 0.6], [255, 255]];
      const lut = makeLut((v) => {
        let k = 0;
        while (k < pts.length - 2 && v > pts[k + 1][0]) k++;
        const p0 = pts[Math.max(0, k - 1)], p1 = pts[k], p2 = pts[k + 1], p3 = pts[Math.min(pts.length - 1, k + 2)];
        const t = (v - p1[0]) / (p2[0] - p1[0]);
        const t2 = t * t, t3 = t2 * t;
        return 0.5 * (2 * p1[1] + (-p0[1] + p2[1]) * t + (2 * p0[1] - 5 * p1[1] + 4 * p2[1] - p3[1]) * t2 + (-p0[1] + 3 * p1[1] - 3 * p2[1] + p3[1]) * t3);
      });
      return lutOp(src, lut);
    },
  },
  exposure: {
    name: 'Exposição',
    group: 'adjust',
    params: [
      { key: 'exposure', label: 'Exposição (EV)', min: -5, max: 5, step: 0.01, value: 0 },
      { key: 'offset', label: 'Deslocamento', min: -0.5, max: 0.5, step: 0.001, value: 0 },
      { key: 'gamma', label: 'Correção de gama', min: 0.1, max: 3, step: 0.01, value: 1 },
    ],
    run(src, { exposure, offset, gamma }) {
      const m = Math.pow(2, exposure);
      return lutOp(src, makeLut((v) => {
        const lin = Math.pow(v / 255, 2.2) * m + offset;
        return Math.pow(clamp(lin, 0, 1), 1 / 2.2 / gamma) * 255;
      }));
    },
  },
  hueSaturation: {
    name: 'Matiz/Saturação',
    group: 'adjust',
    params: [
      { key: 'hue', label: 'Matiz', min: -180, max: 180, value: 0 },
      { key: 'saturation', label: 'Saturação', min: -100, max: 100, value: 0 },
      { key: 'lightness', label: 'Luminosidade', min: -100, max: 100, value: 0 },
      { key: 'colorize', label: 'Colorir', type: 'checkbox', value: false },
    ],
    run(src, { hue, saturation, lightness, colorize }) {
      return pixelOp(src, (d) => {
        const hs = hue / 360, sm = saturation / 100, lm = lightness / 100;
        for (let i = 0; i < d.length; i += 4) {
          if (d[i + 3] === 0) continue;
          let [h, s, l] = rgbToHsl(d[i], d[i + 1], d[i + 2]);
          if (colorize) {
            h = ((hue + 180) / 360) % 1;
            s = clamp(0.25 + sm * 0.75, 0, 1);
          } else {
            h = (h + hs + 1) % 1;
            s = clamp(sm > 0 ? s + (1 - s) * sm * s : s * (1 + sm), 0, 1);
          }
          l = clamp(lm > 0 ? l + (1 - l) * lm : l * (1 + lm), 0, 1);
          const [r, g, b] = hslToRgb(h, s, l);
          d[i] = r; d[i + 1] = g; d[i + 2] = b;
        }
      });
    },
  },
  vibrance: {
    name: 'Vibração',
    group: 'adjust',
    params: [
      { key: 'vibrance', label: 'Vibração', min: -100, max: 100, value: 0 },
      { key: 'saturation', label: 'Saturação', min: -100, max: 100, value: 0 },
    ],
    run(src, { vibrance, saturation }) {
      return pixelOp(src, (d) => {
        const v = vibrance / 100, s = saturation / 100;
        for (let i = 0; i < d.length; i += 4) {
          const r = d[i], g = d[i + 1], b = d[i + 2];
          const max = Math.max(r, g, b), avg = (r + g + b) / 3;
          const amt = (s + v * (1 - ((max - avg) * 2) / 255));
          const gray = 0.299 * r + 0.587 * g + 0.114 * b;
          d[i] = r + (r - gray) * amt;
          d[i + 1] = g + (g - gray) * amt;
          d[i + 2] = b + (b - gray) * amt;
        }
      });
    },
  },
  colorBalance: {
    name: 'Equilíbrio de cores',
    group: 'adjust',
    params: [
      { key: 'cr', label: 'Ciano ⟷ Vermelho', min: -100, max: 100, value: 0 },
      { key: 'mg', label: 'Magenta ⟷ Verde', min: -100, max: 100, value: 0 },
      { key: 'yb', label: 'Amarelo ⟷ Azul', min: -100, max: 100, value: 0 },
      { key: 'temperature', label: 'Temperatura', min: -100, max: 100, value: 0 },
    ],
    run(src, { cr, mg, yb, temperature }) {
      const r = cr + temperature * 0.6, b = yb - temperature * 0.6;
      return lutOp(src, makeLut((v) => v + r * 0.8), makeLut((v) => v + mg * 0.8), makeLut((v) => v + b * 0.8));
    },
  },
  blackWhite: {
    name: 'Preto e branco',
    group: 'adjust',
    params: [
      { key: 'red', label: 'Vermelhos', min: -200, max: 300, value: 40 },
      { key: 'green', label: 'Verdes', min: -200, max: 300, value: 40 },
      { key: 'blue', label: 'Azuis', min: -200, max: 300, value: 20 },
    ],
    run(src, { red, green, blue }) {
      return pixelOp(src, (d) => {
        const sum = red + green + blue || 1;
        const kr = red / sum, kg = green / sum, kb = blue / sum;
        for (let i = 0; i < d.length; i += 4) {
          const v = d[i] * kr + d[i + 1] * kg + d[i + 2] * kb;
          d[i] = d[i + 1] = d[i + 2] = v;
        }
      });
    },
  },
  threshold: {
    name: 'Limiar',
    group: 'adjust',
    params: [{ key: 'level', label: 'Nível', min: 1, max: 255, value: 128 }],
    run(src, { level }) {
      return pixelOp(src, (d) => {
        for (let i = 0; i < d.length; i += 4) {
          const v = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2] >= level ? 255 : 0;
          d[i] = d[i + 1] = d[i + 2] = v;
        }
      });
    },
  },
  posterize: {
    name: 'Posterizar',
    group: 'adjust',
    params: [{ key: 'levels', label: 'Níveis', min: 2, max: 64, value: 6 }],
    run(src, { levels }) {
      const n = levels - 1;
      return lutOp(src, makeLut((v) => (Math.round((v / 255) * n) / n) * 255));
    },
  },
  invert: {
    name: 'Inverter',
    group: 'adjust',
    run: (src) => lutOp(src, makeLut((v) => 255 - v)),
  },
  desaturate: {
    name: 'Dessaturar',
    group: 'adjust',
    run: (src) =>
      pixelOp(src, (d) => {
        for (let i = 0; i < d.length; i += 4) {
          const v = 0.299 * d[i] + 0.587 * d[i + 1] + 0.114 * d[i + 2];
          d[i] = d[i + 1] = d[i + 2] = v;
        }
      }),
  },
  sepia: {
    name: 'Sépia',
    group: 'adjust',
    params: [{ key: 'amount', label: 'Intensidade', min: 0, max: 100, value: 100 }],
    run(src, { amount }) {
      const t = amount / 100;
      return pixelOp(src, (d) => {
        for (let i = 0; i < d.length; i += 4) {
          const r = d[i], g = d[i + 1], b = d[i + 2];
          d[i] = r + ((0.393 * r + 0.769 * g + 0.189 * b) - r) * t;
          d[i + 1] = g + ((0.349 * r + 0.686 * g + 0.168 * b) - g) * t;
          d[i + 2] = b + ((0.272 * r + 0.534 * g + 0.131 * b) - b) * t;
        }
      });
    },
  },
  autoTone: {
    name: 'Tom automático',
    group: 'adjust',
    run(src) {
      const h = histogram(src);
      const total = h.r.reduce((a, b) => a + b, 0);
      const clip = total * 0.005;
      const ends = (arr) => {
        let lo = 0, hi = 255, acc = 0;
        while (lo < 255 && (acc += arr[lo]) < clip) lo++;
        acc = 0;
        while (hi > 0 && (acc += arr[hi]) < clip) hi--;
        return [lo, Math.max(lo + 1, hi)];
      };
      const mk = ([lo, hi]) => makeLut((v) => ((v - lo) / (hi - lo)) * 255);
      return lutOp(src, mk(ends(h.r)), mk(ends(h.g)), mk(ends(h.b)));
    },
  },

  // ---------- Filtros ----------
  gaussianBlur: {
    name: 'Desfoque gaussiano',
    group: 'blur',
    params: [{ key: 'radius', label: 'Raio (px)', min: 0.1, max: 100, step: 0.1, value: 4 }],
    run: (src, { radius }) => blurCanvas(src, radius),
  },
  motionBlur: {
    name: 'Desfoque de movimento',
    group: 'blur',
    params: [
      { key: 'angle', label: 'Ângulo', min: -180, max: 180, value: 0 },
      { key: 'distance', label: 'Distância (px)', min: 1, max: 200, value: 20 },
    ],
    run(src, { angle, distance }) {
      const out = createCanvas(src.width, src.height);
      const ctx = out.getContext('2d');
      const steps = Math.max(2, Math.min(64, Math.round(distance)));
      const a = (angle * Math.PI) / 180;
      for (let i = 0; i < steps; i++) {
        const t = i / (steps - 1) - 0.5;
        ctx.globalAlpha = 1 / (i + 1);
        ctx.drawImage(src, Math.cos(a) * distance * t, Math.sin(a) * distance * t);
      }
      return out;
    },
  },
  radialBlur: {
    name: 'Desfoque radial (zoom)',
    group: 'blur',
    params: [{ key: 'amount', label: 'Intensidade', min: 1, max: 100, value: 20 }],
    run(src, { amount }) {
      const out = createCanvas(src.width, src.height);
      const ctx = out.getContext('2d');
      const steps = 24, cx = src.width / 2, cy = src.height / 2;
      for (let i = 0; i < steps; i++) {
        const s = 1 + (amount / 400) * (i / (steps - 1));
        ctx.globalAlpha = 1 / (i + 1);
        ctx.setTransform(s, 0, 0, s, cx - cx * s, cy - cy * s);
        ctx.drawImage(src, 0, 0);
      }
      return out;
    },
  },
  sharpen: {
    name: 'Máscara de nitidez',
    group: 'sharpen',
    params: [
      { key: 'amount', label: 'Quantidade (%)', min: 1, max: 500, value: 100 },
      { key: 'radius', label: 'Raio (px)', min: 0.1, max: 50, step: 0.1, value: 1.5 },
      { key: 'threshold', label: 'Limiar', min: 0, max: 255, value: 0 },
    ],
    run(src, { amount, radius, threshold }) {
      const blurred = blurCanvas(src, radius).getContext('2d').getImageData(0, 0, src.width, src.height).data;
      const k = amount / 100;
      return pixelOp(src, (d) => {
        for (let i = 0; i < d.length; i += 4) {
          for (let c = 0; c < 3; c++) {
            const diff = d[i + c] - blurred[i + c];
            if (Math.abs(diff) >= threshold) d[i + c] = d[i + c] + diff * k;
          }
        }
      });
    },
  },
  noise: {
    name: 'Adicionar ruído',
    group: 'noise',
    params: [
      { key: 'amount', label: 'Quantidade (%)', min: 1, max: 100, value: 12 },
      { key: 'mono', label: 'Monocromático', type: 'checkbox', value: true },
    ],
    run(src, { amount, mono }) {
      const a = amount * 2.55;
      return pixelOp(src, (d) => {
        for (let i = 0; i < d.length; i += 4) {
          if (mono) {
            const n = (Math.random() - 0.5) * a;
            d[i] += n; d[i + 1] += n; d[i + 2] += n;
          } else {
            d[i] += (Math.random() - 0.5) * a;
            d[i + 1] += (Math.random() - 0.5) * a;
            d[i + 2] += (Math.random() - 0.5) * a;
          }
        }
      });
    },
  },
  median: {
    name: 'Reduzir ruído (mediana)',
    group: 'noise',
    params: [{ key: 'radius', label: 'Raio', min: 1, max: 4, value: 1 }],
    run(src, { radius }) {
      const w = src.width, h = src.height;
      const sd = src.getContext('2d').getImageData(0, 0, w, h).data;
      return pixelOp(src, (d) => {
        const r = radius, vals = [[], [], []];
        for (let y = 0; y < h; y++) {
          for (let x = 0; x < w; x++) {
            vals[0].length = vals[1].length = vals[2].length = 0;
            for (let dy = -r; dy <= r; dy++) {
              const yy = clamp(y + dy, 0, h - 1);
              for (let dx = -r; dx <= r; dx++) {
                const i = (yy * w + clamp(x + dx, 0, w - 1)) * 4;
                vals[0].push(sd[i]); vals[1].push(sd[i + 1]); vals[2].push(sd[i + 2]);
              }
            }
            const i = (y * w + x) * 4, m = vals[0].length >> 1;
            for (let c = 0; c < 3; c++) d[i + c] = vals[c].sort((a, b) => a - b)[m];
          }
        }
      });
    },
  },
  pixelate: {
    name: 'Pixelizar (mosaico)',
    group: 'stylize',
    params: [{ key: 'size', label: 'Tamanho da célula', min: 2, max: 200, value: 12 }],
    run(src, { size }) {
      const sw = Math.max(1, Math.ceil(src.width / size)), sh = Math.max(1, Math.ceil(src.height / size));
      const small = createCanvas(sw, sh);
      const s = small.getContext('2d');
      s.imageSmoothingQuality = 'high';
      s.drawImage(src, 0, 0, sw, sh);
      const out = createCanvas(src.width, src.height);
      const ctx = out.getContext('2d');
      ctx.imageSmoothingEnabled = false;
      ctx.drawImage(small, 0, 0, sw * size, sh * size);
      return out;
    },
  },
  edges: {
    name: 'Detectar bordas',
    group: 'stylize',
    run: (src) => convolve(src, [-1, -1, -1, -1, 8, -1, -1, -1, -1], { divisor: 1 }),
  },
  emboss: {
    name: 'Relevo',
    group: 'stylize',
    run: (src) => convolve(src, [-2, -1, 0, -1, 1, 1, 0, 1, 2], { divisor: 1, offset: 0 }),
  },
  vignette: {
    name: 'Vinheta',
    group: 'stylize',
    params: [
      { key: 'amount', label: 'Intensidade', min: 0, max: 100, value: 60 },
      { key: 'size', label: 'Tamanho', min: 10, max: 100, value: 55 },
    ],
    run(src, { amount, size }) {
      const out = cloneCanvas(src);
      const ctx = out.getContext('2d');
      const w = src.width, h = src.height, r = Math.hypot(w, h) / 2;
      const g = ctx.createRadialGradient(w / 2, h / 2, (r * size) / 100 * 0.6, w / 2, h / 2, r);
      g.addColorStop(0, 'rgba(0,0,0,0)');
      g.addColorStop(1, `rgba(0,0,0,${amount / 100})`);
      ctx.globalCompositeOperation = 'source-atop';
      ctx.fillStyle = g;
      ctx.fillRect(0, 0, w, h);
      return out;
    },
  },
  oilPaint: {
    name: 'Pintura a óleo',
    group: 'stylize',
    params: [
      { key: 'radius', label: 'Raio', min: 1, max: 6, value: 3 },
      { key: 'levels', label: 'Níveis de intensidade', min: 4, max: 40, value: 20 },
    ],
    run(src, { radius, levels }) {
      const w = src.width, h = src.height;
      const sd = src.getContext('2d').getImageData(0, 0, w, h).data;
      const lv = new Uint8Array(w * h);
      for (let p = 0, i = 0; p < lv.length; p++, i += 4) lv[p] = (((sd[i] + sd[i + 1] + sd[i + 2]) / 3) * (levels - 1)) / 255;
      return pixelOp(src, (d) => {
        const cnt = new Uint32Array(levels), sr = new Uint32Array(levels), sg = new Uint32Array(levels), sb = new Uint32Array(levels);
        for (let y = 0; y < h; y++) {
          for (let x = 0; x < w; x++) {
            cnt.fill(0); sr.fill(0); sg.fill(0); sb.fill(0);
            for (let dy = -radius; dy <= radius; dy++) {
              const yy = clamp(y + dy, 0, h - 1);
              for (let dx = -radius; dx <= radius; dx++) {
                const p = yy * w + clamp(x + dx, 0, w - 1), i = p * 4, l = lv[p];
                cnt[l]++; sr[l] += sd[i]; sg[l] += sd[i + 1]; sb[l] += sd[i + 2];
              }
            }
            let best = 0;
            for (let k = 1; k < levels; k++) if (cnt[k] > cnt[best]) best = k;
            const i = (y * w + x) * 4;
            d[i] = sr[best] / cnt[best]; d[i + 1] = sg[best] / cnt[best]; d[i + 2] = sb[best] / cnt[best];
          }
        }
      });
    },
  },
  dropShadow: {
    name: 'Sombra projetada',
    group: 'stylize',
    params: [
      { key: 'dx', label: 'Deslocamento X', min: -100, max: 100, value: 8 },
      { key: 'dy', label: 'Deslocamento Y', min: -100, max: 100, value: 8 },
      { key: 'blur', label: 'Desfoque', min: 0, max: 60, value: 10 },
      { key: 'opacity', label: 'Opacidade (%)', min: 0, max: 100, value: 60 },
      { key: 'color', label: 'Cor', type: 'color', value: '#000000' },
    ],
    run(src, { dx, dy, blur, opacity, color }) {
      const out = createCanvas(src.width, src.height);
      const ctx = out.getContext('2d');
      const sh = createCanvas(src.width, src.height);
      const s = sh.getContext('2d');
      s.drawImage(src, 0, 0);
      s.globalCompositeOperation = 'source-in';
      s.fillStyle = color;
      s.fillRect(0, 0, sh.width, sh.height);
      ctx.globalAlpha = opacity / 100;
      ctx.filter = blur > 0 ? `blur(${blur}px)` : 'none';
      ctx.drawImage(sh, dx, dy);
      ctx.filter = 'none';
      ctx.globalAlpha = 1;
      ctx.drawImage(src, 0, 0);
      return out;
    },
  },
  strokeOutline: {
    name: 'Contorno (traçado)',
    group: 'stylize',
    params: [
      { key: 'size', label: 'Espessura (px)', min: 1, max: 50, value: 4 },
      { key: 'color', label: 'Cor', type: 'color', value: '#000000' },
    ],
    run(src, { size, color }) {
      const out = createCanvas(src.width, src.height);
      const ctx = out.getContext('2d');
      const tint = createCanvas(src.width, src.height);
      const t = tint.getContext('2d');
      t.drawImage(src, 0, 0);
      t.globalCompositeOperation = 'source-in';
      t.fillStyle = color;
      t.fillRect(0, 0, tint.width, tint.height);
      const steps = Math.max(8, Math.round(size * 4));
      for (let i = 0; i < steps; i++) {
        const a = (i / steps) * Math.PI * 2;
        ctx.drawImage(tint, Math.cos(a) * size, Math.sin(a) * size);
      }
      ctx.drawImage(src, 0, 0);
      return out;
    },
  },
};

export const FILTER_MENU = {
  adjust: [
    'brightnessContrast', 'levels', 'curves', 'exposure', '-', 'vibrance', 'hueSaturation', 'colorBalance', 'blackWhite', '-',
    'invert', 'posterize', 'threshold', 'sepia', '-', 'desaturate', 'autoTone',
  ],
  filters: [
    ['Desfoque', ['gaussianBlur', 'motionBlur', 'radialBlur']],
    ['Nitidez', ['sharpen']],
    ['Ruído', ['noise', 'median']],
    ['Estilização', ['pixelate', 'edges', 'emboss', 'oilPaint', 'vignette']],
    ['Efeitos de camada', ['dropShadow', 'strokeOutline']],
  ],
};
