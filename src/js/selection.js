// Seleções baseadas em máscara (canvas com alfa), contorno "formigas marchando" e preenchimento por inundação.
import { createCanvas, cloneCanvas } from './utils.js';

export function maskFromPath(w, h, path, feather = 0) {
  const m = createCanvas(w, h);
  const ctx = m.getContext('2d');
  if (feather > 0) ctx.filter = `blur(${feather}px)`;
  ctx.fillStyle = '#fff';
  ctx.fill(path);
  return m;
}

export function maskFull(w, h) {
  const m = createCanvas(w, h);
  const ctx = m.getContext('2d');
  ctx.fillStyle = '#fff';
  ctx.fillRect(0, 0, w, h);
  return m;
}

// mode: 'new' | 'add' | 'sub' | 'int'
export function combineMasks(oldMask, newMask, mode) {
  if (!oldMask || mode === 'new') return isEmptyMask(newMask) ? null : newMask;
  const out = cloneCanvas(oldMask);
  const ctx = out.getContext('2d');
  ctx.globalCompositeOperation =
    mode === 'add' ? 'source-over' : mode === 'sub' ? 'destination-out' : 'destination-in';
  ctx.drawImage(newMask, 0, 0);
  return isEmptyMask(out) ? null : out;
}

export function invertMask(mask, w, h) {
  const out = maskFull(w, h);
  if (!mask) return null;
  const ctx = out.getContext('2d');
  ctx.globalCompositeOperation = 'destination-out';
  ctx.drawImage(mask, 0, 0);
  return isEmptyMask(out) ? null : out;
}

export function translateMask(mask, dx, dy) {
  const out = createCanvas(mask.width, mask.height);
  out.getContext('2d').drawImage(mask, Math.round(dx), Math.round(dy));
  return out;
}

export function isEmptyMask(mask) {
  return !maskBounds(mask);
}

export function alphaBounds(canvas) {
  const { width: w, height: h } = canvas;
  const d = canvas.getContext('2d').getImageData(0, 0, w, h).data;
  let x0 = w, y0 = h, x1 = -1, y1 = -1;
  for (let y = 0; y < h; y++) {
    let row = y * w * 4 + 3;
    for (let x = 0; x < w; x++, row += 4) {
      if (d[row] > 0) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        y1 = y;
      }
    }
  }
  if (x1 < 0) return null;
  return { x: x0, y: y0, w: x1 - x0 + 1, h: y1 - y0 + 1 };
}

export const maskBounds = alphaBounds;

// Constrói um Path2D com as bordas da máscara (em coordenadas do documento).
export function buildAntsPath(mask) {
  if (!mask) return null;
  if (mask._ants) return mask._ants;
  const { width: w, height: h } = mask;
  const d = mask.getContext('2d').getImageData(0, 0, w, h).data;
  const sel = new Uint8Array(w * h);
  for (let i = 0, j = 3; i < sel.length; i++, j += 4) sel[i] = d[j] > 127 ? 1 : 0;
  const p = new Path2D();
  const at = (x, y) => (x < 0 || y < 0 || x >= w || y >= h ? 0 : sel[y * w + x]);
  // Bordas horizontais
  for (let y = 0; y <= h; y++) {
    let start = -1;
    for (let x = 0; x <= w; x++) {
      const edge = x < w && at(x, y - 1) !== at(x, y);
      if (edge && start < 0) start = x;
      else if (!edge && start >= 0) {
        p.moveTo(start, y);
        p.lineTo(x, y);
        start = -1;
      }
    }
  }
  // Bordas verticais
  for (let x = 0; x <= w; x++) {
    let start = -1;
    for (let y = 0; y <= h; y++) {
      const edge = y < h && at(x - 1, y) !== at(x, y);
      if (edge && start < 0) start = y;
      else if (!edge && start >= 0) {
        p.moveTo(x, start);
        p.lineTo(x, y);
        start = -1;
      }
    }
  }
  mask._ants = p;
  return p;
}

// Retorna Uint8Array (1 = pixel selecionado) para pixels semelhantes à cor em (sx, sy).
export function floodRegion(imageData, sx, sy, tolerance, contiguous = true) {
  const { width: w, height: h, data } = imageData;
  const out = new Uint8Array(w * h);
  if (sx < 0 || sy < 0 || sx >= w || sy >= h) return out;
  const i0 = (sy * w + sx) * 4;
  const r0 = data[i0], g0 = data[i0 + 1], b0 = data[i0 + 2], a0 = data[i0 + 3];
  const tol = tolerance * 2.55;
  const match = (p) => {
    const i = p * 4;
    const a = data[i + 3];
    if (a0 === 0 && a === 0) return true;
    return (
      Math.abs(data[i] - r0) <= tol &&
      Math.abs(data[i + 1] - g0) <= tol &&
      Math.abs(data[i + 2] - b0) <= tol &&
      Math.abs(a - a0) <= tol
    );
  };
  if (!contiguous) {
    for (let p = 0; p < w * h; p++) if (match(p)) out[p] = 1;
    return out;
  }
  const seen = new Uint8Array(w * h);
  const stack = [sx, sy];
  while (stack.length) {
    const y = stack.pop();
    let x = stack.pop();
    let p = y * w + x;
    while (x >= 0 && !seen[p] && match(p)) { x--; p--; }
    x++; p++;
    let up = false, down = false;
    while (x < w && !seen[p] && match(p)) {
      out[p] = 1;
      seen[p] = 1;
      if (y > 0) {
        const q = p - w;
        if (!seen[q] && match(q)) {
          if (!up) { stack.push(x, y - 1); up = true; }
        } else up = false;
      }
      if (y < h - 1) {
        const q = p + w;
        if (!seen[q] && match(q)) {
          if (!down) { stack.push(x, y + 1); down = true; }
        } else down = false;
      }
      x++; p++;
    }
  }
  return out;
}

export function maskFromBits(bits, w, h) {
  const m = createCanvas(w, h);
  const ctx = m.getContext('2d');
  const id = ctx.createImageData(w, h);
  const d = id.data;
  for (let i = 0, j = 0; i < bits.length; i++, j += 4) {
    if (bits[i]) {
      d[j] = d[j + 1] = d[j + 2] = 255;
      d[j + 3] = 255;
    }
  }
  ctx.putImageData(id, 0, 0);
  return m;
}

// Copia do canvas apenas a área da máscara.
export function maskedCopy(src, mask) {
  const out = cloneCanvas(src);
  if (!mask) return out;
  const ctx = out.getContext('2d');
  ctx.globalCompositeOperation = 'destination-in';
  ctx.drawImage(mask, 0, 0);
  return out;
}

// Remove do canvas a área da máscara (retorna novo canvas).
export function maskedHole(src, mask) {
  const out = cloneCanvas(src);
  const ctx = out.getContext('2d');
  if (!mask) {
    ctx.clearRect(0, 0, out.width, out.height);
    return out;
  }
  ctx.globalCompositeOperation = 'destination-out';
  ctx.drawImage(mask, 0, 0);
  return out;
}

export function pointInMask(mask, x, y) {
  if (!mask) return false;
  x = Math.floor(x); y = Math.floor(y);
  if (x < 0 || y < 0 || x >= mask.width || y >= mask.height) return false;
  return mask.getContext('2d').getImageData(x, y, 1, 1).data[3] > 127;
}
