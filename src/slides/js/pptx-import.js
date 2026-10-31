// Importador de .pptx: slides, herança de layout/mestre, textos com formatação, formas, imagens (com recorte),
// tabelas, gráficos, grupos, SmartArt, fundos, anotações, transições e animações de entrada.
import {
  openPackage, kids, kid, path, attr, num, first, all, rid, readColor, cssColor, readTheme, cssFont, isWebImage,
  unwrapAlternate, EMU_PER_PT, hexRgb, rgbHex,
} from '../../shared/ooxml.js';
import { escapeHtml } from '../../shared/ui.js';
import { W, H, uid } from './themes.js';

const DEG = 60000;

const GEOMETRY = {
  rect: 'rect', roundRect: 'round', round1Rect: 'round', round2SameRect: 'round', snipRoundRect: 'round', flowChartAlternateProcess: 'round',
  ellipse: 'ellipse', flowChartConnector: 'ellipse', donut: 'ellipse',
  triangle: 'triangle', rtTriangle: 'triangle', flowChartExtract: 'triangle',
  diamond: 'diamond', flowChartDecision: 'diamond',
  hexagon: 'hexagon', octagon: 'hexagon', flowChartPreparation: 'hexagon',
  star5: 'star', star4: 'star', star6: 'star', star7: 'star', star8: 'star', star10: 'star', star12: 'star',
  rightArrow: 'arrow', leftArrow: 'arrow', notchedRightArrow: 'arrow', stripedRightArrow: 'arrow',
  chevron: 'chevron', homePlate: 'chevron',
  wedgeRoundRectCallout: 'callout', wedgeRectCallout: 'callout', wedgeEllipseCallout: 'callout', cloudCallout: 'callout',
  line: 'line', straightConnector1: 'line', bentConnector2: 'line', bentConnector3: 'line', curvedConnector3: 'line',
};

const TRANSITIONS = {
  fade: 'fade', dissolve: 'fade', random: 'fade', cut: 'none', push: 'push', pull: 'push', wipe: 'cover', cover: 'cover',
  split: 'cover', reveal: 'cover', zoom: 'zoom', newsflash: 'zoom', flip: 'flip', cube: 'flip', rotate: 'flip', ferris: 'flip',
  gallery: 'flip', conveyor: 'push', pan: 'push', wheel: 'zoom', circle: 'zoom', morph: 'fade', vortex: 'zoom', shred: 'fade',
  prism: 'flip', doors: 'cover', window: 'cover', glitter: 'fade', honeycomb: 'fade', flash: 'fade', ripple: 'zoom',
};

export async function importPptx(data, fileTitle = 'Apresentação') {
  const pkg = await openPackage(data);
  const root = await pkg.rels('');
  const presPath = [...root.values()].find((r) => r.type === 'officeDocument')?.target || 'ppt/presentation.xml';
  const presXml = await pkg.xml(presPath);
  if (!presXml) throw new Error('Arquivo do PowerPoint inválido.');
  const presRels = await pkg.rels(presPath);
  const sz = first(presXml, 'sldSz');
  const cx = num(sz, 'cx', 12192000), cy = num(sz, 'cy', 6858000);
  const scale = Math.min(W / cx, H / cy);
  const imp = new Importer(pkg, scale, (W - cx * scale) / 2, (H - cy * scale) / 2, cx, cy);
  imp.defaultText = kid(presXml.documentElement, 'defaultTextStyle');
  const tableStyles = [...presRels.values()].find((r) => r.type === 'tableStyles');
  if (tableStyles) imp.tableStyles = await pkg.xml(tableStyles.target);

  const slides = [];
  for (const sid of all(first(presXml, 'sldIdLst'), 'sldId')) {
    const rel = presRels.get(rid(sid));
    if (!rel) continue;
    try {
      slides.push(await imp.slide(rel.target));
    } catch (err) {
      console.warn('Slide ignorado:', rel.target, err);
    }
  }
  const core = await pkg.xml('docProps/core.xml');
  const title = core ? first(core, 'title')?.textContent?.trim() : '';
  const pres = {
    title: title || fileTitle,
    theme: 'custom',
    customTheme: imp.firstTheme || undefined,
    slides: slides.length ? slides : [],
    assets: imp.assets,
  };
  if (!pres.customTheme) pres.theme = 'claro';
  return pres;
}

class Importer {
  constructor(pkg, scale, offX, offY, cx, cy) {
    this.pkg = pkg;
    this.s = scale;
    this.offX = offX;
    this.offY = offY;
    this.cx = cx;
    this.cy = cy;
    this.assets = {};
    this.assetByPath = new Map();
    this.masters = new Map();
    this.layouts = new Map();
  }

  // ---- Mestre, layout e tema ----------------------------------------------------

  async master(path_) {
    if (this.masters.has(path_)) return this.masters.get(path_);
    const xml = await this.pkg.xml(path_);
    const rels = await this.pkg.rels(path_);
    const themeRel = [...rels.values()].find((r) => r.type === 'theme');
    const theme = await readTheme(this.pkg, themeRel?.target);
    const themeXml = themeRel ? await this.pkg.xml(themeRel.target) : null;
    const clrMap = {};
    const cm = first(xml, 'clrMap');
    if (cm) for (const a of cm.attributes) clrMap[a.localName] = a.value;
    const m = {
      path: path_, xml, rels, theme, themeXml, clrMap,
      txStyles: first(xml, 'txStyles'),
      ph: this.indexPlaceholders(xml),
    };
    this.masters.set(path_, m);
    if (!this.firstTheme) this.firstTheme = await this.customTheme(m);
    return m;
  }

  async layout(path_) {
    if (this.layouts.has(path_)) return this.layouts.get(path_);
    const xml = await this.pkg.xml(path_);
    const rels = await this.pkg.rels(path_);
    const masterRel = [...rels.values()].find((r) => r.type === 'slideMaster');
    const l = { path: path_, xml, rels, master: await this.master(masterRel.target), ph: this.indexPlaceholders(xml) };
    this.layouts.set(path_, l);
    return l;
  }

  indexPlaceholders(xml) {
    const byIdx = new Map(), byType = new Map();
    const tree = first(xml, 'spTree');
    for (const sp of kids(tree)) {
      const ph = first(kid(sp, 'nvSpPr') || kid(sp, 'nvPicPr') || kid(sp, 'nvGraphicFramePr'), 'ph');
      if (!ph) continue;
      const type = attr(ph, 'type') || 'body';
      const idx = attr(ph, 'idx');
      if (idx !== null && !byIdx.has(idx)) byIdx.set(idx, sp);
      if (!byType.has(type)) byType.set(type, sp);
    }
    return { byIdx, byType };
  }

  findPh(index, ph) {
    if (!ph || !index) return null;
    const type = attr(ph, 'type') || 'body';
    const idx = attr(ph, 'idx');
    const alt = { ctrTitle: 'title', subTitle: 'body', obj: 'body' }[type];
    return (idx !== null && index.byIdx.get(idx)) || index.byType.get(type) || (alt && index.byType.get(alt)) || null;
  }

  async customTheme(m) {
    const t = m.theme;
    const c = (name) => t.colors[m.clrMap[name] || name] || null;
    const bg = await this.background([m.xml], m, null);
    const titleLvl = path(m.txStyles, 'titleStyle', 'lvl1pPr', 'defRPr');
    const titleColor = titleLvl ? cssColor(readColor(kid(titleLvl, 'solidFill'), t, m.clrMap)) : null;
    const accent = t.colors.accent1 || '#3b6cf6';
    const [r, g, b] = hexRgb(accent);
    return {
      name: first(m.themeXml, 'theme') ? attr(m.themeXml.documentElement, 'name') || 'Tema importado' : 'Tema importado',
      bg: bg || { type: 'color', color: c('bg1') || '#ffffff' },
      title: titleColor || c('tx1') || '#1d2129',
      text: c('tx1') || '#333333',
      accent,
      accent2: rgbHex(r + (255 - r) * 0.82, g + (255 - g) * 0.82, b + (255 - b) * 0.82),
      head: cssFont(t.major).replace(/"/g, "'"),
      body: cssFont(t.minor).replace(/"/g, "'"),
      imported: true,
    };
  }

  // ---- Slide ----------------------------------------------------------------------

  async slide(slidePath) {
    const xml = await this.pkg.xml(slidePath);
    const rels = await this.pkg.rels(slidePath);
    const layoutRel = [...rels.values()].find((r) => r.type === 'slideLayout');
    const layout = layoutRel ? await this.layout(layoutRel.target) : null;
    const master = layout?.master;
    const clrMap = { ...(master?.clrMap || {}) };
    const ovr = first(first(xml, 'clrMapOvr'), 'overrideClrMapping');
    if (ovr) for (const a of ovr.attributes) clrMap[a.localName] = a.value;
    const ctx = { slidePath, rels, layout, master, theme: master?.theme || { colors: {} }, clrMap, elements: [], spidMap: new Map() };

    const sld = xml.documentElement;
    const show = (n) => attr(first(n, 'cSld')?.parentNode, 'showMasterSp') !== '0';
    // Formas decorativas do mestre e do layout (não marcadores de posição)
    if (master && layout && show(xml) && attr(layout.xml.documentElement, 'showMasterSp') !== '0') await this.tree(first(master.xml, 'spTree'), { ...ctx, rels: master.rels, level: 'master' }, identity);
    if (layout && show(xml)) await this.tree(first(layout.xml, 'spTree'), { ...ctx, rels: layout.rels, level: 'layout' }, identity);
    await this.tree(first(xml, 'spTree'), { ...ctx, level: 'slide' }, identity);

    const slide = {
      id: uid('s'),
      layout: 'blank',
      bg: await this.background([sld, layout?.xml, master?.xml].filter(Boolean), master, ctx),
      elements: ctx.elements,
      notes: await this.notes(rels),
      transition: this.transition(sld),
    };
    if (attr(sld, 'show') === '0') slide.hidden = true;
    this.animations(sld, ctx);
    return slide;
  }

  async background(xmls, master, ctx) {
    for (const x of xmls) {
      const bg = first(first(x, 'cSld'), 'bg');
      if (!bg) continue;
      const clrMap = ctx?.clrMap || master?.clrMap;
      const theme = master?.theme;
      const bgPr = kid(bg, 'bgPr');
      const rels = x === xmls[0] && ctx ? ctx.rels : await this.relsOfXml(x, master, ctx);
      if (bgPr) {
        const f = await this.fill(bgPr, theme, clrMap, rels);
        if (f) return f;
      }
      const ref = kid(bg, 'bgRef');
      if (ref && master?.themeXml) {
        const idx = num(ref, 'idx', 0);
        const phClr = readColor(ref, theme, clrMap)?.hex;
        const lst = first(master.themeXml, idx >= 1001 ? 'bgFillStyleLst' : 'fillStyleLst');
        const item = kids(lst)[idx >= 1001 ? idx - 1001 : idx - 1];
        if (item) {
          const wrap = { children: [item], localName: 'wrap' };
          const f = await this.fill(wrap, theme, clrMap, master.rels, phClr);
          if (f) return f;
        }
        if (phClr) return { type: 'color', color: phClr };
      }
    }
    return null;
  }

  async relsOfXml(x, master, ctx) {
    if (ctx?.layout?.xml === x) return ctx.layout.rels;
    if (master?.xml === x) return master.rels;
    return ctx?.rels || new Map();
  }

  // Preenchimento DrawingML -> { type: 'color'|'gradient'|'image', ... } (ou null)
  async fill(node, theme, clrMap, rels, phClr = null) {
    for (const f of kids(node)) {
      switch (f.localName) {
        case 'solidFill': {
          const c = readColor(f, theme, clrMap, phClr);
          return c ? { type: 'color', color: cssColor(c) } : null;
        }
        case 'gradFill': {
          const stops = kids(kid(f, 'gsLst'), 'gs').map((g) => ({ pos: num(g, 'pos', 0), c: readColor(g, theme, clrMap, phClr) })).filter((x) => x.c).sort((a, b) => a.pos - b.pos);
          if (!stops.length) return null;
          const lin = kid(f, 'lin');
          const angle = lin ? Math.round(num(lin, 'ang', 0) / DEG + 90) : 180;
          return { type: 'gradient', from: cssColor(stops[0].c), to: cssColor(stops[stops.length - 1].c), angle };
        }
        case 'blipFill': {
          const asset = await this.blipAsset(f, rels);
          return asset ? { type: 'image', asset } : null;
        }
        case 'pattFill': {
          const c = readColor(kid(f, 'fgClr'), theme, clrMap, phClr);
          return c ? { type: 'color', color: cssColor(c) } : null;
        }
        case 'noFill': return { type: 'none' };
        default: break;
      }
    }
    return null;
  }

  async blipAsset(blipFill, rels, crop = true) {
    const blip = kid(blipFill, 'blip');
    if (!blip) return null;
    // Preferir SVG quando houver
    const svg = first(blip, 'svgBlip');
    const id = (svg && attr(svg, 'embed')) || attr(blip, 'embed') || attr(blip, 'link');
    const rel = rels.get(id);
    if (!rel || rel.external || !isWebImage(rel.target)) return null;
    let key = rel.target;
    const src = kid(blipFill, 'srcRect');
    const rect = src && crop ? ['l', 't', 'r', 'b'].map((k) => num(src, k, 0) / 100000) : null;
    const cropping = rect && rect.some((v) => Math.abs(v) > 0.001);
    if (cropping) key += `#${rect.join(',')}`;
    if (this.assetByPath.has(key)) return this.assetByPath.get(key);
    let url = await this.pkg.dataUrl(rel.target);
    if (!url) return null;
    if (cropping) url = await cropImage(url, rect).catch(() => url);
    const asset = uid('a');
    this.assets[asset] = url;
    this.assetByPath.set(key, asset);
    return asset;
  }

  // ---- Árvore de formas ---------------------------------------------------------------

  async tree(tree, ctx, T) {
    if (!tree) return;
    for (let n of kids(tree)) {
      n = unwrapAlternate(n);
      if (!n) continue;
      const ln = n.localName;
      if (ln === 'Choice' || ln === 'Fallback') { await this.tree(n, ctx, T); continue; }
      try {
        if (ln === 'sp') await this.shape(n, ctx, T);
        else if (ln === 'pic') await this.picture(n, ctx, T);
        else if (ln === 'cxnSp') await this.shape(n, ctx, T, true);
        else if (ln === 'grpSp') {
          const x = path(n, 'grpSpPr', 'xfrm');
          const G = groupTransform(x, T);
          await this.tree(n, ctx, G);
        } else if (ln === 'graphicFrame') await this.frame(n, ctx, T);
      } catch (err) {
        console.warn('Forma ignorada:', err);
      }
    }
  }

  // Posição final em px do slide
  box(xfrm, T) {
    const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext');
    if (!off || !ext) return null;
    const r = T({ x: num(off, 'x', 0), y: num(off, 'y', 0), w: num(ext, 'cx', 0), h: num(ext, 'cy', 0) });
    return {
      x: Math.round(this.offX + r.x * this.s), y: Math.round(this.offY + r.y * this.s),
      w: Math.max(1, Math.round(r.w * this.s)), h: Math.max(1, Math.round(r.h * this.s)),
      rot: Math.round((num(xfrm, 'rot', 0) / DEG) * 10) / 10,
      flipH: attr(xfrm, 'flipH') === '1', flipV: attr(xfrm, 'flipV') === '1',
    };
  }

  // Marcadores de posição herdados do layout e do mestre
  inherited(n, ctx) {
    const nv = kid(n, 'nvSpPr') || kid(n, 'nvPicPr') || kid(n, 'nvGraphicFramePr') || kid(n, 'nvCxnSpPr');
    const ph = first(nv, 'ph');
    if (!ph) return { ph: null, chain: [] };
    const chain = [];
    if (ctx.level === 'slide') {
      const l = this.findPh(ctx.layout?.ph, ph);
      if (l) chain.push(l);
      const m = this.findPh(ctx.master?.ph, l ? first(kid(l, 'nvSpPr') || kid(l, 'nvPicPr'), 'ph') || ph : ph);
      if (m) chain.push(m);
    } else if (ctx.level === 'layout') {
      const m = this.findPh(ctx.master?.ph, ph);
      if (m) chain.push(m);
    }
    return { ph, chain };
  }

  async shape(n, ctx, T, connector = false) {
    const { ph, chain } = this.inherited(n, ctx);
    // Marcadores de posição do mestre/layout não aparecem nos slides
    if (ph && ctx.level !== 'slide') return;
    const phType = ph ? attr(ph, 'type') || 'body' : null;
    if (phType && ['dt', 'ftr', 'sldNum', 'hdr'].includes(phType) && !first(kid(n, 'txBody'), 't')) return;
    const spPr = kid(n, 'spPr');
    const xfrm = kid(spPr, 'xfrm') || chain.map((c) => path(c, 'spPr', 'xfrm')).find(Boolean);
    const b = xfrm && this.box(xfrm, T);
    if (!b) return;
    const { theme, clrMap } = ctx;
    const style = kid(n, 'style');
    const geomNode = kid(spPr, 'prstGeom') || chain.map((c) => path(c, 'spPr', 'prstGeom')).find(Boolean);
    const prst = geomNode ? attr(geomNode, 'prst') : kid(spPr, 'custGeom') ? 'rect' : 'rect';

    // Preenchimento
    let fill = await this.fill(spPr, theme, clrMap, ctx.rels);
    if (!fill) for (const c of chain) { fill = await this.fill(kid(c, 'spPr'), theme, clrMap, ctx.rels); if (fill) break; }
    if (!fill && style) {
      const ref = kid(style, 'fillRef');
      if (ref && num(ref, 'idx', 0) > 0) { const c = readColor(ref, theme, clrMap); if (c) fill = { type: 'color', color: cssColor(c) }; }
    }
    // Contorno
    const ln = kid(spPr, 'ln') || chain.map((c) => path(c, 'spPr', 'ln')).find(Boolean);
    let stroke = null, strokeW = 0;
    if (ln && !kid(ln, 'noFill')) {
      const c = readColor(kid(ln, 'solidFill'), theme, clrMap) || (kid(ln, 'gradFill') ? readColor(first(ln, 'gs'), theme, clrMap) : null);
      if (c) stroke = cssColor(c);
      strokeW = num(ln, 'w', 12700) * this.s;
    }
    if (!stroke && style && !(ln && kid(ln, 'noFill'))) {
      const ref = kid(style, 'lnRef');
      if (ref && num(ref, 'idx', 0) > 0) { const c = readColor(ref, theme, clrMap); if (c) { stroke = cssColor(c); strokeW = strokeW || Math.max(1.5, 9525 * num(ref, 'idx', 1) * this.s); } }
    }
    if (stroke && !strokeW) strokeW = 12700 * this.s;

    const kind = GEOMETRY[prst] || (connector ? 'line' : 'rect');
    const txBody = kid(n, 'txBody');
    const text = txBody ? this.text(txBody, chain, ctx, phType, style) : null;
    const hasFill = fill && fill.type !== 'none';
    const id = attr(first(n, 'cNvPr'), 'id');

    // Forma preenchida com imagem
    if (fill?.type === 'image') {
      const e = { id: uid(), type: 'image', ...pos(b), rot: b.rot, opacity: 1, image: { asset: fill.asset, fit: 'fill', radius: kind === 'ellipse' ? Math.round(Math.min(b.w, b.h) / 2) : kind === 'round' ? 16 : 0 } };
      this.add(ctx, e, id);
      if (text?.html) this.add(ctx, { id: uid(), type: 'text', ...pos(b), rot: b.rot, opacity: 1, role: 'body', text: text.props }, null);
      return;
    }

    if (kind === 'line' || connector) {
      const color = stroke || (hasFill ? fill.color : null);
      if (!color) return;
      const w = b.w, h = b.h;
      const len = Math.hypot(w, h);
      let ang = (Math.atan2(h, w) * 180) / Math.PI;
      if (b.flipH !== b.flipV) ang = -ang;
      const cxp = b.x + w / 2, cyp = b.y + h / 2;
      const thick = Math.max(2, Math.round(strokeW || 3));
      this.add(ctx, {
        id: uid(), type: 'shape', x: Math.round(cxp - len / 2), y: Math.round(cyp - thick / 2), w: Math.round(len), h: thick,
        rot: Math.round((ang + b.rot) * 10) / 10, opacity: 1,
        shape: { kind: 'line', fill: null, stroke: color, strokeW: thick, radius: 0 },
      }, id);
      return;
    }

    if (!hasFill && !stroke) {
      if (!text?.html) return;
      this.add(ctx, { id: uid(), type: 'text', role: phType === 'title' || phType === 'ctrTitle' ? 'title' : phType === 'subTitle' ? 'subtitle' : 'body', ...pos(b), rot: b.rot, opacity: 1, text: text.props }, id);
      return;
    }
    let rot = b.rot;
    if (kind === 'arrow' && (prst === 'leftArrow' || b.flipH)) rot += 180;
    if (kind === 'triangle' && b.flipV) rot += 180;
    const adj = parseInt((attr(first(geomNode, 'gd'), 'fmla') || '').replace(/\D/g, ''), 10);
    const radius = kind === 'round' ? Math.round((Math.min(b.w, b.h) * (Number.isFinite(adj) ? adj : 16667)) / 100000) : 0;
    const e = {
      id: uid(), type: 'shape', ...pos(b), rot, opacity: 1,
      shape: { kind, fill: hasFill ? (fill.type === 'gradient' ? `linear-gradient(${fill.angle}deg, ${fill.from}, ${fill.to})` : fill.color) : 'transparent', stroke, strokeW: Math.round(strokeW * 10) / 10, radius },
      text: text?.props || { html: '', size: 36, color: '#ffffff', align: 'center', valign: 'middle', bold: false, font: null, line: 1.2 },
    };
    if (kind === 'round' && !radius) e.shape.radius = 24;
    if (kind === 'rect') e.shape.radius = 0;
    if (first(spPr, 'outerShdw') || first(spPr, 'prstShdw')) e.shadow = true;
    const alpha = readColor(kid(spPr, 'solidFill'), theme, clrMap)?.alpha;
    if (alpha !== undefined && alpha < 1 && !hasFill) e.opacity = alpha;
    this.add(ctx, e, id);
  }

  add(ctx, e, spid) {
    ctx.elements.push(e);
    if (spid && ctx.level === 'slide') ctx.spidMap.set(spid, e);
  }

  async picture(n, ctx, T) {
    const { ph, chain } = this.inherited(n, ctx);
    if (ph && ctx.level !== 'slide') return;
    const spPr = kid(n, 'spPr');
    const xfrm = kid(spPr, 'xfrm') || chain.map((c) => path(c, 'spPr', 'xfrm')).find(Boolean);
    const b = xfrm && this.box(xfrm, T);
    if (!b) return;
    const blipFill = kid(n, 'blipFill');
    const asset = blipFill ? await this.blipAsset(blipFill, ctx.rels) : null;
    const prst = attr(kid(spPr, 'prstGeom'), 'prst');
    const amt = first(blipFill, 'alphaModFix');
    const e = {
      id: uid(), type: 'image', ...pos(b), rot: b.rot, opacity: amt ? num(amt, 'amt', 100000) / 100000 : 1,
      image: { asset, fit: 'fill', radius: prst === 'ellipse' ? Math.round(Math.min(b.w, b.h) / 2) : prst === 'roundRect' ? 16 : 0 },
    };
    if (b.flipH) e.image.flipH = true;
    if (b.flipV) e.image.flipV = true;
    if (!asset) e.ph = 'Imagem em formato não suportado (EMF/WMF/TIFF)';
    this.add(ctx, e, attr(first(n, 'cNvPr'), 'id'));
  }

  async frame(n, ctx, T) {
    const xfrm = kid(n, 'xfrm');
    const b = xfrm && this.box(xfrm, T);
    const gd = first(n, 'graphicData');
    if (!gd || !b) return;
    const uri = attr(gd, 'uri') || '';
    const id = attr(first(n, 'cNvPr'), 'id');
    if (uri.endsWith('/table')) {
      const e = this.table(first(gd, 'tbl'), b, ctx);
      if (e) this.add(ctx, e, id);
    } else if (uri.endsWith('/chart')) {
      const rel = ctx.rels.get(attr(first(gd, 'chart'), 'id'));
      const e = rel ? await this.chart(rel.target, b, ctx) : null;
      if (e) this.add(ctx, e, id);
    } else if (uri.endsWith('/diagram')) {
      // SmartArt: usa o desenho já calculado pelo PowerPoint (ppt/diagrams/drawingN.xml)
      const dm = first(gd, 'relIds');
      const dataRel = dm ? ctx.rels.get(attr(dm, 'dm')) : null;
      let drawingPath = null;
      if (dataRel) {
        const dataXml = await this.pkg.xml(dataRel.target);
        const ext = first(dataXml, 'dataModelExt');
        const relId = ext && attr(ext, 'relId');
        if (relId) drawingPath = ctx.rels.get(relId)?.target;
      }
      drawingPath ||= [...ctx.rels.values()].find((r) => r.type === 'diagramDrawing')?.target;
      const dxml = drawingPath ? await this.pkg.xml(drawingPath) : null;
      if (!dxml) return;
      const off = kid(xfrm, 'off');
      const ox = num(off, 'x', 0), oy = num(off, 'y', 0);
      const G = (r) => T({ x: ox + r.x, y: oy + r.y, w: r.w, h: r.h });
      await this.tree(first(dxml, 'spTree'), { ...ctx, rels: await this.pkg.rels(drawingPath) }, G);
    } else {
      // Objetos OLE trazem uma imagem de pré-visualização (p:pic)
      const pic = first(gd, 'pic');
      if (pic) await this.picture(pic, ctx, T);
    }
  }

  // ---- Texto -----------------------------------------------------------------------

  levelChain(chain, ctx, phType, lvl) {
    // Do mais genérico para o mais específico
    const m = ctx.master;
    const lv = `lvl${lvl + 1}pPr`;
    const out = [];
    const isTitle = phType === 'title' || phType === 'ctrTitle';
    if (!phType) {
      out.push(path(this.defaultText, lv));
      out.push(path(m?.txStyles, 'otherStyle', lv));
    } else out.push(path(m?.txStyles, isTitle ? 'titleStyle' : ['dt', 'ftr', 'sldNum', 'hdr'].includes(phType) ? 'otherStyle' : 'bodyStyle', lv));
    for (const c of [...chain].reverse()) out.push(path(c, 'txBody', 'lstStyle', lv));
    return out.filter(Boolean);
  }

  text(txBody, chain, ctx, phType, style) {
    const { theme, clrMap } = ctx;
    const bodyPrs = [...chain.map((c) => path(c, 'txBody', 'bodyPr')).reverse(), kid(txBody, 'bodyPr')].filter(Boolean);
    const bp = {};
    for (const x of bodyPrs) {
      for (const k of ['anchor', 'lIns', 'tIns', 'rIns', 'bIns', 'wrap', 'vert']) if (attr(x, k) !== null) bp[k] = attr(x, k);
      const fa = kid(x, 'normAutofit');
      if (fa) { bp.fontScale = num(fa, 'fontScale', 100000) / 100000; bp.lnRed = num(fa, 'lnSpcReduction', 0) / 100000; }
    }
    const fs = bp.fontScale || 1;
    const fontRefColor = style ? readColor(kid(style, 'fontRef'), theme, clrMap) : null;
    const local = kid(txBody, 'lstStyle');
    const paras = [];
    for (const p of kids(txBody, 'p')) {
      const pPr = kid(p, 'pPr');
      const lvl = num(pPr, 'lvl', 0);
      const inheritedLv = this.levelChain(chain, ctx, phType, lvl);
      const levels = [...inheritedLv, path(local, `lvl${lvl + 1}pPr`), pPr].filter(Boolean);
      const pp = { algn: 'l', bullet: 'none', marL: 0, indent: 0, lnSpc: null, spcBef: 0, spcAft: 0, def: {} };
      levels.forEach((lv, li) => {
        // A cor de p:style/fontRef vale acima dos estilos herdados e abaixo da formatação local
        if (li === inheritedLv.length && fontRefColor) pp.def.color = cssColor(fontRefColor);
        if (attr(lv, 'algn')) pp.algn = attr(lv, 'algn');
        if (attr(lv, 'marL') !== null) pp.marL = num(lv, 'marL', 0);
        if (attr(lv, 'indent') !== null) pp.indent = num(lv, 'indent', 0);
        if (kid(lv, 'buNone')) pp.bullet = 'none';
        if (kid(lv, 'buChar')) { pp.bullet = 'char'; pp.buChar = attr(kid(lv, 'buChar'), 'char'); }
        if (kid(lv, 'buAutoNum')) { pp.bullet = 'num'; pp.buType = attr(kid(lv, 'buAutoNum'), 'type'); }
        if (kid(lv, 'buClr')) pp.buClr = cssColor(readColor(kid(lv, 'buClr'), theme, clrMap));
        const ls = kid(lv, 'lnSpc');
        if (ls) pp.lnSpc = kid(ls, 'spcPct') ? { pct: num(kid(ls, 'spcPct'), 'val', 100000) / 100000 } : { pts: num(kid(ls, 'spcPts'), 'val', 0) / 100 };
        for (const k of ['spcBef', 'spcAft']) {
          const sp = kid(lv, k);
          if (sp) pp[k] = kid(sp, 'spcPts') ? { pts: num(kid(sp, 'spcPts'), 'val', 0) / 100 } : { pct: num(kid(sp, 'spcPct'), 'val', 0) / 100000 };
        }
        const d = kid(lv, 'defRPr');
        if (d) pp.def = { ...pp.def, ...this.rPr(d, theme, clrMap) };
      });
      if (levels.length === inheritedLv.length && fontRefColor) pp.def.color = cssColor(fontRefColor);
      const runs = [];
      for (const r of kids(p)) {
        if (r.localName === 'r' || r.localName === 'fld') {
          const rp = { ...pp.def, ...this.rPr(kid(r, 'rPr'), theme, clrMap) };
          const t = kid(r, 't')?.textContent ?? '';
          const link = first(kid(r, 'rPr'), 'hlinkClick');
          const href = link ? ctx.rels.get(attr(link, 'id'))?.target : null;
          runs.push({ text: t, p: rp, href: href && /^(https?:|mailto:)/i.test(href) ? href : null });
        } else if (r.localName === 'br') runs.push({ br: true, p: pp.def });
      }
      const endP = kid(p, 'endParaRPr');
      paras.push({ pp, runs, lvl, end: { ...pp.def, ...this.rPr(endP, theme, clrMap) } });
    }
    const hasText = paras.some((p) => p.runs.some((r) => r.text?.trim()));
    if (!hasText) return null;

    // Base do elemento = primeiro trecho com texto
    const firstRun = paras.flatMap((p) => p.runs).find((r) => r.text?.trim())?.p || paras[0].end;
    const px = (pt) => Math.round(pt * EMU_PER_PT * this.s * fs * 10) / 10;
    const baseSize = px(firstRun.sz || 18);
    const baseColor = firstRun.color || (phType === 'title' || phType === 'ctrTitle' ? this.firstTheme?.title : this.firstTheme?.text) || '#000000';
    const baseFont = firstRun.font ? cssFont(this.fontName(firstRun.font, theme)).replace(/"/g, "'") : null;
    const textRuns = paras.flatMap((p) => p.runs).filter((r) => r.text?.trim());
    const baseBold = textRuns.length > 0 && textRuns.every((r) => r.p.b);
    const baseAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[paras.find((p) => p.runs.length)?.pp.algn || 'l'] || 'left';
    const firstLn = paras[0].pp.lnSpc;
    const line = firstLn?.pct ? Math.round(firstLn.pct * 1.2 * (1 - (bp.lnRed || 0)) * 100) / 100 : 1.2;

    let html = '';
    let list = null;
    const closeList = () => { if (list) { html += `</${list}>`; list = null; } };
    for (const para of paras) {
      const { pp } = para;
      let inner = '';
      for (const r of para.runs) {
        if (r.br) { inner += '<br>'; continue; }
        if (!r.text) continue;
        let h = escapeHtml(r.text).replace(/ {2}/g, '  ');
        const rp = r.p;
        const css = [];
        if (rp.sz && Math.abs(px(rp.sz) - baseSize) > 0.5) css.push(`font-size:${px(rp.sz)}px`);
        if (rp.color && rp.color !== baseColor) css.push(`color:${rp.color}`);
        if (rp.font) { const f = cssFont(this.fontName(rp.font, theme)).replace(/"/g, "'"); if (f !== baseFont) css.push(`font-family:${f}`); }
        if (rp.cap === 'all') css.push('text-transform:uppercase');
        if (rp.cap === 'small') css.push('font-variant:small-caps');
        if (rp.spc) css.push(`letter-spacing:${Math.round(rp.spc * EMU_PER_PT * this.s) / 100}px`);
        if (rp.hl) css.push(`background-color:${rp.hl}`);
        if (baseBold && !rp.b) css.push('font-weight:400');
        if (css.length) h = `<span style="${css.join(';')}">${h}</span>`;
        if (rp.b && !baseBold) h = `<b>${h}</b>`;
        if (rp.i) h = `<i>${h}</i>`;
        if (rp.u) h = `<u>${h}</u>`;
        if (rp.strike) h = `<s>${h}</s>`;
        if (rp.baseline > 0) h = `<sup>${h}</sup>`;
        if (rp.baseline < 0) h = `<sub>${h}</sub>`;
        if (r.href) h = `<a href="${escapeHtml(r.href)}">${h}</a>`;
        inner += h;
      }
      const css = [];
      const al = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[pp.algn] || 'left';
      if (al !== baseAlign) css.push(`text-align:${al}`);
      const sizePt = para.runs.find((r) => r.p?.sz)?.p.sz || para.end.sz || firstRun.sz || 18;
      const spc = (s) => (s?.pts ? px(s.pts) : s?.pct ? px(sizePt) * s.pct * 1.2 : 0);
      if (pp.spcBef) css.push(`margin-top:${Math.round(spc(pp.spcBef))}px`);
      if (pp.spcAft) css.push(`margin-bottom:${Math.round(spc(pp.spcAft))}px`);
      if (pp.lnSpc?.pts) css.push(`line-height:${px(pp.lnSpc.pts)}px`);
      else if (pp.lnSpc?.pct && Math.abs(pp.lnSpc.pct * 1.2 - line) > 0.02) css.push(`line-height:${Math.round(pp.lnSpc.pct * 1.2 * 100) / 100}`);
      const isList = pp.bullet !== 'none' && para.runs.some((r) => r.text?.trim());
      if (isList) {
        const tag = pp.bullet === 'num' ? 'ol' : 'ul';
        if (list !== tag) {
          closeList();
          const lt = pp.bullet === 'num' ? ({ alphaLcParenR: 'lower-alpha', alphaLcPeriod: 'lower-alpha', alphaUcPeriod: 'upper-alpha', alphaUcParenR: 'upper-alpha', romanLcPeriod: 'lower-roman', romanUcPeriod: 'upper-roman' }[pp.buType] || '') : ({ 'o': 'circle', '§': 'square', '▪': 'square', '■': 'square', '–': '"– "', '-': '"- "', '✓': '"✓ "', 'Ø': '"➢ "', 'ü': '"✓ "' }[pp.buChar] || '');
          html += `<${tag}${lt ? ` style="list-style-type:${lt}"` : ''}>`;
          list = tag;
        }
        if (para.lvl > 0 || pp.marL - 342900 * (para.lvl + 1) > 30000) css.push(`margin-left:${Math.round((para.lvl * 342900 + Math.max(0, pp.marL - 342900 * (para.lvl + 1))) * this.s)}px`);
        css.push(`--bullet:${pp.buClr || 'currentColor'}`);
        html += `<li${css.length ? ` style="${css.join(';')}"` : ''}>${inner}</li>`;
      } else {
        closeList();
        if (pp.marL) css.push(`margin-left:${Math.round(pp.marL * this.s)}px`);
        if (pp.indent && pp.marL) css.push(`text-indent:${Math.round(pp.indent * this.s)}px`);
        html += `<p${css.length ? ` style="${css.join(';')}"` : ''}>${inner || '<br>'}</p>`;
      }
    }
    closeList();
    const ins = (k, d) => (bp[k] !== undefined ? +bp[k] : d) * this.s;
    return {
      html,
      props: {
        html, size: baseSize, color: baseColor, align: baseAlign,
        valign: { t: 'top', ctr: 'middle', b: 'bottom' }[bp.anchor || (phType === 'title' || phType === 'ctrTitle' ? 'ctr' : 't')] || 'top',
        bold: baseBold, italic: !!firstRun.i && paras.every((p) => p.runs.every((r) => !r.text?.trim() || r.p.i)), font: baseFont, line,
        pad: [ins('tIns', 45720), ins('rIns', 91440), ins('bIns', 45720), ins('lIns', 91440)].map((v) => Math.round(v)),
      },
    };
  }

  fontName(f, theme) {
    if (f === '+mj-lt' || f === '+mj-ea') return theme.major;
    if (f === '+mn-lt' || f === '+mn-ea') return theme.minor;
    return f;
  }

  rPr(n, theme, clrMap) {
    const o = {};
    if (!n) return o;
    if (attr(n, 'sz') !== null) o.sz = num(n, 'sz', 1800) / 100;
    if (attr(n, 'b') !== null) o.b = attr(n, 'b') === '1' || attr(n, 'b') === 'true';
    if (attr(n, 'i') !== null) o.i = attr(n, 'i') === '1' || attr(n, 'i') === 'true';
    if (attr(n, 'u') !== null) o.u = attr(n, 'u') !== 'none';
    if (attr(n, 'strike') !== null) o.strike = attr(n, 'strike') !== 'noStrike';
    if (attr(n, 'baseline') !== null) o.baseline = num(n, 'baseline', 0);
    if (attr(n, 'cap') !== null) o.cap = attr(n, 'cap');
    if (attr(n, 'spc') !== null) o.spc = num(n, 'spc', 0);
    const fill = kid(n, 'solidFill') || (kid(n, 'gradFill') ? first(n, 'gs') : null);
    if (fill) { const c = readColor(fill, theme, clrMap); if (c) o.color = cssColor(c); }
    const hl = kid(n, 'highlight');
    if (hl) { const c = readColor(hl, theme, clrMap); if (c) o.hl = cssColor(c); }
    const latin = kid(n, 'latin');
    if (latin && attr(latin, 'typeface')) o.font = attr(latin, 'typeface');
    return o;
  }

  // ---- Tabelas -------------------------------------------------------------------------

  table(tbl, b, ctx) {
    if (!tbl) return null;
    const { theme, clrMap } = ctx;
    const tblPr = kid(tbl, 'tblPr');
    const firstRow = attr(tblPr, 'firstRow') === '1';
    const bandRow = attr(tblPr, 'bandRow') === '1';
    const styleId = first(tblPr, 'tableStyleId')?.textContent?.trim();
    const ts = styleId && this.tableStyles ? all(this.tableStyles, 'tblStyle').find((s) => attr(s, 'styleId') === styleId) : null;
    const part = (name) => {
      const p = kid(ts, name);
      if (!p) return {};
      const fillC = readColor(path(p, 'tcStyle', 'fill', 'solidFill'), theme, clrMap);
      const txt = kid(p, 'tcTxStyle');
      const txC = txt ? readColor(txt, theme, clrMap) : null;
      return { bg: fillC ? cssColor(fillC) : null, color: txC ? cssColor(txC) : null, b: txt ? attr(txt, 'b') === 'on' : undefined };
    };
    let whole = part('wholeTbl'), head = part('firstRow'), band1 = part('band1H');
    if (styleId && !ts) {
      // Estilo embutido do PowerPoint (ex.: "Estilo Médio 2 - Ênfase 1"), não gravado no arquivo
      const acc = theme.colors.accent1 || '#4472c4';
      const [r, g, b_] = hexRgb(acc);
      const tint = (k) => rgbHex(r + (255 - r) * k, g + (255 - g) * k, b_ + (255 - b_) * k);
      whole = { bg: tint(0.8), color: null };
      head = { bg: acc, color: '#ffffff', b: true };
      band1 = { bg: tint(0.6), color: null };
    }
    const grid = kids(kid(tbl, 'tblGrid'), 'gridCol').map((g) => num(g, 'w', 0));
    const total = grid.reduce((a, c) => a + c, 0) || 1;
    const rows = [], cells = [];
    let sizePt = null;
    kids(tbl, 'tr').forEach((tr, r) => {
      const row = [], st = [];
      for (const tc of kids(tr, 'tc')) {
        if (attr(tc, 'hMerge') === '1' || attr(tc, 'vMerge') === '1') { row.push(''); st.push(null); continue; }
        const txBody = kid(tc, 'txBody');
        let html = '';
        let cellColor = null, cellB;
        if (txBody) {
          const parts = [];
          for (const p of kids(txBody, 'p')) {
            let inner = '';
            for (const run of kids(p)) {
              if (run.localName === 'br') { inner += '<br>'; continue; }
              if (run.localName !== 'r' && run.localName !== 'fld') continue;
              const rp = this.rPr(kid(run, 'rPr'), theme, clrMap);
              sizePt ??= rp.sz;
              if (rp.color) cellColor ??= rp.color;
              if (rp.b !== undefined) cellB ??= rp.b;
              let h = escapeHtml(kid(run, 't')?.textContent || '');
              if (rp.b) h = `<b>${h}</b>`;
              if (rp.i) h = `<i>${h}</i>`;
              inner += h;
            }
            parts.push(inner);
          }
          html = parts.filter(Boolean).join('<br>');
        }
        const tcPr = kid(tc, 'tcPr');
        const own = readColor(kid(tcPr, 'solidFill'), theme, clrMap);
        const isHead = firstRow && r === 0;
        const band = !isHead && bandRow && ((r - (firstRow ? 1 : 0)) % 2 === 0) ? band1 : {};
        const zone = isHead ? head : band;
        const bg = own ? cssColor(own) : kid(tcPr, 'noFill') ? 'transparent' : zone.bg || whole.bg || null;
        const color = cellColor || zone.color || whole.color || null;
        row.push(html);
        st.push(bg || color ? { bg, color } : null);
      }
      rows.push(row);
      cells.push(st);
    });
    if (!rows.length) return null;
    const ncol = Math.max(...rows.map((r) => r.length));
    for (const r of rows) while (r.length < ncol) r.push('');
    return {
      id: uid(), type: 'table', x: b.x, y: b.y, w: b.w, h: b.h, rot: b.rot, opacity: 1,
      table: {
        rows, header: firstRow, size: Math.round((sizePt || 18) * EMU_PER_PT * this.s), cells,
        cols: grid.length === ncol ? grid.map((g) => g / total) : undefined, imported: true,
      },
    };
  }

  // ---- Gráficos --------------------------------------------------------------------------

  async chart(chartPath, b, ctx) {
    const xml = await this.pkg.xml(chartPath);
    if (!xml) return null;
    const plot = first(xml, 'plotArea');
    const typeNode = kids(plot).find((k) => /Chart$/.test(k.localName));
    if (!typeNode) return null;
    const tn = typeNode.localName;
    let kind = 'column';
    if (tn.startsWith('bar')) kind = attr(kid(typeNode, 'barDir'), 'val') === 'bar' ? 'bar' : 'column';
    else if (tn.startsWith('line') || tn.startsWith('scatter') || tn.startsWith('radar') || tn.startsWith('stock')) kind = 'line';
    else if (tn.startsWith('area')) kind = 'area';
    else if (tn.startsWith('doughnut')) kind = 'doughnut';
    else if (tn.startsWith('pie') || tn.startsWith('ofPie')) kind = 'pie';
    const strs = (n) => {
      const cache = first(n, 'strCache') || first(n, 'numCache') || first(n, 'lvl');
      if (!cache) return n ? [first(n, 'v')?.textContent || ''] : [];
      const count = num(kid(cache, 'ptCount'), 'val', 0);
      const out = Array.from({ length: count }, () => '');
      for (const pt of kids(cache, 'pt')) out[num(pt, 'idx', 0)] = kid(pt, 'v')?.textContent ?? '';
      return out;
    };
    const { theme, clrMap } = ctx;
    const series = [], colors = [];
    let labels = [];
    for (const ser of kids(typeNode, 'ser')) {
      const name = strs(kid(ser, 'tx'))[0] || `Série ${series.length + 1}`;
      const cat = kid(ser, 'cat') || kid(ser, 'xVal');
      const val = kid(ser, 'val') || kid(ser, 'yVal');
      if (!labels.length && cat) labels = strs(cat);
      series.push({ name, values: strs(val).map((v) => (v === '' ? 0 : +v || 0)) });
      const c = readColor(path(ser, 'spPr', 'solidFill'), theme, clrMap) || readColor(path(ser, 'spPr', 'ln', 'solidFill'), theme, clrMap);
      if (c) colors.push(cssColor(c));
    }
    if (!series.length) return null;
    if (!labels.length) labels = series[0].values.map((_, i) => String(i + 1));
    // Pizza: cores por ponto
    if (kind === 'pie' || kind === 'doughnut') {
      colors.length = 0;
      for (const dp of kids(kids(typeNode, 'ser')[0], 'dPt')) {
        const c = readColor(path(dp, 'spPr', 'solidFill'), theme, clrMap);
        if (c) colors[num(kid(dp, 'idx'), 'val', 0)] = cssColor(c);
      }
    }
    const titleNode = first(xml, 'title');
    const deleted = attr(first(xml, 'autoTitleDeleted'), 'val') === '1';
    const title = titleNode && !deleted ? all(titleNode, 't').map((t) => t.textContent).join('') : '';
    const accents = ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6'].map((a) => theme.colors[a]).filter(Boolean);
    return {
      id: uid(), type: 'chart', x: b.x, y: b.y, w: b.w, h: b.h, rot: b.rot, opacity: 1,
      chart: { kind, title, labels, series, colors: colors.filter(Boolean).length ? colors.map((c, i) => c || accents[i % accents.length]) : accents.length ? accents : undefined },
    };
  }

  // ---- Anotações, transições e animações ---------------------------------------------------

  async notes(rels) {
    const rel = [...rels.values()].find((r) => r.type === 'notesSlide');
    if (!rel) return '';
    const xml = await this.pkg.xml(rel.target);
    if (!xml) return '';
    for (const sp of all(xml, 'sp')) {
      const ph = first(sp, 'ph');
      if (ph && attr(ph, 'type') === 'body') {
        return kids(kid(sp, 'txBody'), 'p').map((p) => all(p, 't').map((t) => t.textContent).join('')).join('\n').trim();
      }
    }
    return '';
  }

  transition(sld) {
    let tr = kid(sld, 'transition');
    if (!tr) {
      const alt = kids(sld, 'AlternateContent')[0];
      if (alt) tr = first(alt, 'transition');
    }
    if (!tr) return 'none';
    const fx = kids(tr).find((k) => k.localName !== 'sndAc' && k.localName !== 'extLst');
    return fx ? TRANSITIONS[fx.localName] || 'fade' : 'fade';
  }

  animations(sld, ctx) {
    const timing = kid(sld, 'timing');
    if (!timing) return;
    for (const ctn of all(timing, 'cTn')) {
      if (attr(ctn, 'presetClass') !== 'entr') continue;
      const tgt = first(ctn, 'spTgt');
      const e = tgt && ctx.spidMap.get(attr(tgt, 'spid'));
      if (!e || e.anim) continue;
      const id = num(ctn, 'presetID', 10);
      const sub = num(ctn, 'presetSubtype', 0);
      let type = 'fade';
      if (id === 2 || id === 47 || id === 42) type = sub === 8 ? 'left' : 'up';
      else if (id === 53 || id === 23 || id === 21 || id === 31) type = 'zoom';
      else if (id === 26 || id === 49) type = 'pop';
      e.anim = { type };
    }
  }
}

const identity = (r) => r;

function groupTransform(xfrm, parent) {
  const off = kid(xfrm, 'off'), ext = kid(xfrm, 'ext'), chOff = kid(xfrm, 'chOff'), chExt = kid(xfrm, 'chExt');
  if (!off || !ext || !chOff || !chExt) return parent;
  const ox = num(off, 'x', 0), oy = num(off, 'y', 0);
  const sx = num(ext, 'cx', 1) / (num(chExt, 'cx', 1) || 1), sy = num(ext, 'cy', 1) / (num(chExt, 'cy', 1) || 1);
  const cx = num(chOff, 'x', 0), cy = num(chOff, 'y', 0);
  return (r) => parent({ x: ox + (r.x - cx) * sx, y: oy + (r.y - cy) * sy, w: r.w * sx, h: r.h * sy });
}

const pos = (b) => ({ x: b.x, y: b.y, w: b.w, h: b.h });

function cropImage(url, [l, t, r, b]) {
  return new Promise((resolve, reject) => {
    const img = new Image();
    img.onload = () => {
      const w = img.naturalWidth, h = img.naturalHeight;
      const sx = w * l, sy = h * t, sw = w * (1 - l - r), sh = h * (1 - t - b);
      if (sw <= 0 || sh <= 0) { resolve(url); return; }
      const c = document.createElement('canvas');
      c.width = Math.round(sw);
      c.height = Math.round(sh);
      c.getContext('2d').drawImage(img, sx, sy, sw, sh, 0, 0, c.width, c.height);
      resolve(c.toDataURL(/jpe?g/.test(url.slice(5, 20)) ? 'image/jpeg' : 'image/png', 0.92));
    };
    img.onerror = reject;
    img.src = url;
  });
}
