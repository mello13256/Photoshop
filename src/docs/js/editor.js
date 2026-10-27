// Núcleo do editor de texto: contentEditable com histórico próprio, formatação, blocos, tabelas e imagens.

const BLOCK_TAGS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'UL', 'OL', 'TABLE', 'HR', 'DIV', 'FIGURE']);
const TEXT_BLOCKS = new Set(['P', 'H1', 'H2', 'H3', 'H4', 'H5', 'H6', 'BLOCKQUOTE', 'PRE', 'DIV']);
export const ZWSP = '​';

export class History {
  constructor(html = '') {
    this.stack = [{ html, sel: null }];
    this.index = 0;
  }
}

export class Editor {
  constructor(root, { onChange, onSelection } = {}) {
    this.root = root;
    this.onChange = onChange || (() => {});
    this.onSelection = onSelection || (() => {});
    this.history = new History(root.innerHTML);
    this.pending = 0;
    root.contentEditable = 'true';
    root.spellcheck = true;
    document.execCommand('defaultParagraphSeparator', false, 'p');
    root.addEventListener('beforeinput', (e) => this.onBeforeInput(e));
    root.addEventListener('input', (e) => this.onInput(e));
    root.addEventListener('keydown', (e) => this.onKeyDown(e));
    root.addEventListener('click', (e) => this.onClick(e));
    root.addEventListener('paste', (e) => this.onPaste(e));
    document.addEventListener('selectionchange', () => {
      if (this.hasFocus()) this.onSelection();
    });
    this.ensureBlock();
  }

  // ---- Estado e histórico --------------------------------------------------

  hasFocus() {
    const s = getSelection();
    return s.rangeCount > 0 && this.root.contains(s.anchorNode);
  }

  setContent(html, history = null) {
    this.root.innerHTML = html || '<p><br></p>';
    this.ensureBlock();
    this.history = history || new History(this.root.innerHTML);
    clearTimeout(this.pending);
    this.pending = 0;
  }

  html() {
    return cleanHTML(this.root);
  }

  // Registra a mudança no histórico (imediato ou com atraso enquanto digita).
  changed(immediate = false) {
    clearTimeout(this.pending);
    if (immediate) this.commit();
    else this.pending = setTimeout(() => this.commit(), 450);
    this.onChange();
  }

  commit() {
    this.pending = 0;
    const html = this.root.innerHTML;
    const h = this.history;
    if (h.stack[h.index].html === html) return;
    h.stack.splice(h.index + 1);
    h.stack.push({ html, sel: this.saveSelection() });
    if (h.stack.length > 300) h.stack.shift();
    h.index = h.stack.length - 1;
  }

  undo() {
    if (this.pending) this.commit();
    const h = this.history;
    if (h.index === 0) return;
    const sel = h.stack[h.index].sel;
    h.index--;
    this.root.innerHTML = h.stack[h.index].html;
    this.restoreSelection(h.stack[h.index].sel || sel);
    this.onChange();
  }

  redo() {
    const h = this.history;
    if (h.index >= h.stack.length - 1) return;
    h.index++;
    this.root.innerHTML = h.stack[h.index].html;
    this.restoreSelection(h.stack[h.index].sel);
    this.onChange();
  }

  pathOf(node) {
    const path = [];
    while (node && node !== this.root) {
      const p = node.parentNode;
      if (!p) return null;
      path.unshift(Array.prototype.indexOf.call(p.childNodes, node));
      node = p;
    }
    return node === this.root ? path : null;
  }

  nodeAt(path) {
    let n = this.root;
    for (const i of path || []) {
      if (!n.childNodes[i]) return null;
      n = n.childNodes[i];
    }
    return n;
  }

  saveSelection() {
    const s = getSelection();
    if (!s.rangeCount || !this.root.contains(s.anchorNode)) return null;
    const r = s.getRangeAt(0);
    const a = this.pathOf(r.startContainer), b = this.pathOf(r.endContainer);
    if (!a || !b) return null;
    return { a, ao: r.startOffset, b, bo: r.endOffset };
  }

  restoreSelection(sel) {
    this.root.focus({ preventScroll: true });
    const s = getSelection();
    try {
      const a = sel && this.nodeAt(sel.a), b = sel && this.nodeAt(sel.b);
      const r = document.createRange();
      if (a && b) {
        r.setStart(a, Math.min(sel.ao, maxOffset(a)));
        r.setEnd(b, Math.min(sel.bo, maxOffset(b)));
      } else {
        r.selectNodeContents(this.root);
        r.collapse(false);
      }
      s.removeAllRanges();
      s.addRange(r);
    } catch { /* seleção inválida */ }
  }

  range() {
    const s = getSelection();
    if (!s.rangeCount) return null;
    const r = s.getRangeAt(0);
    return this.root.contains(r.commonAncestorContainer) ? r : null;
  }

  focus() {
    if (!this.hasFocus()) {
      this.root.focus({ preventScroll: true });
      if (this.lastRange) {
        const s = getSelection();
        s.removeAllRanges();
        s.addRange(this.lastRange);
      }
    }
  }

  rememberRange() {
    const r = this.range();
    if (r) this.lastRange = r.cloneRange();
  }

  // Garante que o editor sempre tenha pelo menos um parágrafo e nenhum texto solto na raiz.
  ensureBlock() {
    const root = this.root;
    if (!root.firstChild) root.innerHTML = '<p><br></p>';
    for (const n of [...root.childNodes]) {
      if (n.nodeType === 3 && n.textContent.trim()) {
        const p = document.createElement('p');
        n.replaceWith(p);
        p.append(n);
      } else if (n.nodeType === 1 && !BLOCK_TAGS.has(n.tagName) && n.tagName !== 'BR') {
        const p = document.createElement('p');
        n.replaceWith(p);
        p.append(n);
      }
    }
  }

  // ---- Utilidades de bloco -------------------------------------------------

  topBlock(node) {
    while (node && node.parentNode !== this.root) node = node.parentNode;
    return node && node.parentNode === this.root ? node : null;
  }

  // Bloco "textual" mais próximo (p, h1… ou li/td).
  currentBlock(node = this.range()?.startContainer) {
    while (node && node !== this.root) {
      if (node.nodeType === 1 && (TEXT_BLOCKS.has(node.tagName) || node.tagName === 'LI' || node.tagName === 'TD' || node.tagName === 'TH')) return node;
      node = node.parentNode;
    }
    return null;
  }

  selectedBlocks() {
    const r = this.range();
    if (!r) return [];
    const out = [];
    const walker = document.createTreeWalker(this.root, NodeFilter.SHOW_ELEMENT, {
      acceptNode: (n) => ((TEXT_BLOCKS.has(n.tagName) && !n.closest('table, ul, ol, .toc, .callout-icon')) || n.tagName === 'LI'
        ? NodeFilter.FILTER_ACCEPT : NodeFilter.FILTER_SKIP),
    });
    for (let n = walker.nextNode(); n; n = walker.nextNode()) if (r.intersectsNode(n)) out.push(n);
    if (!out.length) {
      const b = this.currentBlock();
      if (b) out.push(b);
    }
    // Remove blocos que contêm outros blocos selecionados (ex.: DIV de callout)
    return out.filter((b) => !out.some((o) => o !== b && b.contains(o)));
  }

  // ---- Formatação ----------------------------------------------------------

  exec(cmd, value = null) {
    this.focus();
    document.execCommand('styleWithCSS', false, !['bold', 'italic', 'underline', 'strikeThrough'].includes(cmd));
    document.execCommand(cmd, false, value);
    this.changed(true);
  }

  // Aplica estilo inline à seleção via <font size=7> temporário.
  styleSelection(styleFn) {
    this.focus();
    const r = this.range();
    if (!r || r.collapsed) return false;
    document.execCommand('styleWithCSS', false, false);
    document.execCommand('fontSize', false, '7');
    for (const f of this.root.querySelectorAll('font[size="7"]')) {
      const span = document.createElement('span');
      styleFn(span.style);
      // Remove o mesmo estilo em descendentes para não conflitar
      for (const inner of f.querySelectorAll('span')) {
        const tmp = document.createElement('span');
        styleFn(tmp.style);
        for (const prop of tmp.style) inner.style.removeProperty(prop);
      }
      while (f.firstChild) span.append(f.firstChild);
      f.replaceWith(span);
    }
    this.changed(true);
    return true;
  }

  setFontSize(pt) {
    this.styleSelection((s) => { s.fontSize = `${pt}pt`; });
  }

  setFontFamily(family) {
    this.styleSelection((s) => { s.fontFamily = family; });
  }

  setColor(c) {
    this.exec('foreColor', c);
  }

  setHighlight(c) {
    this.exec('hiliteColor', c || 'transparent');
  }

  clearFormat() {
    this.exec('removeFormat');
    this.exec('unlink');
  }

  setBlock(type) {
    this.focus();
    const sel = this.saveSelection();
    const map = { p: ['P', ''], h1: ['H1', ''], h2: ['H2', ''], h3: ['H3', ''], h4: ['H4', ''], title: ['P', 'title'], subtitle: ['P', 'subtitle'], blockquote: ['BLOCKQUOTE', ''], pre: ['PRE', ''] };
    const [tag, cls] = map[type] || map.p;
    for (const b of this.selectedBlocks()) {
      if (b.tagName === 'LI' || b.tagName === 'TD' || b.tagName === 'TH') continue;
      if (b.classList.contains('callout') || b.classList.contains('toc')) continue;
      const n = document.createElement(tag);
      if (cls) n.className = cls;
      for (const p of ['textAlign', 'lineHeight', 'marginLeft']) if (b.style[p]) n.style[p] = b.style[p];
      if (b.id) n.id = b.id;
      while (b.firstChild) n.append(b.firstChild);
      b.replaceWith(n);
    }
    this.restoreSelection(sel);
    this.changed(true);
  }

  blockType() {
    const b = this.currentBlock();
    if (!b) return 'p';
    if (b.tagName === 'P') return b.classList.contains('title') ? 'title' : b.classList.contains('subtitle') ? 'subtitle' : 'p';
    if (/^H[1-4]$/.test(b.tagName)) return b.tagName.toLowerCase();
    if (b.tagName === 'BLOCKQUOTE') return 'blockquote';
    if (b.tagName === 'PRE') return 'pre';
    return 'p';
  }

  setBlockStyle(prop, value) {
    this.focus();
    for (const b of this.selectedBlocks()) b.style[prop] = value;
    this.changed(true);
  }

  setAlign(a) {
    const cmd = { left: 'justifyLeft', center: 'justifyCenter', right: 'justifyRight', justify: 'justifyFull' }[a];
    this.focus();
    document.execCommand('styleWithCSS', false, true);
    document.execCommand(cmd);
    this.changed(true);
  }

  indent(dir) {
    this.focus();
    if (this.closest('LI')) {
      document.execCommand(dir > 0 ? 'indent' : 'outdent');
    } else {
      for (const b of this.selectedBlocks()) {
        const cur = parseFloat(b.style.marginLeft) || 0;
        const next = Math.max(0, cur + dir * 1.25);
        b.style.marginLeft = next ? `${next}cm` : '';
      }
    }
    this.changed(true);
  }

  closest(tag) {
    let n = this.range()?.startContainer;
    while (n && n !== this.root) {
      if (n.nodeType === 1 && n.tagName === tag) return n;
      n = n.parentNode;
    }
    return null;
  }

  list(kind) {
    this.focus();
    const inList = this.closest('UL') || this.closest('OL');
    const isCheck = inList?.classList.contains('checklist');
    if (kind === 'check') {
      if (inList && isCheck) { document.execCommand('insertUnorderedList'); this.changed(true); return; }
      if (!inList || inList.tagName !== 'UL') document.execCommand('insertUnorderedList');
      const ul = this.closest('UL');
      ul?.classList.add('checklist');
    } else {
      if (isCheck && kind === 'bullet') { inList.classList.remove('checklist'); this.changed(true); return; }
      document.execCommand(kind === 'number' ? 'insertOrderedList' : 'insertUnorderedList');
      this.closest('UL')?.classList.remove('checklist');
    }
    this.changed(true);
  }

  insertHTML(html) {
    this.focus();
    document.execCommand('insertHTML', false, html);
    this.ensureBlock();
    this.changed(true);
  }

  // Insere um elemento de bloco após o bloco atual.
  insertBlock(node, { focusInside = true } = {}) {
    this.focus();
    const cur = this.topBlock(this.range()?.startContainer) || this.root.lastChild;
    const isEmpty = cur && cur.tagName === 'P' && !cur.textContent.trim() && !cur.querySelector('img');
    if (cur && isEmpty) cur.replaceWith(node);
    else if (cur) cur.after(node);
    else this.root.append(node);
    if (!node.nextSibling || node.nextSibling.nodeType !== 1) node.after(paragraph());
    const target = focusInside ? firstEditable(node) || node.nextSibling : node.nextSibling;
    placeCaret(target, false);
    this.changed(true);
  }

  insertLink(url, text) {
    this.focus();
    const r = this.range();
    if (r && !r.collapsed) {
      document.execCommand('createLink', false, url);
    } else {
      const a = document.createElement('a');
      a.href = url;
      a.textContent = text || url;
      r.insertNode(a);
      r.setStartAfter(a);
      r.collapse(true);
      const s = getSelection();
      s.removeAllRanges();
      s.addRange(r);
    }
    this.changed(true);
  }

  insertImage(src, alt = '') {
    const p = document.createElement('p');
    p.style.textIndent = '0';
    const img = document.createElement('img');
    img.src = src;
    img.alt = alt;
    img.className = 'align-center';
    p.append(img);
    this.insertBlock(p, { focusInside: false });
  }

  insertTable(rows, cols, header = true) {
    const t = document.createElement('table');
    const tb = document.createElement('tbody');
    for (let r = 0; r < rows; r++) {
      const tr = document.createElement('tr');
      for (let c = 0; c < cols; c++) {
        const cell = document.createElement(header && r === 0 ? 'th' : 'td');
        cell.innerHTML = '<br>';
        tr.append(cell);
      }
      tb.append(tr);
    }
    t.append(tb);
    this.insertBlock(t);
  }

  // ---- Tabelas -------------------------------------------------------------

  tableOp(op) {
    const cell = this.closest('TD') || this.closest('TH');
    if (!cell) return;
    const tr = cell.parentElement, table = tr.closest('table');
    const ci = [...tr.children].indexOf(cell);
    const rows = [...table.rows];
    const mk = (tag) => { const c = document.createElement(tag); c.innerHTML = '<br>'; return c; };
    switch (op) {
      case 'rowAbove':
      case 'rowBelow': {
        const nr = document.createElement('tr');
        for (let i = 0; i < tr.children.length; i++) nr.append(mk('td'));
        if (op === 'rowAbove' && tr.querySelector('th')) tr.after(nr);
        else if (op === 'rowAbove') tr.before(nr);
        else tr.after(nr);
        placeCaret(nr.children[ci], true);
        break;
      }
      case 'colLeft':
      case 'colRight':
        for (const row of rows) {
          const ref = row.children[ci];
          const c = mk(ref?.tagName === 'TH' ? 'th' : 'td');
          if (op === 'colLeft') ref.before(c); else ref.after(c);
        }
        break;
      case 'delRow':
        if (rows.length <= 1) { table.remove(); break; }
        placeCaret((tr.nextElementSibling || tr.previousElementSibling).children[0], true);
        tr.remove();
        break;
      case 'delCol':
        if (tr.children.length <= 1) { table.remove(); break; }
        for (const row of rows) row.children[ci]?.remove();
        break;
      case 'delTable':
        placeCaret(table.nextElementSibling || table.previousElementSibling, true);
        table.remove();
        break;
      case 'header': {
        const first = rows[0];
        const toTag = first.children[0].tagName === 'TH' ? 'td' : 'th';
        for (const c of [...first.children]) {
          const n = document.createElement(toTag);
          while (c.firstChild) n.append(c.firstChild);
          for (const a of c.attributes) n.setAttribute(a.name, a.value);
          c.replaceWith(n);
        }
        break;
      }
      case 'striped':
        table.classList.toggle('striped');
        break;
      default:
    }
    this.ensureBlock();
    this.changed(true);
  }

  // ---- Eventos -------------------------------------------------------------

  onBeforeInput(e) {
    if (e.inputType === 'historyUndo') { e.preventDefault(); this.undo(); }
    else if (e.inputType === 'historyRedo') { e.preventDefault(); this.redo(); }
  }

  onInput(e) {
    if (e.inputType === 'insertText' && e.data) {
      if (e.data === ' ') this.markdownBlock();
      this.markdownInline();
    }
    if (!this.root.firstElementChild) this.ensureBlock();
    this.changed(false);
  }

  onKeyDown(e) {
    const mod = e.ctrlKey || e.metaKey;
    if (mod && !e.altKey && e.key.toLowerCase() === 'z') { e.preventDefault(); if (e.shiftKey) this.redo(); else this.undo(); return; }
    if (mod && !e.altKey && e.key.toLowerCase() === 'y') { e.preventDefault(); this.redo(); return; }
    if (e.key === 'Tab') {
      const cell = this.closest('TD') || this.closest('TH');
      if (cell) {
        e.preventDefault();
        this.moveCell(cell, e.shiftKey ? -1 : 1);
        return;
      }
      e.preventDefault();
      if (this.closest('LI')) { this.indent(e.shiftKey ? -1 : 1); return; }
      if (!e.shiftKey) { document.execCommand('insertText', false, '\t'); this.changed(); }
      return;
    }
    if (e.key === 'Enter' && !e.shiftKey) {
      const b = this.currentBlock();
      if (b && b.tagName === 'P' && b.textContent === '---') {
        e.preventDefault();
        const hr = document.createElement('hr');
        const p = paragraph();
        b.replaceWith(hr);
        hr.after(p);
        placeCaret(p, true);
        this.changed(true);
        return;
      }
      if (b && b.tagName === 'P' && b.textContent === '```') {
        e.preventDefault();
        const pre = document.createElement('pre');
        pre.innerHTML = '<br>';
        b.replaceWith(pre);
        placeCaret(pre, true);
        this.changed(true);
        return;
      }
      // Enter em título volta para parágrafo normal
      if (b && (/^H[1-4]$/.test(b.tagName) || b.classList.contains('title') || b.classList.contains('subtitle')) && atEnd(this.range(), b)) {
        e.preventDefault();
        const p = paragraph();
        b.after(p);
        placeCaret(p, true);
        this.changed(true);
        return;
      }
      if (b && b.tagName === 'PRE') {
        e.preventDefault();
        if (b.textContent.endsWith('\n\n') && atEnd(this.range(), b)) {
          b.textContent = b.textContent.replace(/\n+$/, '');
          const p = paragraph();
          b.after(p);
          placeCaret(p, true);
        } else {
          document.execCommand('insertText', false, '\n');
        }
        this.changed(true);
        return;
      }
    }
    if (e.key === 'Backspace') {
      const b = this.currentBlock();
      const r = this.range();
      // Backspace no início de um título/citação vazio: volta a parágrafo
      if (b && r?.collapsed && atStart(r, b) && b.tagName !== 'P' && b.tagName !== 'LI' && b.tagName !== 'TD' && b.tagName !== 'TH' && !b.textContent) {
        e.preventDefault();
        this.setBlock('p');
      }
    }
  }

  moveCell(cell, dir) {
    const cells = [...cell.closest('table').querySelectorAll('td, th')];
    const i = cells.indexOf(cell) + dir;
    if (i >= cells.length) {
      this.tableOp('rowBelow');
      const rows = cell.closest('table').rows;
      placeCaret(rows[rows.length - 1].children[0], true);
      return;
    }
    if (i < 0) return;
    const target = cells[i];
    const r = document.createRange();
    r.selectNodeContents(target);
    const s = getSelection();
    s.removeAllRanges();
    s.addRange(r);
  }

  onClick(e) {
    const li = e.target.closest?.('ul.checklist > li');
    if (li && e.offsetX < 22 && e.target === li) {
      li.dataset.checked = li.dataset.checked === 'true' ? 'false' : 'true';
      this.changed(true);
    }
    const a = e.target.closest?.('a[href]');
    if (a && (e.ctrlKey || e.metaKey)) {
      e.preventDefault();
      const href = a.getAttribute('href');
      if (href.startsWith('#')) document.getElementById(href.slice(1))?.scrollIntoView({ behavior: 'smooth', block: 'start' });
      else window.open(href, '_blank');
    }
  }

  onPaste(e) {
    const dt = e.clipboardData;
    if (!dt) return;
    const img = [...dt.items].find((i) => i.type.startsWith('image/'));
    if (img && !dt.types.includes('text/html')) {
      e.preventDefault();
      const reader = new FileReader();
      reader.onload = () => this.insertImage(reader.result);
      reader.readAsDataURL(img.getAsFile());
      return;
    }
    const html = dt.getData('text/html');
    if (html && this.sanitize) {
      e.preventDefault();
      this.insertHTML(this.sanitize(html));
    }
  }

  // ---- Markdown ------------------------------------------------------------

  markdownBlock() {
    const b = this.currentBlock();
    if (!b || b.tagName !== 'P') return;
    const t = b.textContent.replace(/ /g, ' ');
    const rules = [
      [/^# $/, () => this.setBlock('h1')], [/^## $/, () => this.setBlock('h2')], [/^### $/, () => this.setBlock('h3')],
      [/^#### $/, () => this.setBlock('h4')], [/^> $/, () => this.setBlock('blockquote')],
      [/^[-*] $/, () => this.list('bullet')], [/^1[.)] $/, () => this.list('number')], [/^\[ ?\] $/, () => this.list('check')],
    ];
    for (const [re, fn] of rules) {
      if (re.test(t)) {
        b.textContent = '';
        b.innerHTML = '<br>';
        placeCaret(b, true);
        fn();
        return;
      }
    }
  }

  markdownInline() {
    const s = getSelection();
    const node = s.anchorNode;
    if (!node || node.nodeType !== 3 || node.parentElement.closest('pre, code')) return;
    const text = node.textContent.slice(0, s.anchorOffset);
    const rules = [
      [/\*\*([^*\n]+)\*\*$/, 'strong'], [/(?<![_\w])__([^_\n]+)__$/, 'strong'],
      [/(?<![*\w])\*([^*\s][^*\n]*)\*$/, 'em'], [/(?<![_\w])_([^_\s][^_\n]*)_$/, 'em'],
      [/`([^`\n]+)`$/, 'code'], [/~~([^~\n]+)~~$/, 's'],
    ];
    for (const [re, tag] of rules) {
      const m = re.exec(text);
      if (!m) continue;
      const inner = m[1];
      const full = m[0];
      const start = s.anchorOffset - full.length;
      const r = document.createRange();
      r.setStart(node, start);
      r.setEnd(node, s.anchorOffset);
      r.deleteContents();
      const elx = document.createElement(tag);
      elx.textContent = inner;
      r.insertNode(elx);
      const after = document.createTextNode(ZWSP);
      elx.after(after);
      const nr = document.createRange();
      nr.setStart(after, 1);
      nr.collapse(true);
      s.removeAllRanges();
      s.addRange(nr);
      return;
    }
  }

  // ---- Consultas -----------------------------------------------------------

  text() {
    return this.root.innerText.replace(/​/g, '');
  }

  headings() {
    return [...this.root.querySelectorAll('h1, h2, h3')].filter((h) => !h.closest('.toc'));
  }
}

// ---------------------------------------------------------------------------
// Funções auxiliares

export function paragraph() {
  const p = document.createElement('p');
  p.innerHTML = '<br>';
  return p;
}

function maxOffset(n) {
  return n.nodeType === 3 ? n.length : n.childNodes.length;
}

function firstEditable(node) {
  if (node.matches?.('td, th, p, h1, h2, h3, h4, li, pre, blockquote')) return node;
  return node.querySelector?.('td, th, p, h1, h2, h3, h4, li, pre') || null;
}

export function placeCaret(node, atStartOfNode = true) {
  if (!node) return;
  const r = document.createRange();
  r.selectNodeContents(node);
  r.collapse(atStartOfNode);
  const s = getSelection();
  s.removeAllRanges();
  s.addRange(r);
}

function atEnd(range, block) {
  if (!range) return false;
  const r = document.createRange();
  r.selectNodeContents(block);
  r.setStart(range.endContainer, range.endOffset);
  return !r.toString().replace(/​/g, '').length;
}

function atStart(range, block) {
  const r = document.createRange();
  r.selectNodeContents(block);
  r.setEnd(range.startContainer, range.startOffset);
  return !r.toString().length;
}

// Remove marcas de edição do HTML (classes temporárias, espaços de largura zero, seleção de imagens).
export function cleanHTML(root) {
  const clone = root.cloneNode(true);
  clone.querySelectorAll('.active-block').forEach((n) => n.classList.remove('active-block'));
  clone.querySelectorAll('.img-selected').forEach((n) => n.classList.remove('img-selected'));
  clone.querySelectorAll('[class=""]').forEach((n) => n.removeAttribute('class'));
  const walker = document.createTreeWalker(clone, NodeFilter.SHOW_TEXT);
  const texts = [];
  for (let n = walker.nextNode(); n; n = walker.nextNode()) texts.push(n);
  for (const t of texts) if (t.textContent.includes(ZWSP)) t.textContent = t.textContent.replace(/​/g, '');
  return clone.innerHTML;
}
