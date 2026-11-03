/* Piezas compartidas por los tres editores: menús emergentes y búsqueda en texto editable. */
(function () {
  const OF = window.OF;
  OF.checkpoint = OF.checkpoint || (() => {});
  OF.typingCheckpoint = OF.typingCheckpoint || (() => {});

  // Menú emergente anclado a un elemento. build(p, close)
  let current = null;
  OF.closePopup = function () { if (current) { current.remove(); current = null; } };
  OF.popup = function (anchor, build, toggleOff) {
    const same = current && current._anchor === anchor;
    OF.closePopup();
    if (same && toggleOff !== false) return null;
    const p = OF.el('div', { class: 'pop', role: 'menu' });
    p._anchor = anchor;
    current = p;
    build(p, OF.closePopup);
    p.style.maxHeight = innerHeight - 16 + 'px';
    document.body.append(p);
    const r = anchor.getBoundingClientRect();
    const w = p.offsetWidth, h = p.offsetHeight;
    p.style.left = Math.max(8, Math.min(innerWidth - w - 8, r.left)) + 'px';
    const below = r.bottom + 6;
    p.style.top = (below + h > innerHeight - 8 && r.top - h - 6 > 8 ? r.top - h - 6 : Math.max(8, Math.min(below, innerHeight - h - 8))) + 'px';
    const first = p.querySelector('input, button');
    if (first && first.tagName === 'INPUT') setTimeout(() => first.focus(), 30);
    return p;
  };
  document.addEventListener('pointerdown', (e) => { if (current && !current.contains(e.target) && !(current._anchor && current._anchor.contains(e.target))) OF.closePopup(); }, true);
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') OF.closePopup(); });
  window.addEventListener('resize', () => OF.closePopup());

  /* --------- copiar el contenido editable sin pasar por el analizador HTML ---------
     (innerHTML partiría párrafos que contienen bloques, como las imágenes de Word) */
  OF.cloneKids = (root) => Array.from(root.childNodes).map((n) => n.cloneNode(true));
  OF.setKids = (root, nodes) => root.replaceChildren(...nodes.map((n) => n.cloneNode(true)));
  OF.serializeKids = (nodes) => { const xs = new XMLSerializer(); return '<div xmlns="http://www.w3.org/1999/xhtml">' + nodes.map((n) => xs.serializeToString(n)).join('') + '</div>'; };
  OF.parseKids = (str) => {
    const doc = new DOMParser().parseFromString(str, 'application/xhtml+xml');
    if (doc.getElementsByTagName('parsererror').length) return null;
    return Array.from(doc.documentElement.childNodes).map((n) => document.importNode(n, true));
  };

  /* ---------------- buscar y reemplazar en áreas de texto (Word y PowerPoint) ---------------- */
  const esc = (q) => q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  OF.textSearchMixin = {
    findMatches(q) {
      const out = [];
      if (!q) return out;
      const re = new RegExp(esc(q), 'gi');
      for (const root of this.searchRoots()) {
        const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
        let n;
        while ((n = w.nextNode())) {
          if (!n.parentElement.closest('[data-p]') || n.parentElement.closest('[data-x]')) continue;
          re.lastIndex = 0;
          let m;
          while ((m = re.exec(n.nodeValue))) { out.push({ node: n, start: m.index, end: m.index + m[0].length }); if (!m[0].length) break; }
        }
      }
      return out;
    },
    selectMatch(m) {
      const r = document.createRange();
      r.setStart(m.node, m.start); r.setEnd(m.node, m.end);
      const ed = m.node.parentElement.closest('[contenteditable=true]');
      if (ed) ed.focus({ preventScroll: true });
      const s = getSelection(); s.removeAllRanges(); s.addRange(r);
      OF.lastRange = r.cloneRange();
      const el = m.node.parentElement;
      const rect = el.getBoundingClientRect(), host = this.host.getBoundingClientRect();
      if (rect.top < host.top + 40 || rect.bottom > host.bottom - 40) el.scrollIntoView({ block: 'center' });
      const wrap = el.closest('.pp-wrap');
      if (wrap && this.markCur) { this.curSlide = +wrap.dataset.i; this.markCur(); }
    },
    find(q, dir) {
      const m = this.findMatches(q);
      if (!m.length) return { i: -1, n: 0 };
      const cur = OF.currentRange();
      let i = dir < 0 ? m.length - 1 : 0;
      if (cur) {
        const cmp = (x) => { const r = document.createRange(); r.setStart(x.node, x.start); return r.compareBoundaryPoints(Range.START_TO_START, cur); };
        if (dir < 0) { for (let k = m.length - 1; k >= 0; k--) if (cmp(m[k]) < 0) { i = k; break; } }
        else { for (let k = 0; k < m.length; k++) { const c = cmp(m[k]); if (c > 0 || (c === 0 && cur.collapsed)) { i = k; break; } } }
      }
      this.selectMatch(m[i]);
      return { i, n: m.length };
    },
    replaceOne(q, rep) {
      const cur = OF.currentRange();
      if (cur && !cur.collapsed && cur.toString().toLowerCase() === q.toLowerCase() && cur.startContainer === cur.endContainer && cur.startContainer.nodeType === 3) {
        OF.checkpoint();
        const n = cur.startContainer, st = cur.startOffset;
        n.nodeValue = n.nodeValue.slice(0, st) + rep + n.nodeValue.slice(cur.endOffset);
        const r = document.createRange(); r.setStart(n, st + rep.length); r.collapse(true);
        const s = getSelection(); s.removeAllRanges(); s.addRange(r); OF.lastRange = r.cloneRange();
        this.markDirty();
        this.afterTextChange && this.afterTextChange();
      }
      return this.find(q, 1);
    },
    replaceAll(q, rep) {
      const m = this.findMatches(q);
      if (!m.length) return 0;
      OF.checkpoint();
      for (let k = m.length - 1; k >= 0; k--) { const x = m[k]; x.node.nodeValue = x.node.nodeValue.slice(0, x.start) + rep + x.node.nodeValue.slice(x.end); }
      this.markDirty();
      this.afterTextChange && this.afterTextChange();
      return m.length;
    },
  };
})();
