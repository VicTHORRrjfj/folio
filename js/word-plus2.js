/* Word: encabezados y pies de página editables, número de página y paginación automática. */
(function () {
  const OF = window.OF;
  const W = OF.NS.w;
  const WD = OF.WordDoc.prototype;
  const wv = (el, a) => (el ? el.getAttributeNS(W, a) || el.getAttribute('w:' + a) : null);
  const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
  const CT = { header: 'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml', footer: 'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml' };
  const key = (s) => OF.cleanText(s).replace(/\s+/g, '');

  /* ------------------------- carga de las partes ------------------------- */
  // Se relee desde el paquete en cada dibujo: el zip es siempre la fuente de verdad.
  WD.loadHF = async function () {
    this.rels = await OF.readRels(this.zip, this.docPath);
    this.hfParts = {};
    for (const [rid, rel] of Object.entries(this.rels)) {
      const m = /\/(header|footer)$/.exec(rel.type);
      if (!m || !this.zip.file(rel.target)) continue;
      this.hfParts[rel.target] = { rid, path: rel.target, kind: m[1], doc: OF.parseXml(await this.zip.file(rel.target).async('string')) };
    }
    // referencias por defecto primero
    const body = OF.kid(this.xml.documentElement, 'body', W);
    this.hfOrder = [];
    for (const ref of Array.from(body.getElementsByTagNameNS(W, '*')).filter((e) => /^(header|footer)Reference$/.test(e.localName))) {
      const rel = this.rels[ref.getAttributeNS(OF.NS.r, 'id')];
      if (rel && this.hfParts[rel.target]) this.hfOrder.push({ path: rel.target, def: (wv(ref, 'type') || 'default') === 'default' });
    }
    const fnRel = Object.values(this.rels).find((r) => /\/footnotes$/.test(r.type));
    this.fnPart = fnRel && this.zip.file(fnRel.target) ? { path: fnRel.target, doc: OF.parseXml(await this.zip.file(fnRel.target).async('string')) } : null;
    // variantes: primera página, pares y predeterminado (de la última sección)
    const sect = OF.kid(body, 'sectPr', W);
    this.hfRefs = { header: {}, footer: {} };
    if (sect) for (const ref of Array.from(sect.children)) {
      const m = /^(header|footer)Reference$/.exec(ref.localName);
      const rel = m && this.rels[ref.getAttributeNS(OF.NS.r, 'id')];
      if (rel && this.hfParts[rel.target]) this.hfRefs[m[1]][wv(ref, 'type') || 'default'] = this.hfParts[rel.target];
    }
    const tp = sect && OF.kid(sect, 'titlePg', W);
    this.titlePg = !!tp && !/^(0|false)$/.test(wv(tp, 'val') || '');
    this.evenOdd = false;
    const setRel = Object.values(this.rels).find((r) => /\/settings$/.test(r.type));
    if (setRel && this.zip.file(setRel.target)) {
      const st = OF.parseXml(await this.zip.file(setRel.target).async('string'));
      const eo = OF.all(st, 'evenAndOddHeaders', W)[0];
      this.evenOdd = !!eo && !/^(0|false)$/.test(wv(eo, 'val') || '');
      this.settingsPath = setRel.target; this.settingsDoc = st;
    }
  };
  // la parte que corresponde a una página (número global)
  WD.hfFor = function (kind, page) {
    const R = (this.hfRefs || {})[kind] || {};
    if (this.titlePg && page === 1) return R.first || null;
    if (this.evenOdd && page % 2 === 0) return R.even || null;
    return R.default || null;
  };
  // dibuja una parte de encabezado/pie que docx-preview no mostró (p. ej. la de páginas pares)
  WD.renderHFPart = function (part) {
    const el = OF.el(part.kind === 'header' ? 'header' : part.kind === 'footer' ? 'footer' : 'div', { class: part.kind === 'footnote' ? 'fn-item' : 'pg-gen' });
    const root = part.root || part.doc.documentElement;
    const paras = Array.from(root.getElementsByTagNameNS(W, 'p')).filter((p) => { for (let a = p.parentNode; a && a !== root; a = a.parentNode) if (['txbxContent', 'Fallback', 'Choice'].includes(a.localName)) return false; return true; });
    for (const xp of paras) {
      const p = OF.el('p');
      const pPr = OF.kid(xp, 'pPr', W);
      const ps = wv(OF.kid(pPr, 'pStyle', W), 'val');
      if (ps) p.className = 'docx_' + ps.toLowerCase();
      const jc = wv(OF.kid(pPr, 'jc', W), 'val');
      if (jc) p.style.textAlign = { both: 'justify', center: 'center', right: 'right', end: 'right' }[jc] || 'left';
      const walk = (node) => {
        for (const c of Array.from(node.children)) {
          if (c.namespaceURI !== W) continue;
          if (c.localName === 'r') {
            const rPr = OF.kid(c, 'rPr', W);
            const st = {};
            const on = (n) => { const e = OF.kid(rPr, n, W); return e && !/^(0|false)$/.test(wv(e, 'val') || ''); };
            if (on('b')) st.fontWeight = '700';
            if (on('i')) st.fontStyle = 'italic';
            const u = OF.kid(rPr, 'u', W); if (u && wv(u, 'val') !== 'none') st.textDecoration = 'underline';
            const sz = wv(OF.kid(rPr, 'sz', W), 'val'); if (sz) st.fontSize = +sz / 2 + 'pt';
            const col = wv(OF.kid(rPr, 'color', W), 'val'); if (col && col !== 'auto') st.color = '#' + col;
            for (const x of Array.from(c.children)) {
              if (x.localName === 't') p.append(OF.el('span', { style: st }, x.textContent));
              else if (x.localName === 'tab') p.append(OF.el('span', { 'data-x': '', style: { display: 'inline-block', width: '3em' } }));
              else if (x.localName === 'br') p.append(OF.el('br'));
              else if (x.localName === 'footnoteRef') p.append(OF.el('sup', { 'data-x': '', class: 'fn-num', contenteditable: 'false' }, String(part.num || '')));
            }
          } else if (c.localName === 'fldSimple') p.append(OF.el('span'));
          else if (['hyperlink', 'smartTag', 'ins', 'sdt', 'sdtContent', 'customXml'].includes(c.localName)) walk(c);
        }
      };
      walk(xp);
      el.append(p);
    }
    return el;
  };
  const syncFor = () => new OF.TextSync({
    ns: W, prefix: 'w', P: 'p', R: 'r', T: 't', PPR: 'pPr',
    wrappers: new Set(['hyperlink', 'smartTag', 'ins', 'sdt', 'sdtContent', 'customXml', 'moveTo']),
    mustHaveParagraph: (parent) => ['hdr', 'ftr', 'tc', 'sdtContent'].includes(parent.localName),
    markProps: (p) => OF.kid(OF.kid(p, 'pPr', W), 'rPr', W),
    applyFmt: OF.WordDoc.runFmt, applyPFmt: OF.WordDoc.paraFmt,
  });
  // runs que muestran el resultado de un campo complejo (PAGE, NUMPAGES)
  function fieldRuns(p) {
    const out = new Map();
    let instr = '', inRes = false, depth = 0;
    for (const r of p.getElementsByTagNameNS(W, 'r')) {
      const fc = OF.kid(r, 'fldChar', W);
      if (fc) {
        const t = wv(fc, 'fldCharType');
        if (t === 'begin') { depth++; if (depth === 1) { instr = ''; inRes = false; } }
        else if (t === 'separate' && depth === 1) inRes = true;
        else if (t === 'end') { depth--; if (!depth) inRes = false; }
        continue;
      }
      const it = OF.kid(r, 'instrText', W);
      if (it && depth === 1 && !inRes) { instr += it.textContent; continue; }
      if (inRes && OF.kid(r, 't', W)) out.set(r, /NUMPAGES|SECTIONPAGES/i.test(instr) ? 'NUMPAGES' : /\bPAGE\b/i.test(instr) ? 'PAGE' : 'OTHER');
    }
    return out;
  }
  const inSimple = (t, stop) => { for (let a = t.parentNode; a && a !== stop; a = a.parentNode) if (a.localName === 'fldSimple') return true; return false; };
  const noDigits = (s) => key(s).replace(/\d+/g, '');
  const partKey = (part) => noDigits(Array.from(part.doc.getElementsByTagNameNS(W, 't')).filter((t) => !inSimple(t, part.doc.documentElement)).map((t) => t.textContent).join(''));
  const domKey = (el) => noDigits(el.textContent);

  /* ------------------- relacionar lo dibujado con cada parte ------------------- */
  WD.mapHeaders = function () {
    this.hfDom = [];
    if (!this.hfParts) return;
    const used = new Set();
    const order = [...this.hfOrder.filter((o) => o.def), ...this.hfOrder.filter((o) => !o.def)].map((o) => this.hfParts[o.path]);
    for (const p of Object.values(this.hfParts)) if (!order.includes(p)) order.push(p);
    for (const sec of this.pages.querySelectorAll('section.docx')) {
      for (const el of sec.querySelectorAll(':scope > header, :scope > footer')) {
        const kind = el.tagName === 'HEADER' ? 'header' : 'footer';
        const dk = domKey(el);
        const part = order.find((p) => p.kind === kind && partKey(p) === dk);
        if (!part) { el.setAttribute('contenteditable', 'false'); continue; }
        el.dataset.hf = part.path;
        if (used.has(part.path)) { el.dataset.mirror = '1'; el.setAttribute('contenteditable', 'false'); this.hfDom.push({ el, part, primary: false }); continue; }
        used.add(part.path);
        this.mapPart(part, el);
        this.hfDom.push({ el, part, primary: true });
      }
    }
  };
  WD.mapPart = function (part, el) {
    const sync = (part.sync = syncFor());
    const root = part.root || part.doc.documentElement;
    const xparas = Array.from(root.getElementsByTagNameNS(W, 'p')).filter((p) => { for (let a = p.parentNode; a && a !== root; a = a.parentNode) if (['txbxContent', 'Fallback', 'Choice'].includes(a.localName)) return false; return true; });
    const dps = Array.from(el.querySelectorAll('p'));
    xparas.forEach((xp, j) => {
      const dp = dps[j];
      if (!dp) return;
      const fr = fieldRuns(xp);
      const xruns = sync.textRunsOf(xp).filter((r) => sync.textOfRun(r) !== '');
      const leaves = Array.from(dp.querySelectorAll('span')).filter((s) => s.childNodes.length > 0 && Array.from(s.childNodes).every((n) => n.nodeType === 3));
      // campos simples: docx-preview los dibuja como <span> vacío (o con el número guardado)
      const simple = Array.from(xp.getElementsByTagNameNS(W, 'fldSimple')).map((f) => (/NUMPAGES|SECTIONPAGES/i.test(wv(f, 'instr') || '') ? 'NUMPAGES' : /\bPAGE\b/i.test(wv(f, 'instr') || '') ? 'PAGE' : null));
      const pairs = [];
      let li = 0, ok = true;
      for (const xr of xruns) {
        const t = OF.cleanText(sync.textOfRun(xr));
        if (fr.has(xr)) {
          while (li < leaves.length && OF.cleanText(leaves[li].textContent) !== t) li++;
          if (li < leaves.length) { const sp = leaves[li++]; sp.setAttribute('data-x', ''); sp.setAttribute('contenteditable', 'false'); if (fr.get(xr) !== 'OTHER') sp.dataset.field = fr.get(xr); }
          continue;
        }
        const from = li;
        while (li < leaves.length && OF.cleanText(leaves[li].textContent) !== t) li++;
        if (li >= leaves.length) { ok = false; break; }
        // lo saltado entre medio son resultados de campos simples
        for (let k = from; k < li; k++) leaves[k].setAttribute('data-x', '');
        pairs.push([xr, leaves[li++]]);
      }
      if (!ok) { dp.setAttribute('contenteditable', 'false'); dp.classList.add('of-ro'); return; }
      // campos simples: se usa el <span> que dibujó docx-preview o se crea uno
      const spanOf = new Map(pairs);
      const mappedRuns = pairs.map((pp) => pp[0]);
      Array.from(xp.getElementsByTagNameNS(W, 'fldSimple')).forEach((f, fi) => {
        let anchor = null;
        for (const xr of mappedRuns) if (xr.compareDocumentPosition(f) & Node.DOCUMENT_POSITION_FOLLOWING) anchor = spanOf.get(xr);
        let sp = anchor ? anchor.nextElementSibling : dp.firstElementChild;
        if (!(sp && sp.tagName === 'SPAN' && !sp.hasAttribute('data-r') && !sp.querySelector('span') && /^\d*$/.test(OF.cleanText(sp.textContent).trim()))) {
          sp = OF.el('span', {});
          if (anchor) anchor.after(sp); else dp.prepend(sp);
        }
        sp.setAttribute('data-x', ''); sp.setAttribute('contenteditable', 'false');
        if (simple[fi]) sp.dataset.field = simple[fi];
      });
      const pid = sync.addPara(xp);
      sync.mapped.add(pid);
      dp.setAttribute('data-p', pid);
      for (const [xr, span] of pairs) span.setAttribute('data-r', sync.addRun(pid, xr));
    });
    el.setAttribute('contenteditable', 'true');
    el.setAttribute('spellcheck', 'true');
    el.setAttribute('lang', 'es');
    el.classList.add('of-hf');
    el.removeAttribute('title');
    el.dataset.label = { header: 'Encabezado', footer: 'Pie de página', footnote: 'Nota al pie' }[part.kind] || '';
    OF.plainPaste(el, this);
    el.addEventListener('beforeinput', (e) => OF.typingCheckpoint && OF.typingCheckpoint(e));
    el.addEventListener('input', () => this.markDirty());
  };
  // vuelca lo editado en los encabezados/pies al XML de sus partes
  WD.applyHeaders = function () {
    for (const h of this.hfDom || []) {
      if (!h.primary || !h.part.sync) continue;
      if (h.part.sync.apply(Array.from(h.el.querySelectorAll('p'))) || h.part.sync.result) this.zip.file(h.part.path, OF.serializeXml(h.part.doc));
    }
  };

  /* --------------------------- integración --------------------------- */
  const baseRender = WD.render;
  WD.render = async function () {
    await baseRender.call(this);
    await this.loadHF();
    this.mapHeaders();
    this.setupFootnotes();
    this.paginate();
    if (!this._pgBound) {
      this._pgBound = true;
      this.host.addEventListener('input', () => this.schedulePaginate());
    }
  };
  const baseStructural = WD.structural;
  WD.structural = async function (fn) {
    this.applyHeaders();
    return baseStructural.call(this, fn);
  };
  const baseSave = WD.save;
  WD.save = async function () {
    if (this.dirty) this.applyHeaders();
    return baseSave.call(this);
  };
  const baseRoots = WD.searchRoots;
  WD.searchRoots = function () { return [...baseRoots.call(this), ...(this.hfDom || []).filter((h) => h.primary).map((h) => h.el)]; };
  WD.afterRestore = function () { this.paginate(); };

  /* ------------------------- paginación automática ------------------------- */
  WD.schedulePaginate = function () { clearTimeout(this._pgT); this._pgT = setTimeout(() => this.paginate(), 350); };
  WD.paginate = function () {
    if (!this.pages) return;
    const sects = Array.from(this.pages.querySelectorAll('section.docx'));
    let global = 0;
    this.pageMap = [];
    const placed = new Set();
    this.pages.querySelectorAll('.pg-gen').forEach((g) => { g.style.display = 'none'; });
    for (const sec of sects) {
      sec.querySelectorAll(':scope > .pg-band, :scope > .pg-clone').forEach((e) => e.remove());
      const art = sec.querySelector(':scope > article');
      if (!art) continue;
      art.querySelectorAll('tr.pg-sp').forEach((e) => e.remove());
      art.querySelectorAll('.pg-gap').forEach((g) => { const p = g.parentNode; g.remove(); p.normalize(); });
      art.querySelectorAll(':scope > [data-pgm]').forEach((el) => { el.style.marginTop = el.dataset.pgm === '_' ? '' : el.dataset.pgm; delete el.dataset.pgm; });
      if (!sec.dataset.minh) sec.dataset.minh = sec.style.minHeight || '792pt';
      sec.style.minHeight = sec.dataset.minh;
      sec.style.position = 'relative';
      const probe = OF.el('div', { style: { position: 'absolute', visibility: 'hidden', height: sec.dataset.minh, width: '1px' } });
      sec.append(probe);
      const pageH = probe.offsetHeight;
      probe.remove();
      if (!pageH) continue;
      const cs = getComputedStyle(sec);
      const padT = parseFloat(cs.paddingTop) || 0, padB = parseFloat(cs.paddingBottom) || 0, padL = parseFloat(cs.paddingLeft) || 0, padR = parseFloat(cs.paddingRight) || 0;
      const contentH = pageH - padT - padB;
      const secRect = () => sec.getBoundingClientRect();
      const scaleOf = () => secRect().height / sec.offsetHeight || 1;
      const pageOf = (y) => Math.floor(y / pageH);
      const reserve = {};
      const limitOf = (k) => k * pageH + pageH - padB - (reserve[k] || 0);
      const pushTo = (k, y) => (k + 1) * pageH + padT - y;
      // líneas de un párrafo: [{top, bottom, node, offset}] en coordenadas de la sección
      const lines = (el) => {
        const out = [];
        const sr = secRect(), z = scaleOf();
        const w = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
        let n;
        const r = document.createRange();
        while ((n = w.nextNode())) {
          if (n.parentElement.closest('[data-x]') && !n.parentElement.closest('[data-r]')) continue;
          for (let i = 0; i < n.length; i++) {
            if (n.data[i] === ' ' && i > 0) continue;
            r.setStart(n, i); r.setEnd(n, i + 1);
            const rc = r.getClientRects()[0];
            if (!rc) continue;
            const top = (rc.top - sr.top) / z, bottom = (rc.bottom - sr.top) / z;
            const last = out[out.length - 1];
            if (!last || top >= last.bottom - 2) out.push({ top, bottom, node: n, offset: i });
            else last.bottom = Math.max(last.bottom, bottom);
          }
        }
        return out;
      };
      const gap = (node, offset, h) => {
        const g = OF.el('span', { class: 'pg-gap', 'data-x': '', contenteditable: 'false', style: { display: 'block', height: h + 'px' } });
        const rest = offset > 0 ? node.splitText(offset) : node;
        rest.parentNode.insertBefore(g, rest);
      };
      const els = Array.from(art.children);
      // notas al pie: se reserva su alto al final de la página donde está la llamada
      const fnItems = (this.fnDom || []).filter((f) => sec.contains(f.ref));
      fnItems.forEach((f) => { f.el.style.display = ''; f.el.style.visibility = 'hidden'; f.el.style.top = '0px'; if (f.el.parentNode !== sec) sec.append(f.el); f.h = f.el.offsetHeight; f.page = null; });
      const FN_SEP = 14;
      const account = (el) => {
        let changed = false;
        for (const f of fnItems) {
          if (f.page != null || !el.contains(f.ref)) continue;
          const rr = f.ref.getBoundingClientRect(), sr = secRect();
          f.page = pageOf((rr.top - sr.top) / scaleOf());
          reserve[f.page] = (reserve[f.page] || FN_SEP) + f.h;
          changed = true;
        }
        return changed;
      };
      const handle = (el) => {
        const top = el.offsetTop, h = el.offsetHeight;
        if (!h) return;
        const k = pageOf(top), start = k * pageH + padT, limit = limitOf(k);
        if (top + h <= limit + 1) return;
        if (el.tagName === 'TABLE') {
          // tablas: pasan a la página siguiente fila por fila
          let guard = 0;
          for (const tr of Array.from(el.rows)) {
            if (tr.classList.contains('pg-sp') || tr.parentElement.closest('table') !== el) continue;
            const t0 = el.offsetTop + tr.offsetTop, th = tr.offsetHeight;
            const kk = pageOf(t0);
            if (t0 + th > limitOf(kk) + 1 && t0 > kk * pageH + padT + 2 && th <= contentH && guard++ < 200) {
              const sp = OF.el('tr', { class: 'pg-sp', 'data-x': '', contenteditable: 'false' }, OF.el('td', { colspan: 99, style: { height: pushTo(kk, t0) + 'px', padding: 0, border: 'none' } }));
              tr.parentNode.insertBefore(sp, tr);
            }
          }
          return;
        }
        if (el.tagName === 'P' && el.querySelector('[data-r]')) {
          // párrafos: se parten por líneas (al menos 2 líneas a cada lado, como Word)
          let guard = 0;
          while (guard++ < 20) {
            const ls = lines(el);
            const i = ls.findIndex((l) => l.bottom > limitOf(pageOf(l.top)) + 1 && l.top > pageOf(l.top) * pageH + padT - 1);
            if (i < 0) break;
            const kk = pageOf(ls[i].top);
            const before = ls.filter((l, j) => j < i && pageOf(l.top) === kk).length;
            const after = ls.length - i;
            const tall = ls[ls.length - 1].bottom - ls[0].top > contentH;
            if ((before >= 2 && after >= 2) || (tall && before >= 1)) { gap(ls[i].node, ls[i].offset, pushTo(kk, ls[i].top)); continue; }
            if (before === 0 && i === 0) break;
            // pocas líneas: el párrafo entero pasa a la página siguiente
            if (el.offsetTop > start + 2 && h <= contentH) { if (!el.dataset.pgm) el.dataset.pgm = el.style.marginTop || '_'; el.style.marginTop = (parseFloat(getComputedStyle(el).marginTop) || 0) + pushTo(k, top) + 'px'; }
            break;
          }
          return;
        }
        if (top > start + 2 && h <= contentH) {
          if (!el.dataset.pgm) el.dataset.pgm = el.style.marginTop || '_';
          el.style.marginTop = (parseFloat(getComputedStyle(el).marginTop) || 0) + pushTo(k, top) + 'px';
        }
            };
      for (const el of els) {
        handle(el);
        // si la nota de este bloque achica la página, se vuelve a ubicar el bloque
        if (account(el)) {
          handle(el);
          for (const f of fnItems) if (el.contains(f.ref) && f.page != null) { reserve[f.page] -= f.h; if (reserve[f.page] <= FN_SEP) delete reserve[f.page]; f.page = null; }
          account(el);
        }
      }
      const byPage = {};
      for (const f of fnItems) if (f.page != null) (byPage[f.page] = byPage[f.page] || []).push(f);
      sec.querySelectorAll(':scope > .fn-sep').forEach((e) => e.remove());
      for (const k in byPage) {
        let y = +k * pageH + pageH - padB - reserve[k] + FN_SEP;
        sec.append(OF.el('div', { class: 'fn-sep', 'aria-hidden': 'true', style: { top: y - 8 + 'px', left: padL + 'px' } }));
        for (const f of byPage[k]) { Object.assign(f.el.style, { top: y + 'px', left: padL + 'px', width: sec.clientWidth - padL - padR + 'px', visibility: '' }); y += f.h; }
      }
      const bottom = art.offsetTop + art.offsetHeight;
      const n = Math.max(1, Math.ceil((bottom + padB - 2) / pageH));
      sec.style.minHeight = n * pageH + 'px';
      sec.dataset.pages = n;
      const hdr = sec.querySelector(':scope > header'), ftr = sec.querySelector(':scope > footer');
      for (let k = 1; k < n; k++) sec.append(OF.el('div', { class: 'pg-band', 'aria-hidden': 'true', style: { top: k * pageH - 9 + 'px' } }));
      for (const [kind, dom] of [['header', hdr], ['footer', ftr]]) {
        if (!dom) continue;
        dom.style.visibility = '';
        const domPage = kind === 'header' ? 0 : n - 1;
        const off = dom.offsetTop - domPage * pageH;
        for (let k = 0; k < n; k++) {
          const page = global + k + 1;
          const want = this.hfFor(kind, page);
          const top = k * pageH + off;
          if (k === domPage) {
            if (want && dom.dataset.hf === want.path) { dom.querySelectorAll('[data-field=PAGE]').forEach((x) => { x.textContent = String(page); }); continue; }
            dom.style.visibility = 'hidden';
          }
          if (!want) continue;
          // la copia editable de esa parte: la del documento (si está visible) o una dibujada aquí
          let inst = (this.hfDom || []).find((h) => h.primary && h.part.path === want.path && (h.el.classList.contains('pg-gen') || h.el.style.visibility !== 'hidden'));
          if (!inst || (inst.el.classList.contains('pg-gen') && !placed.has(want.path))) {
            if (!inst) {
              for (const h of this.hfDom || []) if (h.primary && h.part.path === want.path) { h.primary = false; h.el.setAttribute('contenteditable', 'false'); }
              const g = this.renderHFPart(want);
              this.mapPart(want, g);
              g.classList.add('of-hf');
              inst = { el: g, part: want, primary: true };
              (this.hfDom = this.hfDom || []).push(inst);
            }
            if (!placed.has(want.path)) {
              placed.add(want.path);
              sec.append(inst.el);
              inst.el.style.display = '';
              Object.assign(inst.el.style, { position: 'absolute', top: top + 'px', left: dom.offsetLeft + 'px', width: dom.offsetWidth + 'px', margin: '0', minHeight: dom.style.minHeight });
              inst.el.querySelectorAll('[data-field=PAGE]').forEach((x) => { x.textContent = String(page); });
              continue;
            }
          }
          if (inst.el.textContent.trim()) sec.append(this.hfClone(inst.el, top, page, dom));
        }
      }
      this.pageMap.push({ sec, pageH, n, first: global + 1 });
      global += n;
    }
    this.totalPages = global;
    this.pages.querySelectorAll('[data-field=NUMPAGES]').forEach((s) => { s.textContent = String(global); });
    this.onselect && this.onselect();
  };
  WD.hfClone = function (src, top, page, dom) {
    const c = src.cloneNode(true);
    c.classList.remove('pg-gen', 'of-hf');
    c.classList.add('pg-clone');
    c.style.visibility = '';
    if (dom) src = dom;
    c.removeAttribute('contenteditable'); c.removeAttribute('data-hf'); c.removeAttribute('title');
    c.querySelectorAll('[contenteditable],[data-p],[data-r]').forEach((e) => { e.removeAttribute('contenteditable'); e.removeAttribute('data-p'); e.removeAttribute('data-r'); });
    c.querySelectorAll('[data-field=PAGE]').forEach((s) => { s.textContent = String(page); });
    Object.assign(c.style, { position: 'absolute', top: top + 'px', left: src.offsetLeft + 'px', width: src.offsetWidth + 'px', margin: '0' });
    return c;
  };
  WD.statusText = function () {
    const text = Array.from(this.pages.querySelectorAll('section.docx > article')).map((a) => a.innerText).join(' ');
    const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’.-]*/gu) || []).length;
    let cur = 1;
    const r = OF.currentRange();
    if (r && this.pageMap) {
      const node = r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement;
      const pm = this.pageMap.find((x) => x.sec.contains(node));
      if (pm) {
        const rect = (r.getClientRects()[0] || node.getBoundingClientRect());
        const sr = pm.sec.getBoundingClientRect();
        const scale = sr.height / pm.sec.offsetHeight || 1;
        cur = pm.first + Math.min(pm.n - 1, Math.max(0, Math.floor((rect.top - sr.top) / scale / pm.pageH)));
      }
    }
    return `Página ${cur} de ${this.totalPages || 1} · ${words.toLocaleString('es-CL')} palabra${words === 1 ? '' : 's'}`;
  };
  const basePrint = WD.printSetup;
  WD.printSetup = function () { const css = basePrint.call(this); const sec = this.pages.querySelector('section.docx'); return sec && sec.dataset.minh ? `@page { size: ${sec.style.width || '612pt'} ${sec.dataset.minh}; margin: 0 }` : css; };


  /* ------------------------------ notas al pie ------------------------------ */
  // docx-preview las lista al final (y repetidas); aquí se dibujan editables y van al pie de su página
  WD.setupFootnotes = function () {
    this.fnDom = [];
    this.pages.querySelectorAll('section.docx > ol').forEach((o) => o.remove());
    const body = OF.kid(this.xml.documentElement, 'body', W);
    const refs = Array.from(body.getElementsByTagNameNS(W, 'footnoteReference'));
    if (!refs.length || !this.fnPart) return;
    const sups = Array.from(this.pages.querySelectorAll('section.docx > article sup')).filter((x) => !x.querySelector('sup') && !x.closest('[data-r]') && /^\s*\d+\s*$/.test(x.textContent));
    if (sups.length !== refs.length) return;
    const notes = new Map(OF.all(this.fnPart.doc, 'footnote', W).map((f) => [wv(f, 'id'), f]));
    refs.forEach((ref, i) => {
      const id = wv(ref, 'id'), fn = notes.get(id);
      const sup = sups[i];
      sup.textContent = String(i + 1);
      const box = sup.closest('span') || sup;
      box.setAttribute('data-x', ''); box.setAttribute('contenteditable', 'false'); box.classList.add('fn-ref');
      if (!fn) return;
      const part = { kind: 'footnote', path: this.fnPart.path, doc: this.fnPart.doc, root: fn, num: i + 1 };
      const el = this.renderHFPart(part);
      this.mapPart(part, el);
      el.style.position = 'absolute';
      this.fnDom.push({ el, ref: sup, part });
      (this.hfDom = this.hfDom || []).push({ el, part, primary: true });
    });
  };
  // insertar una nota al pie en el cursor
  WD.insertFootnote = async function () {
    const r = OF.currentRange();
    if (!r || !r.startContainer.parentElement || !r.startContainer.parentElement.closest('section.docx > article [data-p]')) { OF.toast('Toca el texto donde va la llamada de la nota.'); return; }
    const MARK = '\u2063';
    r.collapse(false);
    r.insertNode(document.createTextNode(MARK));
    let newId = null;
    await this.structural(async () => {
      const body = OF.kid(this.xml.documentElement, 'body', W);
      const t = Array.from(body.getElementsByTagNameNS(W, 't')).find((x) => x.textContent.includes(MARK));
      if (!t) return false;
      if (!this.fnPart) {
        const path = 'word/footnotes.xml';
        const doc = OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:footnotes xmlns:w="${W}"><w:footnote w:type="separator" w:id="-1"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:separator/></w:r></w:p></w:footnote><w:footnote w:type="continuationSeparator" w:id="0"><w:p><w:pPr><w:spacing w:after="0" w:line="240" w:lineRule="auto"/></w:pPr><w:r><w:continuationSeparator/></w:r></w:p></w:footnote></w:footnotes>`);
        this.zip.file(path, OF.serializeXml(doc));
        await OF.addRel(this.zip, this.docPath, REL + 'footnotes', path);
        await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, path, 'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml'));
        this.fnPart = { path, doc };
      }
      const doc = this.fnPart.doc;
      newId = Math.max(0, ...OF.all(doc, 'footnote', W).map((f) => +wv(f, 'id') || 0)) + 1;
      const st = Object.values(this.styles || {}).find((x) => x.type === 'paragraph' && x.name === 'footnote text');
      const rst = Object.values(this.styles || {}).find((x) => x.type === 'character' && x.name === 'footnote reference');
      const refPr = rst ? `<w:rPr><w:rStyle w:val="${rst.id}"/></w:rPr>` : '<w:rPr><w:vertAlign w:val="superscript"/></w:rPr>';
      doc.documentElement.appendChild(OF.fragment(doc, `<w:footnote xmlns:w="${W}" w:id="${newId}"><w:p><w:pPr>${st ? `<w:pStyle w:val="${st.id}"/>` : '<w:spacing w:after="0" w:line="240" w:lineRule="auto"/>'}</w:pPr><w:r>${refPr}<w:footnoteRef/></w:r><w:r><w:rPr>${st ? '' : '<w:sz w:val="18"/>'}</w:rPr><w:t xml:space="preserve"> Escribe aquí la nota.</w:t></w:r></w:p></w:footnote>`));
      this.zip.file(this.fnPart.path, OF.serializeXml(doc));
      // partir el fragmento donde estaba el cursor y meter la llamada
      const run = t.parentNode, txt = t.textContent, at = txt.indexOf(MARK);
      const after = run.cloneNode(true);
      OF.setText(t, txt.slice(0, at));
      OF.setText(OF.kid(after, 't', W), txt.slice(at + 1));
      const refRun = OF.fragment(this.xml, `<w:r xmlns:w="${W}">${refPr}<w:footnoteReference w:id="${newId}"/></w:r>`);
      run.after(refRun, after);
      if (!txt.slice(0, at)) run.remove();
      if (!txt.slice(at + 1)) after.remove();
      // las notas quedan en el mismo orden que sus llamadas (algunos programas lo esperan)
      const order = Array.from(body.getElementsByTagNameNS(W, 'footnoteReference')).map((x) => wv(x, 'id'));
      const fns = OF.all(doc, 'footnote', W).filter((f) => !wv(f, 'type') || wv(f, 'type') === 'normal');
      fns.sort((a, b) => order.indexOf(wv(a, 'id')) - order.indexOf(wv(b, 'id'))).forEach((f) => doc.documentElement.appendChild(f));
      // e identificadores correlativos 1, 2, 3… en ese orden
      const map = new Map(order.map((id, i) => [id, String(i + 1)]));
      for (const x of body.getElementsByTagNameNS(W, 'footnoteReference')) x.setAttributeNS(W, 'w:id', map.get(wv(x, 'id')));
      for (const f of fns) if (map.has(wv(f, 'id'))) f.setAttributeNS(W, 'w:id', map.get(wv(f, 'id')));
      newId = map.get(String(newId)) || newId;
      this.zip.file(this.fnPart.path, OF.serializeXml(doc));
    });
    const f = (this.fnDom || []).find((x) => x.part.root && wv(x.part.root, 'id') === String(newId));
    if (f) {
      const sp = f.el.querySelector('[data-r]');
      f.el.focus();
      const rr = document.createRange();
      if (sp && sp.firstChild && sp.firstChild.nodeType === 3) { rr.setStart(sp.firstChild, /^\s/.test(sp.firstChild.data) ? 1 : 0); rr.setEnd(sp.firstChild, sp.firstChild.length); } else rr.selectNodeContents(sp || f.el);
      const s = getSelection(); s.removeAllRanges(); s.addRange(rr); OF.lastRange = rr.cloneRange();
      f.el.scrollIntoView({ block: 'center' });
    }
  };
  /* --------------------- insertar encabezado, pie y número --------------------- */
  WD.lastSectPr = function () {
    const body = OF.kid(this.xml.documentElement, 'body', W);
    let s = OF.kid(body, 'sectPr', W);
    if (!s) { s = this.xml.createElementNS(W, 'w:sectPr'); body.appendChild(s); }
    return s;
  };
  WD.ensureHF = async function (kind, type) {
    type = type || 'default';
    const sect = this.lastSectPr();
    const ref = OF.kids(sect, kind + 'Reference', W).find((r) => (wv(r, 'type') || 'default') === type);
    if (ref) { const rel = this.rels[ref.getAttributeNS(OF.NS.r, 'id')]; if (rel && this.hfParts[rel.target]) return this.hfParts[rel.target]; }
    const path = OF.uniquePart(this.zip, 'word', kind, 'xml');
    const tag = kind === 'header' ? 'hdr' : 'ftr';
    const st = Object.values(this.styles || {}).find((s) => s.type === 'paragraph' && s.name === kind);
    const doc = OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:${tag} xmlns:w="${W}" xmlns:r="${OF.NS.r}"><w:p>${st ? `<w:pPr><w:pStyle w:val="${st.id}"/></w:pPr>` : ''}</w:p></w:${tag}>`);
    this.zip.file(path, OF.serializeXml(doc));
    const rid = await OF.addRel(this.zip, this.docPath, REL + kind, path);
    await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, path, CT[kind]));
    const e = this.xml.createElementNS(W, `w:${kind}Reference`);
    e.setAttributeNS(W, 'w:type', type);
    e.setAttributeNS(OF.NS.r, 'r:id', rid);
    // headerReference y footerReference van al comienzo de sectPr
    let before = sect.firstElementChild;
    while (before && /^(headerReference|footerReference)$/.test(before.localName)) before = before.nextElementSibling;
    sect.insertBefore(e, before);
    this.rels = await OF.readRels(this.zip, this.docPath);
    const part = { rid, path, kind, doc };
    this.hfParts[path] = part;
    return part;
  };
  const SECT_AFTER_TITLE = ['textDirection', 'bidi', 'rtlGutter', 'docGrid', 'printerSettings', 'sectPrChange'];
  WD.toggleFirstPage = async function () {
    const on = !this.titlePg;
    await this.structural(async () => {
      const sect = this.lastSectPr();
      OF.removeKids(sect, ['titlePg'], W);
      if (on) {
        const e = this.xml.createElementNS(W, 'w:titlePg');
        sect.insertBefore(e, Array.from(sect.children).find((c) => SECT_AFTER_TITLE.includes(c.localName)) || null);
        await this.ensureHF('header', 'first'); await this.ensureHF('footer', 'first');
      }
    });
    OF.toast(on ? 'Primera página con encabezado y pie propios (vacíos): escríbelos en la página 1.' : 'La primera página usa el mismo encabezado y pie que las demás.', 4000);
  };
  const SET_AFTER_EVEN = ['bookFoldRevPrinting', 'bookFoldPrinting', 'bookFoldPrintingSheets', 'drawingGridHorizontalSpacing', 'drawingGridVerticalSpacing', 'displayHorizontalDrawingGridEvery', 'displayVerticalDrawingGridEvery', 'doNotUseMarginsForDrawingGridOrigin', 'drawingGridHorizontalOrigin', 'drawingGridVerticalOrigin', 'doNotShadeFormData', 'noPunctuationKerning', 'characterSpacingControl', 'printTwoOnOne', 'strictFirstAndLastChars', 'noLineBreaksAfter', 'noLineBreaksBefore', 'savePreviewPicture', 'doNotValidateAgainstSchema', 'saveInvalidXml', 'ignoreMixedContent', 'alwaysShowPlaceholderText', 'doNotDemarcateInvalidXml', 'saveXmlDataOnly', 'useXSLTWhenSaving', 'saveThroughXslt', 'showXMLTags', 'alwaysMergeEmptyNamespace', 'updateFields', 'hdrShapeDefaults', 'footnotePr', 'endnotePr', 'compat', 'docVars', 'rsids', 'mathPr', 'attachedSchema', 'themeFontLang', 'clrSchemeMapping', 'doNotIncludeSubdocsInStats', 'doNotAutoCompressPictures', 'forceUpgrade', 'captions', 'readModeInkLockDown', 'smartTagType', 'schemaLibrary', 'shapeDefaults', 'doNotEmbedSmartTags', 'decimalSymbol', 'listSeparator'];
  WD.toggleEvenOdd = async function () {
    if (!this.settingsDoc) { OF.toast('Este documento no tiene configuración; no se puede cambiar.'); return; }
    const on = !this.evenOdd;
    await this.structural(async () => {
      const root = this.settingsDoc.documentElement;
      OF.removeKids(root, ['evenAndOddHeaders'], W);
      if (on) {
        root.insertBefore(this.settingsDoc.createElementNS(W, 'w:evenAndOddHeaders'), Array.from(root.children).find((c) => SET_AFTER_EVEN.includes(c.localName)) || null);
        await this.ensureHF('header', 'even'); await this.ensureHF('footer', 'even');
      }
      this.zip.file(this.settingsPath, OF.serializeXml(this.settingsDoc));
    });
    OF.toast(on ? 'Páginas pares con encabezado y pie propios (vacíos): escríbelos en la página 2.' : 'Pares e impares usan el mismo encabezado y pie.', 4000);
  };
  WD.editHF = async function (kind) {
    const has = (this.hfDom || []).find((h) => h.primary && h.part.kind === kind);
    if (!has) {
      await this.structural(async () => { await this.ensureHF(kind); });
    }
    const h = (this.hfDom || []).find((x) => x.primary && x.part.kind === kind);
    if (!h) return;
    const p = h.el.querySelector('[data-p]') || h.el.querySelector('p');
    h.el.focus();
    const r = document.createRange();
    r.selectNodeContents(p || h.el); r.collapse(false);
    const s = getSelection(); s.removeAllRanges(); s.addRange(r); OF.lastRange = r.cloneRange();
    h.el.scrollIntoView({ block: 'center' });
    OF.toast(kind === 'header' ? 'Escribe el encabezado; se repite en todas las páginas.' : 'Escribe el pie de página; se repite en todas las páginas.', 3500);
  };
  WD.pageNumber = async function (where) {
    await this.structural(async () => {
      const part = await this.ensureHF(where || 'footer');
      const root = part.doc.documentElement;
      const doc = part.doc;
      const p = OF.fragment(doc, `<w:p xmlns:w="${W}"><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t xml:space="preserve">Página </w:t></w:r><w:fldSimple w:instr=" PAGE "><w:r><w:t>1</w:t></w:r></w:fldSimple><w:r><w:t xml:space="preserve"> de </w:t></w:r><w:fldSimple w:instr=" NUMPAGES "><w:r><w:t>1</w:t></w:r></w:fldSimple></w:p>`);
      const ps = OF.kids(root, 'p', W);
      const emptyOnly = ps.length === 1 && !ps[0].textContent.trim() && !OF.all(ps[0], 'drawing').length;
      if (emptyOnly) root.replaceChild(p, ps[0]); else root.appendChild(p);
      this.zip.file(part.path, OF.serializeXml(doc));
    });
    OF.toast(where === 'header' ? 'Número de página agregado al encabezado.' : 'Número de página agregado al pie de página.');
  };
  WD.removeHF = async function (kind) {
    await this.structural(async () => {
      const sect = this.lastSectPr();
      OF.removeKids(sect, [kind + 'Reference'], W);
    });
  };
})();

/* Word: ajuste de texto de las imágenes (en línea, cuadrado izquierda/derecha, arriba y abajo). */
(function () {
  const OF = window.OF;
  const W = OF.NS.w;
  const WP = 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing';
  const WD = OF.WordDoc.prototype;
  WD.imageWrapMenu = function (i, btn) {
    const dr = this.drawings && this.drawings[i];
    if (!dr) return;
    const box = dr.getElementsByTagNameNS(WP, 'anchor')[0];
    const cur = !box ? 'inline' : box.getElementsByTagNameNS(WP, 'wrapNone')[0] ? 'front' : box.getElementsByTagNameNS(WP, 'wrapTopAndBottom')[0] ? 'center' : ((box.getElementsByTagNameNS(WP, 'positionH')[0] || {}).textContent || '').trim() === 'right' ? 'right' : 'left';
    OF.popup(btn, (p, close) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Ajuste del texto · arrastra la imagen para moverla'));
      for (const [k, l] of [['inline', 'En línea con el texto'], ['left', 'Cuadrado, a la izquierda'], ['right', 'Cuadrado, a la derecha'], ['center', 'Arriba y abajo, centrada'], ['front', 'Delante del texto (posición libre)']]) {
        p.append(OF.el('button', { class: 'pop-item' + (k === cur ? ' on' : ''), onclick: () => { close(); this.imageWrap(i, k); } }, (k === cur ? '✓ ' : '') + l));
      }
    }, false);
  };
  WD.imageWrap = async function (i, mode, opts) {
    const dr = this.drawings && this.drawings[i];
    if (!dr) return;
    OF.closeImageBar();
    await this.structural(() => {
      const old = dr.getElementsByTagNameNS(WP, 'inline')[0] || dr.getElementsByTagNameNS(WP, 'anchor')[0];
      if (!old) return false;
      const doc = this.xml;
      const pick = (n) => old.getElementsByTagNameNS(WP, n)[0];
      const keep = ['extent', 'effectExtent', 'docPr', 'cNvGraphicFramePr'].map(pick).filter(Boolean);
      const graphic = Array.from(old.children).find((c) => c.localName === 'graphic');
      const mk = (n, attrs) => { const e = doc.createElementNS(WP, 'wp:' + n); for (const k in attrs || {}) e.setAttribute(k, attrs[k]); return e; };
      let nu;
      if (mode === 'inline') {
        nu = mk('inline', { distT: '0', distB: '0', distL: '0', distR: '0' });
        for (const e of keep) nu.appendChild(e);
      } else {
        const ids = Array.from(doc.getElementsByTagNameNS(WP, 'anchor')).map((a) => +a.getAttribute('relativeHeight') || 0);
        nu = mk('anchor', { distT: '0', distB: '0', distL: '114300', distR: '114300', simplePos: '0', relativeHeight: String(Math.max(251658240, ...ids) + 1024), behindDoc: '0', locked: '0', layoutInCell: '1', allowOverlap: '1' });
        nu.appendChild(mk('simplePos', { x: '0', y: '0' }));
        const ph = mk('positionH', { relativeFrom: mode === 'center' ? 'margin' : 'column' });
        if (mode === 'front') { const o = mk('posOffset'); o.textContent = String(Math.round((opts && opts.x) || 0)); ph.appendChild(o); }
        else { const al = mk('align'); al.textContent = mode === 'right' ? 'right' : mode === 'center' ? 'center' : 'left'; ph.appendChild(al); }
        const pv = mk('positionV', { relativeFrom: 'paragraph' }), po = mk('posOffset'); po.textContent = String(Math.round((opts && opts.y) || 0)); pv.appendChild(po);
        nu.append(ph, pv);
        const ext = keep.find((e) => e.localName === 'extent'); if (ext) nu.appendChild(ext);
        nu.appendChild(keep.find((e) => e.localName === 'effectExtent') || mk('effectExtent', { l: '0', t: '0', r: '0', b: '0' }));
        nu.appendChild(mode === 'center' ? mk('wrapTopAndBottom') : mode === 'front' ? mk('wrapNone') : mk('wrapSquare', { wrapText: 'bothSides' }));
        for (const e of keep) if (e.localName === 'docPr' || e.localName === 'cNvGraphicFramePr') nu.appendChild(e);
      }
      if (graphic) nu.appendChild(graphic);
      old.parentNode.replaceChild(nu, old);
      // el objeto anclado va al comienzo de su párrafo (o del que se eligió al arrastrar)
      if (mode !== 'inline') {
        let r = nu; while (r && r.localName !== 'r') r = r.parentNode;
        const p = (opts && opts.para) || (r && r.parentNode);
        if (r && p && p.localName === 'p') { const ppr = Array.from(p.children).find((c) => c.localName === 'pPr'); p.insertBefore(r, ppr ? ppr.nextSibling : p.firstChild); }
      }
    });
  };
  // arrastrar una imagen flotante: libre si está delante del texto; si el texto la rodea, cambia de párrafo y de lado
  WD.bindImageDrag = function () {
    if (this._imgDrag === this.pages) return;
    this._imgDrag = this.pages;
    this.pages.addEventListener('pointerdown', (e) => {
      const im = e.target.closest('img.of-img');
      if (!im || e.button) return;
      const i = +im.dataset.di, dr = this.drawings && this.drawings[i];
      const anc = dr && dr.getElementsByTagNameNS(WP, 'anchor')[0];
      if (!anc) return;
      e.preventDefault();
      const sec = im.closest('section.docx');
      const z = sec.getBoundingClientRect().width / sec.offsetWidth || 1;
      const st = { x: e.clientX, y: e.clientY, moved: false, dx: 0, dy: 0 };
      const move = (ev) => {
        st.dx = (ev.clientX - st.x) / z; st.dy = (ev.clientY - st.y) / z;
        if (!st.moved && Math.abs(st.dx) + Math.abs(st.dy) > 4) { st.moved = true; OF.closeImageBar(); im.classList.add('dragging'); }
        if (st.moved) im.style.transform = `translate(${st.dx}px, ${st.dy}px)`;
      };
      const up = async (ev) => {
        window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
        if (!st.moved) return;
        const free = !!anc.getElementsByTagNameNS(WP, 'wrapNone')[0];
        const ir = im.getBoundingClientRect();
        im.style.visibility = 'hidden';
        const under = document.elementFromPoint(ev.clientX, ev.clientY);
        im.style.visibility = ''; im.style.transform = ''; im.classList.remove('dragging');
        const dp = under && under.closest('section.docx > article [data-p]');
        const para = dp ? this.sync.paras[+dp.dataset.p] : null;
        if (free) {
          // posición relativa al párrafo donde se suelta
          const pr = (dp || im.closest('p') || sec).getBoundingClientRect();
          const art = sec.querySelector('article').getBoundingClientRect();
          await this.imageWrap(i, 'front', { x: ((ir.left - art.left) / z) * 9525, y: ((ir.top - pr.top) / z) * 9525, para });
        } else {
          const art = sec.querySelector('article').getBoundingClientRect();
          const mode = anc.getElementsByTagNameNS(WP, 'wrapTopAndBottom')[0] ? 'center' : ev.clientX < art.left + art.width / 2 ? 'left' : 'right';
          await this.imageWrap(i, mode, { para });
        }
      };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    });
  };
  const baseRender3 = WD.render;
  WD.render = async function () {
    await baseRender3.call(this);
    this.bindImageDrag();
    this.pages.querySelectorAll('img.of-img').forEach((im) => { const dr = this.drawings && this.drawings[+im.dataset.di]; if (dr && dr.getElementsByTagNameNS(WP, 'anchor')[0]) {
      im.classList.add('of-float'); im.title = 'Arrastra para mover';
      // «delante del texto»: fuera del flujo, relativo al párrafo (como Word, sin dejar un hueco)
      const box = im.parentElement;
      if (dr.getElementsByTagNameNS(WP, 'wrapNone')[0] && box && box.style.width === '0px') { box.style.position = 'absolute'; box.style.zIndex = '1'; const p = box.closest('p'); if (p) p.style.position = 'relative'; }
    } });
  };
})();
