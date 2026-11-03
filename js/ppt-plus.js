/* PowerPoint: miniaturas, mover/redimensionar cualquier objeto, cuadros de texto, formas, tablas,
   diseño al crear diapositiva, fondo, viñetas, interlineado, fuentes, buscar, deshacer, imprimir. */
(function () {
  const OF = window.OF;
  const A = OF.NS.a, P = OF.NS.p, R = OF.NS.r;
  const EMU = 9525;
  const PD = OF.PptDoc.prototype;
  const PPR_ORDER = ['lnSpc', 'spcBef', 'spcAft', 'buClrTx', 'buClr', 'buSzTx', 'buSzPct', 'buSzPts', 'buFontTx', 'buFont', 'buNone', 'buAutoNum', 'buChar', 'buBlip', 'tabLst', 'defRPr', 'extLst'];
  const LAYOUT_ES = { 'title slide': 'Diapositiva de título', 'title and content': 'Título y contenido', 'section header': 'Encabezado de sección', 'two content': 'Dos contenidos', comparison: 'Comparación', 'title only': 'Solo el título', blank: 'En blanco', 'content with caption': 'Contenido con título', 'picture with caption': 'Imagen con título', 'title and vertical text': 'Título y texto vertical', 'vertical title and text': 'Título vertical y texto' };
  const frag = (doc, xml) => OF.fragment(doc, xml);
  const NSD = `xmlns:a="${A}" xmlns:p="${P}" xmlns:r="${R}"`;

  /* ------------------------------ vista ------------------------------ */
  const baseRender = PD.render;
  PD.render = async function () {
    await baseRender.call(this);
    this.buildThumbs();
    this.afterRender();
  };
  PD.fit = function () {
    if (!this.list) return;
    this.enableEditing();
    this.bindPictures();
    const thumbs = this.thumbs && this.thumbs.offsetParent ? this.thumbs.offsetWidth + 8 : 0;
    const availW = Math.max(200, this.host.clientWidth - 60 - thumbs);
    const availH = Math.max(200, this.host.clientHeight - 50);
    const z = Math.min(1.6, availW / this.W, Math.max(availH / this.H, 0.55 * availW / this.W)) * (this.userZoom || 1);
    this.list.querySelectorAll('.pp-frame').forEach((f) => { f.style.width = this.W * z + 'px'; f.style.height = this.H * z + 'px'; f.firstChild.style.transform = `scale(${z})`; });
    this.scale = z;
    if (this.selPic && this.selPic.isConnected) this.placeFrame();
  };
  PD.enableEditing = function () {
    this.list.querySelectorAll('[data-slide-edit]').forEach((tx) => {
      if (tx.dataset.bound) return;
      tx.dataset.bound = '1';
      tx.setAttribute('contenteditable', 'true');
      tx.setAttribute('spellcheck', 'true');
      tx.setAttribute('lang', 'es');
      OF.plainPaste(tx, this);
      tx.addEventListener('beforeinput', (e) => OF.typingCheckpoint && OF.typingCheckpoint(e));
      tx.addEventListener('input', () => {
        this.markDirty();
        if (tx.textContent.trim()) tx.classList.remove('pp-empty');
        const w = tx.closest('.pp-wrap');
        if (w) this.updateThumb(+w.dataset.i);
      });
      tx.addEventListener('focus', () => { const sh = tx.closest('.pp-obj'); if (sh) this.selectPic(sh, true); });
    });
  };
  PD.afterRender = function () {
    let target = null;
    if (this.caretTarget) {
      this.syncs.forEach((sync, si) => { const pid = sync.paras.indexOf(this.caretTarget); if (pid >= 0) target = this.slides[si].box.querySelector(`[data-slide-edit] [data-p="${pid}"]`); });
    }
    if (target) {
      const tx = target.closest('[data-slide-edit]');
      const r = document.createRange();
      const spans = target.querySelectorAll('[data-r]');
      if (spans.length) { r.selectNodeContents(spans[spans.length - 1]); r.collapse(false); } else { r.selectNodeContents(target); r.collapse(true); }
      try { tx.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r); OF.lastRange = r.cloneRange(); } catch (e) {}
    }
    if (this.curSlide != null) { const w = this.list.querySelector(`.pp-wrap[data-i="${this.curSlide}"]`); if (w && this.scrollToCur) { this._lockCur = Date.now() + 600; w.scrollIntoView({ block: 'center' }); } this.markCur(); }
    this.scrollToCur = false;
    this.caretTarget = null;
  };

  /* ---------------------------- miniaturas ---------------------------- */
  PD.buildThumbs = function () {
    const host = this.host;
    host.classList.add('pp-host');
    if (this.thumbs) this.thumbs.remove();
    const t = (this.thumbs = OF.el('nav', { class: 'pp-thumbs', 'aria-label': 'Diapositivas' }));
    host.insertBefore(t, this.list);
    this.slides.forEach((s, i) => {
      const b = OF.el('button', { class: 'pp-th' + (i === (this.curSlide || 0) ? ' on' : '') + (s.hidden ? ' hid' : ''), 'data-i': i, type: 'button', 'aria-label': 'Diapositiva ' + (i + 1) }, OF.el('span', { class: 'pp-th-n' }, String(i + 1)), OF.el('div', { class: 'pp-th-box' }));
      b.addEventListener('click', () => { this.curSlide = i; this.markCur(); this._lockCur = Date.now() + 900; const w = this.list.querySelector(`.pp-wrap[data-i="${i}"]`); if (w) w.scrollIntoView({ block: 'center', behavior: 'smooth' }); });
      b.addEventListener('contextmenu', (e) => { e.preventDefault(); this.curSlide = i; this.markCur(); OF.popup(b, (p, close) => {
        const it = (label, fn, danger) => p.append(OF.el('button', { class: 'pop-item' + (danger ? ' danger' : ''), onclick: () => { close(); fn(); } }, label));
        it('Nueva diapositiva', () => this.newSlideMenu(b));
        it('Duplicar', () => this.slideOp('dup'));
        it('Mover hacia arriba', () => this.slideOp('up'));
        it('Mover hacia abajo', () => this.slideOp('down'));
        it('Eliminar', () => this.slideOp('delete'), true);
      }, true); });
      t.append(b);
      this.updateThumb(i, true);
    });
    // la diapositiva visible se marca al desplazarse
    if (this._io) this._io.disconnect();
    this._vis = {};
    this._io = new IntersectionObserver((ents) => {
      if (this._lockCur && Date.now() < this._lockCur) { for (const e of ents) this._vis[e.target.dataset.i] = e.intersectionRatio; return; }
      for (const e of ents) this._vis[e.target.dataset.i] = e.intersectionRatio;
      let best = null, br = 0;
      for (const k in this._vis) if (this._vis[k] > br) { br = this._vis[k]; best = +k; }
      if (best != null && best !== this.curSlide) { this.curSlide = best; this.markCur(true); }
    }, { root: host, threshold: [0, 0.25, 0.5, 0.75, 1] });
    this.list.querySelectorAll('.pp-wrap').forEach((w) => this._io.observe(w));
  };
  PD.markCur = function (fromScroll) {
    if (!this.thumbs) return;
    this.thumbs.querySelectorAll('.pp-th').forEach((b) => b.classList.toggle('on', +b.dataset.i === this.curSlide));
    this.list.querySelectorAll('.pp-wrap').forEach((w) => w.classList.toggle('cur', +w.dataset.i === this.curSlide));
    const on = this.thumbs.querySelector('.pp-th.on');
    if (on && fromScroll) { const r = on.getBoundingClientRect(), t = this.thumbs.getBoundingClientRect(); if (r.top < t.top || r.bottom > t.bottom) on.scrollIntoView({ block: 'nearest' }); }
    this.onselect && this.onselect();
  };
  PD.updateThumb = function (i, now) {
    if (!this.thumbs) return;
    clearTimeout((this._thT = this._thT || {})[i]);
    const go = () => {
      const b = this.thumbs && this.thumbs.querySelector(`.pp-th[data-i="${i}"] .pp-th-box`);
      const s = this.slides[i];
      if (!b || !s || !s.box) return;
      const k = 150 / this.W;
      const c = s.box.cloneNode(true);
      c.querySelectorAll('[contenteditable]').forEach((e) => e.removeAttribute('contenteditable'));
      c.querySelectorAll('.pp-handle, .pp-grip, .pp-selframe').forEach((e) => e.remove());
      c.querySelectorAll('.pp-selpic, .pp-empty').forEach((e) => e.classList.remove('pp-selpic', 'pp-empty'));
      c.style.transform = `scale(${k})`;
      b.style.width = 150 + 'px'; b.style.height = this.H * k + 'px';
      b.innerHTML = '';
      b.append(c);
    };
    if (now) go(); else this._thT[i] = setTimeout(go, 500);
  };

  /* --------------------- seleccionar, mover y cambiar tamaño --------------------- */
  PD.regShape = function (d, el, sync, T, rot, hasText) {
    if (!sync) return;
    if (T.inGroup) {
      // hijo directo de un grupo de primer nivel: se puede seleccionar y mover por separado
      if (!T.grp) return;
      d._grp = T.grp;
      d.dataset.inGrp = '1';
    } else if (T.sx !== 1 || T.sy !== 1 || T.x(0) !== 0) return;
    d.dataset.pe = this.picReg.push({ el, sync }) - 1;
    d.classList.add(hasText ? 'pp-obj' : 'pp-pic');
    if (rot) d.dataset.rot = rot;
  };
  PD.bindPictures = function () {
    if (this._boundList === this.list) return;
    this._boundList = this.list;
    this.list.addEventListener('pointerdown', (e) => {
      const grip = e.target.closest('.pp-grip'), handle = e.target.closest('.pp-handle');
      if (e.target.closest('.pp-grip-x')) return;
      if (e.target.closest('.pp-rot')) { this.rotateDrag(e); return; }
      const fr = e.target.closest('.pp-selframe');
      let d = fr ? fr._for : e.target.closest('.pp-pic');
      // primer toque en un objeto agrupado selecciona el grupo; el segundo, el objeto
      if (d && !fr && d.dataset.inGrp) { const g = d.parentElement.closest('.pp-grpbox'); if (g && this.selPic !== d && !(this.selPic && (this.selPic === g || this.selPic.parentElement === d.parentElement))) d = g; }
      if (!d) { if (!e.target.closest('.imgbar') && !e.target.closest('.pp-obj')) this.selectPic(null); return; }
      // texto dentro de un grupo: se edita; el grupo queda seleccionado para moverlo con «Mover»
      if (!fr && d.dataset.grp && e.target.closest('[data-slide-edit]')) { this.selectPic(d); return; }
      e.preventDefault();
      if (!grip && !handle) this.selectPic(d);
      const z = this.scale || 1;
      const gbox = d.dataset.inGrp && d.parentElement.closest('.pp-grpbox');
      const st = { gr: gbox ? ((+gbox.dataset.rot || 0) * Math.PI) / 180 : 0, x: e.clientX, y: e.clientY, l: parseFloat(d.style.left), t: parseFloat(d.style.top), w: parseFloat(d.style.width), h: parseFloat(d.style.height), moved: false };
      const keepRatio = d.classList.contains('pp-pic') && !!this.picReg[+d.dataset.pe] && this.picReg[+d.dataset.pe].el.localName === 'pic';
      const move = (ev) => {
        let dx = (ev.clientX - st.x) / z, dy = (ev.clientY - st.y) / z;
        if (st.gr) { const c = Math.cos(st.gr), s2 = Math.sin(st.gr); [dx, dy] = [dx * c + dy * s2, -dx * s2 + dy * c]; }
        if (!st.moved && Math.abs(dx) + Math.abs(dy) > 2) { st.moved = true; OF.checkpoint && OF.checkpoint(); }
        if (handle) {
          // en objetos girados se mide en los ejes del objeto y se mantiene fija la esquina opuesta
          const a = ((+d.dataset.rot || 0) * Math.PI) / 180, cos = Math.cos(a), sin = Math.sin(a);
          const lx = dx * cos + dy * sin, ly = -dx * sin + dy * cos;
          const w = Math.max(16, st.w + lx);
          const h = keepRatio ? (w * st.h) / st.w : Math.max(12, st.h + ly);
          d.style.width = w + 'px';
          d.style.height = h + 'px';
          if (a) {
            const ddx = (w - st.w) / 2, ddy = (h - st.h) / 2;
            const cx = st.l + st.w / 2 + ddx * cos - ddy * sin, cy = st.t + st.h / 2 + ddx * sin + ddy * cos;
            d.style.left = cx - w / 2 + 'px'; d.style.top = cy - h / 2 + 'px';
          }
          const svg = d.querySelector(':scope > svg');
          if (svg && svg.getAttribute('viewBox')) { svg.setAttribute('width', d.style.width); svg.setAttribute('height', d.style.height); }
        } else { d.style.left = st.l + dx + 'px'; d.style.top = st.t + dy + 'px'; }
        this.placeFrame();
      };
      const up = () => {
        window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
        if (st.moved) {
          this.writePic(d);
          const w = d.closest('.pp-wrap'); if (w) this.updateThumb(+w.dataset.i);
          // un grupo cambiado de tamaño se vuelve a dibujar (sus hijos se escalan)
          if (handle && d.dataset.grp) this.redrawKeepSel(d);
        }
      };
      window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
    });
    if (!this._keyBound) {
      this._keyBound = true;
      document.addEventListener('keydown', (e) => {
        if (!this.selPic || !this.selPic.isConnected) return;
        if ((e.key === 'Delete' || e.key === 'Backspace') && !e.target.closest('[contenteditable=true]')) { e.preventDefault(); this.picAction('del'); }
        if (e.key === 'Escape') this.selectPic(null);
      });
    }
  };
  PD.selectPic = function (d, fromFocus) {
    if (this.selPic === d && d) { this.placeFrame(); return; }
    if (this.selPic) { this.selPic.classList.remove('pp-selpic'); }
    if (this.frame) { this.frame.remove(); this.frame = null; }
    this.selPic = d;
    OF.closeImageBar();
    if (!d) return;
    d.classList.add('pp-selpic');
    const isText = d.classList.contains('pp-obj') || !!d.dataset.grp;
    const f = (this.frame = OF.el('div', { class: 'pp-selframe', 'aria-hidden': 'true' }));
    if (isText) {
      const grip = OF.el('div', { class: 'pp-grip', title: 'Arrastra para mover' }, OF.el('span', null, '⠿ Mover'));
      const x = OF.el('button', { class: 'pp-grip-x', type: 'button', title: 'Eliminar cuadro', 'aria-label': 'Eliminar cuadro' }, '×');
      x.addEventListener('mousedown', (e) => e.preventDefault());
      x.addEventListener('click', (e) => { e.stopPropagation(); this.picAction('del'); });
      grip.append(x);
      f.append(grip);
      f.classList.add('has-grip');
    }
    f.append(OF.el('div', { class: 'pp-handle' }));
    const reg = this.picReg[+d.dataset.pe];
    if (reg && reg.el.localName !== 'graphicFrame') f.append(OF.el('div', { class: 'pp-rot', title: 'Girar' }));
    // los controles viven en una capa aparte para no tocar el contenido del objeto
    f._for = d;
    d.parentNode.append(f);
    this.placeFrame();
    if (!isText) OF.imageBar(d, d.querySelector('img') ? [['Más pequeña', () => this.picAction(0.8)], ['Más grande', () => this.picAction(1.25)], ['Eliminar', () => this.picAction('del'), true]] : [['Más pequeño', () => this.picAction(0.8)], ['Más grande', () => this.picAction(1.25)], ['Eliminar', () => this.picAction('del'), true]]);
    void fromFocus;
  };
  PD.placeFrame = function () {
    const d = this.selPic, f = this.frame;
    if (!d || !f) return;
    Object.assign(f.style, { left: d.style.left, top: d.style.top, width: d.style.width, height: d.style.height, transform: d.style.transform || '' });
    const inv = 1 / (this.scale || 1);
    f.style.setProperty('--inv', inv);
    f.classList.toggle('grip-below', parseFloat(d.style.top) < 34 * inv);
    OF._imgBar && OF._imgBar.place();
  };
  PD.picXfrm = function (d) {
    const { el } = this.picReg[+d.dataset.pe];
    if (el.localName === 'graphicFrame') {
      let x = OF.kid(el, 'xfrm', P);
      if (!x) { x = frag(el.ownerDocument, `<p:xfrm ${NSD}><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></p:xfrm>`); el.insertBefore(x, OF.kid(el, 'graphic', A)); }
      return x;
    }
    if (el.localName === 'grpSp') return OF.kid(OF.kid(el, 'grpSpPr', P), 'xfrm', A);
    let spPr = OF.kid(el, 'spPr', P);
    if (!spPr) { spPr = el.ownerDocument.createElementNS(P, 'p:spPr'); el.insertBefore(spPr, OF.kid(el, 'style', P) || OF.kid(el, 'txBody', P)); }
    let xfrm = OF.kid(spPr, 'xfrm', A);
    if (!xfrm) { xfrm = frag(el.ownerDocument, `<a:xfrm xmlns:a="${A}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>`); spPr.insertBefore(xfrm, spPr.firstChild); }
    return xfrm;
  };
  PD.writePic = function (d) {
    const x = this.picXfrm(d);
    // dentro de un grupo las medidas están en las coordenadas de los hijos (chOff/chExt)
    const g = d._grp || { cx: 0, cy: 0, sx: 1, sy: 1 };
    OF.kid(x, 'off', A).setAttribute('x', Math.round(g.cx + (parseFloat(d.style.left) * EMU) / g.sx));
    OF.kid(x, 'off', A).setAttribute('y', Math.round(g.cy + (parseFloat(d.style.top) * EMU) / g.sy));
    OF.kid(x, 'ext', A).setAttribute('cx', Math.round((parseFloat(d.style.width) * EMU) / g.sx));
    OF.kid(x, 'ext', A).setAttribute('cy', Math.round((parseFloat(d.style.height) * EMU) / g.sy));
    const rot = ((Math.round(+d.dataset.rot || 0) % 360) + 360) % 360;
    if (rot) x.setAttribute('rot', rot * 60000); else x.removeAttribute('rot');
    this.markDirty();
  };
  PD.setRot = function (d, deg) {
    const rest = (d.style.transform || '').replace(/rotate\([^)]*\)\s*/g, '').trim();
    d.style.transform = ((deg ? `rotate(${deg}deg) ` : '') + rest).trim();
    if (deg) d.dataset.rot = deg; else delete d.dataset.rot;
  };
  PD.rotateDrag = function (e) {
    const d = this.selPic;
    if (!d) return;
    e.preventDefault();
    const r = d.getBoundingClientRect();
    const cx = r.left + r.width / 2, cy = r.top + r.height / 2;
    let moved = false;
    const move = (ev) => {
      if (!moved) { moved = true; OF.checkpoint && OF.checkpoint(); }
      let a = (Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180) / Math.PI + 90;
      a = ((a % 360) + 360) % 360;
      a = ev.shiftKey ? Math.round(a / 15) * 15 : Math.round(a);
      for (const k of [0, 90, 180, 270, 360]) if (Math.abs(a - k) < 4) a = k % 360;
      this.setRot(d, a);
      this.placeFrame();
    };
    const up = () => {
      window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
      if (!moved) return;
      this.writePic(d);
      const w = d.closest('.pp-wrap'); if (w) this.updateThumb(+w.dataset.i);
    };
    window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
  };
  PD.redrawKeepSel = async function (d) {
    const reg = this.picReg[+d.dataset.pe];
    const el = reg && reg.el;
    this.selectPic(null);
    this.commitAll();
    this.flushSlides();
    await this.load();
    await this.render();
    this.markDirty();
    void el;
  };
  PD.picAction = function (a) {
    const d = this.selPic;
    if (!d) return;
    OF.checkpoint && OF.checkpoint();
    const w0 = d.closest('.pp-wrap');
    if (a === 'del') {
      const { el, sync } = this.picReg[+d.dataset.pe];
      // los párrafos del objeto borrado dejan de existir para la sincronización de texto
      sync.paras.forEach((p, pid) => { if (el.contains(p)) sync.mapped.delete(pid); });
      el.parentNode.removeChild(el);
      d.remove();
      this.selectPic(null);
      this.markDirty();
      if (w0) this.updateThumb(+w0.dataset.i);
      return;
    }
    const w = parseFloat(d.style.width), h = parseFloat(d.style.height), l = parseFloat(d.style.left), t = parseFloat(d.style.top);
    const nw = w * a, nh = h * a;
    Object.assign(d.style, { width: nw + 'px', height: nh + 'px', left: l - (nw - w) / 2 + 'px', top: t - (nh - h) / 2 + 'px' });
    const svg = d.querySelector(':scope > svg'); if (svg && svg.getAttribute('viewBox')) { svg.setAttribute('width', nw); svg.setAttribute('height', nh); }
    this.writePic(d);
    this.placeFrame();
    if (w0) this.updateThumb(+w0.dataset.i);
  };

  /* ---------------------- estructura con varios párrafos ---------------------- */
  PD.selectionTargets = function () {
    const range = OF.currentRange();
    if (!range || !this.list.contains(range.startContainer)) return [];
    return OF.parasIn(range).map((dp) => {
      const wrap = dp.closest('.pp-wrap');
      const si = +wrap.dataset.i, pid = +dp.dataset.p;
      const same = Array.from(wrap.querySelectorAll(`[data-slide-edit] [data-p="${pid}"]`));
      return { si, pid, gi: Math.max(0, same.indexOf(dp)) };
    });
  };
  PD.structural = async function (fn) {
    OF.checkpoint && OF.checkpoint();
    const infos = this.selectionTargets();
    const si = infos.length ? infos[0].si : this.slideIndex();
    this.commitAll();
    const targets = [];
    for (const i of infos) { const res = this.syncs[i.si].result && this.syncs[i.si].result.get(i.pid); const t = res ? res[Math.min(i.gi, res.length - 1)] : null; if (t && !targets.includes(t)) targets.push(t); }
    this.caretTarget = targets[0] || null;
    const ok = await fn(si, targets[0] || null, targets);
    // recordar dónde estaba el cursor como ruta dentro del XML (load() vuelve a leer las diapositivas)
    let caretPath = null;
    if (this.caretTarget) {
      const sl = this.slides.find((x) => x.doc.contains(this.caretTarget));
      if (sl) { const idx = []; for (let n = this.caretTarget; n && n !== sl.doc.documentElement; n = n.parentNode) idx.unshift(Array.prototype.indexOf.call(n.parentNode.children, n)); caretPath = { path: sl.path, idx }; }
    }
    this.flushSlides();
    this.zip.file(this.presPath, OF.serializeXml(this.pres));
    await this.load();
    this.caretTarget = null;
    if (caretPath) {
      const sl = this.slides.find((x) => x.path === caretPath.path);
      let n = sl && sl.doc.documentElement;
      for (const i of caretPath.idx) n = n && n.children[i];
      this.caretTarget = n || null;
    }
    if (ok === false) { await this.render(); return; }
    this.dirty = false;
    this.markDirty();
    await this.render();
  };
  PD.flushSlides = function () { for (const s of this.slides) if (this.zip.file(s.path)) this.zip.file(s.path, OF.serializeXml(s.doc)); };
  PD.slideIndex = function () { return this.curSlide != null && this.curSlide < this.slides.length ? this.curSlide : 0; };

  /* ------------------------------ formato ------------------------------ */
  PD.setFont = function (name) { this.format({ font: name }, { fontFamily: `"${name}", Calibri, Carlito, Arial, sans-serif` }); };
  const pPrOf = (p) => { let e = OF.kid(p, 'pPr', A); if (!e) { e = p.ownerDocument.createElementNS(A, 'a:pPr'); p.insertBefore(e, p.firstChild); } return e; };
  PD.toggleList = async function (kind) {
    // la viñeta puede venir del diseño (heredada), así que se mira lo que se ve en pantalla
    const range = OF.currentRange();
    const dps = range && this.list.contains(range.startContainer) ? OF.parasIn(range) : [];
    const off = dps.length > 0 && dps.every((dp) => { const b = dp.querySelector('.pp-bu'); return b && (kind === 'bullet' ? !/\d/.test(b.textContent) : /\d/.test(b.textContent)); });
    await this.structural((si, t, targets) => {
      if (!targets.length) { OF.toast('Toca el texto donde quieres la lista.'); return false; }
      for (const p of targets) {
        const pPr = pPrOf(p);
        OF.removeKids(pPr, ['buNone', 'buChar', 'buAutoNum', 'buBlip', 'buFont', 'buFontTx'], A);
        if (off) { OF.ordered(pPr, A, 'a', 'buNone', PPR_ORDER); pPr.setAttribute('marL', '0'); pPr.setAttribute('indent', '0'); }
        else {
          if (kind === 'bullet') { OF.ordered(pPr, A, 'a', 'buFont', PPR_ORDER).setAttribute('typeface', 'Arial'); OF.ordered(pPr, A, 'a', 'buChar', PPR_ORDER).setAttribute('char', '•'); }
          else OF.ordered(pPr, A, 'a', 'buAutoNum', PPR_ORDER).setAttribute('type', 'arabicPeriod');
          const lvl = +(pPr.getAttribute('lvl') || 0);
          pPr.setAttribute('marL', 342900 + lvl * 457200); pPr.setAttribute('indent', '-342900');
        }
      }
    });
  };
  PD.indent = async function (dir) {
    await this.structural((si, t, targets) => {
      if (!targets.length) return false;
      for (const p of targets) {
        const pPr = pPrOf(p);
        const lvl = Math.max(0, Math.min(8, +(pPr.getAttribute('lvl') || 0) + dir));
        if (lvl) pPr.setAttribute('lvl', lvl); else pPr.removeAttribute('lvl');
        if (pPr.hasAttribute('marL')) pPr.setAttribute('marL', Math.max(0, +pPr.getAttribute('marL') + dir * 457200));
      }
    });
  };
  PD.lineSpacing = async function (mult) {
    await this.structural((si, t, targets) => {
      if (!targets.length) return false;
      for (const p of targets) {
        const pPr = pPrOf(p);
        OF.removeKids(pPr, ['lnSpc'], A);
        OF.ordered(pPr, A, 'a', 'lnSpc', PPR_ORDER).appendChild(frag(p.ownerDocument, `<a:spcPct xmlns:a="${A}" val="${Math.round(mult * 100000)}"/>`));
      }
    });
  };
  PD.formatState = function () {
    const r = OF.currentRange();
    if (!r || !this.list || !this.list.contains(r.startContainer)) return { slide: true };
    const el = r.startContainer.nodeType === 3 ? r.startContainer.parentElement : r.startContainer;
    const dp = el.closest('[data-p]');
    const cs = getComputedStyle(el);
    let u = false;
    for (let a = el; a && a !== dp; a = a.parentElement) if (/underline/.test(getComputedStyle(a).textDecorationLine)) u = true;
    const st = { b: +cs.fontWeight >= 600, i: cs.fontStyle === 'italic', u, font: cs.fontFamily.split(',')[0].replace(/["']/g, '').trim(), size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2 };
    if (dp) { const ta = getComputedStyle(dp).textAlign; st.align = /center/.test(ta) ? 'center' : /right|end/.test(ta) ? 'right' : /justify/.test(ta) ? 'justify' : 'left'; st.list = !!dp.querySelector('.pp-bu'); }
    return st;
  };
  PD.statusText = function () { return `Diapositiva ${(this.curSlide || 0) + 1} de ${this.slides.length}`; };

  /* ------------------------------ insertar ------------------------------ */
  PD.nextId = function (slide) { return Math.max(1, ...OF.all(slide.doc, 'cNvPr').map((e) => +e.getAttribute('id') || 0)) + 1; };
  PD.addToTree = function (slide, xml) {
    const tree = OF.all(slide.doc, 'spTree', P)[0];
    const el = frag(slide.doc, xml);
    tree.insertBefore(el, OF.kid(tree, 'extLst', P));
    return el;
  };
  PD.insertTextBox = async function () {
    await this.structural((si) => {
      const s = this.slides[si], id = this.nextId(s);
      const W = this.W * EMU, H = this.H * EMU;
      const cx = Math.round(W * 0.5), cy = Math.round(H * 0.12);
      const el = this.addToTree(s, `<p:sp ${NSD}><p:nvSpPr><p:cNvPr id="${id}" name="CuadroTexto ${id}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round((W - cx) / 2)}" y="${Math.round((H - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="es-CL" sz="2000" dirty="0"/></a:p></p:txBody></p:sp>`);
      this.caretTarget = OF.all(el, 'p', A)[0];
      this.curSlide = si;
    });
  };
  PD.insertShape = async function (prst) {
    await this.structural((si) => {
      const s = this.slides[si], id = this.nextId(s);
      const W = this.W * EMU, H = this.H * EMU;
      const line = prst === 'line';
      const cx = Math.round(W * (line ? 0.3 : 0.22)), cy = line ? 0 : Math.round(prst === 'rightArrow' ? H * 0.16 : W * 0.22 * (prst === 'rect' || prst === 'roundRect' ? 0.6 : 1));
      const off = `<a:off x="${Math.round((W - cx) / 2)}" y="${Math.round((H - cy) / 2)}"/>`;
      if (line) {
        this.addToTree(s, `<p:cxnSp ${NSD}><p:nvCxnSpPr><p:cNvPr id="${id}" name="Línea ${id}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm>${off}<a:ext cx="${cx}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="28575"><a:solidFill><a:schemeClr val="accent1"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`);
      } else {
        const el = this.addToTree(s, `<p:sp ${NSD}><p:nvSpPr><p:cNvPr id="${id}" name="Forma ${id}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm>${off}<a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom></p:spPr><p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="es-CL" dirty="0"/></a:p></p:txBody></p:sp>`);
        void el;
      }
      this.curSlide = si;
    });
  };
  PD.insertTable = async function (rows, cols) {
    await this.structural((si) => {
      const s = this.slides[si], id = this.nextId(s);
      const W = this.W * EMU, H = this.H * EMU;
      const cx = Math.round(W * 0.8), rh = 370840, cw = Math.floor(cx / cols);
      let grid = '', trs = '';
      for (let c = 0; c < cols; c++) grid += `<a:gridCol w="${cw}"/>`;
      for (let r = 0; r < rows; r++) { trs += `<a:tr h="${rh}">`; for (let c = 0; c < cols; c++) trs += '<a:tc><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-CL" dirty="0"/></a:p></a:txBody><a:tcPr/></a:tc>'; trs += '</a:tr>'; }
      const el = this.addToTree(s, `<p:graphicFrame ${NSD}><p:nvGraphicFramePr><p:cNvPr id="${id}" name="Tabla ${id}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr><p:xfrm><a:off x="${Math.round((W - cx) / 2)}" y="${Math.round(H * 0.3)}"/><a:ext cx="${cw * cols}" cy="${rh * rows}"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1" bandRow="1"><a:tableStyleId>{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}</a:tableStyleId></a:tblPr><a:tblGrid>${grid}</a:tblGrid>${trs}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`);
      this.caretTarget = OF.all(el, 'p', A)[0];
      this.curSlide = si;
    });
  };
  PD.tableCol = async function (op) {
    await this.structural((si, t) => {
      let tc = t; while (tc && tc.localName !== 'tc') tc = tc.parentNode;
      if (!tc) { OF.toast('Toca una celda de la tabla primero.'); return false; }
      const tr = tc.parentNode, tbl = tr.parentNode;
      const idx = OF.kids(tr, 'tc', A).indexOf(tc);
      const grid = OF.kid(tbl, 'tblGrid', A), gc = OF.kids(grid, 'gridCol', A);
      if (op === 'del' && gc.length <= 1) { OF.toast('La tabla debe tener al menos una columna.'); return false; }
      for (const row of OF.kids(tbl, 'tr', A)) {
        const c = OF.kids(row, 'tc', A)[idx];
        if (!c) continue;
        if (op === 'del') { row.removeChild(c); continue; }
        const nc = c.cloneNode(true);
        const tb = OF.kid(nc, 'txBody', A);
        if (tb) { const ps = OF.kids(tb, 'p', A); ps.slice(1).forEach((p) => tb.removeChild(p)); if (ps[0]) for (const r of Array.from(ps[0].children)) if (['r', 'br', 'fld'].includes(r.localName)) ps[0].removeChild(r); }
        row.insertBefore(nc, op === 'left' ? c : c.nextSibling);
      }
      const total = gc.reduce((a, g) => a + +g.getAttribute('w'), 0);
      if (op === 'del') grid.removeChild(gc[idx]); else grid.insertBefore(gc[idx].cloneNode(true), op === 'left' ? gc[idx] : gc[idx].nextSibling);
      const cols = OF.kids(grid, 'gridCol', A);
      cols.forEach((g) => g.setAttribute('w', Math.floor(total / cols.length)));
    });
  };

  /* ---------------------------- diapositivas ---------------------------- */
  PD.layoutsFor = async function (slide) {
    const master = slide.master;
    if (!master) return [];
    const out = [];
    for (const rel of Object.values(master.rels)) {
      if (!/slideLayout$/.test(rel.type) || !this.zip.file(rel.target)) continue;
      const part = await this.part(rel.target);
      const root = part.doc.documentElement;
      const name = (OF.all(part.doc, 'cSld', P)[0] || root).getAttribute('name') || root.getAttribute('type') || 'Diseño';
      out.push({ path: rel.target, part, name, label: LAYOUT_ES[name.toLowerCase()] || name, type: root.getAttribute('type') });
    }
    const order = ['title', 'obj', 'secHead', 'twoObj', 'twoTxTwoObj', 'titleOnly', 'blank'];
    return out.sort((a, b) => ((order.indexOf(a.type) + 1) || 99) - ((order.indexOf(b.type) + 1) || 99));
  };
  PD.newSlideMenu = async function (btn) {
    const s = this.slides[this.slideIndex()];
    const layouts = await this.layoutsFor(s);
    if (!layouts.length) return this.slideOp('new');
    OF.popup(btn, (p, close) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Diseño de la nueva diapositiva'));
      for (const l of layouts) p.append(OF.el('button', { class: 'pop-item' + (s.layout && s.layout.path === l.path ? ' cur' : ''), onclick: () => { close(); this.newSlide(l.part); } }, l.label));
    });
  };
  PD.newSlide = async function (layout) {
    await this.structural(async (si) => {
      const shapes = [];
      let id = 2;
      const blank = layout && layout.doc.documentElement.getAttribute('type') === 'blank';
      const keep = (t) => !['dt', 'ftr', 'sldNum', 'hdr', 'sldImg'].includes(t);
      if (layout && !blank) for (const sp of OF.all(OF.all(layout.doc, 'spTree', P)[0], 'sp', P)) {
        const ph = OF.all(sp.firstElementChild, 'ph', P)[0];
        if (!ph || !keep(ph.getAttribute('type') || 'obj')) continue;
        const attrs = Array.from(ph.attributes).map((a) => `${a.name}="${a.value}"`).join(' ');
        shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Marcador ${id++}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-CL"/></a:p></p:txBody></p:sp>`);
      }
      const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld ${NSD}><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
      const path = OF.uniquePart(this.zip, 'ppt/slides', 'slide', 'xml');
      const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OF.NS.rel}">${layout ? `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${OF.relTarget(path, layout.path)}"/>` : ''}</Relationships>`;
      const { rid } = await this.newSlidePart(xml, rels);
      this.insertSldId(rid, si);
      this.curSlide = si + 1;
      this.scrollToCur = true;
    });
  };
  const baseSlideOp = PD.slideOp;
  PD.slideOp = async function (op) {
    if (op === 'new') { const s = this.slides[this.slideIndex()]; return this.newSlide(s && s.layout); }
    this.scrollToCur = true;
    return baseSlideOp.call(this, op);
  };
  PD.setBackground = async function (color, all) {
    await this.structural((si) => {
      const list = all ? this.slides : [this.slides[si]];
      for (const s of list) {
        const cSld = OF.all(s.doc, 'cSld', P)[0];
        OF.removeKids(cSld, ['bg'], P);
        if (color) cSld.insertBefore(frag(s.doc, `<p:bg ${NSD}><p:bgPr><a:solidFill><a:srgbClr val="${color.replace('#', '').toUpperCase()}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>`), cSld.firstChild);
      }
      this.curSlide = si;
    });
  };
  PD.toggleHidden = async function () {
    await this.structural((si) => {
      const root = this.slides[si].doc.documentElement;
      if (root.getAttribute('show') === '0') root.removeAttribute('show'); else root.setAttribute('show', '0');
      this.curSlide = si;
    });
  };

  /* ---------------------- buscar, deshacer, imprimir ---------------------- */
  PD.searchRoots = function () { return Array.from(this.list.querySelectorAll('[data-slide-edit]')); };
  PD.afterTextChange = function () { this.slides.forEach((x, i) => this.updateThumb(i)); };
  Object.assign(PD, OF.textSearchMixin);
  PD.snapshot = function () {
    this._flushCache = this._flushCache || {};
    for (const s of this.slides) { const x = OF.serializeXml(s.doc); if (this._flushCache[s.path] !== x || !this.zip.file(s.path)) { this.zip.file(s.path, x); this._flushCache[s.path] = x; } }
    return { files: Object.assign({}, this.zip.files), html: this.searchRoots().map((t) => OF.cloneKids(t)), cur: this.curSlide, scroll: this.host.scrollTop };
  };
  PD.restore = async function (s) {
    const same = Object.keys(s.files).length === Object.keys(this.zip.files).length && Object.keys(s.files).every((k) => this.zip.files[k] === s.files[k]);
    if (!same) {
      const z = new JSZip();
      for (const k in s.files) z.files[k] = s.files[k];
      this.zip = z;
      this.parts = {};
      this._flushCache = {};
      this.pres = OF.parseXml(await z.file(this.presPath).async('string'));
      await this.load();
      this.curSlide = s.cur;
      await this.render();
    }
    const roots = this.searchRoots();
    if (roots.length === s.html.length) roots.forEach((t, i) => OF.setKids(t, s.html[i]));
    this.host.scrollTop = s.scroll;
    this.slides.forEach((x, i) => this.updateThumb(i, true));
  };
  PD.printSetup = function () {
    this.selectPic(null);
    this.list.querySelectorAll('.pp-frame').forEach((f) => { f.style.width = this.W + 'px'; f.style.height = this.H + 'px'; f.firstChild.style.transform = 'none'; });
    return `@page { size: ${this.W}px ${this.H}px; margin: 0 }`;
  };
})();
