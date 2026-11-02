/* Editor de PowerPoint: renderizador propio de DrawingML (respeta patrón, diseño y tema)
   con edición de texto directa sobre el XML de cada diapositiva. */
(function () {
  const OF = window.OF;
  const A = OF.NS.a, P = OF.NS.p, R = OF.NS.r;
  const EMU = 9525; // EMU por px a 96 ppp
  const px = (v) => (+v || 0) / EMU;

  const PRST = {
    triangle: [[0.5, 0], [1, 1], [0, 1]],
    rtTriangle: [[0, 0], [1, 1], [0, 1]],
    diamond: [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]],
    parallelogram: [[0.25, 0], [1, 0], [0.75, 1], [0, 1]],
    trapezoid: [[0.25, 0], [0.75, 0], [1, 1], [0, 1]],
    pentagon: [[0.5, 0], [1, 0.38], [0.81, 1], [0.19, 1], [0, 0.38]],
    hexagon: [[0.25, 0], [0.75, 0], [1, 0.5], [0.75, 1], [0.25, 1], [0, 0.5]],
    octagon: [[0.29, 0], [0.71, 0], [1, 0.29], [1, 0.71], [0.71, 1], [0.29, 1], [0, 0.71], [0, 0.29]],
    homePlate: [[0, 0], [0.8, 0], [1, 0.5], [0.8, 1], [0, 1]],
    chevron: [[0, 0], [0.8, 0], [1, 0.5], [0.8, 1], [0, 1], [0.2, 0.5]],
    rightArrow: [[0, 0.25], [0.7, 0.25], [0.7, 0], [1, 0.5], [0.7, 1], [0.7, 0.75], [0, 0.75]],
    leftArrow: [[1, 0.25], [0.3, 0.25], [0.3, 0], [0, 0.5], [0.3, 1], [0.3, 0.75], [1, 0.75]],
    upArrow: [[0.25, 1], [0.25, 0.3], [0, 0.3], [0.5, 0], [1, 0.3], [0.75, 0.3], [0.75, 1]],
    downArrow: [[0.25, 0], [0.25, 0.7], [0, 0.7], [0.5, 1], [1, 0.7], [0.75, 0.7], [0.75, 0]],
    plus: [[0.33, 0], [0.67, 0], [0.67, 0.33], [1, 0.33], [1, 0.67], [0.67, 0.67], [0.67, 1], [0.33, 1], [0.33, 0.67], [0, 0.67], [0, 0.33], [0.33, 0.33]],
    star5: [[0.5, 0], [0.62, 0.38], [1, 0.38], [0.69, 0.62], [0.81, 1], [0.5, 0.76], [0.19, 1], [0.31, 0.62], [0, 0.38], [0.38, 0.38]],
    flowChartDecision: [[0.5, 0], [1, 0.5], [0.5, 1], [0, 0.5]],
  };

  function rgb2hsl(r, g, b) {
    r /= 255; g /= 255; b /= 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0; const l = (max + min) / 2;
    if (max !== min) { const d = max - min; s = l > 0.5 ? d / (2 - max - min) : d / (max + min); h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4; h /= 6; }
    return [h, s, l];
  }
  function hsl2rgb(h, s, l) {
    const f = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    if (s === 0) return [l * 255, l * 255, l * 255];
    const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    return [f(p, q, h + 1 / 3) * 255, f(p, q, h) * 255, f(p, q, h - 1 / 3) * 255];
  }

  class PptDoc {
    constructor() { this.kind = 'ppt'; this.dirty = false; this.urls = []; }

    async open(buffer, host) {
      this.original = buffer;
      this.zip = await JSZip.loadAsync(buffer);
      await this.load();
      this.host = host;
      host.addEventListener('focusin', (e) => { const w = e.target.closest && e.target.closest('.pp-wrap'); if (w) this.curSlide = +w.dataset.i; });
      host.addEventListener('pointerdown', (e) => { const w = e.target.closest && e.target.closest('.pp-wrap'); if (w) this.curSlide = +w.dataset.i; });
      await this.render();
    }

    async load() {
      const zip = this.zip;
      const rels = await OF.readRels(zip, '');
      this.presPath = (Object.values(rels).find((r) => /officeDocument$/.test(r.type)) || {}).target || 'ppt/presentation.xml';
      this.pres = OF.parseXml(await zip.file(this.presPath).async('string'));
      const presRels = await OF.readRels(zip, this.presPath);
      const sz = OF.all(this.pres, 'sldSz')[0];
      this.W = px(sz ? sz.getAttribute('cx') : 9144000);
      this.H = px(sz ? sz.getAttribute('cy') : 6858000);
      this.defaultTextStyle = OF.all(this.pres, 'defaultTextStyle')[0] || null;
      this.parts = {};
      this.slides = [];
      for (const id of OF.all(this.pres, 'sldId')) {
        const rid = id.getAttributeNS(R, 'id');
        const rel = presRels[rid];
        if (!rel || !zip.file(rel.target)) continue;
        const slide = await this.part(rel.target);
        const layoutRel = Object.values(slide.rels).find((r) => /slideLayout$/.test(r.type));
        const layout = layoutRel ? await this.part(layoutRel.target) : null;
        const masterRel = layout && Object.values(layout.rels).find((r) => /slideMaster$/.test(r.type));
        const master = masterRel ? await this.part(masterRel.target) : null;
        if (master && !master.theme) {
          const themeRel = Object.values(master.rels).find((r) => /theme$/.test(r.type));
          master.theme = themeRel ? this.readTheme(OF.parseXml(await zip.file(themeRel.target).async('string'))) : this.readTheme(null);
        }
        this.slides.push({ ...slide, layout, master, rid, hidden: slide.doc.documentElement.getAttribute('show') === '0' });
      }
    }

    async part(path) {
      if (this.parts[path]) return this.parts[path];
      const doc = OF.parseXml(await this.zip.file(path).async('string'));
      const rels = await OF.readRels(this.zip, path);
      return (this.parts[path] = { path, doc, rels });
    }

    readTheme(doc) {
      const t = { colors: { dk1: '000000', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72' }, major: 'Calibri Light', minor: 'Calibri' };
      if (!doc) return t;
      const cs = OF.all(doc, 'clrScheme')[0];
      if (cs) for (const c of Array.from(cs.children)) { const v = c.firstElementChild; if (v) t.colors[c.localName] = (v.localName === 'sysClr' ? v.getAttribute('lastClr') || (v.getAttribute('val') === 'window' ? 'FFFFFF' : '000000') : v.getAttribute('val')) || t.colors[c.localName]; }
      const fs = OF.all(doc, 'fontScheme')[0];
      if (fs) {
        const mj = OF.path(fs, 'majorFont', 'latin'), mn = OF.path(fs, 'minorFont', 'latin');
        if (mj) t.major = mj.getAttribute('typeface');
        if (mn) t.minor = mn.getAttribute('typeface');
      }
      return t;
    }

    /* ----------------------------- colores ----------------------------- */
    color(el, ctx) {
      // el: elemento contenedor (solidFill, etc.) o el propio *Clr
      if (!el) return null;
      const isClr = (x) => /^(srgbClr|schemeClr|sysClr|prstClr|scrgbClr|hslClr)$/.test(x.localName);
      let c = isClr(el) ? el : Array.from(el.children).find(isClr);
      if (!c) return null;
      let hex;
      const th = ctx.theme.colors;
      switch (c.localName) {
        case 'srgbClr': hex = c.getAttribute('val'); break;
        case 'sysClr': hex = c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000'); break;
        case 'schemeClr': {
          let v = c.getAttribute('val');
          if (v === 'phClr') { hex = ctx.phClr || '000000'; break; }
          v = ctx.clrMap[v] || v;
          hex = th[v] || '000000';
          break;
        }
        case 'prstClr': hex = { black: '000000', white: 'FFFFFF', red: 'FF0000', green: '008000', blue: '0000FF', yellow: 'FFFF00', gray: '808080' }[c.getAttribute('val')] || '000000'; break;
        case 'scrgbClr': hex = ['r', 'g', 'b'].map((k) => Math.round((+c.getAttribute(k) / 100000) * 255).toString(16).padStart(2, '0')).join(''); break;
        default: return null;
      }
      let r = parseInt(hex.slice(0, 2), 16), g = parseInt(hex.slice(2, 4), 16), b = parseInt(hex.slice(4, 6), 16), a = 1;
      for (const m of Array.from(c.children)) {
        const val = +m.getAttribute('val') / 100000;
        if (m.localName === 'alpha') a = val;
        else if (m.localName === 'lumMod' || m.localName === 'lumOff') {
          const [h, s, l] = rgb2hsl(r, g, b);
          const nl = m.localName === 'lumMod' ? l * val : l + val;
          [r, g, b] = hsl2rgb(h, s, Math.max(0, Math.min(1, nl)));
        } else if (m.localName === 'tint') { r = r + (255 - r) * (1 - val); g = g + (255 - g) * (1 - val); b = b + (255 - b) * (1 - val); }
        else if (m.localName === 'shade') { r *= val; g *= val; b *= val; }
        else if (m.localName === 'satMod') { const [h, s, l] = rgb2hsl(r, g, b); [r, g, b] = hsl2rgb(h, Math.min(1, s * val), l); }
      }
      r = Math.round(r); g = Math.round(g); b = Math.round(b);
      return a < 1 ? `rgba(${r},${g},${b},${a.toFixed(3)})` : `rgb(${r},${g},${b})`;
    }

    fill(spPr, ctx, styleEl) {
      if (spPr) {
        for (const c of Array.from(spPr.children)) {
          if (c.localName === 'noFill') return { none: true };
          if (c.localName === 'solidFill') return { css: this.color(c, ctx) };
          if (c.localName === 'gradFill') {
            const stops = OF.kids(OF.kid(c, 'gsLst', A), 'gs', A).map((g) => `${this.color(g, ctx)} ${+g.getAttribute('pos') / 1000}%`);
            const lin = OF.kid(c, 'lin', A);
            const ang = lin ? +lin.getAttribute('ang') / 60000 + 90 : 180;
            if (stops.length) return { css: OF.kid(c, 'path', A) ? `radial-gradient(${stops.join(',')})` : `linear-gradient(${ang}deg,${stops.join(',')})`, grad: true };
          }
          if (c.localName === 'blipFill') { const u = this.blipUrl(c, ctx); if (u) return { img: u }; }
          if (c.localName === 'pattFill') {
            const fg = this.color(OF.kid(c, 'fgClr', A), ctx) || '#000', bg = this.color(OF.kid(c, 'bgClr', A), ctx) || '#fff';
            const pr = c.getAttribute('prst') || '';
            const pat = /Grid|cross/i.test(pr)
              ? `linear-gradient(${fg} 1px, transparent 1px) 0 0/8px 8px, linear-gradient(90deg, ${fg} 1px, transparent 1px) 0 0/8px 8px, ${bg}`
              : /Horz/i.test(pr) ? `repeating-linear-gradient(0deg, ${fg} 0 1px, ${bg} 1px 5px)`
              : /Vert/i.test(pr) ? `repeating-linear-gradient(90deg, ${fg} 0 1px, ${bg} 1px 5px)`
              : /Diag|wdDn|wdUp|ltDn|ltUp|dkDn|dkUp/i.test(pr) ? `repeating-linear-gradient(45deg, ${fg} 0 1px, ${bg} 1px 6px)`
              : `radial-gradient(${fg} 0.8px, transparent 1.2px) 0 0/5px 5px, ${bg}`;
            return { css: pat, grad: true };
          }
          if (c.localName === 'grpFill' && ctx.grpFill) return ctx.grpFill;
        }
      }
      const ref = styleEl && OF.path(styleEl, 'fillRef');
      if (ref && ref.getAttribute('idx') !== '0') { const col = this.color(ref, ctx); if (col) return { css: col }; }
      return null;
    }
    line(spPr, ctx, styleEl) {
      const ln = spPr && OF.kid(spPr, 'ln', A);
      const w = ln && ln.getAttribute('w') ? Math.max(0.75, px(ln.getAttribute('w'))) : 1;
      if (ln) {
        if (OF.kid(ln, 'noFill', A)) return null;
        const sf = OF.kid(ln, 'solidFill', A);
        if (sf) return { color: this.color(sf, ctx), w, dash: OF.kid(ln, 'prstDash', A) ? OF.kid(ln, 'prstDash', A).getAttribute('val') : 'solid' };
      }
      const ref = styleEl && OF.path(styleEl, 'lnRef');
      if (ref && ref.getAttribute('idx') !== '0') { const col = this.color(ref, ctx); if (col) return { color: col, w }; }
      return null;
    }
    blipUrl(blipFill, ctx) {
      const blip = OF.kid(blipFill, 'blip', A);
      if (!blip) return null;
      const rid = blip.getAttributeNS(R, 'embed');
      const rel = rid && ctx.rels[rid];
      if (!rel || rel.external) return null;
      if (/\.(emf|wmf)$/i.test(rel.target)) return 'unsupported';
      const key = rel.target;
      if (!this.imgCache) this.imgCache = {};
      if (this.imgCache[key]) return this.imgCache[key];
      const f = this.zip.file(rel.target);
      if (!f) return null;
      const p = f.async('blob').then((b) => { const u = URL.createObjectURL(new Blob([b], { type: OF.mimeFor(rel.target) })); this.urls.push(u); return u; });
      this.imgCache[key] = p;
      return p;
    }

    /* ----------------------- herencia de marcadores ----------------------- */
    phOf(sp) {
      const nv = sp.firstElementChild;
      const ph = nv && OF.all(nv, 'ph')[0];
      if (!ph) return null;
      return { type: ph.getAttribute('type') || 'body', idx: ph.getAttribute('idx') };
    }
    findPh(partDoc, ph, byTypeOnly) {
      if (!partDoc) return null;
      const shapes = OF.all(OF.all(partDoc, 'spTree')[0], 'sp', P);
      const norm = (t) => (t === 'ctrTitle' ? 'title' : t);
      if (!byTypeOnly && ph.idx != null) {
        const m = shapes.find((s) => { const p = this.phOf(s); return p && p.idx === ph.idx; });
        if (m) return m;
      }
      const tt = norm(ph.type);
      return shapes.find((s) => { const p = this.phOf(s); return p && norm(p.type) === tt; }) || (['subTitle', 'obj'].includes(ph.type) ? shapes.find((s) => { const p = this.phOf(s); return p && p.type === 'body'; }) : null);
    }

    ctxFor(slide, level) {
      const master = slide.master;
      const theme = master ? master.theme : this.readTheme(null);
      const cm = master && OF.all(master.doc, 'clrMap')[0];
      const clrMap = {};
      if (cm) for (const at of Array.from(cm.attributes)) clrMap[at.name] = at.value;
      else Object.assign(clrMap, { bg1: 'lt1', tx1: 'dk1', bg2: 'lt2', tx2: 'dk2' });
      const ovr = level === 'slide' && OF.all(slide.doc, 'overrideClrMapping')[0];
      if (ovr) for (const at of Array.from(ovr.attributes)) clrMap[at.name] = at.value;
      const part = level === 'slide' ? slide : level === 'layout' ? slide.layout : master;
      return { theme, clrMap, rels: part.rels, slide, level };
    }

    /* ------------------------------ render ------------------------------ */
    async render() {
      const host = this.host;
      const scroll = host.scrollTop;
      host.innerHTML = '';
      this.urls.forEach(URL.revokeObjectURL); this.urls = []; this.imgCache = {};
      this.list = OF.el('div', { class: 'pp-list' });
      host.append(this.list);
      this.syncs = [];
      this.pending = [];
      this.picReg = [];
      OF.closeImageBar();
      for (let i = 0; i < this.slides.length; i++) {
        const slide = this.slides[i];
        slide.num = i + 1;
        const box = OF.el('div', { class: 'pp-slide', style: { width: this.W + 'px', height: this.H + 'px' } });
        const wrap = OF.el('div', { class: 'pp-wrap' + (slide.hidden ? ' hidden-slide' : '') + (i === this.curSlide ? ' cur' : ''), 'data-i': i }, OF.el('div', { class: 'pp-num' }, String(i + 1)), OF.el('div', { class: 'pp-frame' }, box));
        this.list.append(wrap);
        slide.box = box;
        this.renderSlide(slide, box, i);
      }
      await Promise.all(this.pending);
      this.fit();
      host.scrollTop = scroll;
    }

    renderSlide(slide, box, si) {
      const sync = (this.syncs[si] = new OF.TextSync({
        ns: A, prefix: 'a', P: 'p', R: 'r', T: 't', PPR: 'pPr', wrappers: new Set(),
        markProps: (p) => OF.kid(p, 'endParaRPr', A),
        applyFmt: pptRunFmt,
        applyPFmt: pptParaFmt,
      }));
      // fondo
      const bgFrom = [slide, slide.layout, slide.master];
      const levels = ['slide', 'layout', 'master'];
      for (let k = 0; k < 3; k++) {
        const part = bgFrom[k];
        const bg = part && OF.all(part.doc, 'bg', P)[0];
        if (!bg) continue;
        const ctx = this.ctxFor(slide, levels[k]);
        const bgPr = OF.kid(bg, 'bgPr', P), bgRef = OF.kid(bg, 'bgRef', P);
        let f = bgPr ? this.fill(bgPr, ctx) : bgRef ? { css: this.color(bgRef, ctx) } : null;
        if (f && !f.none) { this.applyFill(box, f); break; }
        if (f && f.none) break;
      }
      if (!box.style.background) box.style.background = '#fff';
      const showMaster = (p) => !p || p.doc.documentElement.getAttribute('showMasterSp') !== '0';
      if (slide.master && showMaster(slide.layout) && showMaster(slide)) this.renderTree(slide.master, slide, 'master', box, null);
      if (slide.layout && showMaster(slide)) this.renderTree(slide.layout, slide, 'layout', box, null);
      this.renderTree(slide, slide, 'slide', box, sync);
    }

    renderTree(part, slide, level, box, sync) {
      const tree = OF.all(part.doc, 'spTree', P)[0];
      if (!tree) return;
      const ctx = this.ctxFor(slide, level);
      const T = { x: (v) => px(v), y: (v) => px(v), sx: 1, sy: 1 };
      this.renderChildren(tree, ctx, box, sync, T, level !== 'slide');
    }

    renderChildren(parent, ctx, box, sync, T, skipPh) {
      for (const el of Array.from(parent.children)) {
        const n = el.localName;
        if (n === 'AlternateContent') { const ch = OF.kid(el, 'Choice') ; const fb = OF.kid(el, 'Fallback'); this.renderChildren(fb || ch, ctx, box, sync, T, skipPh); continue; }
        if (!['sp', 'pic', 'grpSp', 'graphicFrame', 'cxnSp'].includes(n)) continue;
        const ph = n === 'sp' || n === 'pic' ? this.phOf(el) : null;
        if (skipPh && ph) continue;
        try { this.renderShape(el, n, ph, ctx, box, sync, T); } catch (e) { console.warn('forma no dibujada', e); }
      }
    }

    xfrmOf(el, n, ph, ctx) {
      const prName = n === 'grpSp' ? 'grpSpPr' : n === 'graphicFrame' ? null : 'spPr';
      let xfrm = n === 'graphicFrame' ? OF.kid(el, 'xfrm', P) : OF.kid(OF.kid(el, prName), 'xfrm', A);
      if ((!xfrm || !OF.kid(xfrm, 'off', A)) && ph) {
        const lay = this.findPh(ctx.slide.layout && ctx.slide.layout.doc, ph);
        const lx = lay && OF.path(lay, 'spPr', 'xfrm');
        if (lx && OF.kid(lx, 'off', A)) xfrm = lx;
        else { const ms = this.findPh(ctx.slide.master && ctx.slide.master.doc, ph, true); const mx = ms && OF.path(ms, 'spPr', 'xfrm'); if (mx) xfrm = mx; }
      }
      return xfrm;
    }

    renderShape(el, n, ph, ctx, box, sync, T) {
      const xfrm = this.xfrmOf(el, n, ph, ctx);
      const off = xfrm && OF.kid(xfrm, 'off', A), ext = xfrm && OF.kid(xfrm, 'ext', A);
      if (!off || !ext) { if (n !== 'grpSp') return; }
      const x = T.x(+off.getAttribute('x')), y = T.y(+off.getAttribute('y'));
      const w = px(ext.getAttribute('cx')) * T.sx, h = px(ext.getAttribute('cy')) * T.sy;
      const rot = xfrm.getAttribute('rot') ? +xfrm.getAttribute('rot') / 60000 : 0;
      const flipH = xfrm.getAttribute('flipH') === '1', flipV = xfrm.getAttribute('flipV') === '1';
      const d = OF.el('div', { class: 'pp-sh', style: { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' } });
      const tf = [];
      if (rot) tf.push(`rotate(${rot}deg)`);
      if (flipH) tf.push('scaleX(-1)');
      if (flipV) tf.push('scaleY(-1)');
      if (tf.length) d.style.transform = tf.join(' ');

      if (n === 'grpSp') {
        const chOff = OF.kid(xfrm, 'chOff', A), chExt = OF.kid(xfrm, 'chExt', A);
        const sx = chExt && +chExt.getAttribute('cx') ? +ext.getAttribute('cx') / +chExt.getAttribute('cx') : 1;
        const sy = chExt && +chExt.getAttribute('cy') ? +ext.getAttribute('cy') / +chExt.getAttribute('cy') : 1;
        const ox = +off.getAttribute('x'), oy = +off.getAttribute('y');
        const cx = chOff ? +chOff.getAttribute('x') : 0, cy = chOff ? +chOff.getAttribute('y') : 0;
        const T2 = { x: (v) => T.x(ox + (v - cx) * sx), y: (v) => T.y(oy + (v - cy) * sy), sx: T.sx * sx, sy: T.sy * sy };
        const gctx = { ...ctx, grpFill: this.fill(OF.kid(el, 'grpSpPr', P), ctx) };
        if (tf.length) {
          // grupo rotado: los hijos se dibujan dentro de un contenedor rotado
          const inner = OF.el('div', { class: 'pp-grp', style: { left: 0, top: 0, width: this.W + 'px', height: this.H + 'px', transformOrigin: `${x + w / 2}px ${y + h / 2}px`, transform: tf.join(' ') } });
          box.append(inner);
          this.renderChildren(el, gctx, inner, sync, T2, false);
        } else this.renderChildren(el, gctx, box, sync, T2, false);
        return;
      }

      const styleEl = OF.kid(el, 'style', P);
      if (n === 'pic') {
        const bf = OF.kid(el, 'blipFill', P);
        const u = bf && this.blipUrl(bf, ctx);
        const src = OF.kid(bf, 'srcRect', A);
        const img = OF.el('img', { class: 'pp-img', alt: '', draggable: 'false' });
        if (src) {
          const l = +(src.getAttribute('l') || 0) / 1e5, t = +(src.getAttribute('t') || 0) / 1e5, r = +(src.getAttribute('r') || 0) / 1e5, b = +(src.getAttribute('b') || 0) / 1e5;
          const iw = w / (1 - l - r || 1), ih = h / (1 - t - b || 1);
          Object.assign(img.style, { position: 'absolute', width: iw + 'px', height: ih + 'px', left: -l * iw + 'px', top: -t * ih + 'px', maxWidth: 'none' });
          d.style.overflow = 'hidden';
        }
        if (sync && T.sx === 1 && T.sy === 1 && T.x(0) === 0 && !rot) { d.classList.add('pp-pic'); d.dataset.pe = this.picReg.push({ el, sync }) - 1; }
        if (u === 'unsupported') d.append(OF.el('div', { class: 'pp-unsup' }, 'Imagen EMF/WMF'));
        else if (u) { this.pending.push(Promise.resolve(u).then((url) => (img.src = url))); d.append(img); }
        const ln = this.line(OF.kid(el, 'spPr', P), ctx, styleEl);
        if (ln) d.style.outline = `${ln.w}px solid ${ln.color}`;
        const prst = OF.path(el, 'spPr', 'prstGeom');
        if (prst && prst.getAttribute('prst') === 'ellipse') { d.style.borderRadius = '50%'; d.style.overflow = 'hidden'; }
        box.append(d);
        return;
      }
      if (n === 'graphicFrame') {
        const gd = OF.all(el, 'graphicData', A)[0];
        const tbl = gd && OF.kid(gd, 'tbl', A);
        if (tbl) this.renderTable(tbl, ctx, d, sync);
        else {
          const kind = gd && /chart/.test(gd.getAttribute('uri')) ? 'Gráfico' : gd && /diagram/.test(gd.getAttribute('uri')) ? 'SmartArt' : 'Objeto';
          d.append(OF.el('div', { class: 'pp-unsup' }, `${kind} (se conserva, no editable aquí)`));
        }
        box.append(d);
        return;
      }

      // sp / cxnSp
      const spPr = OF.kid(el, 'spPr', P);
      const lay = ph && this.findPh(ctx.slide.layout && ctx.slide.layout.doc, ph);
      const mas = ph && this.findPh(ctx.slide.master && ctx.slide.master.doc, ph, true);
      let fill = this.fill(spPr, ctx, styleEl);
      if (!fill && ph) fill = (lay && this.fill(OF.kid(lay, 'spPr', P), this.ctxFor(ctx.slide, 'layout'))) || null;
      let ln = this.line(spPr, ctx, styleEl);
      const geomEl = spPr && (OF.kid(spPr, 'prstGeom', A) || OF.kid(spPr, 'custGeom', A));
      const prst = geomEl && geomEl.localName === 'prstGeom' ? geomEl.getAttribute('prst') : geomEl ? 'cust' : 'rect';
      const isLine = n === 'cxnSp' || /^(line|straightConnector1|bentConnector\d|curvedConnector\d)$/.test(prst);
      if (isLine) {
        ln = ln || { color: '#000', w: 1 };
        const svg = svgEl('svg', { width: Math.max(w, 1), height: Math.max(h, 1), style: 'position:absolute;left:0;top:0;overflow:visible' });
        const dash = { dash: '6 4', sysDash: '3 2', dot: '1 2', sysDot: '1 2', lgDash: '10 4', dashDot: '6 3 1 3' }[ln.dash];
        svg.append(svgEl('line', { x1: 0, y1: 0, x2: w, y2: h, stroke: ln.color, 'stroke-width': ln.w, 'stroke-dasharray': dash || '' }));
        d.append(svg);
        box.append(d);
        return;
      }
      if (prst === 'rect' || prst === 'roundRect' || prst === 'ellipse' || prst === 'snip1Rect' || prst === 'round2SameRect') {
        if (fill) this.applyFill(d, fill);
        if (ln) d.style.border = `${ln.w}px ${ln.dash && ln.dash !== 'solid' ? 'dashed' : 'solid'} ${ln.color}`;
        if (prst === 'ellipse') d.style.borderRadius = '50%';
        if (prst === 'roundRect') { const adj = OF.all(geomEl, 'gd', A)[0]; const a = adj ? +adj.getAttribute('fmla').split(' ')[1] / 100000 : 0.16667; d.style.borderRadius = Math.min(w, h) * a + 'px'; }
        if (ln) d.style.boxSizing = 'border-box';
      } else if (fill || ln) {
        const svg = svgEl('svg', { width: w, height: h, viewBox: `0 0 ${w} ${h}`, style: 'position:absolute;left:0;top:0;overflow:visible', preserveAspectRatio: 'none' });
        let pathD = null;
        if (prst === 'cust') pathD = this.custPath(geomEl, w, h);
        else { const pts = PRST[prst] || [[0, 0], [1, 0], [1, 1], [0, 1]]; pathD = 'M' + pts.map(([a, b]) => `${a * w},${b * h}`).join('L') + 'Z'; }
        const attrs = { d: pathD, fill: 'none' };
        if (fill && fill.css && !fill.grad) attrs.fill = fill.css;
        else if (fill && (fill.grad || fill.img)) { attrs.fill = 'transparent'; this.applyFill(d, fill); d.style.clipPath = `path('${pathD}')`; }
        if (ln) { attrs.stroke = ln.color; attrs['stroke-width'] = ln.w; }
        svg.append(svgEl('path', attrs));
        d.append(svg);
      }
      // texto
      const txBody = OF.kid(el, 'txBody', P);
      if (txBody) this.renderText(txBody, el, ph, lay, mas, ctx, d, sync, styleEl);
      box.append(d);
    }

    custPath(geom, w, h) {
      let out = '';
      for (const path of OF.all(geom, 'path', A)) {
        const pw = +path.getAttribute('w') || w * EMU, ph = +path.getAttribute('h') || h * EMU;
        const X = (v) => (+v / pw) * w, Y = (v) => (+v / ph) * h;
        for (const c of Array.from(path.children)) {
          const pts = OF.kids(c, 'pt', A).map((p) => `${X(p.getAttribute('x'))},${Y(p.getAttribute('y'))}`);
          if (c.localName === 'moveTo') out += 'M' + pts[0];
          else if (c.localName === 'lnTo') out += 'L' + pts[0];
          else if (c.localName === 'cubicBezTo') out += 'C' + pts.join(' ');
          else if (c.localName === 'quadBezTo') out += 'Q' + pts.join(' ');
          else if (c.localName === 'close') out += 'Z';
        }
      }
      return out || `M0,0L${w},0L${w},${h}L0,${h}Z`;
    }

    applyFill(el, f) {
      if (!f || f.none) return;
      if (f.img) {
        this.pending.push(Promise.resolve(f.img).then((u) => { if (u && u !== 'unsupported') { el.style.backgroundImage = `url("${u}")`; el.style.backgroundSize = '100% 100%'; } }));
      } else if (f.css) el.style.background = f.css;
    }

    /* ------------------------------- texto ------------------------------- */
    lvlProps(listEl, lvl) {
      if (!listEl) return null;
      return OF.kid(listEl, `lvl${lvl + 1}pPr`, A) || null;
    }

    renderText(txBody, sp, ph, lay, mas, ctx, d, sync, styleEl) {
      const bodyPr = OF.kid(txBody, 'bodyPr', A);
      const chainBody = [mas && OF.path(mas, 'txBody', 'bodyPr'), lay && OF.path(lay, 'txBody', 'bodyPr'), bodyPr].filter(Boolean);
      const bp = (name, def) => { let v = def; for (const b of chainBody) if (b.hasAttribute(name)) v = b.getAttribute(name); return v; };
      const lIns = px(bp('lIns', 91440)), rIns = px(bp('rIns', 91440)), tIns = px(bp('tIns', 45720)), bIns = px(bp('bIns', 45720));
      const anchor = bp('anchor', 't');
      const wrap = bp('wrap', 'square');
      let fontScale = 1, lnRed = 0;
      for (const b of chainBody) { const na = OF.kid(b, 'normAutofit', A); if (na) { fontScale = na.getAttribute('fontScale') ? +na.getAttribute('fontScale') / 100000 : 1; lnRed = na.getAttribute('lnSpcReduction') ? +na.getAttribute('lnSpcReduction') / 100000 : 0; } }

      // cadena de estilos de lista (de menor a mayor prioridad)
      const t = ph ? ph.type : null;
      const txStyles = ctx.slide.master && OF.all(ctx.slide.master.doc, 'txStyles', P)[0];
      let base = null;
      if (ph) base = txStyles && OF.kid(txStyles, t === 'title' || t === 'ctrTitle' ? 'titleStyle' : ['body', 'subTitle', 'obj'].includes(t) || !t ? 'bodyStyle' : 'otherStyle', P);
      else base = this.defaultTextStyle;
      if (ph && ['dt', 'ftr', 'sldNum', 'hdr'].includes(t)) base = txStyles && OF.kid(txStyles, 'otherStyle', P);
      const lists = [base, mas && OF.path(mas, 'txBody', 'lstStyle'), lay && OF.path(lay, 'txBody', 'lstStyle'), OF.kid(txBody, 'lstStyle', A)];
      const masCtx = this.ctxFor(ctx.slide, 'master');
      const fontRef = styleEl && OF.kid(styleEl, 'fontRef', A);
      const shapeColor = fontRef ? this.color(fontRef, ctx) : null;

      const tx = OF.el('div', { class: 'pp-tx', 'data-tx': '' });
      Object.assign(tx.style, { padding: `${tIns}px ${rIns}px ${bIns}px ${lIns}px`, justifyContent: anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start' });
      if (wrap === 'none') tx.style.whiteSpace = 'nowrap';
      const vert = bp('vert', 'horz');
      if (vert === 'vert' || vert === 'eaVert') tx.style.writingMode = 'vertical-rl';
      if (vert === 'vert270') { tx.style.writingMode = 'vertical-rl'; tx.style.transform = 'rotate(180deg)'; }

      const autoNum = {};
      const editable = !!sync;
      for (const p of OF.kids(txBody, 'p', A)) {
        const pPr = OF.kid(p, 'pPr', A);
        const lvl = pPr && pPr.hasAttribute('lvl') ? +pPr.getAttribute('lvl') : 0;
        const chain = lists.map((l) => this.lvlProps(l, lvl)).filter(Boolean);
        if (pPr) chain.push(pPr);
        const pa = (name, def) => { let v = def; for (const c of chain) if (c.hasAttribute(name)) v = c.getAttribute(name); return v; };
        const pk = (name) => { let v = null; for (const c of chain) { const k = OF.kid(c, name, A); if (k) v = k; } return v; };
        const defR = chain.map((c) => OF.kid(c, 'defRPr', A)).filter(Boolean);

        const pd = OF.el('div', { class: 'pp-p' });
        const algn = pa('algn', 'l');
        pd.style.textAlign = { l: 'left', ctr: 'center', r: 'right', just: 'justify', dist: 'justify' }[algn] || 'left';
        const marL = px(pa('marL', 0)), indent = px(pa('indent', 0));
        pd.style.paddingLeft = marL + 'px';
        pd.style.textIndent = indent + 'px';
        const sp = (el) => { if (!el) return null; const pts = OF.kid(el, 'spcPts', A), pct = OF.kid(el, 'spcPct', A); return pts ? { pt: +pts.getAttribute('val') / 100 } : pct ? { pct: +pct.getAttribute('val') / 100000 } : null; };
        const ls = sp(pk('lnSpc'));
        // tamaño de letra efectivo del párrafo (para espaciados)
        const firstR = OF.all(p, 'rPr', A)[0];
        const szOf = (rPr) => { let s = 18; for (const dr of defR) if (dr.hasAttribute('sz')) s = +dr.getAttribute('sz') / 100; if (rPr && rPr.hasAttribute('sz')) s = +rPr.getAttribute('sz') / 100; return s * fontScale; };
        const pSz = szOf(firstR || OF.kid(p, 'endParaRPr', A));
        if (ls && ls.pct) pd.style.lineHeight = (1.2 * ls.pct * (1 - lnRed)).toFixed(3);
        else if (ls && ls.pt) pd.style.lineHeight = ls.pt + 'pt';
        else pd.style.lineHeight = (1.2 * (1 - lnRed)).toFixed(3);
        const sb = sp(pk('spcBef')), sa = sp(pk('spcAft'));
        if (sb) pd.style.marginTop = sb.pt != null ? sb.pt + 'pt' : sb.pct * pSz + 'pt';
        if (sa) pd.style.marginBottom = sa.pt != null ? sa.pt + 'pt' : sa.pct * pSz + 'pt';

        // estilo de fragmento
        const runCss = (rPr) => {
          const rs = [...defR];
          if (rPr) rs.push(rPr);
          const ra = (name) => { let v = null; for (const c of rs) if (c.hasAttribute(name)) v = c.getAttribute(name); return v; };
          const rk = (name) => { let v = null; for (const c of rs) { const k = OF.kid(c, name, A); if (k) v = { el: k, from: c }; } return v; };
          const st = {};
          st.fontSize = szOf(rPr) + 'pt';
          if (ra('b') === '1' || ra('b') === 'true') st.fontWeight = '700';
          if (ra('i') === '1' || ra('i') === 'true') st.fontStyle = 'italic';
          const u = ra('u'); const strike = ra('strike');
          const deco = [u && u !== 'none' ? 'underline' : '', strike && strike !== 'noStrike' ? 'line-through' : ''].filter(Boolean).join(' ');
          if (deco) st.textDecoration = deco;
          const latin = rk('latin');
          let face = latin ? latin.el.getAttribute('typeface') : '+mn-lt';
          if (face === '+mj-lt') face = ctx.theme.major; else if (face === '+mn-lt') face = ctx.theme.minor;
          if (!latin && (t === 'title' || t === 'ctrTitle')) face = ctx.theme.major;
          st.fontFamily = `"${face}", ${face && /light/i.test(face) ? '"Calibri Light",' : ''} Calibri, Carlito, Arial, sans-serif`;
          const fill = rk('solidFill');
          let col = fill ? this.color(fill.el, fill.from === rPr ? ctx : masCtx) : null;
          // el estilo de la forma (fontRef) manda sobre los estilos de lista heredados, pero no sobre el formato directo
          if (shapeColor && (!fill || fill.from !== rPr)) col = shapeColor;
          if (!col) col = this.color(OF.parseXml(`<a:schemeClr xmlns:a="${A}" val="tx1"/>`).documentElement, ctx);
          if (rPr && OF.kid(rPr, 'hlinkClick', A) && !fill) { col = '#' + ctx.theme.colors.hlink; st.textDecoration = 'underline'; }
          st.color = col;
          const cap = ra('cap'); if (cap === 'all') st.textTransform = 'uppercase'; if (cap === 'small') st.fontVariant = 'small-caps';
          const bl = ra('baseline'); if (bl && +bl) { st.verticalAlign = +bl > 0 ? 'super' : 'sub'; st.fontSize = szOf(rPr) * 0.65 + 'pt'; }
          const spc = ra('spc'); if (spc) st.letterSpacing = +spc / 100 + 'pt';
          return st;
        };

        // viñetas
        const runs = Array.from(p.children).filter((c) => ['r', 'br', 'fld'].includes(c.localName));
        const hasText = runs.some((r) => r.localName !== 'br' && OF.kid(r, 't', A) && OF.kid(r, 't', A).textContent);
        const buNone = pk('buNone'), buChar = pk('buChar'), buAuto = pk('buAutoNum');
        let bullet = null;
        // el último elemento de viñeta definido gana
        let lastBu = null;
        for (const c of chain) for (const k of ['buNone', 'buChar', 'buAutoNum', 'buBlip']) if (OF.kid(c, k, A)) lastBu = OF.kid(c, k, A);
        if (lastBu && hasText) {
          if (lastBu.localName === 'buChar') bullet = lastBu.getAttribute('char');
          else if (lastBu.localName === 'buAutoNum') {
            const key = lvl + (lastBu.getAttribute('type') || '');
            autoNum[key] = (autoNum[key] || +(lastBu.getAttribute('startAt') || 1) - 1) + 1;
            const nn = autoNum[key], ty = lastBu.getAttribute('type') || 'arabicPeriod';
            const val = /alphaLc/.test(ty) ? String.fromCharCode(96 + nn) : /alphaUc/.test(ty) ? String.fromCharCode(64 + nn) : /romanLc/.test(ty) ? roman(nn).toLowerCase() : /romanUc/.test(ty) ? roman(nn) : String(nn);
            bullet = /Paren(R|Both)/.test(ty) ? (/Both/.test(ty) ? '(' : '') + val + ')' : /Period/.test(ty) ? val + '.' : val;
          }
        }
        if (bullet != null) {
          const bst = runCss(OF.kid(runs.find((r) => r.localName === 'r') || p, 'rPr', A));
          const buClr = pk('buClr'); if (buClr) bst.color = this.color(buClr, ctx);
          const buSz = pk('buSzPct'); if (buSz) bst.fontSize = parseFloat(bst.fontSize) * (+buSz.getAttribute('val') / 100000) + 'pt';
          const buFont = pk('buFont'); if (buFont) { const tf = buFont.getAttribute('typeface'); bst.fontFamily = /wingdings|symbol/i.test(tf) ? 'inherit' : `"${tf}"`; if (/wingdings/i.test(tf)) bullet = mapWingdings(bullet); }
          bst.fontWeight = bst.fontWeight || 'normal'; bst.fontStyle = 'normal'; bst.textDecoration = 'none';
          const bs = OF.el('span', { class: 'pp-bu', 'data-x': '', contenteditable: 'false', style: bst }, bullet);
          bs.style.display = 'inline-block';
          bs.style.minWidth = Math.max(0, -indent) + 'px';
          bs.style.textIndent = '0';
          pd.append(bs);
        }

        let pid = null;
        if (editable) { pid = sync.addPara(p); sync.mapped.add(pid); pd.setAttribute('data-p', pid); }
        for (const r of runs) {
          if (r.localName === 'br') { pd.append(OF.el('br', { 'data-x': '' })); continue; }
          const tEl = OF.kid(r, 't', A);
          const text = tEl ? tEl.textContent : '';
          if (r.localName === 'fld') {
            const ty = r.getAttribute('type') || '';
            const shown = /slidenum/.test(ty) ? String(ctx.slide.num) : text;
            pd.append(OF.el('span', { 'data-x': '', contenteditable: 'false', style: runCss(OF.kid(r, 'rPr', A)) }, shown));
            continue;
          }
          if (!tEl || text === '') continue;
          const span = OF.el('span', { style: runCss(OF.kid(r, 'rPr', A)) }, text);
          if (editable) span.setAttribute('data-r', sync.addRun(pid, r));
          pd.append(span);
        }
        if (!hasText) {
          // párrafo vacío: mantener altura según endParaRPr
          const st = runCss(OF.kid(p, 'endParaRPr', A));
          pd.style.fontSize = st.fontSize;
          pd.append(OF.el('br', { 'data-x': '' }));
        } else pd.style.fontSize = pSz + 'pt';
        tx.append(pd);
      }
      if (editable) tx.setAttribute('data-slide-edit', '');
      if (editable && ph && !OF.all(txBody, 't', A).some((t) => t.textContent)) {
        tx.classList.add('pp-empty');
        const first = tx.querySelector('.pp-p');
        if (first) first.dataset.hint = t === 'title' || t === 'ctrTitle' ? 'Toca para agregar un título' : t === 'subTitle' ? 'Toca para agregar un subtítulo' : 'Toca para agregar texto';
      }
      d.append(tx);
    }

    renderTable(tbl, ctx, d, sync) {
      const grid = OF.kids(OF.kid(tbl, 'tblGrid', A), 'gridCol', A).map((g) => px(g.getAttribute('w')));
      const tblPr = OF.kid(tbl, 'tblPr', A);
      const firstRow = tblPr && tblPr.getAttribute('firstRow') === '1', bandRow = tblPr && tblPr.getAttribute('bandRow') === '1';
      const t = OF.el('table', { class: 'pp-tbl' });
      const cg = OF.el('colgroup');
      grid.forEach((w) => cg.append(OF.el('col', { style: { width: w + 'px' } })));
      t.append(cg);
      const acc = this.color(OF.parseXml(`<a:schemeClr xmlns:a="${A}" val="accent1"/>`).documentElement, ctx);
      const accLight = (pct) => this.color(OF.parseXml(`<a:schemeClr xmlns:a="${A}" val="accent1"><a:tint val="${pct}"/></a:schemeClr>`).documentElement, ctx);
      OF.kids(tbl, 'tr', A).forEach((tr, ri) => {
        const row = OF.el('tr', { style: { height: px(tr.getAttribute('h')) + 'px' } });
        OF.kids(tr, 'tc', A).forEach((tc) => {
          if (tc.getAttribute('hMerge') === '1' || tc.getAttribute('vMerge') === '1') return;
          const td = OF.el('td');
          if (tc.getAttribute('gridSpan')) td.colSpan = +tc.getAttribute('gridSpan');
          if (tc.getAttribute('rowSpan')) td.rowSpan = +tc.getAttribute('rowSpan');
          const tcPr = OF.kid(tc, 'tcPr', A);
          const f = tcPr && this.fill(tcPr, ctx);
          let header = false;
          if (f && !f.none) this.applyFill(td, f);
          else if (!f && firstRow && ri === 0) { td.style.background = acc; header = true; }
          else if (!f && bandRow && ri % 2 === (firstRow ? 1 : 0)) td.style.background = accLight(40000);
          else if (!f && bandRow) td.style.background = accLight(20000);
          for (const [side, css] of [['lnL', 'borderLeft'], ['lnR', 'borderRight'], ['lnT', 'borderTop'], ['lnB', 'borderBottom']]) {
            const l = tcPr && OF.kid(tcPr, side, A);
            if (!l) continue;
            if (OF.kid(l, 'noFill', A)) td.style[css] = 'none';
            else { const sf = OF.kid(l, 'solidFill', A); if (sf) td.style[css] = `${Math.max(1, px(l.getAttribute('w') || 12700))}px solid ${this.color(sf, ctx)}`; }
          }
          if (tcPr) {
            const mar = (n, def) => px(tcPr.getAttribute(n) || def);
            td.style.padding = `${mar('marT', 45720)}px ${mar('marR', 91440)}px ${mar('marB', 45720)}px ${mar('marL', 91440)}px`;
            const an = tcPr.getAttribute('anchor'); td.style.verticalAlign = an === 'ctr' ? 'middle' : an === 'b' ? 'bottom' : 'top';
          }
          const txb = OF.kid(tc, 'txBody', A);
          if (txb) {
            const holder = OF.el('div', { class: 'pp-cell' });
            const fakeSp = null;
            this.renderText(txb, fakeSp, null, null, null, ctx, holder, sync, null);
            const tx = holder.firstChild;
            tx.style.padding = '0';
            tx.style.position = 'static';
            if (header) tx.querySelectorAll('span[data-r], .pp-p').forEach((s) => { if (!s.style.color || s.dataset.r) { s.style.color = '#fff'; s.style.fontWeight = '700'; } });
            td.append(tx);
          }
          row.append(td);
        });
        t.append(row);
      });
      d.append(t);
    }

    /* ------------------------------ edición ------------------------------ */
    enableEditing() {
      this.list.querySelectorAll('[data-slide-edit]').forEach((tx) => {
        tx.setAttribute('contenteditable', 'true');
        tx.setAttribute('spellcheck', 'true');
        tx.setAttribute('lang', 'es');
        OF.plainPaste(tx, this);
        tx.addEventListener('input', () => { this.markDirty(); if (tx.textContent.trim()) tx.classList.remove('pp-empty'); });
      });
    }

    markDirty() { if (!this.dirty) { this.dirty = true; this.onchange && this.onchange(); } }

    /* ---- imágenes de la diapositiva: tocar para seleccionar, arrastrar, esquina para tamaño ---- */
    bindPictures() {
      if (this._picBound) return;
      this._picBound = true;
      this.list.addEventListener('pointerdown', (e) => {
        const d = e.target.closest('.pp-pic');
        if (!d) { if (!e.target.closest('.imgbar')) this.selectPic(null); return; }
        e.preventDefault();
        const handle = e.target.closest('.pp-handle');
        this.selectPic(d);
        const z = this.scale || 1;
        const st = { x: e.clientX, y: e.clientY, l: parseFloat(d.style.left), t: parseFloat(d.style.top), w: parseFloat(d.style.width), h: parseFloat(d.style.height), moved: false };
        const move = (ev) => {
          const dx = (ev.clientX - st.x) / z, dy = (ev.clientY - st.y) / z;
          if (Math.abs(dx) + Math.abs(dy) > 2) st.moved = true;
          if (handle) { const w = Math.max(20, st.w + dx); d.style.width = w + 'px'; d.style.height = (w * st.h) / st.w + 'px'; }
          else { d.style.left = st.l + dx + 'px'; d.style.top = st.t + dy + 'px'; }
          OF._imgBar && OF._imgBar.place();
        };
        const up = () => {
          window.removeEventListener('pointermove', move); window.removeEventListener('pointerup', up);
          if (st.moved) this.writePic(d);
        };
        window.addEventListener('pointermove', move); window.addEventListener('pointerup', up);
      });
      document.addEventListener('keydown', (e) => {
        if (this.selPic && this.selPic.isConnected && (e.key === 'Delete' || e.key === 'Backspace') && !e.target.closest('[contenteditable=true]')) { e.preventDefault(); this.picAction('del'); }
      });
    }
    selectPic(d) {
      if (this.selPic) { this.selPic.classList.remove('pp-selpic'); this.selPic.querySelectorAll('.pp-handle').forEach((h) => h.remove()); }
      this.selPic = d;
      if (!d) { OF.closeImageBar(); return; }
      d.classList.add('pp-selpic');
      d.append(OF.el('div', { class: 'pp-handle', 'aria-hidden': 'true' }));
      OF.imageBar(d, [['Más pequeña', () => this.picAction(0.8)], ['Más grande', () => this.picAction(1.25)], ['Eliminar', () => this.picAction('del'), true]]);
    }
    picXfrm(d) {
      const { el } = this.picReg[+d.dataset.pe];
      const spPr = OF.kid(el, 'spPr', P);
      let xfrm = OF.kid(spPr, 'xfrm', A);
      if (!xfrm) { xfrm = OF.fragment(el.ownerDocument, `<a:xfrm xmlns:a="${A}"><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>`); spPr.insertBefore(xfrm, spPr.firstChild); }
      return xfrm;
    }
    writePic(d) {
      const x = this.picXfrm(d);
      OF.kid(x, 'off', A).setAttribute('x', Math.round(parseFloat(d.style.left) * EMU));
      OF.kid(x, 'off', A).setAttribute('y', Math.round(parseFloat(d.style.top) * EMU));
      OF.kid(x, 'ext', A).setAttribute('cx', Math.round(parseFloat(d.style.width) * EMU));
      OF.kid(x, 'ext', A).setAttribute('cy', Math.round(parseFloat(d.style.height) * EMU));
      this.markDirty();
    }
    picAction(a) {
      const d = this.selPic;
      if (!d) return;
      if (a === 'del') {
        const { el } = this.picReg[+d.dataset.pe];
        el.parentNode.removeChild(el);
        d.remove();
        this.selectPic(null);
        this.markDirty();
        return;
      }
      const w = parseFloat(d.style.width), h = parseFloat(d.style.height), l = parseFloat(d.style.left), t = parseFloat(d.style.top);
      const nw = w * a, nh = h * a;
      Object.assign(d.style, { width: nw + 'px', height: nh + 'px', left: l - (nw - w) / 2 + 'px', top: t - (nh - h) / 2 + 'px' });
      this.writePic(d);
      OF._imgBar && OF._imgBar.place();
    }

    fit() {
      if (!this.list) return;
      this.enableEditing();
      this._picBound = false;
      this.bindPictures();
      const availW = Math.max(200, this.host.clientWidth - 60);
      const availH = Math.max(200, this.host.clientHeight - 50);
      const z = Math.min(1.6, availW / this.W, Math.max(availH / this.H, 0.55 * availW / this.W)) * (this.userZoom || 1);
      this.list.querySelectorAll('.pp-frame').forEach((f) => { f.style.width = this.W * z + 'px'; f.style.height = this.H * z + 'px'; f.firstChild.style.transform = `scale(${z})`; });
      this.scale = z;
    }
    zoom(f) { this.userZoom = Math.max(0.3, Math.min(3, (this.userZoom || 1) * f)); this.fit(); }

    async save() {
      if (!this.dirty) return new Blob([this.original], { type: PptDoc.mime });
      this.slides.forEach((slide, i) => {
        const dps = Array.from(slide.box.querySelectorAll('[data-slide-edit] .pp-p, [data-slide-edit] > div'));
        const uniq = Array.from(new Set(dps)).filter((e) => e.closest('[data-slide-edit]'));
        this.syncs[i].apply(uniq);
        this.zip.file(slide.path, OF.serializeXml(slide.doc));
      });
      const blob = await this.zip.generateAsync({ type: 'blob', mimeType: PptDoc.mime, compression: 'DEFLATE' });
      this.original = await blob.arrayBuffer();
      this.dirty = false;
      await this.render();
      return blob;
    }

    /* ---------------------------- presentación ---------------------------- */
    present(start) {
      const visible = this.slides.map((s, i) => i).filter((i) => !this.slides[i].hidden);
      let k = Math.max(0, visible.indexOf(start || 0));
      const ov = OF.el('div', { class: 'pp-show', tabindex: '0' });
      const stage = OF.el('div', { class: 'pp-stage' });
      ov.append(stage);
      document.body.append(ov);
      const show = () => {
        stage.innerHTML = '';
        const src = this.slides[visible[k]].box;
        const clone = src.cloneNode(true);
        clone.querySelectorAll('[contenteditable]').forEach((e) => e.removeAttribute('contenteditable'));
        clone.querySelectorAll('.pp-empty').forEach((e) => e.classList.remove('pp-empty'));
        clone.querySelectorAll('.pp-handle').forEach((e) => e.remove());
        clone.querySelectorAll('.pp-selpic').forEach((e) => e.classList.remove('pp-selpic'));
        const z = Math.min(innerWidth / this.W, innerHeight / this.H);
        clone.style.transform = `scale(${z})`;
        clone.style.transformOrigin = '0 0';
        stage.style.width = this.W * z + 'px'; stage.style.height = this.H * z + 'px';
        stage.append(clone);
      };
      const close = () => { ov.remove(); document.removeEventListener('keydown', key); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); };
      const go = (dk) => { k += dk; if (k < 0) k = 0; if (k >= visible.length) return close(); show(); };
      const key = (e) => { if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(e.key)) go(1); else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) go(-1); else if (e.key === 'Escape') close(); };
      ov.addEventListener('click', (e) => go(e.clientX < innerWidth / 3 ? -1 : 1));
      document.addEventListener('keydown', key);
      if (ov.requestFullscreen) ov.requestFullscreen().catch(() => {});
      window.addEventListener('resize', show, { once: true });
      show();
    }
  }

  function svgEl(tag, attrs) {
    const e = document.createElementNS('http://www.w3.org/2000/svg', tag);
    for (const k in attrs) e.setAttribute(k, attrs[k]);
    return e;
  }
  function roman(n) { const m = [[1000, 'M'], [900, 'CM'], [500, 'D'], [400, 'CD'], [100, 'C'], [90, 'XC'], [50, 'L'], [40, 'XL'], [10, 'X'], [9, 'IX'], [5, 'V'], [4, 'IV'], [1, 'I']]; let s = ''; for (const [v, r] of m) while (n >= v) { s += r; n -= v; } return s; }
  function mapWingdings(c) { return { '§': '■', 'Ø': '➢', 'ü': '✓', 'q': '❑', 'v': '❖', 'n': '■', 'l': '●', 'Ü': '➢', 'à': '→' }[c] || '•'; }

  /* ------------------------- formato en el XML de PowerPoint ------------------------- */
  function pptRunFmt(r, f) {
    const doc = r.ownerDocument;
    let rPr = OF.kid(r, 'rPr', A);
    if (!rPr) { rPr = doc.createElementNS(A, 'a:rPr'); rPr.setAttribute('lang', 'es-CL'); r.insertBefore(rPr, r.firstChild); }
    if ('b' in f) rPr.setAttribute('b', f.b ? '1' : '0');
    if ('i' in f) rPr.setAttribute('i', f.i ? '1' : '0');
    if ('u' in f) rPr.setAttribute('u', f.u ? 'sng' : 'none');
    if ('sz' in f) rPr.setAttribute('sz', Math.round(f.sz * 100));
    if ('color' in f) {
      OF.removeKids(rPr, ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'], A);
      if (f.color) OF.ordered(rPr, A, 'a', 'solidFill', PRPR).appendChild(OF.fragment(doc, `<a:srgbClr xmlns:a="${A}" val="${f.color.replace('#', '').toUpperCase()}"/>`));
    }
    if (f.font) { const l = OF.ordered(rPr, A, 'a', 'latin', PRPR); l.setAttribute('typeface', f.font); OF.removeKids(rPr, ['cs'], A); }
  }
  const PRPR = ['ln', 'noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill', 'effectLst', 'effectDag', 'highlight', 'uLnTx', 'uLn', 'uFillTx', 'uFill', 'latin', 'ea', 'cs', 'sym', 'hlinkClick', 'hlinkMouseOver', 'rtl', 'extLst'];
  function pptParaFmt(p, f) {
    let pPr = OF.kid(p, 'pPr', A);
    if (!pPr) { pPr = p.ownerDocument.createElementNS(A, 'a:pPr'); p.insertBefore(pPr, p.firstChild); }
    if (f.jc) pPr.setAttribute('algn', { left: 'l', center: 'ctr', right: 'r', both: 'just' }[f.jc] || 'l');
  }
  const SLIDE_CT = 'application/vnd.openxmlformats-officedocument.presentationml.slide+xml';
  const REL_SLIDE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide';

  // El formato de texto es igual que en Word (selección -> fragmentos marcados)
  Object.assign(PptDoc.prototype, {
    format: OF.WordDoc.prototype.format,
    toggle: OF.WordDoc.prototype.toggle,
    setColor(hex) { this.format({ color: hex }, { color: hex || '#000' }); },
    setSize(pt) { this.format({ sz: pt }, { fontSize: pt + 'pt' }); },
    align: OF.WordDoc.prototype.align,

    slideIndex() {
      if (this.curSlide != null && this.curSlide < this.slides.length) return this.curSlide;
      const mid = this.host.getBoundingClientRect().top + this.host.clientHeight / 2;
      const wraps = Array.from(this.list.querySelectorAll('.pp-wrap'));
      const hit = wraps.findIndex((w) => { const r = w.getBoundingClientRect(); return r.top <= mid && r.bottom >= mid; });
      return Math.max(0, hit);
    },
    caretInfo() {
      const range = OF.currentRange();
      const node = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
      const dp = node && node.closest && node.closest('[data-p]');
      const wrap = dp && dp.closest('.pp-wrap');
      if (!dp || !wrap) return null;
      const si = +wrap.dataset.i, pid = +dp.dataset.p;
      const same = Array.from(wrap.querySelectorAll(`[data-slide-edit] [data-p="${pid}"]`));
      return { si, pid, gi: Math.max(0, same.indexOf(dp)) };
    },
    // Vuelca todas las ediciones de texto al XML de cada diapositiva
    commitAll() {
      this.slides.forEach((slide, i) => {
        const dps = Array.from(slide.box.querySelectorAll('[data-slide-edit] .pp-p'));
        this.syncs[i].apply(dps);
        this.zip.file(slide.path, OF.serializeXml(slide.doc));
      });
    },
    async structural(fn) {
      const info = this.caretInfo();
      const si = info ? info.si : this.slideIndex();
      this.commitAll();
      let target = null;
      if (info) { const res = this.syncs[info.si].result && this.syncs[info.si].result.get(info.pid); target = res ? res[Math.min(info.gi, res.length - 1)] : null; }
      const ok = await fn(si, target);
      this.zip.file(this.presPath, OF.serializeXml(this.pres));
      await this.load();
      if (ok === false) { await this.render(); return; }
      this.dirty = false;
      this.markDirty();
      await this.render();
    },
    sldIds() { return OF.kids(OF.all(this.pres, 'sldIdLst', P)[0], 'sldId', P); },
    async newSlidePart(xml, relsXml) {
      const path = OF.uniquePart(this.zip, 'ppt/slides', 'slide', 'xml');
      this.zip.file(path, xml);
      this.zip.file(OF.relsPathFor(path), relsXml);
      await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, path, SLIDE_CT));
      const rid = await OF.addRel(this.zip, this.presPath, REL_SLIDE, path);
      return { path, rid };
    },
    insertSldId(rid, afterIdx) {
      const list = OF.all(this.pres, 'sldIdLst', P)[0];
      const ids = this.sldIds();
      const el = this.pres.createElementNS(P, 'p:sldId');
      el.setAttribute('id', Math.max(255, ...ids.map((e) => +e.getAttribute('id'))) + 1);
      el.setAttributeNS(R, 'r:id', rid);
      const ref = ids[afterIdx];
      list.insertBefore(el, ref ? ref.nextSibling : null);
    },
    async presRelsDoc() { const p = OF.relsPathFor(this.presPath); return { p, doc: OF.parseXml(await this.zip.file(p).async('string')) }; },

    async slideOp(op) {
      await this.structural(async (si) => {
        const slide = this.slides[si];
        if (op === 'up' || op === 'down') {
          const ids = this.sldIds(), el = ids[si];
          if (op === 'up' && si > 0) el.parentNode.insertBefore(el, ids[si - 1]);
          else if (op === 'down' && si < ids.length - 1) el.parentNode.insertBefore(ids[si + 1], el);
          else return false;
          this.curSlide = op === 'up' ? si - 1 : si + 1;
          return;
        }
        if (op === 'delete') {
          if (this.slides.length <= 1) { OF.toast('La presentación debe tener al menos una diapositiva.'); return false; }
          const ids = this.sldIds();
          const rid = ids[si].getAttributeNS(R, 'id');
          ids[si].parentNode.removeChild(ids[si]);
          const { p, doc } = await this.presRelsDoc();
          for (const r of OF.all(doc, 'Relationship')) if (r.getAttribute('Id') === rid) r.parentNode.removeChild(r);
          this.zip.file(p, OF.serializeXml(doc));
          const gone = [slide.path];
          for (const rel of Object.values(slide.rels)) if (/notesSlide$/.test(rel.type)) gone.push(rel.target);
          for (const g of gone) { this.zip.remove(g); this.zip.remove(OF.relsPathFor(g)); }
          await OF.contentTypes(this.zip, (ct) => { for (const o of OF.all(ct, 'Override')) if (gone.includes(o.getAttribute('PartName').slice(1))) o.parentNode.removeChild(o); });
          this.curSlide = Math.max(0, si - 1);
          return;
        }
        if (op === 'dup') {
          const xml = OF.serializeXml(slide.doc);
          const relDoc = OF.parseXml(await this.zip.file(OF.relsPathFor(slide.path)).async('string'));
          for (const r of OF.all(relDoc, 'Relationship')) if (/(notesSlide|comments)$/.test(r.getAttribute('Type'))) r.parentNode.removeChild(r);
          const { rid } = await this.newSlidePart(xml, OF.serializeXml(relDoc));
          this.insertSldId(rid, si);
          this.curSlide = si + 1;
          return;
        }
        if (op === 'new') {
          const layout = slide.layout;
          const shapes = [];
          let id = 2;
          const keep = (t) => !['dt', 'ftr', 'sldNum', 'hdr', 'pic', 'chart', 'tbl', 'media', 'clipArt', 'dgm', 'sldImg'].includes(t);
          if (layout) for (const sp of OF.all(OF.all(layout.doc, 'spTree', P)[0], 'sp', P)) {
            const ph = OF.all(sp.firstElementChild, 'ph', P)[0];
            if (!ph || !keep(ph.getAttribute('type') || 'obj')) continue;
            const attrs = Array.from(ph.attributes).map((a) => `${a.name}="${a.value}"`).join(' ');
            shapes.push(`<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Marcador ${id++}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph ${attrs}/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="es-CL"/></a:p></p:txBody></p:sp>`);
          }
          if (!shapes.length) {
            // diseño sin marcadores de texto: título y cuerpo como cuadros de texto
            const W = Math.round(this.W * EMU), H = Math.round(this.H * EMU);
            const box = (y, h, sz, b) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Texto ${id++}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="${Math.round(W * 0.06)}" y="${Math.round(H * y)}"/><a:ext cx="${Math.round(W * 0.88)}" cy="${Math.round(H * h)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr><p:txBody><a:bodyPr wrap="square"><a:normAutofit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="es-CL" sz="${sz}"${b ? ' b="1"' : ''}/></a:p></p:txBody></p:sp>`;
            shapes.push(box(0.06, 0.16, 3200, true), box(0.25, 0.65, 2000, false));
          }
          const xml = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><p:sld xmlns:a="${A}" xmlns:r="${R}" xmlns:p="${P}"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>${shapes.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>`;
          const path = OF.uniquePart(this.zip, 'ppt/slides', 'slide', 'xml');
          const rels = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OF.NS.rel}">${layout ? `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="${OF.relTarget(path, layout.path)}"/>` : ''}</Relationships>`;
          const { rid } = await this.newSlidePart(xml, rels);
          this.insertSldId(rid, si);
          this.curSlide = si + 1;
        }
      });
      const w = this.list.querySelector(`.pp-wrap[data-i="${this.curSlide}"]`);
      if (w) w.scrollIntoView({ block: 'center' });
    },

    async insertImage(file) {
      const im = await OF.prepareImage(file);
      await this.structural(async (si) => {
        const slide = this.slides[si];
        const media = OF.uniquePart(this.zip, 'ppt/media', 'imagen', im.ext);
        this.zip.file(media, im.bytes);
        const rId = await OF.addRel(this.zip, slide.path, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', media);
        await OF.contentTypes(this.zip, (ct) => OF.ensureDefaultCT(ct, im.ext, im.mime));
        const tree = OF.all(slide.doc, 'spTree', P)[0];
        const ids = OF.all(slide.doc, 'cNvPr').map((e) => +e.getAttribute('id') || 0);
        const id = Math.max(1, ...ids) + 1;
        const W = this.W * EMU, H = this.H * EMU;
        const k = Math.min((W * 0.6) / im.w, (H * 0.6) / im.h);
        const cx = Math.round(im.w * k), cy = Math.round(im.h * k);
        const pic = OF.fragment(slide.doc, `<p:pic xmlns:p="${P}" xmlns:a="${A}" xmlns:r="${R}"><p:nvPicPr><p:cNvPr id="${id}" name="Imagen ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr><a:xfrm><a:off x="${Math.round((W - cx) / 2)}" y="${Math.round((H - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`);
        const ext = OF.kid(tree, 'extLst', P);
        tree.insertBefore(pic, ext);
        this.zip.file(slide.path, OF.serializeXml(slide.doc));
        this.curSlide = si;
      });
    },

    async tableRow(op) {
      await this.structural((si, target) => {
        let tr = target;
        while (tr && tr.localName !== 'tr') tr = tr.parentNode;
        if (!tr) { OF.toast('Toca una celda de la tabla primero.'); return false; }
        const tbl = tr.parentNode;
        const frame = (() => { let g = tbl; while (g && g.localName !== 'graphicFrame') g = g.parentNode; return g; })();
        const ext = frame && OF.path(frame, 'xfrm', 'ext');
        const h = +tr.getAttribute('h') || 370840;
        if (op === 'del') {
          if (OF.kids(tbl, 'tr', A).length <= 1) { OF.toast('La tabla debe tener al menos una fila.'); return false; }
          tbl.removeChild(tr);
          if (ext) ext.setAttribute('cy', Math.max(0, +ext.getAttribute('cy') - h));
        } else {
          const nr = tr.cloneNode(true);
          for (const tc of OF.kids(nr, 'tc', A)) {
            const tb = OF.kid(tc, 'txBody', A);
            if (!tb) continue;
            const ps = OF.kids(tb, 'p', A);
            ps.slice(1).forEach((p) => tb.removeChild(p));
            const p = ps[0];
            if (!p) continue;
            const r0 = OF.kid(p, 'r', A);
            for (const el of Array.from(p.children)) if (['r', 'br', 'fld'].includes(el.localName)) p.removeChild(el);
            if (!OF.kid(p, 'endParaRPr', A) && r0 && OF.kid(r0, 'rPr', A)) {
              const e = p.ownerDocument.createElementNS(A, 'a:endParaRPr');
              for (const a of Array.from(OF.kid(r0, 'rPr', A).attributes)) e.setAttribute(a.name, a.value);
              p.appendChild(e);
            }
          }
          tbl.insertBefore(nr, op === 'above' ? tr : tr.nextSibling);
          if (ext) ext.setAttribute('cy', +ext.getAttribute('cy') + h);
        }
        this.zip.file(this.slides[si].path, OF.serializeXml(this.slides[si].doc));
        this.curSlide = si;
      });
    },
  });

  PptDoc.mime = 'application/vnd.openxmlformats-officedocument.presentationml.presentation';
  OF.PptDoc = PptDoc;
})();
