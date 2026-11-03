/* PowerPoint: presentación con animaciones de entrada, salida y énfasis, y transiciones simples. */
(function () {
  const OF = window.OF;
  const P = OF.NS.p;
  const PD = OF.PptDoc.prototype;

  // Lee la secuencia principal (mainSeq) de una diapositiva: pasos que avanzan con cada clic
  PD.animSteps = function (slide) {
    const seq = OF.all(slide.doc, 'cTn', P).find((c) => c.getAttribute('nodeType') === 'mainSeq');
    if (!seq) return [];
    const list = OF.kid(seq, 'childTnLst', P);
    const steps = [];
    for (const par of OF.kids(list, 'par', P)) {
      const effects = [];
      for (const c of par.getElementsByTagNameNS(P, 'cTn')) {
        const cls = c.getAttribute('presetClass');
        if (!cls) continue;
        const tgt = c.getElementsByTagNameNS(P, 'spTgt')[0];
        if (!tgt) continue;
        const dur = c.getElementsByTagNameNS(P, 'cTn');
        let ms = 500;
        for (const x of dur) { const v = +x.getAttribute('dur'); if (v) { ms = v; break; } }
        effects.push({ cls, id: +(c.getAttribute('presetID') || 1), sub: +(c.getAttribute('presetSubtype') || 0), spid: tgt.getAttribute('spid'), node: c.getAttribute('nodeType'), ms: Math.min(ms, 4000), path: cls === 'path' ? this.motionOf(c) : null });
      }
      if (effects.length) steps.push({ auto: effects[0].node && effects[0].node !== 'clickEffect', effects });
    }
    return steps;
  };
  // animación CSS según el efecto de PowerPoint
  function frames(ef, entering) {
    const fly = { 1: 'translateY(-110%)', 2: 'translateX(110%)', 4: 'translateY(110%)', 8: 'translateX(-110%)' }[ef.sub] || 'translateY(110%)';
    let k;
    switch (ef.id) {
      case 1: k = null; break; // aparecer / desaparecer
      case 2: case 3: k = [{ transform: fly, opacity: 1 }, { transform: 'none', opacity: 1 }]; break; // desplazar hacia dentro
      case 21: case 22: k = [{ clipPath: 'inset(0 100% 0 0)' }, { clipPath: 'inset(0 0 0 0)' }]; break; // barrido
      case 23: case 53: case 31: k = [{ transform: 'scale(.3)', opacity: 0 }, { transform: 'none', opacity: 1 }]; break; // zoom / crecer
      default: k = [{ opacity: 0 }, { opacity: 1 }]; // desvanecer y el resto
    }
    if (k && !entering) k = k.slice().reverse();
    return k;
  }
  function run(el, ef, instant, W, H) {
    if (!el) return;
    const t = instant ? 0 : ef.ms;
    if (ef.cls === 'path' && ef.path) {
      // la forma recorre la trayectoria y se queda en el punto final
      const base = (el.dataset.baseTf != null ? el.dataset.baseTf : (el.dataset.baseTf = el.style.transform || '')).trim();
      const prev = el._pathEnd || [0, 0];
      const tf = ([x, y]) => `translate(${(prev[0] + x) * W}px, ${(prev[1] + y) * H}px) ${base}`.trim();
      const end = ef.path[ef.path.length - 1];
      if (t) el.animate([[0, 0], ...ef.path].map((p) => ({ transform: tf(p) })), { duration: t, easing: 'ease-in-out' });
      el.style.transform = tf(end);
      el._pathEnd = [prev[0] + end[0], prev[1] + end[1]];
      return;
    }
    if (ef.cls === 'entr') {
      el.style.visibility = 'visible';
      const k = frames(ef, true);
      if (k && t) el.animate(k, { duration: t, easing: 'ease-out' });
    } else if (ef.cls === 'exit') {
      const k = frames(ef, false);
      if (k && t) el.animate(k, { duration: t, easing: 'ease-in' }).onfinish = () => { el.style.visibility = 'hidden'; };
      else el.style.visibility = 'hidden';
    } else if (ef.cls === 'emph' && t) {
      el.animate([{ transform: el.style.transform || 'none' }, { transform: (el.style.transform || '') + ' scale(1.08)' }, { transform: el.style.transform || 'none' }], { duration: Math.max(t, 600) });
    }
  }

  PD.present = function (start) {
    const visible = this.slides.map((s, i) => i).filter((i) => !this.slides[i].hidden);
    let k = Math.max(0, visible.indexOf(start || 0));
    let steps = [], si = 0, root = null, timer = null;
    const ov = OF.el('div', { class: 'pp-show', tabindex: '0' });
    const stage = OF.el('div', { class: 'pp-stage' });
    const counter = OF.el('div', { class: 'pp-show-n' });
    ov.append(stage, counter);
    document.body.append(ov);
    const find = (spid) => root.querySelector(`[data-spid="${spid}"]`);
    const play = (instant) => {
      const st = steps[si++];
      for (const ef of st.effects) run(find(ef.spid), ef, instant, this.W, this.H);
      // los pasos «después de» / «con» el anterior se encadenan solos
      clearTimeout(timer);
      if (!instant && si < steps.length && steps[si].auto) timer = setTimeout(() => play(false), Math.max(...st.effects.map((e) => e.ms)) + 60);
    };
    const show = (allDone, transition) => {
      clearTimeout(timer);
      stage.innerHTML = '';
      const slide = this.slides[visible[k]];
      const clone = slide.box.cloneNode(true);
      clone.querySelectorAll('[contenteditable]').forEach((e) => e.removeAttribute('contenteditable'));
      clone.querySelectorAll('.pp-empty').forEach((e) => e.classList.remove('pp-empty'));
      clone.querySelectorAll('.pp-handle, .pp-selframe').forEach((e) => e.remove());
      clone.querySelectorAll('.pp-selpic').forEach((e) => e.classList.remove('pp-selpic'));
      const z = Math.min(innerWidth / this.W, innerHeight / this.H);
      clone.style.transform = `scale(${z})`;
      clone.style.transformOrigin = '0 0';
      stage.style.width = this.W * z + 'px'; stage.style.height = this.H * z + 'px';
      stage.append(clone);
      root = clone;
      steps = this.animSteps(slide);
      si = 0;
      // lo que entra con animación empieza oculto
      const seen = new Set();
      for (const st of steps) for (const ef of st.effects) { if (seen.has(ef.spid)) continue; seen.add(ef.spid); if (ef.cls === 'entr') { const el = find(ef.spid); if (el) el.style.visibility = 'hidden'; } }
      if (allDone) while (si < steps.length) play(true);
      else if (steps.length && steps[0].auto) play(false);
      if (transition) stage.animate([{ opacity: 0 }, { opacity: 1 }], { duration: 250 });
      counter.textContent = `${k + 1} / ${visible.length}`;
    };
    const close = () => { clearTimeout(timer); ov.remove(); document.removeEventListener('keydown', key); if (document.fullscreenElement) document.exitFullscreen().catch(() => {}); };
    const go = (dk) => {
      if (dk > 0 && si < steps.length) { play(false); return; }
      k += dk;
      if (k < 0) { k = 0; return; }
      if (k >= visible.length) return close();
      show(dk < 0, true);
    };
    const key = (e) => { if (['ArrowRight', 'ArrowDown', ' ', 'PageDown', 'Enter'].includes(e.key)) { e.preventDefault(); go(1); } else if (['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'].includes(e.key)) { e.preventDefault(); go(-1); } else if (e.key === 'Escape') close(); };
    ov.addEventListener('click', (e) => go(e.clientX < innerWidth / 3 ? -1 : 1));
    document.addEventListener('keydown', key);
    if (ov.requestFullscreen) ov.requestFullscreen().catch(() => {});
    window.addEventListener('resize', () => show(true), { once: true });
    show(false);
    this._show = { go, close, state: () => ({ k, si, steps: steps.length }) };
  };

  /* ------------------------- texto de SmartArt ------------------------- */
  const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';
  const DGM = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';
  const A = OF.NS.a;
  PD.commitSmart = function () {
    for (const sp of this.smartParts || []) {
      if (!sp.box.isConnected) continue;
      const dps = Array.from(sp.box.querySelectorAll('[data-slide-edit] .pp-p'));
      const before = OF.serializeXml(sp.doc);
      sp.sync.apply(dps);
      const after = OF.serializeXml(sp.doc);
      if (after === before) continue;
      // el dibujo vuelve a su espacio de nombres original
      this.zip.file(sp.path, after.split(P).join(DSP));
      // PowerPoint vuelve a armar el SmartArt desde sus datos: se copia el texto también ahí
      if (!sp.dataDoc) continue;
      const data = sp.dataDoc;
      const pts = OF.all(data, 'pt', DGM);
      const byId = new Map(pts.map((x) => [x.getAttribute('modelId'), x]));
      for (const shape of OF.all(sp.doc, 'sp', P)) {
        const mid = shape.getAttribute('modelId');
        const tx = OF.kid(shape, 'txBody', P);
        if (!mid || !tx) continue;
        const pres = byId.get(mid);
        const prSet = pres && OF.kid(pres, 'prSet', DGM);
        const node = prSet && byId.get(prSet.getAttribute('presAssocID'));
        const t = node && OF.kid(node, 't', DGM);
        if (!t) continue;
        for (const old of OF.kids(t, 'p', A)) t.removeChild(old);
        for (const np of OF.kids(tx, 'p', A)) t.appendChild(data.importNode(np, true));
      }
      this.zip.file(sp.dataPath, OF.serializeXml(data));
    }
  };
  const baseSave = PD.save;
  PD.save = async function () {
    const r0 = this.dirty;
    void r0;
    return baseSave.call(this);
  };

  /* ------------------------- trayectorias de animación ------------------------- */
  // p:animMotion path="M 0 0 L 0.25 0.1 E": coordenadas relativas al tamaño de la diapositiva
  PD.motionOf = function (cTn) {
    const am = cTn.getElementsByTagNameNS(P, 'animMotion')[0];
    if (!am) return null;
    const nums = (am.getAttribute('path') || '').replace(/[A-DF-Za-df-z]/g, ' ').trim().split(/[\s,]+/).map(Number).filter((n) => !isNaN(n));
    const pts = [];
    for (let i = 0; i + 1 < nums.length; i += 2) pts.push([nums[i], nums[i + 1]]);
    return pts.length > 1 ? pts : null;
  };
})();
