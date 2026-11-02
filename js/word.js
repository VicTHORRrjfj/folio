/* Editor de Word: vista fiel con docx-preview + edición de texto sobre el XML original. */
(function () {
  const OF = window.OF;
  const W = OF.NS.w;
  const SKIP_ANCESTORS = new Set(['txbxContent', 'Fallback', 'del', 'moveFrom', 'comment', 'footnote', 'endnote']);
  const CONTENT_SKIP = new Set(['rPr']);

  class WordDoc {
    constructor() {
      this.kind = 'word';
      this.dirty = false;
    }

    async open(buffer, host) {
      this.original = buffer;
      this.zip = await JSZip.loadAsync(buffer);
      const rels = await OF.readRels(this.zip, '');
      const main = Object.values(rels).find((r) => /officeDocument$/.test(r.type));
      this.docPath = main ? main.target : 'word/document.xml';
      this.xml = OF.parseXml(await this.zip.file(this.docPath).async('string'));
      this.normalize();
      this.host = host;
      await this.render();
    }

    // Separa fragmentos con varios contenidos (texto+tabulación, etc.) en fragmentos de un solo
    // contenido con el mismo formato. Es equivalente para Word y permite un mapeo exacto.
    normalize() {
      const runs = Array.from(this.xml.getElementsByTagNameNS(W, 'r'));
      for (const r of runs) {
        const content = Array.from(r.children).filter((c) => !CONTENT_SKIP.has(c.localName));
        if (content.length < 2 || !content.some((c) => c.localName === 't')) continue;
        const rpr = OF.kid(r, 'rPr', W);
        for (const c of content) {
          const nr = r.cloneNode(false);
          if (rpr) nr.appendChild(rpr.cloneNode(true));
          nr.appendChild(c);
          r.parentNode.insertBefore(nr, r);
        }
        r.parentNode.removeChild(r);
      }
    }

    async render() {
      const host = this.host;
      const scroll = host.scrollTop;
      if (this.urls) this.urls.forEach(URL.revokeObjectURL);
      host.innerHTML = '';
      const styleBox = OF.el('div', { class: 'docx-styles' });
      const pages = OF.el('div', { class: 'docx-pages' });
      host.append(styleBox, pages);
      this.zip.file(this.docPath, OF.serializeXml(this.xml));
      const blob = await this.zip.generateAsync({ type: 'blob' });
      await docx.renderAsync(blob, pages, styleBox, {
        className: 'docx',
        inWrapper: true,
        breakPages: true,
        ignoreLastRenderedPageBreak: true,
        renderHeaders: true,
        renderFooters: true,
        renderFootnotes: true,
        renderEndnotes: true,
        useBase64URL: false,
        experimental: true,
      });
      this.pages = pages;
      this.fixBullets(styleBox);
      this.map();
      this.enableEditing();
      this.fit();
      host.scrollTop = scroll;
    }

    // Las viñetas de Word usan las fuentes Symbol/Wingdings, que no existen en tablets: usar Unicode
    fixBullets(styleBox) {
      const sym = ['•', '◦', '▪', '•', '◦', '▪', '•', '◦', '▪'];
      for (const st of styleBox.querySelectorAll('style')) {
        const sheet = st.sheet;
        if (!sheet) continue;
        for (const rule of Array.from(sheet.cssRules)) {
          const m = /docx-num-\d+-(\d+)::?before/.exec(rule.selectorText || '');
          if (!m || !rule.style) continue;
          const content = rule.style.getPropertyValue('content');
          const ff = rule.style.getPropertyValue('font-family');
          if (/counter/.test(content)) continue;
          const visible = content.replace(/^"|"$/g, '').replace(/\\9|\s|\\[0-9a-f]{1,6}/gi, '');
          if (/symbol|wingdings/i.test(ff) || !visible) {
            rule.style.setProperty('content', `"${sym[+m[1]] || '•'}"`);
            rule.style.setProperty('font-family', 'Arial, sans-serif');
            rule.style.setProperty('display', 'inline-block');
            rule.style.setProperty('min-width', '18pt');
            rule.style.setProperty('text-indent', '0');
          }
        }
      }
    }

    map() {
      const sync = (this.sync = new OF.TextSync({
        ns: W, prefix: 'w', P: 'p', R: 'r', T: 't', PPR: 'pPr',
        wrappers: new Set(['hyperlink', 'smartTag', 'ins', 'fldSimple', 'sdt', 'sdtContent', 'customXml', 'moveTo', 'bdo', 'dir']),
        mustHaveParagraph: (parent) => parent.localName === 'tc' || parent.localName === 'body',
        markProps: (p) => OF.kid(OF.kid(p, 'pPr', W), 'rPr', W),
        applyFmt: applyRunFmt,
        applyPFmt: applyParaFmt,
      }));
      const body = OF.kid(this.xml.documentElement, 'body', W);
      const xparas = Array.from(body.getElementsByTagNameNS(W, 'p')).filter((p) => {
        for (let a = p.parentNode; a && a !== body; a = a.parentNode) if (SKIP_ANCESTORS.has(a.localName)) return false;
        return true;
      });
      const articles = Array.from(this.pages.querySelectorAll('section.docx > article'));
      const dps = articles.flatMap((a) => Array.from(a.querySelectorAll('p')));
      const key = (s) => OF.cleanText(s).replace(/\s+/g, '');
      const dkeys = dps.map((p) => key(p.textContent));
      let j = 0;
      let mappedCount = 0;
      for (const xp of xparas) {
        const xruns = sync.textRunsOf(xp).filter((r) => sync.textOfRun(r) !== '');
        const xk = key(xruns.map((r) => sync.textOfRun(r)).join(''));
        let found = -1;
        for (let k = j; k < Math.min(dps.length, j + 40); k++) if (dkeys[k] === xk) { found = k; break; }
        if (found < 0) continue;
        const dp = dps[found];
        // emparejar fragmentos con <span> hoja en orden
        const leaves = Array.from(dp.querySelectorAll('span')).filter(
          (s) => s.childNodes.length > 0 && Array.from(s.childNodes).every((n) => n.nodeType === 3) && !s.closest('[data-x]')
        );
        const pairs = [];
        let li = 0, ok = true;
        for (const xr of xruns) {
          const t = OF.cleanText(sync.textOfRun(xr));
          while (li < leaves.length && OF.cleanText(leaves[li].textContent) !== t) li++;
          if (li >= leaves.length) { ok = false; break; }
          pairs.push([xr, leaves[li++]]);
        }
        j = found + 1;
        if (!ok) continue;
        const pid = sync.addPara(xp);
        sync.mapped.add(pid);
        dp.setAttribute('data-p', pid);
        for (const [xr, span] of pairs) span.setAttribute('data-r', sync.addRun(pid, xr));
        mappedCount++;
      }
      // todo lo no mapeado queda de solo lectura (se conserva tal cual)
      for (const dp of dps) if (!dp.hasAttribute('data-p')) { dp.setAttribute('contenteditable', 'false'); dp.classList.add('of-ro'); }
      for (const a of articles) {
        a.querySelectorAll('img, svg, canvas').forEach((m) => {
          const box = m.closest('span:not([data-r])') || m.parentElement;
          box.setAttribute('data-x', '');
          box.setAttribute('contenteditable', 'false');
        });
        a.querySelectorAll('span:not([data-r])').forEach((s) => {
          if (s.closest('[data-p]') && !s.querySelector('[data-r]') && s.textContent.trim() === '' && s.childElementCount === 0) s.setAttribute('data-x', '');
        });
      }
      this.stats = { paras: xparas.length, mapped: mappedCount };
      // imágenes del cuerpo <-> elementos w:drawing (mismo orden)
      const drs = Array.from(body.getElementsByTagNameNS(W, 'drawing')).filter((d) => { for (let a = d.parentNode; a && a !== body; a = a.parentNode) if (SKIP_ANCESTORS.has(a.localName)) return false; return true; });
      const imgs = articles.flatMap((a) => Array.from(a.querySelectorAll('img')));
      this.drawings = drs.length === imgs.length ? drs : [];
      if (this.drawings.length) imgs.forEach((im, i) => { im.dataset.di = i; im.classList.add('of-img'); });
    }

    enableEditing() {
      const articles = this.pages.querySelectorAll('section.docx > article');
      articles.forEach((a) => {
        a.setAttribute('contenteditable', 'true');
        a.setAttribute('spellcheck', 'true');
        a.setAttribute('lang', 'es');
        OF.plainPaste(a);
        a.addEventListener('input', () => this.markDirty());
      });
      OF.closeImageBar();
      this.pages.addEventListener('click', (e) => {
        const im = e.target.closest('img.of-img');
        this.pages.querySelectorAll('.of-img.sel').forEach((x) => x.classList.remove('sel'));
        if (!im) { if (!e.target.closest('.imgbar')) OF.closeImageBar(); return; }
        im.classList.add('sel');
        const i = +im.dataset.di;
        OF.imageBar(im, [['Más pequeña', () => this.imageAction(i, 0.8)], ['Más grande', () => this.imageAction(i, 1.25)], ['Eliminar', () => this.imageAction(i, 'del'), true]]);
      });
      // encabezados/pies: solo lectura en esta versión
      this.pages.querySelectorAll('section.docx > header, section.docx > footer').forEach((h) => h.setAttribute('title', 'Encabezado/pie: solo lectura'));
    }

    markDirty() {
      if (!this.dirty) { this.dirty = true; this.onchange && this.onchange(); }
    }

    fit() {
      const sec = this.pages.querySelector('section.docx');
      if (!sec) return;
      const wrap = this.pages.querySelector('.docx-wrapper') || this.pages;
      wrap.style.zoom = 1;
      const avail = this.host.clientWidth - 24;
      const z = Math.min(1, avail / sec.offsetWidth);
      this.baseZoom = z;
      wrap.style.zoom = z * (this.userZoom || 1);
    }
    zoom(f) {
      this.userZoom = Math.max(0.4, Math.min(3, (this.userZoom || 1) * f));
      this.fit();
    }

    async save() {
      if (!this.dirty) return new Blob([this.original], { type: WordDoc.mime });
      const dps = Array.from(this.pages.querySelectorAll('section.docx > article p'));
      this.sync.apply(dps);
      this.zip.file(this.docPath, OF.serializeXml(this.xml));
      const blob = await this.zip.generateAsync({ type: 'blob', mimeType: WordDoc.mime, compression: 'DEFLATE' });
      this.original = await blob.arrayBuffer();
      this.dirty = false;
      await this.render();
      return blob;
    }
  }
  /* ------------------------- formato en el XML de Word ------------------------- */
  const RPR = ['rStyle', 'rFonts', 'b', 'bCs', 'i', 'iCs', 'caps', 'smallCaps', 'strike', 'dstrike', 'outline', 'shadow', 'emboss', 'imprint', 'noProof', 'snapToGrid', 'vanish', 'webHidden', 'color', 'spacing', 'w', 'kern', 'position', 'sz', 'szCs', 'highlight', 'u', 'effect', 'bdr', 'shd', 'fitText', 'vertAlign', 'rtl', 'cs', 'em', 'lang', 'eastAsianLayout', 'specVanish', 'oMath'];
  const PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  const setVal = (el, v) => el.setAttributeNS(W, 'w:val', v);
  function firstKid(parent, name) {
    let el = OF.kid(parent, name, W);
    if (!el) { el = parent.ownerDocument.createElementNS(W, 'w:' + name); parent.insertBefore(el, parent.firstChild); }
    return el;
  }
  function applyRunFmt(r, f) {
    const rPr = firstKid(r, 'rPr');
    const set = (name, v) => { const e = OF.ordered(rPr, W, 'w', name, RPR); for (const a of Array.from(e.attributes)) e.removeAttributeNode(a); setVal(e, v); };
    if ('b' in f) { set('b', f.b ? '1' : '0'); set('bCs', f.b ? '1' : '0'); }
    if ('i' in f) { set('i', f.i ? '1' : '0'); set('iCs', f.i ? '1' : '0'); }
    if ('u' in f) set('u', f.u ? 'single' : 'none');
    if ('color' in f) set('color', f.color ? f.color.replace('#', '').toUpperCase() : 'auto');
    if ('sz' in f) { set('sz', String(Math.round(f.sz * 2))); set('szCs', String(Math.round(f.sz * 2))); }
    if ('hl' in f) { if (f.hl) set('highlight', f.hl); else OF.removeKids(rPr, ['highlight'], W); }
  }
  function applyParaFmt(p, f) {
    const pPr = firstKid(p, 'pPr');
    if (f.jc) { const e = OF.ordered(pPr, W, 'w', 'jc', PPR); setVal(e, f.jc); }
  }
  function stripIds(el) {
    for (const n of [el, ...el.getElementsByTagName('*')]) for (const a of Array.from(n.attributes)) if (/^(paraId|textId)$/.test(a.localName)) n.removeAttributeNode(a);
  }

  Object.assign(WordDoc.prototype, {
    /* ---- formato de texto ---- */
    format(fmt, css) {
      const range = OF.currentRange();
      if (!range || range.collapsed) { OF.toast('Selecciona el texto que quieres cambiar.'); return; }
      if (OF.formatSelection(range, fmt, css)) this.markDirty();
    },
    toggle(prop) {
      const range = OF.currentRange();
      if (!range || range.collapsed) { OF.toast('Selecciona el texto que quieres cambiar.'); return; }
      const tests = { b: (cs) => +cs.fontWeight >= 600, i: (cs) => cs.fontStyle === 'italic', u: (cs) => /underline/.test(cs.textDecorationLine) };
      const on = !OF.selectionHas(range, tests[prop]);
      const css = { b: { fontWeight: on ? '700' : '400' }, i: { fontStyle: on ? 'italic' : 'normal' }, u: { textDecoration: on ? 'underline' : 'none' } }[prop];
      this.format({ [prop]: on }, css);
    },
    setColor(hex) { this.format({ color: hex }, { color: hex || '#000' }); },
    setHighlight(name, css) { this.format({ hl: name }, { backgroundColor: css || 'transparent' }); },
    setSize(pt) { this.format({ sz: pt }, { fontSize: pt + 'pt', minHeight: '' }); },
    align(a) {
      const range = OF.currentRange();
      if (!range) { OF.toast('Toca el párrafo que quieres alinear.'); return; }
      const jc = { left: 'left', center: 'center', right: 'right', justify: 'both' }[a];
      for (const p of OF.parasIn(range)) {
        p.dataset.pfmt = JSON.stringify({ ...(p.dataset.pfmt ? JSON.parse(p.dataset.pfmt) : {}), jc });
        p.style.textAlign = a;
      }
      this.markDirty();
    },

    /* ---- cambios de estructura: se vuelcan las ediciones, se modifica el XML y se vuelve a dibujar ---- */
    caretInfo() {
      const range = OF.currentRange();
      const node = range && (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement);
      const dp = node && node.closest && node.closest('[data-p]');
      if (!dp) return null;
      const pid = +dp.dataset.p;
      const same = Array.from(this.pages.querySelectorAll(`article [data-p="${pid}"]`));
      return { pid, gi: Math.max(0, same.indexOf(dp)) };
    },
    async structural(fn) {
      const info = this.caretInfo();
      const dps = Array.from(this.pages.querySelectorAll('section.docx > article p'));
      this.sync.apply(dps);
      const res = info && this.sync.result.get(info.pid);
      const target = res ? res[Math.min(info.gi, res.length - 1)] || res[0] : null;
      const ok = await fn(target);
      if (ok === false) { await this.render(); return; }
      this.dirty = false;
      this.markDirty();
      await this.render();
    },
    async insertImage(file) {
      const im = await OF.prepareImage(file);
      await this.structural(async (target) => {
        const body = OF.kid(this.xml.documentElement, 'body', W);
        const media = OF.uniquePart(this.zip, 'word/media', 'imagen', im.ext);
        this.zip.file(media, im.bytes);
        const rId = await OF.addRel(this.zip, this.docPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', media);
        await OF.contentTypes(this.zip, (ct) => OF.ensureDefaultCT(ct, im.ext, im.mime));
        // ancho útil de la página
        const sect = OF.all(body, 'sectPr').pop();
        const pg = sect && OF.kid(sect, 'pgSz', W), mar = sect && OF.kid(sect, 'pgMar', W);
        const tw = (el, a, d) => (el && el.getAttributeNS(W, a) ? +el.getAttributeNS(W, a) : d);
        let avail = (tw(pg, 'w', 12240) - tw(mar, 'left', 1440) - tw(mar, 'right', 1440)) * 635;
        const inCell = target && target.parentNode && target.parentNode.localName === 'tc';
        if (inCell) avail = Math.min(avail, 2286000);
        const cx = Math.round(Math.min(avail, im.w * 9525)), cy = Math.round(cx * im.h / im.w);
        const ids = OF.all(this.xml, 'docPr').map((d) => +d.getAttribute('id') || 0);
        const id = Math.max(0, ...ids) + 1;
        const xml = `<w:p xmlns:w="${W}" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="${OF.NS.a}" xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" xmlns:r="${OF.NS.r}"><w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0"><wp:extent cx="${cx}" cy="${cy}"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="${id}" name="Imagen ${id}"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="${id}" name="${media.split('/').pop()}"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>`;
        const np = OF.fragment(this.xml, xml);
        if (target && target.parentNode) target.parentNode.insertBefore(np, target.nextSibling);
        else { const last = body.lastElementChild; body.insertBefore(np, last && last.localName === 'sectPr' ? last : null); }
      });
    },
    async imageAction(i, a) {
      const dr = this.drawings[i];
      if (!dr) return;
      OF.closeImageBar();
      await this.structural(() => {
        if (a === 'del') { let r = dr; while (r && r.localName !== 'r') r = r.parentNode; (r || dr).parentNode.removeChild(r || dr); return; }
        const body = OF.kid(this.xml.documentElement, 'body', W);
        const sect = OF.all(body, 'sectPr').pop();
        const pg = sect && OF.kid(sect, 'pgSz', W), mar = sect && OF.kid(sect, 'pgMar', W);
        const tw = (el, at, d) => (el && el.getAttributeNS(W, at) ? +el.getAttributeNS(W, at) : d);
        const maxW = (tw(pg, 'w', 12240) - tw(mar, 'left', 1440) - tw(mar, 'right', 1440)) * 635;
        const exts = Array.from(dr.getElementsByTagName('*')).filter((e) => (e.localName === 'extent' || e.localName === 'ext') && e.hasAttribute('cx'));
        const ext0 = exts.find((e) => e.localName === 'extent') || exts[0];
        if (!ext0) return false;
        let f = a;
        const cx0 = +ext0.getAttribute('cx');
        if (cx0 * f > maxW) f = maxW / cx0;
        if (cx0 * f < 95250) f = 95250 / cx0;
        for (const e of exts) { e.setAttribute('cx', Math.round(+e.getAttribute('cx') * f)); e.setAttribute('cy', Math.round(+e.getAttribute('cy') * f)); }
      });
    },
    async tableRow(op) {
      await this.structural((target) => {
        let tr = target;
        while (tr && tr.localName !== 'tr') tr = tr.parentNode;
        if (!tr) { OF.toast('Toca una celda de la tabla primero.'); return false; }
        if (op === 'del') {
          const tbl = tr.parentNode;
          if (OF.kids(tbl, 'tr', W).length > 1) tbl.removeChild(tr);
          else {
            const parent = tbl.parentNode;
            parent.removeChild(tbl);
            if (parent.localName === 'tc' && !OF.kids(parent, 'p', W).length) parent.appendChild(this.xml.createElementNS(W, 'w:p'));
          }
          return;
        }
        const nr = tr.cloneNode(true);
        stripIds(nr);
        for (const tc of OF.kids(nr, 'tc', W)) {
          const ps = OF.kids(tc, 'p', W);
          for (const el of Array.from(tc.children)) if (el.localName !== 'tcPr' && el !== ps[0]) tc.removeChild(el);
          const p = ps[0] || tc.appendChild(this.xml.createElementNS(W, 'w:p'));
          for (const el of Array.from(p.children)) if (el.localName !== 'pPr') p.removeChild(el);
          // mantener el formato del texto de la celda para lo que se escriba
          const r0 = OF.all(OF.kids(tr, 'tc', W)[OF.kids(nr, 'tc', W).indexOf(tc)], 'r', W).find((r) => OF.kid(r, 'rPr', W));
          if (r0) { const pPr = firstKid(p, 'pPr'); OF.removeKids(pPr, ['rPr'], W); const rp = OF.kid(r0, 'rPr', W).cloneNode(true); OF.ordered(pPr, W, 'w', 'rPr', PPR); pPr.replaceChild(rp, OF.kid(pPr, 'rPr', W)); }
        }
        tr.parentNode.insertBefore(nr, op === 'above' ? tr : tr.nextSibling);
      });
    },
  });

  WordDoc.mime = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  OF.WordDoc = WordDoc;
})();
