/* Word: estilos, listas, sangría, interlineado, tablas nuevas, saltos de página, fuentes,
   estado de formato, buscar/reemplazar, deshacer (instantáneas) y barra de estado. */
(function () {
  const OF = window.OF;
  const W = OF.NS.w;
  const WD = OF.WordDoc.prototype;
  const PPR = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr', 'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap', 'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd', 'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc', 'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr', 'sectPr', 'pPrChange'];
  const wv = (el, a) => (el ? el.getAttributeNS(W, a) || el.getAttribute('w:' + a) : null);
  const setW = (el, a, v) => el.setAttributeNS(W, 'w:' + a, v);
  const pPrOf = (p) => { let e = OF.kid(p, 'pPr', W); if (!e) { e = p.ownerDocument.createElementNS(W, 'w:pPr'); p.insertBefore(e, p.firstChild); } return e; };
  const wfrag = (doc, xml) => OF.fragment(doc, xml.replace(/^<(w:\w+)/, `<$1 xmlns:w="${W}"`));

  const STYLE_LABEL = { normal: 'Normal', title: 'Título', subtitle: 'Subtítulo', 'heading 1': 'Título 1', 'heading 2': 'Título 2', 'heading 3': 'Título 3', 'heading 4': 'Título 4', quote: 'Cita', 'intense quote': 'Cita destacada', 'list paragraph': 'Párrafo de lista', 'no spacing': 'Sin espaciado', 'list bullet': 'Lista con viñetas', 'list number': 'Lista numerada', 'list bullet 2': 'Lista con viñetas 2', 'list number 2': 'Lista numerada 2', caption: 'Descripción', header: 'Encabezado', footer: 'Pie de página', 'intense emphasis': 'Énfasis intenso', 'body text': 'Texto independiente', 'table grid': 'Tabla con cuadrícula', 'título 1': 'Título 1', 'título 2': 'Título 2', 'título 3': 'Título 3', 'título': 'Título' };
  const STYLE_DEFS = {
    title: { id: 'Title', ppr: '<w:spacing w:after="80" w:line="240" w:lineRule="auto"/><w:contextualSpacing/>', rpr: '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:spacing w:val="-10"/><w:kern w:val="28"/><w:sz w:val="56"/><w:szCs w:val="56"/>' },
    subtitle: { id: 'Subtitle', ppr: '<w:spacing w:after="160"/>', rpr: '<w:color w:val="595959"/><w:spacing w:val="15"/><w:sz w:val="28"/><w:szCs w:val="28"/>' },
    'heading 1': { id: 'Heading1', ppr: '<w:keepNext/><w:keepLines/><w:spacing w:before="360" w:after="80"/><w:outlineLvl w:val="0"/>', rpr: '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:color w:val="2F5496"/><w:sz w:val="40"/><w:szCs w:val="40"/>' },
    'heading 2': { id: 'Heading2', ppr: '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="1"/>', rpr: '<w:rFonts w:asciiTheme="majorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:color w:val="2F5496"/><w:sz w:val="32"/><w:szCs w:val="32"/>' },
    'heading 3': { id: 'Heading3', ppr: '<w:keepNext/><w:keepLines/><w:spacing w:before="160" w:after="80"/><w:outlineLvl w:val="2"/>', rpr: '<w:rFonts w:eastAsiaTheme="majorEastAsia" w:cstheme="majorBidi"/><w:color w:val="1F3763"/><w:sz w:val="28"/><w:szCs w:val="28"/>' },
    quote: { id: 'Quote', ppr: '<w:spacing w:before="160"/><w:ind w:left="864" w:right="864"/><w:jc w:val="center"/>', rpr: '<w:i/><w:iCs/><w:color w:val="404040"/>' },
  };
  const MENU_STYLES = ['normal', 'title', 'subtitle', 'heading 1', 'heading 2', 'heading 3', 'quote'];
  const COMMON_FONTS = ['Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Times New Roman', 'Cambria', 'Georgia', 'Garamond', 'Verdana', 'Tahoma', 'Segoe UI', 'Century Gothic', 'Trebuchet MS', 'Courier New'];

  /* ------------------------------ apertura ------------------------------ */
  const baseOpen = WD.open;
  WD.open = async function (buffer, host) {
    this.zip = null;
    await baseOpen.call(this, buffer, host);
    await this.afterLoad();
  };
  WD.afterLoad = async function () {
    this._num = null;
    this.rels = await OF.readRels(this.zip, this.docPath);
    const st = Object.values(this.rels).find((r) => /\/styles$/.test(r.type));
    this.stylesPath = st && this.zip.file(st.target) ? st.target : null;
    this.stylesDoc = this.stylesPath ? OF.parseXml(await this.zip.file(this.stylesPath).async('string')) : null;
    this.readStyles();
    const ft = Object.values(this.rels).find((r) => /fontTable$/.test(r.type));
    const docFonts = [];
    if (ft && this.zip.file(ft.target)) for (const f of OF.all(OF.parseXml(await this.zip.file(ft.target).async('string')), 'font')) { const n = wv(f, 'name'); if (n && !/^(Symbol|Wingdings|Times New Roman)$/.test(n) && !/[^\x20-\x7eáéíóúñ]/i.test(n)) docFonts.push(n); }
    this.fonts = Array.from(new Set([...docFonts, ...COMMON_FONTS]));
  };
  WD.readStyles = function () {
    this.styles = {}; this.styleByName = {};
    if (!this.stylesDoc) return;
    for (const s of OF.all(this.stylesDoc, 'style')) {
      const id = wv(s, 'styleId'), type = wv(s, 'type');
      const name = (wv(OF.kid(s, 'name', W), 'val') || id || '').toLowerCase();
      this.styles[id] = { id, type, name, el: s, isDefault: wv(s, 'default') === '1' || wv(s, 'default') === 'true' };
      if (type === 'paragraph') this.styleByName[name] = id;
      if (type === 'paragraph' && this.styles[id].isDefault) this.normalId = id;
    }
  };
  WD.ensureStyle = function (name) {
    if (this.styleByName[name]) return this.styleByName[name];
    if (!this.stylesDoc) return null;
    const def = STYLE_DEFS[name];
    if (!def) return null;
    let id = def.id; while (this.styles[id]) id += 'F';
    const based = this.normalId ? `<w:basedOn w:val="${this.normalId}"/><w:next w:val="${this.normalId}"/>` : '';
    const el = wfrag(this.stylesDoc, `<w:style w:type="paragraph" w:styleId="${id}"><w:name w:val="${name}"/>${based}<w:uiPriority w:val="9"/><w:qFormat/><w:pPr>${def.ppr}</w:pPr><w:rPr>${def.rpr}</w:rPr></w:style>`);
    this.stylesDoc.documentElement.appendChild(el);
    this.zip.file(this.stylesPath, OF.serializeXml(this.stylesDoc));
    this.readStyles();
    return id;
  };
  WD.styleMenu = function () {
    const out = MENU_STYLES.map((n) => ({ key: n, label: STYLE_LABEL[n] }));
    for (const s of Object.values(this.styles)) if (s.type === 'paragraph' && !MENU_STYLES.includes(s.name) && this.usedStyles && this.usedStyles.has(s.id)) out.push({ key: s.name, label: STYLE_LABEL[s.name] || s.name.replace(/^./, (c) => c.toUpperCase()) });
    return out;
  };

  /* ------------------- cambios de estructura con varios párrafos ------------------- */
  WD.selectionTargets = function () {
    const range = OF.currentRange();
    if (!range || !this.pages.contains(range.startContainer)) return [];
    const dps = OF.parasIn(range).filter((d) => d.closest('section.docx > article'));
    return dps.map((dp) => { const pid = +dp.dataset.p; const same = Array.from(this.pages.querySelectorAll(`article [data-p="${pid}"]`)); return { pid, gi: Math.max(0, same.indexOf(dp)) }; });
  };
  WD.structural = async function (fn) {
    OF.checkpoint && OF.checkpoint();
    const infos = this.selectionTargets();
    const dps = Array.from(this.pages.querySelectorAll('section.docx > article p'));
    this.sync.apply(dps);
    const targets = [];
    for (const i of infos) { const res = this.sync.result.get(i.pid); const t = res ? res[Math.min(i.gi, res.length - 1)] || res[0] : null; if (t && !targets.includes(t)) targets.push(t); }
    this.caretTarget = targets[0] || null;
    const ok = await fn(targets[0] || null, targets);
    if (ok === false) { await this.render(); return; }
    this.dirty = false;
    this.markDirty();
    await this.render();
  };
  const baseRender = WD.render;
  WD.render = async function () {
    await baseRender.call(this);
    this.usedStyles = new Set(OF.all(this.xml, 'pStyle').map((e) => wv(e, 'val')));
    this.afterRender();
  };
  WD.afterRender = function () {
    const arts = Array.from(this.pages.querySelectorAll('section.docx > article'));
    const empty = arts.every((a) => !a.textContent.trim()) && !this.pages.querySelector('article img, article table');
    arts.forEach((a, i) => {
      a.classList.toggle('empty-doc', empty && i === 0);
      a.addEventListener('input', () => a.classList.toggle('empty-doc', empty && i === 0 && !a.textContent.trim()), { passive: true });
      a.addEventListener('beforeinput', (e) => OF.typingCheckpoint && OF.typingCheckpoint(e), { passive: true });
    });
    // devolver el cursor al párrafo trabajado
    let dp = null;
    if (this.caretTarget) { const pid = this.sync.paras.indexOf(this.caretTarget); if (pid >= 0) dp = this.pages.querySelector(`article [data-p="${pid}"]`); }
    if (!dp && empty) dp = this.pages.querySelector('article [data-p]');
    if (dp) {
      const art = dp.closest('[contenteditable=true]');
      const r = document.createRange();
      const spans = dp.querySelectorAll('[data-r]');
      if (spans.length) { const last = spans[spans.length - 1]; r.selectNodeContents(last); r.collapse(false); } else { r.selectNodeContents(dp); r.collapse(true); }
      try { art.focus({ preventScroll: true }); const s = getSelection(); s.removeAllRanges(); s.addRange(r); OF.lastRange = r.cloneRange(); } catch (e) {}
      if (this.caretTarget) { const rect = dp.getBoundingClientRect(), h = this.host.getBoundingClientRect(); if (rect.top < h.top || rect.bottom > h.bottom) dp.scrollIntoView({ block: 'center' }); }
    }
    this.caretTarget = null;
  };

  /* ------------------------------- formato ------------------------------- */
  WD.setFont = function (name) { this.format({ font: name }, { fontFamily: `"${name}", Calibri, Carlito, Arial, sans-serif` }); };
  WD.setStyle = async function (name) {
    await this.structural((t, targets) => {
      if (!targets.length) { OF.toast('Toca el párrafo al que quieres dar estilo.'); return false; }
      const id = name === 'normal' ? null : this.ensureStyle(name);
      if (name !== 'normal' && !id) { OF.toast('Este documento no permite agregar estilos.'); return false; }
      for (const p of targets) {
        const pPr = pPrOf(p);
        OF.removeKids(pPr, ['pStyle'], W);
        if (id) setW(OF.ordered(pPr, W, 'w', 'pStyle', PPR), 'val', id);
      }
    });
  };
  WD.lineSpacing = async function (mult) {
    await this.structural((t, targets) => {
      if (!targets.length) return false;
      for (const p of targets) { const sp = OF.ordered(pPrOf(p), W, 'w', 'spacing', PPR); setW(sp, 'line', Math.round(240 * mult)); setW(sp, 'lineRule', 'auto'); }
    });
  };
  WD.paraSpacing = async function (on) {
    await this.structural((t, targets) => {
      if (!targets.length) return false;
      for (const p of targets) { const sp = OF.ordered(pPrOf(p), W, 'w', 'spacing', PPR); setW(sp, 'before', on ? 120 : 0); setW(sp, 'after', on ? 160 : 0); }
    });
  };

  /* ------------------------------- listas ------------------------------- */
  WD.numbering = async function () {
    if (this._num) return this._num;
    const rel = Object.values(this.rels).find((r) => /\/numbering$/.test(r.type));
    let path, doc;
    if (rel && this.zip.file(rel.target)) { path = rel.target; doc = OF.parseXml(await this.zip.file(path).async('string')); }
    else {
      path = this.zip.file('word/numbering.xml') ? OF.uniquePart(this.zip, 'word', 'numbering', 'xml') : 'word/numbering.xml';
      doc = OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:numbering xmlns:w="${W}"/>`);
      this.zip.file(path, OF.serializeXml(doc));
      await OF.addRel(this.zip, this.docPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering', path);
      await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, '/' === path[0] ? path.slice(1) : path, 'application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml'));
      this.rels = await OF.readRels(this.zip, this.docPath);
    }
    return (this._num = { path, doc });
  };
  WD.listNumId = async function (kind) {
    const { path, doc } = await this.numbering();
    const nsid = kind === 'bullet' ? 'F01A0B01' : 'F01A0D01';
    const root = doc.documentElement;
    let abs = OF.all(doc, 'abstractNum').find((a) => wv(OF.kid(a, 'nsid', W), 'val') === nsid);
    if (!abs) {
      const absId = Math.max(-1, ...OF.all(doc, 'abstractNum').map((a) => +wv(a, 'abstractNumId'))) + 1;
      const chars = ['•', '◦', '▪'], fmts = ['decimal', 'lowerLetter', 'lowerRoman'];
      let lv = '';
      for (let i = 0; i < 9; i++) {
        lv += kind === 'bullet'
          ? `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${chars[i % 3]}"/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr><w:rPr><w:rFonts w:ascii="Arial" w:hAnsi="Arial" w:hint="default"/></w:rPr></w:lvl>`
          : `<w:lvl w:ilvl="${i}"><w:start w:val="1"/><w:numFmt w:val="${fmts[i % 3]}"/><w:lvlText w:val="%${i + 1}."/><w:lvlJc w:val="left"/><w:pPr><w:ind w:left="${720 * (i + 1)}" w:hanging="360"/></w:pPr></w:lvl>`;
      }
      abs = wfrag(doc, `<w:abstractNum w:abstractNumId="${absId}"><w:nsid w:val="${nsid}"/><w:multiLevelType w:val="hybridMultilevel"/>${lv}</w:abstractNum>`);
      const firstNum = OF.kid(root, 'num', W);
      root.insertBefore(abs, firstNum || OF.kid(root, 'numIdMacAtCleanup', W));
    }
    const absId = wv(abs, 'abstractNumId');
    let num = OF.all(doc, 'num').find((n) => wv(OF.kid(n, 'abstractNumId', W), 'val') === absId && !OF.kid(n, 'lvlOverride', W));
    if (!num) {
      const numId = Math.max(0, ...OF.all(doc, 'num').map((n) => +wv(n, 'numId'))) + 1;
      num = wfrag(doc, `<w:num w:numId="${numId}"><w:abstractNumId w:val="${absId}"/></w:num>`);
      root.insertBefore(num, OF.kid(root, 'numIdMacAtCleanup', W));
    }
    this.zip.file(path, OF.serializeXml(doc));
    return wv(num, 'numId');
  };
  WD.toggleList = async function (kind) {
    const numId = await this.listNumId(kind);
    await this.structural((t, targets) => {
      if (!targets.length) { OF.toast('Toca el párrafo donde quieres la lista.'); return false; }
      const cur = (p) => wv(OF.kid(OF.kid(OF.kid(p, 'pPr', W), 'numPr', W), 'numId', W), 'val');
      const off = targets.every((p) => cur(p) === numId);
      for (const p of targets) {
        const pPr = pPrOf(p);
        const old = OF.kid(pPr, 'numPr', W);
        const ilvl = old ? wv(OF.kid(old, 'ilvl', W), 'val') || '0' : '0';
        OF.removeKids(pPr, ['numPr'], W);
        if (off) {
          const ps = wv(OF.kid(pPr, 'pStyle', W), 'val');
          if (ps && /list|lista/i.test((this.styles[ps] || {}).name || ps)) { const np = OF.ordered(pPr, W, 'w', 'numPr', PPR); np.appendChild(wfrag(p.ownerDocument, '<w:ilvl w:val="0"/>')); np.appendChild(wfrag(p.ownerDocument, '<w:numId w:val="0"/>')); }
        } else {
          const np = OF.ordered(pPr, W, 'w', 'numPr', PPR);
          np.appendChild(wfrag(p.ownerDocument, `<w:ilvl w:val="${ilvl}"/>`));
          np.appendChild(wfrag(p.ownerDocument, `<w:numId w:val="${numId}"/>`));
          const ind = OF.kid(pPr, 'ind', W);
          if (ind) pPr.removeChild(ind);
        }
      }
    });
  };
  WD.indent = async function (dir) {
    await this.structural((t, targets) => {
      if (!targets.length) return false;
      for (const p of targets) {
        const pPr = pPrOf(p);
        const np = OF.kid(pPr, 'numPr', W);
        const numId = np && wv(OF.kid(np, 'numId', W), 'val');
        if (numId && numId !== '0') {
          let il = OF.kid(np, 'ilvl', W);
          if (!il) { il = wfrag(p.ownerDocument, '<w:ilvl w:val="0"/>'); np.insertBefore(il, np.firstChild); }
          setW(il, 'val', Math.max(0, Math.min(8, +(wv(il, 'val') || 0) + dir)));
        } else {
          const ind = OF.ordered(pPr, W, 'w', 'ind', PPR);
          const attr = wv(ind, 'start') != null ? 'start' : 'left';
          setW(ind, attr, Math.max(0, +(wv(ind, attr) || 0) + dir * 720));
        }
      }
    });
  };

  /* ------------------------------- insertar ------------------------------- */
  const bodyOf = (self) => OF.kid(self.xml.documentElement, 'body', W);
  const topLevel = (self, p) => { const body = bodyOf(self); let n = p; while (n && n.parentNode !== body) n = n.parentNode; return n; };
  const lastBodyBefore = (self) => { const body = bodyOf(self); const last = body.lastElementChild; return last && last.localName === 'sectPr' ? last : null; };
  WD.textWidth = function () {
    const sect = OF.all(bodyOf(this), 'sectPr').pop();
    const pg = sect && OF.kid(sect, 'pgSz', W), mar = sect && OF.kid(sect, 'pgMar', W);
    const tw = (el, at, d) => (el && wv(el, at) ? +wv(el, at) : d);
    return tw(pg, 'w', 12240) - tw(mar, 'left', 1440) - tw(mar, 'right', 1440);
  };
  WD.insertTable = async function (rows, cols) {
    await this.structural((t) => {
      const doc = this.xml, body = bodyOf(this);
      const width = this.textWidth();
      const cw = Math.floor(width / cols);
      const bd = ['top', 'left', 'bottom', 'right', 'insideH', 'insideV'].map((s) => `<w:${s} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`).join('');
      const tg = Object.values(this.styles).find((s) => s.type === 'table' && s.name === 'table grid');
      let xml = `<w:tbl><w:tblPr>${tg ? `<w:tblStyle w:val="${tg.id}"/>` : ''}<w:tblW w:w="5000" w:type="pct"/><w:tblBorders>${bd}</w:tblBorders><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0" w:firstColumn="1" w:lastColumn="0" w:noHBand="0" w:noVBand="1"/></w:tblPr><w:tblGrid>`;
      for (let c = 0; c < cols; c++) xml += `<w:gridCol w:w="${cw}"/>`;
      xml += '</w:tblGrid>';
      for (let r = 0; r < rows; r++) { xml += '<w:tr>'; for (let c = 0; c < cols; c++) xml += `<w:tc><w:tcPr><w:tcW w:w="${cw}" w:type="dxa"/></w:tcPr><w:p/></w:tc>`; xml += '</w:tr>'; }
      xml += '</w:tbl>';
      const tbl = wfrag(doc, xml);
      const anchor = t ? topLevel(this, t) : null;
      if (anchor) body.insertBefore(tbl, anchor.nextSibling); else body.insertBefore(tbl, lastBodyBefore(this));
      const next = tbl.nextElementSibling;
      if (!next || next.localName !== 'p') body.insertBefore(doc.createElementNS(W, 'w:p'), next);
      this.caretTarget = OF.all(tbl, 'p')[0];
    });
    this.caretTarget = null;
  };
  WD.pageBreak = async function () {
    await this.structural((t) => {
      const doc = this.xml, body = bodyOf(this);
      const br = wfrag(doc, '<w:p><w:r><w:br w:type="page"/></w:r></w:p>');
      const after = doc.createElementNS(W, 'w:p');
      const anchor = t ? topLevel(this, t) : null;
      const ref = anchor ? anchor.nextSibling : lastBodyBefore(this);
      body.insertBefore(br, ref); body.insertBefore(after, ref);
      this.caretTarget = after;
    });
  };
  WD.tableCol = async function (op) {
    await this.structural((t) => {
      let tc = t; while (tc && tc.localName !== 'tc') tc = tc.parentNode;
      if (!tc) { OF.toast('Toca una celda de la tabla primero.'); return false; }
      const tr = tc.parentNode, tbl = tr.parentNode;
      const idx = OF.kids(tr, 'tc', W).indexOf(tc);
      const grid = OF.kid(tbl, 'tblGrid', W);
      for (const row of OF.kids(tbl, 'tr', W)) {
        const cells = OF.kids(row, 'tc', W);
        const c = cells[Math.min(idx, cells.length - 1)];
        if (!c) continue;
        if (op === 'del') { if (cells.length > 1) row.removeChild(c); continue; }
        const nc = c.cloneNode(true);
        for (const el of Array.from(nc.children)) if (el.localName !== 'tcPr') nc.removeChild(el);
        nc.appendChild(this.xml.createElementNS(W, 'w:p'));
        row.insertBefore(nc, op === 'left' ? c : c.nextSibling);
      }
      if (grid) {
        const gc = OF.kids(grid, 'gridCol', W);
        if (op === 'del') { if (gc[idx] && gc.length > 1) grid.removeChild(gc[idx]); }
        else if (gc[idx]) grid.insertBefore(gc[idx].cloneNode(true), op === 'left' ? gc[idx] : gc[idx].nextSibling);
        // repartir el ancho disponible
        const cols = OF.kids(grid, 'gridCol', W), w = Math.floor(this.textWidth() / cols.length);
        cols.forEach((g) => setW(g, 'w', w));
        for (const row of OF.kids(tbl, 'tr', W)) for (const c of OF.kids(row, 'tc', W)) { const tw = OF.kid(OF.kid(c, 'tcPr', W), 'tcW', W); if (tw && wv(tw, 'type') === 'dxa' && !wv(OF.kid(c, 'tcPr', W) && OF.kid(OF.kid(c, 'tcPr', W), 'gridSpan', W), 'val')) setW(tw, 'w', w); }
      }
      if (!OF.kids(tr, 'tc', W).length) tbl.parentNode.removeChild(tbl);
    });
  };
  WD.deleteTable = async function () {
    await this.structural((t) => {
      let tbl = t; while (tbl && tbl.localName !== 'tbl') tbl = tbl.parentNode;
      if (!tbl) { OF.toast('Toca una celda de la tabla primero.'); return false; }
      const parent = tbl.parentNode;
      const p = this.xml.createElementNS(W, 'w:p');
      parent.replaceChild(p, tbl);
      this.caretTarget = p;
    });
  };

  /* --------------------------- estado del formato --------------------------- */
  WD.formatState = function () {
    const r = OF.currentRange();
    if (!r || !this.pages || !this.pages.contains(r.startContainer)) return null;
    let el = r.startContainer.nodeType === 3 ? r.startContainer.parentElement : r.startContainer;
    if (r.startContainer.nodeType === 1 && r.startContainer.childNodes[r.startOffset - 1]) { const c = r.startContainer.childNodes[r.startOffset - 1]; if (c.nodeType === 1) el = c; }
    const dp = el.closest('[data-p]');
    const cs = getComputedStyle(el);
    let u = false;
    for (let a = el; a && a !== dp; a = a.parentElement) if (/underline/.test(getComputedStyle(a).textDecorationLine)) u = true;
    const fam = cs.fontFamily.split(',')[0].replace(/["']/g, '').trim();
    const st = { b: +cs.fontWeight >= 600, i: cs.fontStyle === 'italic', u, font: fam, size: Math.round(parseFloat(cs.fontSize) * 0.75 * 2) / 2 };
    if (dp) {
      const ta = getComputedStyle(dp).textAlign;
      st.align = /center/.test(ta) ? 'center' : /right|end/.test(ta) ? 'right' : /justify/.test(ta) ? 'justify' : 'left';
      const xp = this.sync.paras[+dp.dataset.p];
      const sid = xp && wv(OF.kid(OF.kid(xp, 'pPr', W), 'pStyle', W), 'val');
      const s = sid ? this.styles[sid] : null;
      st.style = s ? STYLE_LABEL[s.name] || s.name.replace(/^./, (c) => c.toUpperCase()) : 'Normal';
      const np = xp && OF.kid(OF.kid(xp, 'pPr', W), 'numPr', W);
      st.list = !!(np && wv(OF.kid(np, 'numId', W), 'val') !== '0') || /docx-num-/.test(dp.className);
    }
    return st;
  };
  WD.statusText = function () {
    const secs = Array.from(this.pages.querySelectorAll('section.docx'));
    const text = Array.from(this.pages.querySelectorAll('section.docx > article')).map((a) => a.innerText).join(' ');
    const words = (text.match(/[\p{L}\p{N}][\p{L}\p{N}'’.-]*/gu) || []).length;
    const r = OF.currentRange();
    let cur = 1;
    if (r) { const s = (r.startContainer.nodeType === 1 ? r.startContainer : r.startContainer.parentElement).closest('section.docx'); if (s) cur = secs.indexOf(s) + 1; }
    return `Página ${cur} de ${secs.length} · ${words.toLocaleString('es-CL')} palabra${words === 1 ? '' : 's'}`;
  };

  /* ------------------------- buscar y reemplazar ------------------------- */
  WD.searchRoots = function () { return Array.from(this.pages.querySelectorAll('section.docx > article')); };
  Object.assign(WD, OF.textSearchMixin);

  /* --------------------------- deshacer / rehacer --------------------------- */
  // La instantánea guarda las partes del paquete (sin copiarlas) y el texto editado en pantalla.
  WD.snapshot = function () {
    return { files: Object.assign({}, this.zip.files), html: this.searchRoots().map((a) => OF.cloneKids(a)), scroll: this.host.scrollTop };
  };
  WD.restore = async function (s) {
    const same = Object.keys(s.files).length === Object.keys(this.zip.files).length && Object.keys(s.files).every((k) => this.zip.files[k] === s.files[k]);
    if (!same) {
      const z = new JSZip();
      for (const k in s.files) z.files[k] = s.files[k];
      this.zip = z;
      this.xml = OF.parseXml(await z.file(this.docPath).async('string'));
      await this.afterLoad();
      await this.render();
    }
    const arts = this.searchRoots();
    if (arts.length === s.html.length) arts.forEach((a, i) => OF.setKids(a, s.html[i]));
    this.host.scrollTop = s.scroll;
    this.afterRestore && this.afterRestore();
  };

  /* -------------------------------- imprimir -------------------------------- */
  WD.printSetup = function () {
    const sec = this.pages.querySelector('section.docx');
    const wrap = this.pages.querySelector('.docx-wrapper');
    if (wrap) wrap.style.zoom = 1;
    return sec ? `@page { size: ${sec.style.width || '612pt'} ${sec.style.minHeight || '792pt'}; margin: 0 }` : '';
  };
})();
