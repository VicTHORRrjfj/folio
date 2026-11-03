/* Editor de Excel: dibuja las hojas con sus estilos reales y edita las celdas directamente
   en el XML original (estilos, gráficos, formatos condicionales, etc. se conservan). */
(function () {
  const OF = window.OF;
  const S = OF.NS.s;

  const INDEXED = ['000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF','800000','008000','000080','808000','800080','008080','C0C0C0','808080','9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF','000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF','00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99','3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696','003366','339966','003300','333300','993300','993366','333399','333333'];
  const BORDER = { thin: '1px solid', hair: '1px dotted', dotted: '1px dotted', dashed: '1px dashed', dashDot: '1px dashed', dashDotDot: '1px dashed', medium: '2px solid', mediumDashed: '2px dashed', mediumDashDot: '2px dashed', mediumDashDotDot: '2px dashed', slantDashDot: '2px dashed', thick: '3px solid', double: '3px double' };
  const BW = { thin: 1, hair: 1, dotted: 1, dashed: 1, dashDot: 1, dashDotDot: 1, medium: 2, mediumDashed: 2, mediumDashDot: 2, mediumDashDotDot: 2, slantDashDot: 2, thick: 3, double: 3 };

  const colName = (c) => { let s = ''; c++; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  const colIndex = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const parseRef = (ref) => { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref); return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null; };
  const addr = (r, c) => colName(c) + (r + 1);
  OF.colName = colName;

  function shiftFormula(f, dr, dc) {
    // mueve referencias relativas (para fórmulas compartidas)
    return f.split(/("(?:[^"]|"")*")/).map((seg, i) => {
      if (i % 2) return seg;
      return seg.replace(/(?<![A-Za-z0-9_.])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (m, d1, col, d2, row) => {
        const c = d1 ? colIndex(col) : colIndex(col) + dc;
        const r = d2 ? +row : +row + dr;
        if (c < 0 || r < 1) return '#REF!';
        return d1 + colName(c) + d2 + r;
      });
    }).join('');
  }

  function applyTint(hex, tint) {
    if (!tint) return hex;
    let r = parseInt(hex.slice(0, 2), 16) / 255, g = parseInt(hex.slice(2, 4), 16) / 255, b = parseInt(hex.slice(4, 6), 16) / 255;
    const max = Math.max(r, g, b), min = Math.min(r, g, b);
    let h = 0, s = 0, l = (max + min) / 2;
    if (max !== min) {
      const d = max - min;
      s = l > 0.5 ? d / (2 - max - min) : d / (max + min);
      h = max === r ? (g - b) / d + (g < b ? 6 : 0) : max === g ? (b - r) / d + 2 : (r - g) / d + 4;
      h /= 6;
    }
    l = tint < 0 ? l * (1 + tint) : l * (1 - tint) + tint;
    const hue = (p, q, t) => { if (t < 0) t += 1; if (t > 1) t -= 1; if (t < 1 / 6) return p + (q - p) * 6 * t; if (t < 1 / 2) return q; if (t < 2 / 3) return p + (q - p) * (2 / 3 - t) * 6; return p; };
    let R, G, B;
    if (s === 0) R = G = B = l;
    else { const q = l < 0.5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q; R = hue(p, q, h + 1 / 3); G = hue(p, q, h); B = hue(p, q, h - 1 / 3); }
    return [R, G, B].map((x) => Math.round(x * 255).toString(16).padStart(2, '0')).join('');
  }

  class ExcelDoc {
    constructor() {
      this.kind = 'excel';
      this.dirty = false;
      this.locale = 'cl';
      this.undo = [];
    }

    async open(buffer, host) {
      this.original = buffer;
      this.zip = await JSZip.loadAsync(buffer);
      const rels = await OF.readRels(this.zip, '');
      this.wbPath = (Object.values(rels).find((r) => /officeDocument$/.test(r.type)) || {}).target || 'xl/workbook.xml';
      this.wb = OF.parseXml(await this.zip.file(this.wbPath).async('string'));
      this.wbRels = await OF.readRels(this.zip, this.wbPath);
      this.date1904 = OF.attr(OF.kid(this.wb.documentElement, 'workbookPr', S), 'date1904') === '1';
      await this.loadTheme();
      await this.loadStrings();
      await this.loadStyles();
      this.sheets = [];
      for (const sh of OF.all(this.wb, 'sheet')) {
        const rid = sh.getAttributeNS(OF.NS.r, 'id') || sh.getAttribute('r:id');
        const rel = this.wbRels[rid];
        if (!rel || !this.zip.file(rel.target)) continue;
        this.sheets.push({ name: sh.getAttribute('name'), hidden: (sh.getAttribute('state') || 'visible') !== 'visible', path: rel.target });
      }
      for (const s of this.sheets) await this.loadSheet(s);
      this.buildEngine();
      this.host = host;
      this.buildUI();
      this.showSheet(this.sheets.findIndex((s) => !s.hidden));
    }

    async loadTheme() {
      this.theme = ['FFFFFF', '000000', 'E7E6E6', '44546A', '4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '0563C1', '954F72'];
      const rel = Object.values(this.wbRels).find((r) => /theme$/.test(r.type));
      if (!rel || !this.zip.file(rel.target)) return;
      const doc = OF.parseXml(await this.zip.file(rel.target).async('string'));
      const cs = OF.all(doc, 'clrScheme')[0];
      if (!cs) return;
      const get = (n) => { const e = OF.kid(cs, n); if (!e) return null; const c = e.firstElementChild; if (!c) return null; return c.localName === 'sysClr' ? c.getAttribute('lastClr') || (c.getAttribute('val') === 'window' ? 'FFFFFF' : '000000') : c.getAttribute('val'); };
      const order = ['lt1', 'dk1', 'lt2', 'dk2', 'accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'];
      order.forEach((n, i) => { const v = get(n); if (v) this.theme[i] = v.slice(-6); });
    }

    color(el) {
      if (!el) return null;
      let hex = null;
      if (el.hasAttribute('rgb')) hex = el.getAttribute('rgb').slice(-6);
      else if (el.hasAttribute('theme')) hex = this.theme[+el.getAttribute('theme')] || '000000';
      else if (el.hasAttribute('indexed')) { const i = +el.getAttribute('indexed'); if (i === 64 || i === 65) return null; hex = INDEXED[i] || '000000'; }
      else return null;
      return '#' + applyTint(hex, parseFloat(el.getAttribute('tint') || '0'));
    }

    async loadStrings() {
      this.ssRel = Object.values(this.wbRels).find((r) => /sharedStrings$/.test(r.type));
      this.strings = [];
      if (!this.ssRel || !this.zip.file(this.ssRel.target)) { this.ssDoc = null; return; }
      this.ssDoc = OF.parseXml(await this.zip.file(this.ssRel.target).async('string'));
      for (const si of OF.kids(this.ssDoc.documentElement, 'si', S)) this.strings.push(siText(si));
    }

    async loadStyles() {
      const rel = Object.values(this.wbRels).find((r) => /styles$/.test(r.type));
      this.numFmts = {};
      this.fonts = [{}]; this.fills = [{}]; this.borders = [{}]; this.xfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }];
      if (!rel || !this.zip.file(rel.target)) return;
      const doc = OF.parseXml(await this.zip.file(rel.target).async('string'));
      this.stylesDoc = doc; this.stylesPath = rel.target;
      const root = doc.documentElement;
      for (const n of OF.kids(OF.kid(root, 'numFmts', S), 'numFmt', S)) this.numFmts[n.getAttribute('numFmtId')] = n.getAttribute('formatCode');
      this.fonts = OF.kids(OF.kid(root, 'fonts', S), 'font', S).map((f) => this.parseFont(f));
      this.fills = OF.kids(OF.kid(root, 'fills', S), 'fill', S).map((f) => this.parseFill(f));
      this.borders = OF.kids(OF.kid(root, 'borders', S), 'border', S).map((b) => {
        const side = (n) => { const e = OF.kid(b, n, S) || (n === 'left' ? OF.kid(b, 'start', S) : n === 'right' ? OF.kid(b, 'end', S) : null); const st = e && e.getAttribute('style'); return st ? { style: st, color: this.color(OF.kid(e, 'color', S)) || '#000' } : null; };
        return { left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') };
      });
      this.xfs = OF.kids(OF.kid(root, 'cellXfs', S), 'xf', S).map((x) => this.parseXf(x));
      if (!this.xfs.length) this.xfs = [{ numFmtId: 0, fontId: 0, fillId: 0, borderId: 0 }];
    }

    parseFont(f) {
      const v = (n) => { const e = OF.kid(f, n, S); return e ? e.getAttribute('val') ?? true : null; };
      const on = (n) => { const e = OF.kid(f, n, S); return !!e && e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false'; };
      return { b: on('b'), i: on('i'), u: !!OF.kid(f, 'u', S) && OF.kid(f, 'u', S).getAttribute('val') !== 'none', strike: on('strike'), sz: parseFloat(v('sz')) || 11, name: v('name') || 'Calibri', color: this.color(OF.kid(f, 'color', S)), va: v('vertAlign') };
    }
    parseFill(f) {
      const pf = OF.kid(f, 'patternFill', S);
      if (pf && pf.getAttribute('patternType') && pf.getAttribute('patternType') !== 'none') {
        const fg = this.color(OF.kid(pf, 'fgColor', S));
        return { bg: pf.getAttribute('patternType') === 'solid' ? fg : fg || this.color(OF.kid(pf, 'bgColor', S)) };
      }
      const gf = OF.kid(f, 'gradientFill', S);
      if (gf) { const st = OF.kids(gf, 'stop', S).map((s) => this.color(OF.kid(s, 'color', S))).filter(Boolean); if (st.length) return { bg: st[0], grad: st }; }
      return {};
    }
    parseXf(x) {
      const al = OF.kid(x, 'alignment', S);
      return {
        numFmtId: +(x.getAttribute('numFmtId') || 0), fontId: +(x.getAttribute('fontId') || 0), fillId: +(x.getAttribute('fillId') || 0), borderId: +(x.getAttribute('borderId') || 0),
        h: al && al.getAttribute('horizontal'), v: al && al.getAttribute('vertical'), wrap: al && al.getAttribute('wrapText') === '1', indent: al ? +(al.getAttribute('indent') || 0) : 0, rot: al ? +(al.getAttribute('textRotation') || 0) : 0,
      };
    }

    fmtCode(xfIdx) {
      const id = (this.xfs[xfIdx] || this.xfs[0]).numFmtId;
      if (this.numFmts[id]) return this.numFmts[id];
      if (this.locale === 'cl') { const cl = { 14: 'dd-mm-yyyy', 15: 'd-mmm-yy', 16: 'd-mmm', 22: 'dd-mm-yyyy h:mm' }; if (cl[id]) return cl[id]; }
      const t = SSF.get_table();
      return t[id] || 'General';
    }

    async loadSheet(sh) {
      const doc = (sh.doc = OF.parseXml(await this.zip.file(sh.path).async('string')));
      const root = doc.documentElement;
      sh.rels = await OF.readRels(this.zip, sh.path);
      const fp = OF.kid(root, 'sheetFormatPr', S);
      sh.defW = fp && fp.getAttribute('defaultColWidth') ? +fp.getAttribute('defaultColWidth') : (fp && fp.getAttribute('baseColWidth') ? +fp.getAttribute('baseColWidth') + 0.71 : 8.43);
      sh.defH = fp && fp.getAttribute('defaultRowHeight') ? +fp.getAttribute('defaultRowHeight') : 15;
      const sv = OF.all(root, 'sheetView')[0];
      sh.grid = !sv || sv.getAttribute('showGridLines') !== '0';
      sh.cols = {};
      for (const c of OF.kids(OF.kid(root, 'cols', S), 'col', S)) {
        const min = +c.getAttribute('min') - 1, max = Math.min(+c.getAttribute('max') - 1, 16383);
        for (let i = min; i <= max && i < min + 1000; i++) sh.cols[i] = { w: c.getAttribute('width') ? +c.getAttribute('width') : null, hidden: c.getAttribute('hidden') === '1', s: c.getAttribute('style') ? +c.getAttribute('style') : null };
      }
      sh.data = OF.kid(root, 'sheetData', S);
      sh.rowEls = new Map();
      sh.rowInfo = {};
      sh.cells = new Map(); // "r,c" -> {el, s}
      let lastR = -1, maxR = -1, maxC = -1;
      for (const row of OF.kids(sh.data, 'row', S)) {
        const r = row.hasAttribute('r') ? +row.getAttribute('r') - 1 : lastR + 1;
        row.setAttribute('r', r + 1);
        lastR = r;
        sh.rowEls.set(r, row);
        sh.rowInfo[r] = { ht: row.getAttribute('ht') ? +row.getAttribute('ht') : null, hidden: row.getAttribute('hidden') === '1', s: row.getAttribute('customFormat') === '1' && row.getAttribute('s') ? +row.getAttribute('s') : null };
        let lastC = -1;
        for (const c of OF.kids(row, 'c', S)) {
          const p = c.hasAttribute('r') ? parseRef(c.getAttribute('r')) : { r, c: lastC + 1 };
          if (!c.hasAttribute('r')) c.setAttribute('r', addr(r, p.c));
          lastC = p.c;
          sh.cells.set(r + ',' + p.c, c);
          const hasContent = OF.kid(c, 'v', S) || OF.kid(c, 'f', S) || OF.kid(c, 'is', S);
          if (hasContent || (c.getAttribute('s') && c.getAttribute('s') !== '0')) { if (r > maxR) maxR = r; if (p.c > maxC) maxC = p.c; }
        }
      }
      sh.merges = OF.kids(OF.kid(root, 'mergeCells', S), 'mergeCell', S).map((m) => {
        const [a, b] = m.getAttribute('ref').split(':');
        const A = parseRef(a), B = parseRef(b || a);
        if (B.r > maxR) maxR = B.r; if (B.c > maxC) maxC = B.c;
        return { r1: A.r, c1: A.c, r2: B.r, c2: B.c };
      });
      sh.maxR = maxR; sh.maxC = maxC;
      // fórmulas compartidas: guardar maestro para poder traducirlas
      sh.shared = {};
      for (const c of sh.cells.values()) {
        const f = OF.kid(c, 'f', S);
        if (f && f.getAttribute('t') === 'shared' && f.getAttribute('ref') && f.textContent) sh.shared[f.getAttribute('si')] = { f: f.textContent, at: parseRef(c.getAttribute('r')) };
      }
    }

    // Contenido "crudo" de una celda en el XML
    raw(sh, r, c) {
      const el = sh.cells.get(r + ',' + c);
      if (!el) return { s: this.defaultStyle(sh, r, c) };
      const s = +(el.getAttribute('s') || 0);
      const t = el.getAttribute('t') || 'n';
      const fEl = OF.kid(el, 'f', S), vEl = OF.kid(el, 'v', S);
      let formula = null;
      if (fEl) {
        if (fEl.textContent) formula = fEl.textContent;
        else if (fEl.getAttribute('t') === 'shared' && sh.shared[fEl.getAttribute('si')]) {
          const m = sh.shared[fEl.getAttribute('si')];
          formula = shiftFormula(m.f, r - m.at.r, c - m.at.c);
        }
      }
      let v = null;
      if (t === 'inlineStr') { const is = OF.kid(el, 'is', S); v = is ? siText(is) : ''; }
      else if (vEl) {
        const x = vEl.textContent;
        if (t === 's') v = this.strings[+x] ?? '';
        else if (t === 'b') v = x === '1';
        else if (t === 'e') v = { error: x };
        else if (t === 'str') v = x;
        else v = x === '' ? null : +x;
      }
      return { s, v, f: formula, el };
    }
    defaultStyle(sh, r, c) {
      const ri = sh.rowInfo[r];
      if (ri && ri.s != null) return ri.s;
      const ci = sh.cols[c];
      return ci && ci.s != null ? ci.s : 0;
    }

    /* ---------------- motor de fórmulas (HyperFormula, licencia GPLv3) ---------------- */
    buildEngine() {
      this.hf = null;
      this.cacheOnly = new Set();
      if (!window.HyperFormula) return;
      try {
        const sheetsData = {};
        for (const sh of this.sheets) {
          const rows = [];
          for (const [key, el] of sh.cells) {
            const [r, c] = key.split(',').map(Number);
            const x = this.raw(sh, r, c);
            let val = null;
            if (x.f != null) val = '=' + x.f;
            else if (x.v != null && typeof x.v !== 'object') val = x.v;
            if (val === null) continue;
            (rows[r] = rows[r] || [])[c] = val;
          }
          for (let i = 0; i < rows.length; i++) if (!rows[i]) rows[i] = [];
          sheetsData[sh.name] = rows;
        }
        this.hf = HyperFormula.buildFromSheets(sheetsData, { licenseKey: 'gpl-v3', leapYear1900: true, smartRounding: true, useArrayArithmetic: true });
        for (const dn of OF.all(this.wb, 'definedName')) {
          if (dn.hasAttribute('localSheetId')) continue;
          try { this.hf.addNamedExpression(dn.getAttribute('name'), '=' + dn.textContent); } catch (e) {}
        }
        // fórmulas que el motor no calcula igual que Excel: se muestran con el valor guardado
        this.sheets.forEach((sh, si) => {
          sh.hfId = this.hf.getSheetId(sh.name);
          for (const [key] of sh.cells) {
            const [r, c] = key.split(',').map(Number);
            const x = this.raw(sh, r, c);
            if (x.f == null) continue;
            const hv = this.hf.getCellValue({ sheet: sh.hfId, row: r, col: c });
            if (hv && typeof hv === 'object' && 'value' in hv && x.v != null && !(typeof x.v === 'object')) this.cacheOnly.add(si + ':' + key);
          }
        });
      } catch (e) {
        console.warn('Motor de fórmulas no disponible:', e);
        this.hf = null;
      }
    }

    value(si, r, c) {
      const sh = this.sheets[si];
      const x = this.raw(sh, r, c);
      if (x.f != null && this.hf && !this.cacheOnly.has(si + ':' + r + ',' + c)) {
        const v = this.hf.getCellValue({ sheet: sh.hfId, row: r, col: c });
        if (v && typeof v === 'object' && 'value' in v) return { ...x, v: { error: v.value } };
        return { ...x, v };
      }
      return x;
    }

    display(x) {
      const v = x.v;
      if (v == null) return '';
      if (typeof v === 'object') return v.error;
      if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
      if (typeof v === 'string') return v;
      const code = this.fmtCode(x.s);
      let out;
      try { out = SSF.format(code, v, { date1904: this.date1904 }); } catch (e) { out = String(v); }
      if (this.locale === 'cl' && !SSF.is_date(code)) out = out.replace(/[.,]/g, (m) => (m === '.' ? ',' : '.'));
      return out;
    }

    /* ------------------------------- UI ------------------------------- */
    buildUI() {
      const host = this.host;
      host.innerHTML = '';
      host.classList.add('xl-host');
      this.fx = OF.el('input', { class: 'xl-fx', spellcheck: 'false', placeholder: 'Contenido de la celda' });
      this.nameBox = OF.el('div', { class: 'xl-name' }, 'A1');
      const bar = OF.el('div', { class: 'xl-bar' }, this.nameBox, OF.el('span', { class: 'xl-fxl' }, 'fx'), this.fx);
      this.gridBox = OF.el('div', { class: 'xl-grid', tabindex: '0' });
      this.tabs = OF.el('div', { class: 'xl-tabs' });
      host.append(bar, this.gridBox, this.tabs);
      this.styleEl = OF.el('style');
      host.append(this.styleEl);
      this.buildStyleSheet();
      this.sheets.forEach((s, i) => {
        if (s.hidden) return;
        this.tabs.append(OF.el('button', { class: 'xl-tab', 'data-i': i, onclick: () => this.showSheet(i) }, s.name));
      });

      this.fx.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') { e.preventDefault(); this.commit(this.fx.value); this.move(1, 0); this.gridBox.focus(); }
        if (e.key === 'Escape') { this.select(this.sel.r, this.sel.c); this.gridBox.focus(); }
      });
      this.gridBox.addEventListener('keydown', (e) => this.onKey(e));
      this.gridBox.addEventListener('click', (e) => {
        const hc = e.target.closest('thead th[data-hc]');
        if (hc) { this.stopEdit(true); this.selectWhole('col', +hc.dataset.hc, e.shiftKey || this.rangeMode); return; }
        const rh = e.target.closest('tbody th');
        if (rh) { this.stopEdit(true); this.selectWhole('row', +rh.parentNode.dataset.r, e.shiftKey || this.rangeMode); return; }
        const td = e.target.closest('td[data-c]');
        if (!td) return;
        const r = +td.parentNode.dataset.r, c = +td.dataset.c;
        const extend = e.shiftKey || this.rangeMode;
        if (!extend && this.sel && this.sel.r === r && this.sel.c === c && !this.editor && !this.isRange()) this.startEdit();
        else { this.stopEdit(true); this.select(r, c, extend); }
      });
      this.gridBox.addEventListener('dblclick', (e) => { if (e.target.closest('td[data-c]')) this.startEdit(); });
      this.gridBox.addEventListener('paste', (e) => {
        if (this.editor || e.target === this.fx) return;
        e.preventDefault();
        const cd = e.clipboardData;
        this.pasteCells(cd.getData('text/html'), cd.getData('text/plain'));
      });
      this.gridBox.addEventListener('copy', (e) => { if (this.editor) return; e.preventDefault(); this.copyCells(e.clipboardData); });
      this.gridBox.addEventListener('cut', (e) => { if (this.editor) return; e.preventDefault(); this.copyCells(e.clipboardData); this.isRange() ? this.clearRange() : this.commit(''); });
    }

    buildStyleSheet() {
      const f0 = this.fonts[0] || {};
      let css = `.xl-table td{font-family:"${f0.name || 'Calibri'}",Calibri,Carlito,Arial,sans-serif;font-size:${f0.sz || 11}pt}`;
      this.xfs.forEach((x, i) => {
        const f = this.fonts[x.fontId] || {}, fl = this.fills[x.fillId] || {};
        let s = '';
        if (f.name) s += `font-family:"${f.name}",Calibri,Carlito,Arial,sans-serif;`;
        if (f.sz) s += `font-size:${f.sz}pt;`;
        if (f.b) s += 'font-weight:700;';
        if (f.i) s += 'font-style:italic;';
        const deco = [f.u && 'underline', f.strike && 'line-through'].filter(Boolean).join(' ');
        if (deco) s += `text-decoration:${deco};`;
        if (f.color) s += `color:${f.color};`;
        if (fl.bg) s += `background:${fl.grad ? `linear-gradient(90deg,${fl.grad.join(',')})` : fl.bg};`;
        if (x.h && x.h !== 'general') s += `text-align:${{ centerContinuous: 'center', distributed: 'justify', fill: 'left' }[x.h] || x.h};`;
        s += `vertical-align:${{ center: 'middle', top: 'top', distributed: 'middle', justify: 'middle' }[x.v] || 'bottom'};`;
        if (x.wrap) s += 'white-space:pre-wrap;word-break:break-word;';
        if (x.indent) s += `padding-left:${3 + x.indent * 9}px;`;
        css += `.xl-table td.x${i}{${s}}`;
      });
      this.styleEl.textContent = css;
    }

    showSheet(i) {
      this.stopEdit(true);
      this.cur = i;
      const sh = this.sheets[i];
      this.tabs.querySelectorAll('.xl-tab').forEach((t) => t.classList.toggle('on', +t.dataset.i === i));
      const nR = Math.max(sh.maxR + 30, 60), nC = Math.max(sh.maxC + 6, 15);
      sh.nR = nR; sh.nC = nC;
      // celdas cubiertas por combinaciones
      sh.covered = new Map();
      for (const m of sh.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) sh.covered.set(r + ',' + c, r === m.r1 && c === m.c1 ? m : 'x');
      const table = OF.el('table', { class: 'xl-table' + (sh.grid ? '' : ' nogrid') });
      const cg = OF.el('colgroup');
      cg.append(OF.el('col', { style: { width: '44px' } }));
      for (let c = 0; c < nC; c++) {
        const ci = sh.cols[c];
        const w = ci && ci.hidden ? 0 : Math.round(((ci && ci.w) || sh.defW) * 7 + 5);
        cg.append(OF.el('col', { style: { width: w + 'px' } }));
      }
      table.append(cg);
      const thead = OF.el('thead');
      const hr = OF.el('tr', null, OF.el('th', { class: 'xl-corner' }));
      for (let c = 0; c < nC; c++) hr.append(OF.el('th', { 'data-hc': c }, colName(c)));
      thead.append(hr);
      table.append(thead);
      this.tbody = OF.el('tbody');
      table.append(this.tbody);
      this.gridBox.innerHTML = '';
      this.gridBox.append(table);
      this.table = table;
      this.rendered = 0;
      this.renderRows(Math.min(nR, 120));
      const sentinel = (this.sentinel = OF.el('div', { class: 'xl-sentinel' }));
      this.gridBox.append(sentinel);
      if (this.io) this.io.disconnect();
      this.io = new IntersectionObserver((ents) => { if (ents[0].isIntersecting && this.rendered < sh.nR) this.renderRows(Math.min(sh.nR, this.rendered + 200)); }, { root: this.gridBox, rootMargin: '600px' });
      this.io.observe(sentinel);
      this.gridBox.scrollTop = 0; this.gridBox.scrollLeft = 0;
      this.anchor = null; this.whole = null;
      this.select(0, 0);
      this.drawPictures();
    }

    renderRows(upto) {
      const sh = this.sheets[this.cur];
      const frag = document.createDocumentFragment();
      for (let r = this.rendered; r < upto; r++) frag.append(this.renderRow(sh, r));
      this.tbody.append(frag);
      this.rendered = upto;
    }

    renderRow(sh, r) {
      const ri = sh.rowInfo[r] || {};
      const h = ri.hidden ? 0 : Math.round((ri.ht || sh.defH) * 4 / 3);
      const tr = OF.el('tr', { 'data-r': r, style: { height: h + 'px' } });
      if (ri.hidden) tr.classList.add('hid');
      tr.append(OF.el('th', null, String(r + 1)));
      for (let c = 0; c < sh.nC; c++) {
        const cov = sh.covered.get(r + ',' + c);
        if (cov === 'x') continue;
        const td = OF.el('td', { 'data-c': c });
        if (cov) { td.rowSpan = cov.r2 - cov.r1 + 1; td.colSpan = cov.c2 - cov.c1 + 1; }
        this.fillCell(td, sh, r, c);
        tr.append(td);
      }
      return tr;
    }

    fillCell(td, sh, r, c) {
      const si = this.sheets.indexOf(sh);
      const x = this.value(si, r, c);
      const xf = this.xfs[x.s] || this.xfs[0];
      td.className = 'x' + x.s;
      td.style.zIndex = '';
      const txt = this.display(x);
      td.textContent = '';
      if (txt !== '') {
        const isNum = typeof x.v === 'number', isBool = typeof x.v === 'boolean' || (x.v && typeof x.v === 'object');
        if (!xf.h || xf.h === 'general') td.style.textAlign = isNum ? 'right' : isBool ? 'center' : '';
        else td.style.textAlign = '';
        const span = OF.el('span', { class: 'xl-v' }, txt);
        // el texto se desborda hacia celdas vacías a la derecha, como en Excel
        if (!isNum && !xf.wrap && (!xf.h || xf.h === 'general' || xf.h === 'left')) {
          const nx = this.raw(sh, r, c + 1);
          if (nx.v == null && nx.f == null && !sh.covered.get(r + ',' + (c + 1))) { span.classList.add('ov'); td.classList.add('ovf'); }
        }
        if (isNum && /#+/.test(txt) === false && txt.length > 0) span.classList.add('num');
        td.append(span);
      }
      // bordes: cada celda dibuja derecha/abajo resolviendo con la vecina; arriba/izquierda solo si es borde de tabla
      const b = this.borders[xf.borderId] || {};
      const nb = (rr, cc) => { const y = this.raw(sh, rr, cc); const xx = this.xfs[y.s] || this.xfs[0]; return this.borders[xx.borderId] || {}; };
      const pick = (a, bb) => (!a ? bb : !bb ? a : (BW[bb.style] || 0) > (BW[a.style] || 0) ? bb : a);
      const span = sh.covered.get(r + ',' + c);
      const right = pick(b.right, nb(r, (span && span !== 'x' ? span.c2 : c) + 1).left);
      const bottom = pick(b.bottom, nb((span && span !== 'x' ? span.r2 : r) + 1, c).top);
      const top = r === 0 ? null : pick(b.top, nb(r - 1, c).bottom);
      const left = c === 0 ? null : pick(b.left, nb(r, c - 1).right);
      td.style.borderRight = right ? `${BORDER[right.style] || '1px solid'} ${right.color}` : '';
      td.style.borderBottom = bottom ? `${BORDER[bottom.style] || '1px solid'} ${bottom.color}` : '';
      // marcar si arriba/izquierda tiene borde explícito (la vecina ya lo dibuja; esto cubre combinaciones)
      td.style.borderTop = top && span ? `${BORDER[top.style] || '1px solid'} ${top.color}` : '';
      td.style.borderLeft = left && span ? `${BORDER[left.style] || '1px solid'} ${left.color}` : '';
      if ((this.fills[xf.fillId] || {}).bg) td.classList.add('filled');
    }

    tdAt(r, c) {
      const tr = this.tbody.querySelector(`tr[data-r="${r}"]`);
      return tr && tr.querySelector(`td[data-c="${c}"]`);
    }

    select(r, c, extend) {
      const sh = this.sheets[this.cur];
      if (!extend) { this.whole = null; }
      r = Math.max(0, Math.min(sh.nR - 1, r)); c = Math.max(0, Math.min(sh.nC - 1, c));
      const cov = sh.covered.get(r + ',' + c);
      if (cov === 'x') { const m = sh.merges.find((m) => r >= m.r1 && r <= m.r2 && c >= m.c1 && c <= m.c2); if (m) { r = m.r1; c = m.c1; } }
      while (r >= this.rendered) this.renderRows(Math.min(sh.nR, this.rendered + 200));
      if (this.selTd) this.selTd.classList.remove('sel');
      this.table.querySelectorAll('.hsel').forEach((e) => e.classList.remove('hsel'));
      this.sel = { r, c };
      const td = this.tdAt(r, c);
      this.selTd = td;
      if (td) { td.classList.add('sel'); td.scrollIntoView({ block: 'nearest', inline: 'nearest' }); }
      const hc = this.table.querySelector(`th[data-hc="${c}"]`); if (hc) hc.classList.add('hsel');
      const rh = this.tbody.querySelector(`tr[data-r="${r}"] > th`); if (rh) rh.classList.add('hsel');
      if (!extend || !this.anchor) this.anchor = { r, c };
      this.paintRange();
      const g = this.rangeRect();
      this.nameBox.textContent = this.isRange() ? addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2) : addr(r, c);
      this.fx.value = this.inputText(r, c);
      this.onselect && this.onselect();
    }

    // Texto para editar: fórmula con "=", o el valor en formato editable
    inputText(r, c) {
      const sh = this.sheets[this.cur];
      const x = this.raw(sh, r, c);
      if (x.f != null) return '=' + x.f;
      const v = x.v;
      if (v == null) return '';
      if (typeof v === 'boolean') return v ? 'VERDADERO' : 'FALSO';
      if (typeof v === 'object') return v.error;
      if (typeof v === 'number') {
        const code = this.fmtCode(x.s);
        if (SSF.is_date(code)) { const d = SSF.parse_date_code(v, { date1904: this.date1904 }); if (d) return `${String(d.d).padStart(2, '0')}-${String(d.m).padStart(2, '0')}-${d.y}` + (d.H || d.M ? ` ${String(d.H).padStart(2, '0')}:${String(d.M).padStart(2, '0')}` : ''); }
        if (/%/.test(code)) return this.num(+(v * 100).toPrecision(15)) + '%';
        return this.num(v);
      }
      return v;
    }
    num(v) { const s = String(+v.toPrecision(15)); return this.locale === 'cl' ? s.replace('.', ',') : s; }

    startEdit(initial) {
      if (this.editor || !this.selTd) return;
      const td = this.selTd;
      const ed = (this.editor = OF.el('textarea', { class: 'xl-ed', spellcheck: 'false', rows: 1 }));
      ed.value = initial != null ? initial : this.inputText(this.sel.r, this.sel.c);
      const rect = td.getBoundingClientRect(), box = this.gridBox.getBoundingClientRect();
      Object.assign(ed.style, { left: rect.left - box.left + this.gridBox.scrollLeft + 'px', top: rect.top - box.top + this.gridBox.scrollTop + 'px', minWidth: rect.width + 'px', height: rect.height + 'px' });
      this.gridBox.append(ed);
      ed.focus();
      ed.setSelectionRange(ed.value.length, ed.value.length);
      ed.addEventListener('input', () => { this.fx.value = ed.value; });
      ed.addEventListener('keydown', (e) => {
        if (e.key === 'Enter' && !e.altKey) { e.preventDefault(); this.stopEdit(true); this.move(1, 0); }
        else if (e.key === 'Tab') { e.preventDefault(); this.stopEdit(true); this.move(0, e.shiftKey ? -1 : 1); }
        else if (e.key === 'Escape') { e.preventDefault(); this.stopEdit(false); }
        e.stopPropagation();
      });
      ed.addEventListener('blur', () => setTimeout(() => { if (this.editor === ed) this.stopEdit(true); }, 0));
    }
    stopEdit(save) {
      const ed = this.editor;
      if (!ed) return;
      this.editor = null;
      const val = ed.value;
      ed.remove();
      if (save) this.commit(val);
      else this.select(this.sel.r, this.sel.c);
      this.gridBox.focus({ preventScroll: true });
    }
    move(dr, dc, extend) {
      const sh = this.sheets[this.cur];
      let { r, c } = this.sel;
      const cov = sh.covered.get(r + ',' + c);
      if (cov && cov !== 'x' && !extend) { if (dr > 0) r = cov.r2; if (dc > 0) c = cov.c2; }
      this.select(r + dr, c + dc, extend);
    }

    onKey(e) {
      if (this.editor || e.target === this.fx) return;
      const k = e.key;
      const mv = { ArrowDown: [1, 0], ArrowUp: [-1, 0], ArrowLeft: [0, -1], ArrowRight: [0, 1], Enter: [1, 0], Tab: [0, e.shiftKey ? -1 : 1] }[k];
      if (mv) { e.preventDefault(); this.move(mv[0], mv[1], e.shiftKey && k.startsWith('Arrow')); return; }
      if ((e.ctrlKey || e.metaKey) && ['z', 'y'].includes(k.toLowerCase())) return;
      if (k === 'Delete' || k === 'Backspace') { e.preventDefault(); if (this.isRange()) this.clearRange(); else this.commit(''); return; }
      if (k === 'F2') { e.preventDefault(); this.startEdit(); return; }
      if (k.length === 1 && !e.ctrlKey && !e.metaKey && !e.altKey) { e.preventDefault(); this.startEdit(k); }
    }

    doUndo() {
      const u = this.undo.pop();
      if (!u) return;
      if (u.sheet !== this.cur) this.showSheet(u.sheet);
      this.select(u.r, u.c);
      this.commit(u.prev, true);
    }

    /* ---------------------- escritura en el XML ---------------------- */
    parseInput(str) {
      const s = str.trim();
      if (s === '') return { type: 'empty' };
      if (s.startsWith('=') && s.length > 1) return { type: 'formula', f: s.slice(1) };
      const up = s.toUpperCase();
      if (up === 'VERDADERO' || up === 'TRUE') return { type: 'bool', v: true };
      if (up === 'FALSO' || up === 'FALSE') return { type: 'bool', v: false };
      const dm = /^(\d{1,2})[-/](\d{1,2})[-/](\d{2,4})(?:\s+(\d{1,2}):(\d{2}))?$/.exec(s);
      if (dm) {
        let y = +dm[3]; if (y < 100) y += 2000;
        const d = Date.UTC(y, +dm[2] - 1, +dm[1], +(dm[4] || 0), +(dm[5] || 0));
        let serial = d / 86400000 + 25569;
        if (this.date1904) serial -= 1462;
        if (!isNaN(serial)) return { type: 'num', v: serial, date: true };
      }
      const pct = /%$/.test(s);
      let n = s.replace(/%$/, '').replace(/^\$\s?|\s/g, '');
      if (this.locale === 'cl') {
        if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(n)) n = n.replace(/\./g, '').replace(',', '.');
        else if (/^-?\d+,\d+$/.test(n)) n = n.replace(',', '.');
      } else n = n.replace(/,(?=\d{3}(\D|$))/g, '');
      if (/^-?(\d+\.?\d*|\.\d+)(e[+-]?\d+)?$/i.test(n)) return { type: 'num', v: pct ? +n / 100 : +n, pct };
      return { type: 'str', v: s };
    }

    cellEl(sh, r, c, create) {
      let el = sh.cells.get(r + ',' + c);
      if (el || !create) return el;
      const doc = sh.doc;
      let row = sh.rowEls.get(r);
      if (!row) {
        row = doc.createElementNS(S, 'row');
        row.setAttribute('r', r + 1);
        let next = null;
        for (const [rr, re] of sh.rowEls) if (rr > r && (!next || rr < next[0])) next = [rr, re];
        sh.data.insertBefore(row, next ? next[1] : null);
        sh.rowEls.set(r, row);
      }
      el = doc.createElementNS(S, 'c');
      el.setAttribute('r', addr(r, c));
      const st = this.defaultStyle(sh, r, c);
      if (st) el.setAttribute('s', st);
      let next = null;
      for (const ce of OF.kids(row, 'c', S)) { const p = parseRef(ce.getAttribute('r')); if (p.c > c) { next = ce; break; } }
      row.insertBefore(el, next);
      if (row.hasAttribute('spans')) row.removeAttribute('spans');
      sh.cells.set(r + ',' + c, el);
      return el;
    }

    expandShared(sh, si) {
      const m = sh.shared[si];
      if (!m) return;
      for (const [key, el] of sh.cells) {
        const f = OF.kid(el, 'f', S);
        if (!f || f.getAttribute('t') !== 'shared' || f.getAttribute('si') !== si) continue;
        const [r, c] = key.split(',').map(Number);
        const text = f.textContent || shiftFormula(m.f, r - m.at.r, c - m.at.c);
        const nf = sh.doc.createElementNS(S, 'f');
        nf.textContent = text;
        el.replaceChild(nf, f);
      }
      delete sh.shared[si];
    }

    addString(text) {
      if (!this.ssDoc) return null;
      const root = this.ssDoc.documentElement;
      const si = this.ssDoc.createElementNS(S, 'si');
      const t = this.ssDoc.createElementNS(S, 't');
      OF.setText(t, text);
      si.appendChild(t);
      root.appendChild(si);
      this.strings.push(text);
      root.setAttribute('uniqueCount', this.strings.length);
      root.setAttribute('count', +(root.getAttribute('count') || 0) + 1);
      this.ssDirty = true;
      return this.strings.length - 1;
    }

    commit(input, isUndo) {
      const sh = this.sheets[this.cur];
      const { r, c } = this.sel;
      const prev = this.inputText(r, c);
      if (input === prev) { this.select(r, c); return; }
      const p = this.parseInput(input);
      const el = this.cellEl(sh, r, c, p.type !== 'empty');
      if (!isUndo) OF.checkpoint && OF.checkpoint();
      if (el) {
        const oldF = OF.kid(el, 'f', S);
        if (oldF && oldF.getAttribute('t') === 'shared' && oldF.getAttribute('ref')) this.expandShared(sh, oldF.getAttribute('si'));
        for (const ch of Array.from(el.children)) if (['f', 'v', 'is'].includes(ch.localName)) el.removeChild(ch);
        el.removeAttribute('t');
        const doc = sh.doc;
        const add = (name, text) => { const e = doc.createElementNS(S, name); e.textContent = text; el.appendChild(e); return e; };
        if (p.type === 'formula') { add('f', p.f); this.formulasTouched = true; }
        else if (p.type === 'num') add('v', String(p.v));
        else if (p.type === 'bool') { el.setAttribute('t', 'b'); add('v', p.v ? '1' : '0'); }
        else if (p.type === 'str') {
          const idx = this.addString(p.v);
          if (idx != null) { el.setAttribute('t', 's'); add('v', String(idx)); }
          else { el.setAttribute('t', 'inlineStr'); const is = add('is', ''); const t = doc.createElementNS(S, 't'); OF.setText(t, p.v); is.appendChild(t); }
        }
        // formato automático para fechas/porcentajes escritos en una celda sin formato
        if ((p.date || p.pct) && this.fmtCode(+(el.getAttribute('s') || 0)) === 'General') {
          const s = this.ensureXf(p.date ? 14 : 9, +(el.getAttribute('s') || 0));
          if (s != null) el.setAttribute('s', s);
        }
        if (r > sh.maxR) sh.maxR = r; if (c > sh.maxC) sh.maxC = c;
      }
      this.cacheOnly.delete(this.cur + ':' + r + ',' + c);
      const changed = new Set([r + ',' + c]);
      if (this.hf) {
        try {
          const val = p.type === 'formula' ? '=' + p.f : p.type === 'empty' ? null : p.v;
          const changes = this.hf.setCellContents({ sheet: sh.hfId, row: r, col: c }, [[val]]);
          for (const ch of changes) if (ch.address && ch.address.sheet === sh.hfId) {
            changed.add(ch.address.row + ',' + ch.address.col);
            // si una celda "solo caché" depende de esto, pasa a calcularse
            this.cacheOnly.delete(this.cur + ':' + ch.address.row + ',' + ch.address.col);
          }
        } catch (e) { console.warn(e); }
      }
      // repintar celdas afectadas (y vecinas por desbordes/bordes)
      for (const key of changed) {
        const [rr, cc] = key.split(',').map(Number);
        for (const [a, b] of [[rr, cc], [rr, cc - 1]]) { const td = this.tdAt(a, b); if (td) this.fillCell(td, sh, a, b); }
      }
      this.markDirty();
      this.select(r, c);
    }

    // Busca o crea un estilo igual al actual pero con otro formato numérico
    ensureXf(numFmtId, baseIdx) {
      const base = this.xfs[baseIdx] || this.xfs[0];
      const i = this.xfs.findIndex((x) => x.numFmtId === numFmtId && x.fontId === base.fontId && x.fillId === base.fillId && x.borderId === base.borderId && x.h === base.h);
      if (i >= 0) return i;
      if (!this.stylesDoc) return null;
      const cellXfs = OF.kid(this.stylesDoc.documentElement, 'cellXfs', S);
      const list = OF.kids(cellXfs, 'xf', S);
      const src = list[baseIdx] || list[0];
      if (!src) return null;
      const nx = src.cloneNode(true);
      nx.setAttribute('numFmtId', numFmtId);
      nx.setAttribute('applyNumberFormat', '1');
      cellXfs.appendChild(nx);
      cellXfs.setAttribute('count', list.length + 1);
      this.xfs.push({ ...base, numFmtId });
      this.stylesDirty = true;
      this.buildStyleSheet();
      return this.xfs.length - 1;
    }

    markDirty() {
      if (!this.dirty) { this.dirty = true; this.onchange && this.onchange(); }
    }

    zoom(f) {
      this.userZoom = Math.max(0.4, Math.min(3, (this.userZoom || 1) * f));
      this.table.style.zoom = this.userZoom;
      if (this.picLayer) this.picLayer.style.zoom = this.userZoom;
    }
    fit() {}

    async save() {
      if (!this.dirty) return new Blob([this.original], { type: ExcelDoc.mime });
      // valores calculados para las fórmulas (caché), así se ven bien en cualquier programa
      this.sheets.forEach((sh, si) => {
        for (const [key, el] of sh.cells) {
          const f = OF.kid(el, 'f', S);
          if (!f || !this.hf || this.cacheOnly.has(si + ':' + key)) continue;
          const [r, c] = key.split(',').map(Number);
          const v = this.hf.getCellValue({ sheet: sh.hfId, row: r, col: c });
          let vEl = OF.kid(el, 'v', S);
          if (v && typeof v === 'object') { if (vEl) el.removeChild(vEl); el.removeAttribute('t'); continue; }
          if (!vEl) { vEl = sh.doc.createElementNS(S, 'v'); el.appendChild(vEl); }
          if (typeof v === 'number') { el.removeAttribute('t'); vEl.textContent = String(v); }
          else if (typeof v === 'boolean') { el.setAttribute('t', 'b'); vEl.textContent = v ? '1' : '0'; }
          else if (v == null) { el.removeChild(vEl); el.removeAttribute('t'); }
          else { el.setAttribute('t', 'str'); vEl.textContent = String(v); }
        }
        this.zip.file(sh.path, OF.serializeXml(sh.doc));
      });
      if (this.stylesDirty) this.zip.file(this.stylesPath, OF.serializeXml(this.stylesDoc));
      if (this.ssDirty) this.zip.file(this.ssRel.target, OF.serializeXml(this.ssDoc));
      // pedir a Excel que recalcule todo al abrir y quitar la cadena de cálculo (Excel la regenera)
      let calcPr = OF.kid(this.wb.documentElement, 'calcPr', S);
      if (!calcPr) {
        calcPr = this.wb.createElementNS(S, 'calcPr');
        const after = ['oleSize', 'customWorkbookViews', 'pivotCaches', 'smartTagPr', 'smartTagTypes', 'webPublishing', 'fileRecoveryPr', 'webPublishObjects', 'extLst'];
        let before = null;
        for (const n of after) { before = OF.kid(this.wb.documentElement, n, S); if (before) break; }
        this.wb.documentElement.insertBefore(calcPr, before);
      }
      calcPr.setAttribute('fullCalcOnLoad', '1');
      this.zip.file(this.wbPath, OF.serializeXml(this.wb));
      await this.dropCalcChain();
      const blob = await this.zip.generateAsync({ type: 'blob', mimeType: ExcelDoc.mime, compression: 'DEFLATE' });
      this.original = await blob.arrayBuffer();
      this.dirty = false;
      return blob;
    }

    async dropCalcChain() {
      const relPath = OF.relsPathFor(this.wbPath);
      const relFile = this.zip.file(relPath);
      if (!relFile) return;
      const doc = OF.parseXml(await relFile.async('string'));
      let removed = false;
      for (const r of OF.all(doc, 'Relationship')) {
        if (/calcChain$/.test(r.getAttribute('Type'))) {
          this.zip.remove(OF.resolvePath(this.wbPath, r.getAttribute('Target')));
          r.parentNode.removeChild(r);
          removed = true;
        }
      }
      if (!removed) return;
      this.zip.file(relPath, OF.serializeXml(doc));
      const ct = OF.parseXml(await this.zip.file('[Content_Types].xml').async('string'));
      for (const o of OF.all(ct, 'Override')) if (/calcChain/.test(o.getAttribute('PartName'))) o.parentNode.removeChild(o);
      this.zip.file('[Content_Types].xml', OF.serializeXml(ct));
    }
  }

  function siText(si) {
    let s = '';
    for (const t of si.getElementsByTagNameNS(S, 't')) {
      let p = t.parentNode, skip = false;
      while (p && p !== si) { if (p.localName === 'rPh') skip = true; p = p.parentNode; }
      if (!skip) s += t.textContent;
    }
    return s;
  }
  function sameValue(h, x) {
    if (h && typeof h === 'object' && 'value' in h) return !!(x && typeof x === 'object');
    if (x == null || x === '') return h == null || h === '' || h === 0;
    if (typeof x === 'number' && typeof h === 'number') return Math.abs(h - x) <= 1e-9 * Math.max(1, Math.abs(x));
    return h === x;
  }

  /* =================== formato, filas/columnas e imágenes =================== */
  const FONT_ORDER = ['b', 'i', 'strike', 'condense', 'extend', 'outline', 'shadow', 'u', 'vertAlign', 'sz', 'color', 'name', 'family', 'charset', 'scheme'];
  const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const REF_RE = /(?<![\p{L}\p{N}_.$'])((?:'(?:[^']|'')+'|[\p{L}_][\p{L}\p{N}_.]*)!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?|\$?[A-Z]{1,3}:\$?[A-Z]{1,3}|\$?\d+:\$?\d+)(?![\p{L}\p{N}_(!])/gu;

  // Desplaza una referencia (A1, A1:B2, A:B, 1:3). Devuelve null si queda eliminada.
  function shiftRef(ref, axis, one, rng) {
    const cell = /^(\$?)([A-Z]{1,3})(\$?)(\d+)$/;
    const parts = ref.split(':');
    if (parts.length === 1) {
      const m = cell.exec(ref);
      if (!m) return ref;
      if (axis === 'row') { const v = one(+m[4] - 1); return v == null ? null : m[1] + m[2] + m[3] + (v + 1); }
      const v = one(colIndex(m[2])); return v == null ? null : m[1] + colName(v) + m[3] + m[4];
    }
    const [A, B] = parts;
    const ma = cell.exec(A), mb = cell.exec(B);
    if (ma && mb) {
      if (axis === 'row') { const x = rng(+ma[4] - 1, +mb[4] - 1); return x ? `${ma[1]}${ma[2]}${ma[3]}${x[0] + 1}:${mb[1]}${mb[2]}${mb[3]}${x[1] + 1}` : null; }
      const x = rng(colIndex(ma[2]), colIndex(mb[2])); return x ? `${ma[1]}${colName(x[0])}${ma[3]}${ma[4]}:${mb[1]}${colName(x[1])}${mb[3]}${mb[4]}` : null;
    }
    const ca = /^(\$?)([A-Z]{1,3})$/.exec(A), cb = /^(\$?)([A-Z]{1,3})$/.exec(B);
    if (ca && cb) { if (axis === 'row') return ref; const x = rng(colIndex(ca[2]), colIndex(cb[2])); return x ? `${ca[1]}${colName(x[0])}:${cb[1]}${colName(x[1])}` : null; }
    const ra = /^(\$?)(\d+)$/.exec(A), rb = /^(\$?)(\d+)$/.exec(B);
    if (ra && rb) { if (axis === 'col') return ref; const x = rng(+ra[2] - 1, +rb[2] - 1); return x ? `${ra[1]}${x[0] + 1}:${rb[1]}${x[1] + 1}` : null; }
    return ref;
  }
  function shiftRefs(f, same, target, axis, one, rng) {
    return f.split(/("(?:[^"]|"")*")/).map((seg, i) => {
      if (i % 2) return seg;
      return seg.replace(REF_RE, (m, pre, ref) => {
        const sheet = pre ? pre.slice(0, -1).replace(/^'|'$/g, '').replace(/''/g, "'") : null;
        if (sheet ? sheet !== target : !same) return m;
        const out = shiftRef(ref, axis, one, rng);
        return (pre || '') + (out == null ? '#REF!' : out);
      });
    }).join('');
  }
  OF._shiftRefs = shiftRefs; // (pruebas)

  Object.assign(ExcelDoc.prototype, {
    /* ---------------------------- selección de rango ---------------------------- */
    isRange() { const g = this.rangeRect(); return g.r1 !== g.r2 || g.c1 !== g.c2; },
    rangeRect() {
      const a = this.anchor || this.sel, b = this.sel;
      const g = { r1: Math.min(a.r, b.r), r2: Math.max(a.r, b.r), c1: Math.min(a.c, b.c), c2: Math.max(a.c, b.c) };
      const sh = this.sheets[this.cur];
      if (this.whole === 'row') { g.c1 = 0; g.c2 = sh.nC - 1; }
      if (this.whole === 'col') { g.r1 = 0; g.r2 = sh.nR - 1; }
      return g;
    },
    paintRange() {
      this.table.querySelectorAll('.rng').forEach((t) => t.classList.remove('rng'));
      const g = this.rangeRect();
      for (let c = g.c1; c <= g.c2; c++) { const h = this.table.querySelector(`th[data-hc="${c}"]`); if (h) h.classList.add('hsel'); }
      for (let r = g.r1; r <= Math.min(g.r2, this.rendered - 1); r++) {
        const tr = this.tbody.querySelector(`tr[data-r="${r}"]`);
        if (!tr) continue;
        tr.firstChild.classList.add('hsel');
        if (!this.isRange()) continue;
        for (const td of tr.querySelectorAll('td[data-c]')) { const c = +td.dataset.c; if (c >= g.c1 && c <= g.c2) td.classList.add('rng'); }
      }
    },
    selectWhole(kind, i, extend) {
      if (!extend || this.whole !== kind) this.anchor = kind === 'col' ? { r: 0, c: i } : { r: i, c: 0 };
      this.whole = kind;
      const keepR = this.sel ? this.sel.r : 0;
      if (kind === 'col') this.select(Math.min(keepR, this.rendered - 1), i, true);
      else this.select(i, 0, true);
    },
    // Celdas a tocar: con fila/columna completa se limita a la zona usada
    workRect() {
      const sh = this.sheets[this.cur];
      const g = { ...this.rangeRect() };
      if (this.whole === 'col') g.r2 = Math.max(g.r1, sh.maxR);
      if (this.whole === 'row') g.c2 = Math.max(g.c1, sh.maxC);
      return g;
    },
    repaintRect(g) {
      const sh = this.sheets[this.cur];
      for (let r = Math.max(0, g.r1 - 1); r <= Math.min(g.r2 + 1, this.rendered - 1); r++)
        for (let c = Math.max(0, g.c1 - 1); c <= Math.min(g.c2 + 1, sh.nC - 1); c++) { const td = this.tdAt(r, c); if (td) this.fillCell(td, sh, r, c); }
      this.select(this.sel.r, this.sel.c, true);
    },
    clearRange() {
      OF.checkpoint && OF.checkpoint();
      const sh = this.sheets[this.cur], g = this.workRect();
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        const el = sh.cells.get(r + ',' + c);
        if (!el) continue;
        const f = OF.kid(el, 'f', S);
        if (f && f.getAttribute('t') === 'shared' && f.getAttribute('ref')) this.expandShared(sh, f.getAttribute('si'));
        for (const ch of Array.from(el.children)) if (['f', 'v', 'is'].includes(ch.localName)) el.removeChild(ch);
        el.removeAttribute('t');
        if (this.hf) try { this.hf.setCellContents({ sheet: sh.hfId, row: r, col: c }, [[null]]); } catch (e) {}
      }
      this.markDirty();
      this.showSheetKeep();
    },
    showSheetKeep() {
      const sel = { ...this.sel }, anchor = this.anchor && { ...this.anchor }, whole = this.whole;
      const st = this.gridBox.scrollTop, sl = this.gridBox.scrollLeft;
      this.showSheet(this.cur);
      while (this.rendered <= sel.r && this.rendered < this.sheets[this.cur].nR) this.renderRows(Math.min(this.sheets[this.cur].nR, this.rendered + 200));
      this.gridBox.scrollTop = st; this.gridBox.scrollLeft = sl;
      this.whole = whole; this.anchor = anchor;
      this.select(sel.r, sel.c, true);
    },

    /* ------------------------------ estilos de celda ------------------------------ */
    numFmtId(code) {
      const t = SSF.get_table();
      for (const id in t) if (+id < 164 && t[id] === code) return +id;
      for (const id in this.numFmts) if (this.numFmts[id] === code) return +id;
      const root = this.stylesDoc.documentElement;
      let nf = OF.kid(root, 'numFmts', S);
      if (!nf) { nf = this.stylesDoc.createElementNS(S, 'numFmts'); root.insertBefore(nf, root.firstElementChild); }
      const id = Math.max(163, ...Object.keys(this.numFmts).map(Number)) + 1;
      const e = this.stylesDoc.createElementNS(S, 'numFmt');
      e.setAttribute('numFmtId', id); e.setAttribute('formatCode', code);
      nf.appendChild(e);
      nf.setAttribute('count', OF.kids(nf, 'numFmt', S).length);
      this.numFmts[id] = code;
      return id;
    },
    newXf(base, ch) {
      const key = base + '|' + JSON.stringify(ch);
      this.xfCache = this.xfCache || {};
      if (this.xfCache[key] != null) return this.xfCache[key];
      const doc = this.stylesDoc, root = doc.documentElement;
      const cellXfs = OF.kid(root, 'cellXfs', S);
      const list = OF.kids(cellXfs, 'xf', S);
      const nx = (list[base] || list[0]).cloneNode(true);
      const xf = this.xfs[base] || this.xfs[0];
      if (ch.font) {
        const fontsEl = OF.kid(root, 'fonts', S), fl = OF.kids(fontsEl, 'font', S);
        const nf = (fl[xf.fontId] || fl[0]).cloneNode(true);
        for (const k of ['b', 'i', 'u']) if (k in ch.font) { OF.removeKids(nf, [k], S); if (ch.font[k]) OF.ordered(nf, S, '', k, FONT_ORDER); }
        if ('color' in ch.font) { OF.removeKids(nf, ['color'], S); if (ch.font.color) OF.ordered(nf, S, '', 'color', FONT_ORDER).setAttribute('rgb', 'FF' + ch.font.color.replace('#', '').toUpperCase()); }
        if ('sz' in ch.font) OF.ordered(nf, S, '', 'sz', FONT_ORDER).setAttribute('val', ch.font.sz);
        if (ch.font.name) { OF.ordered(nf, S, '', 'name', FONT_ORDER).setAttribute('val', ch.font.name); OF.removeKids(nf, ['scheme', 'family', 'charset'], S); }
        fontsEl.appendChild(nf);
        fontsEl.setAttribute('count', fl.length + 1);
        this.fonts.push(this.parseFont(nf));
        nx.setAttribute('fontId', fl.length);
        nx.setAttribute('applyFont', '1');
      }
      if ('fill' in ch) {
        if (!ch.fill) nx.setAttribute('fillId', '0');
        else {
          const fillsEl = OF.kid(root, 'fills', S), fl = OF.kids(fillsEl, 'fill', S);
          const nf = OF.fragment(doc, `<fill xmlns="${S}"><patternFill patternType="solid"><fgColor rgb="FF${ch.fill.replace('#', '').toUpperCase()}"/><bgColor indexed="64"/></patternFill></fill>`);
          fillsEl.appendChild(nf);
          fillsEl.setAttribute('count', fl.length + 1);
          this.fills.push(this.parseFill(nf));
          nx.setAttribute('fillId', fl.length);
        }
        nx.setAttribute('applyFill', '1');
      }
      if (ch.sides) {
        const bordersEl = OF.kid(root, 'borders', S), bl = OF.kids(bordersEl, 'border', S);
        const nb = (bl[xf.borderId] || bl[0]).cloneNode(true);
        const ORDER = ['start', 'left', 'end', 'right', 'top', 'bottom', 'diagonal', 'vertical', 'horizontal'];
        for (const side of ['left', 'right', 'top', 'bottom']) {
          if (!(side in ch.sides)) continue;
          OF.removeKids(nb, [side, side === 'left' ? 'start' : side === 'right' ? 'end' : '__'], S);
          const e = OF.ordered(nb, S, '', side, ORDER);
          const v = ch.sides[side];
          if (v) { e.setAttribute('style', v); const col = doc.createElementNS(S, 'color'); col.setAttribute('auto', '1'); e.appendChild(col); }
        }
        bordersEl.appendChild(nb);
        bordersEl.setAttribute('count', bl.length + 1);
        const side = (n) => { const e = OF.kid(nb, n, S); const st = e && e.getAttribute('style'); return st ? { style: st, color: this.color(OF.kid(e, 'color', S)) || '#000' } : null; };
        this.borders.push({ left: side('left'), right: side('right'), top: side('top'), bottom: side('bottom') });
        nx.setAttribute('borderId', bl.length);
        nx.setAttribute('applyBorder', '1');
      }
      if ('h' in ch || 'wrap' in ch) {
        let al = OF.kid(nx, 'alignment', S);
        if (!al) { al = doc.createElementNS(S, 'alignment'); nx.insertBefore(al, nx.firstChild); }
        if ('h' in ch) al.setAttribute('horizontal', ch.h);
        if ('wrap' in ch) { if (ch.wrap) al.setAttribute('wrapText', '1'); else al.removeAttribute('wrapText'); }
        nx.setAttribute('applyAlignment', '1');
      }
      if ('numFmt' in ch) { nx.setAttribute('numFmtId', this.numFmtId(ch.numFmt)); nx.setAttribute('applyNumberFormat', '1'); }
      for (const a of ['xfId']) if (!nx.hasAttribute(a)) nx.setAttribute(a, '0');
      cellXfs.appendChild(nx);
      cellXfs.setAttribute('count', list.length + 1);
      this.xfs.push(this.parseXf(nx));
      this.stylesDirty = true;
      return (this.xfCache[key] = this.xfs.length - 1);
    },
    applyStyle(ch) {
      if (!this.stylesDoc) { OF.toast('Este archivo no tiene hoja de estilos; no se puede dar formato.'); return; }
      this.stopEdit(true);
      OF.checkpoint && OF.checkpoint();
      const sh = this.sheets[this.cur], g = this.workRect();
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        if (sh.covered.get(r + ',' + c) === 'x') continue;
        const el = this.cellEl(sh, r, c, true);
        el.setAttribute('s', this.newXf(+(el.getAttribute('s') || 0), ch));
      }
      this.buildStyleSheet();
      this.repaintRect(g);
      this.markDirty();
    },
    activeFont() { const x = this.raw(this.sheets[this.cur], this.sel.r, this.sel.c); return this.fonts[(this.xfs[x.s] || this.xfs[0]).fontId] || {}; },
    toggle(prop) { this.applyStyle({ font: { [prop]: !this.activeFont()[prop] } }); },
    setColor(hex) { this.applyStyle({ font: { color: hex } }); },
    setFill(hex) { this.applyStyle({ fill: hex }); },
    setSize(pt) { this.applyStyle({ font: { sz: pt } }); },
    align(a) { this.applyStyle({ h: a === 'justify' ? 'justify' : a }); },
    wrapText() { const x = this.raw(this.sheets[this.cur], this.sel.r, this.sel.c); this.applyStyle({ wrap: !(this.xfs[x.s] || {}).wrap }); },
    numFormat(code) { this.applyStyle({ numFmt: code }); },

    /* --------------------------- filas y columnas --------------------------- */
    async rowOp(op) {
      const g = this.rangeRect();
      const count = this.whole === 'col' ? 1 : g.r2 - g.r1 + 1;
      const at = op === 'below' ? (this.whole === 'col' ? this.sel.r : g.r2) + 1 : this.whole === 'col' ? this.sel.r : g.r1;
      await this.shiftSheet('row', at, op === 'del' ? -count : count);
    },
    async colOp(op) {
      const g = this.rangeRect();
      const count = this.whole === 'row' ? 1 : g.c2 - g.c1 + 1;
      const at = op === 'right' ? (this.whole === 'row' ? this.sel.c : g.c2) + 1 : this.whole === 'row' ? this.sel.c : g.c1;
      await this.shiftSheet('col', at, op === 'del' ? -count : count);
    },
    async sheetParts(sh, typeRe) {
      const out = [];
      for (const rel of Object.values(sh.rels || {})) {
        if (!typeRe.test(rel.type) || rel.external || !this.zip.file(rel.target)) continue;
        out.push({ path: rel.target, doc: OF.parseXml(await this.zip.file(rel.target).async('string')) });
      }
      return out;
    },
    async shiftSheet(axis, at, n) {
      this.stopEdit(true);
      OF.checkpoint && OF.checkpoint();
      const T = this.sheets[this.cur];
      const k = -n;
      const one = (i) => (n > 0 ? (i >= at ? i + n : i) : i < at ? i : i >= at + k ? i - k : null);
      const rng = (a, b) => {
        if (n > 0) return [a >= at ? a + n : a, b >= at ? b + n : b];
        const a2 = a < at ? a : a >= at + k ? a - k : at, b2 = b < at ? b : b >= at + k ? b - k : at - 1;
        return b2 < a2 ? null : [a2, b2];
      };
      const tables = await this.sheetParts(T, /\/table$/);
      // tablas con formato: columnas nuevas o eliminadas dentro de la tabla
      const newHeads = [];
      if (axis === 'col') {
        for (const t of tables) {
          const tr = t.doc.documentElement;
          const [a, b] = tr.getAttribute('ref').split(':').map(parseRef);
          const tcs = OF.kid(tr, 'tableColumns', S);
          const list = OF.kids(tcs, 'tableColumn', S);
          if (n > 0 && at > a.c && at <= b.c) {
            const names = new Set(list.map((e) => (e.getAttribute('name') || '').toLowerCase()));
            let maxId = Math.max(0, ...list.map((e) => +e.getAttribute('id') || 0)), num = 0;
            const before = list[at - a.c];
            for (let i = 0; i < n; i++) {
              let name; do { name = 'Columna' + ++num; } while (names.has(name.toLowerCase()));
              names.add(name.toLowerCase());
              const e = t.doc.createElementNS(S, tr.prefix ? tr.prefix + ':tableColumn' : 'tableColumn');
              e.setAttribute('id', ++maxId); e.setAttribute('name', name);
              tcs.insertBefore(e, before);
              newHeads.push({ r: a.r, c: at + i, name, s: (T.cells.get(a.r + ',' + (at - 1)) || { getAttribute: () => null }).getAttribute('s') });
            }
            tcs.setAttribute('count', OF.kids(tcs, 'tableColumn', S).length);
          } else if (n < 0 && !(at + k - 1 < a.c || at > b.c)) {
            const lo = Math.max(at, a.c) - a.c, hi = Math.min(at + k - 1, b.c) - a.c;
            if (hi - lo + 1 >= list.length) { OF.toast('Para quitar todas las columnas de una tabla con formato, elimina la tabla en Excel.', 5000); return; }
            for (let i = lo; i <= hi; i++) tcs.removeChild(list[i]);
            tcs.setAttribute('count', OF.kids(tcs, 'tableColumn', S).length);
            for (const fc of OF.all(tr, 'filterColumn')) fc.parentNode.removeChild(fc);
          }
        }
      }
      const fx = (text, same) => shiftRefs(text, same, T.name, axis, one, rng);
      const sr = (ref) => shiftRef(ref, axis, one, rng);
      // fórmulas compartidas a fórmulas normales (para poder moverlas)
      for (const sh of this.sheets) for (const si of Object.keys(sh.shared)) this.expandShared(sh, si);
      // fórmulas de todo el libro
      for (const sh of this.sheets) {
        const same = sh === T;
        for (const el of sh.cells.values()) {
          const f = OF.kid(el, 'f', S);
          if (!f) continue;
          if (f.textContent) f.textContent = fx(f.textContent, same);
          if (same && f.getAttribute('ref')) { const nr = sr(f.getAttribute('ref')); if (nr) f.setAttribute('ref', nr); else f.removeAttribute('ref'); }
        }
        for (const name of ['formula', 'formula1', 'formula2']) for (const fe of OF.all(sh.doc, name, S)) fe.textContent = fx(fe.textContent, same);
      }
      const wbSheets = OF.all(this.wb, 'sheet');
      const tIdx = wbSheets.findIndex((e) => e.getAttribute('name') === T.name);
      for (const dn of OF.all(this.wb, 'definedName')) dn.textContent = fx(dn.textContent, dn.getAttribute('localSheetId') === String(tIdx));
      // celdas de la hoja
      for (const row of OF.kids(T.data, 'row', S)) {
        const r = +row.getAttribute('r') - 1;
        if (axis === 'row') {
          const nr = one(r);
          if (nr == null) { T.data.removeChild(row); continue; }
          if (nr !== r) { row.setAttribute('r', nr + 1); for (const c of OF.kids(row, 'c', S)) c.setAttribute('r', addr(nr, parseRef(c.getAttribute('r')).c)); }
        } else {
          row.removeAttribute('spans');
          for (const c of OF.kids(row, 'c', S)) {
            const nc = one(parseRef(c.getAttribute('r')).c);
            if (nc == null) row.removeChild(c);
            else c.setAttribute('r', addr(r, nc));
          }
        }
      }
      // encabezados de las columnas nuevas de una tabla (deben coincidir con su nombre)
      for (const h of newHeads) {
        const row = OF.kids(T.data, 'row', S).find((rw) => +rw.getAttribute('r') === h.r + 1);
        if (!row) continue;
        const c = T.doc.createElementNS(S, T.data.prefix ? T.data.prefix + ':c' : 'c');
        c.setAttribute('r', addr(h.r, h.c)); c.setAttribute('t', 'inlineStr');
        if (h.s) c.setAttribute('s', h.s);
        const is = T.doc.createElementNS(S, T.data.prefix ? T.data.prefix + ':is' : 'is'), tt = T.doc.createElementNS(S, T.data.prefix ? T.data.prefix + ':t' : 't');
        tt.textContent = h.name; is.appendChild(tt); c.appendChild(is);
        let next = null; for (const ce of OF.kids(row, 'c', S)) if (parseRef(ce.getAttribute('r')).c > h.c) { next = ce; break; }
        for (const ce of OF.kids(row, 'c', S)) if (parseRef(ce.getAttribute('r')).c === h.c) row.removeChild(ce);
        row.insertBefore(c, next);
      }
      const root = T.doc.documentElement;
      if (axis === 'col') {
        const cols = OF.kid(root, 'cols', S);
        for (const col of OF.kids(cols, 'col', S)) {
          const x = rng(+col.getAttribute('min') - 1, +col.getAttribute('max') - 1);
          if (!x) cols.removeChild(col);
          else { col.setAttribute('min', x[0] + 1); col.setAttribute('max', x[1] + 1); }
        }
        if (cols && !cols.firstElementChild) root.removeChild(cols);
      }
      // rangos guardados en atributos
      const fixAttr = (el, name, removeIfGone) => {
        const v = el.getAttribute(name);
        if (!v) return;
        const out = v.split(/\s+/).map(sr).filter(Boolean).join(' ');
        if (out) el.setAttribute(name, out);
        else if (removeIfGone) el.parentNode.removeChild(el);
      };
      for (const d of OF.all(root, 'dimension', S)) fixAttr(d, 'ref');
      for (const m of OF.all(root, 'mergeCell', S)) { fixAttr(m, 'ref', true); if (m.parentNode && !/:/.test(m.getAttribute('ref') || '')) m.parentNode.removeChild(m); }
      const mc = OF.kid(root, 'mergeCells', S);
      if (mc) { const cnt = OF.kids(mc, 'mergeCell', S).length; if (cnt) mc.setAttribute('count', cnt); else root.removeChild(mc); }
      for (const h of OF.all(root, 'hyperlink', S)) fixAttr(h, 'ref', true);
      for (const a of OF.all(root, 'autoFilter', S)) fixAttr(a, 'ref', true);
      for (const a of OF.all(root, 'sortState', S)) fixAttr(a, 'ref', true);
      for (const a of OF.all(root, 'conditionalFormatting', S)) fixAttr(a, 'sqref', true);
      for (const a of OF.all(root, 'dataValidation', S)) fixAttr(a, 'sqref', true);
      for (const a of OF.all(root, 'ignoredError', S)) fixAttr(a, 'sqref', true);
      for (const name of ['hyperlinks', 'dataValidations', 'ignoredErrors']) { const e = OF.kid(root, name, S); if (e && !e.firstElementChild) root.removeChild(e); }
      const dv = OF.kid(root, 'dataValidations', S); if (dv) dv.setAttribute('count', OF.kids(dv, 'dataValidation', S).length);
      for (const sel of OF.all(root, 'selection', S)) { sel.setAttribute('activeCell', 'A1'); sel.setAttribute('sqref', 'A1'); }
      // dibujos (imágenes y gráficos) anclados a celdas
      for (const d of await this.sheetParts(T, /\/drawing$/)) {
        for (const anc of ['from', 'to']) for (const e of d.doc.getElementsByTagNameNS(XDR, anc)) {
          const ax = OF.kid(e, axis === 'row' ? 'row' : 'col', XDR);
          if (!ax) continue;
          const v = +ax.textContent;
          ax.textContent = n > 0 ? (v >= at ? v + n : v) : v < at ? v : v >= at + k ? v - k : at;
        }
        this.zip.file(d.path, OF.serializeXml(d.doc));
      }
      // tablas con formato
      for (const t of tables) {
        const tr = t.doc.documentElement;
        fixAttr(tr, 'ref');
        for (const a of OF.all(tr, 'autoFilter')) fixAttr(a, 'ref');
        this.zip.file(t.path, OF.serializeXml(t.doc));
      }
      // gráficos de todo el libro
      for (const path of Object.keys(this.zip.files).filter((p) => /^xl\/charts\/chart\d*\.xml$/.test(p))) {
        const doc = OF.parseXml(await this.zip.file(path).async('string'));
        let touched = false;
        for (const f of OF.all(doc, 'f')) { const nt = fx(f.textContent, false); if (nt !== f.textContent) { f.textContent = nt; touched = true; } }
        if (touched) this.zip.file(path, OF.serializeXml(doc));
      }
      for (const sh of this.sheets) this.zip.file(sh.path, OF.serializeXml(sh.doc));
      this.zip.file(this.wbPath, OF.serializeXml(this.wb));
      await this.loadSheet(T);
      T.pics = null;
      this.buildEngine();
      this.undo = [];
      this.xfCache = {};
      this.markDirty();
      const pos = { ...this.sel };
      if (axis === 'row') pos.r = at; else pos.c = at;
      this.showSheet(this.cur);
      this.select(pos.r, pos.c);
      OF.toast(n > 0 ? `${axis === 'row' ? (n > 1 ? n + ' filas insertadas' : 'Fila insertada') : (n > 1 ? n + ' columnas insertadas' : 'Columna insertada')}` : `${axis === 'row' ? (k > 1 ? k + ' filas eliminadas' : 'Fila eliminada') : (k > 1 ? k + ' columnas eliminadas' : 'Columna eliminada')}`);
    },

    /* ------------------------------ copiar y pegar ------------------------------ */
    copyCells(cd) {
      const sh = this.sheets[this.cur], g = this.workRect();
      const rows = [], html = [];
      for (let r = g.r1; r <= g.r2; r++) {
        const vals = [], tds = [];
        for (let c = g.c1; c <= g.c2; c++) {
          const x = this.value(this.cur, r, c);
          const xf = this.xfs[x.s] || this.xfs[0], f = this.fonts[xf.fontId] || {}, fl = this.fills[xf.fillId] || {};
          const shown = this.display(x);
          vals.push(shown);
          let st = '';
          if (f.b) st += 'font-weight:700;'; if (f.i) st += 'font-style:italic;'; if (f.u) st += 'text-decoration:underline;';
          if (f.color) st += `color:${f.color};`; if (f.sz) st += `font-size:${f.sz}pt;`; if (f.name) st += `font-family:'${f.name}';`;
          if (fl.bg) st += `background:${fl.bg};`; if (xf.h && xf.h !== 'general') st += `text-align:${xf.h};`;
          const num = typeof x.v === 'number' ? ` x:num="${x.v}"` : '';
          const fmla = x.f != null ? ` x:fmla="=${x.f.replace(/"/g, '&quot;')}"` : '';
          tds.push(`<td style="${st}" data-fmt="${this.fmtCode(x.s).replace(/"/g, '&quot;')}"${num}${fmla}>${shown.replace(/&/g, '&amp;').replace(/</g, '&lt;')}</td>`);
        }
        rows.push(vals.join('\t'));
        html.push('<tr>' + tds.join('') + '</tr>');
      }
      cd.setData('text/plain', rows.join('\r\n'));
      cd.setData('text/html', `<table data-folio-src="${g.r1},${g.c1}">${html.join('')}</table>`);
      OF.toast(`Copiado ${g.r2 - g.r1 + 1}×${g.c2 - g.c1 + 1}`);
    },
    // Formato de número deducido del texto mostrado (cuando el origen no lo indica)
    guessFormat(text, code) {
      if (code && code !== 'General') return code;
      const t = (text || '').trim();
      if (/^-?\$/.test(t)) return /,\d{1,2}$|\.\d{1,2}$/.test(t) && !/^\$?-?\d{1,3}(\.\d{3})+$/.test(t.replace('$', '')) ? '"$"#,##0.00' : '"$"#,##0';
      if (/%$/.test(t)) return /[.,]\d+%$/.test(t) ? '0.00%' : '0%';
      if (/^\d{1,2}[-/]\d{1,2}[-/]\d{2,4}$/.test(t)) return 'dd-mm-yyyy';
      return null;
    },
    pasteCells(html, text) {
      this.stopEdit(true);
      OF.checkpoint && OF.checkpoint();
      const sh = this.sheets[this.cur];
      const start = { r: this.rangeRect().r1, c: this.rangeRect().c1 };
      let grid = null;
      if (html && /<table/i.test(html)) {
        try {
          grid = OF.withHtml(html, (d, win) => {
            const table = d.querySelector('table');
            if (!table) return null;
            const out = [];
            const src = table.getAttribute('data-folio-src');
            if (src) { const [sr, sc] = src.split(',').map(Number); out.src = { r: sr, c: sc }; }
            for (const tr of table.querySelectorAll('tr')) {
              const row = [];
              for (const td of tr.querySelectorAll('td,th')) {
                const cs = win.getComputedStyle(td);
                const inner = td.querySelector('*') || td;
                const ics = win.getComputedStyle(inner.closest('font,span,b,i,u,strong,em') || td);
                const st = {
                  b: +ics.fontWeight >= 600 || +cs.fontWeight >= 600, i: ics.fontStyle === 'italic' || cs.fontStyle === 'italic',
                  u: /underline/.test(ics.textDecorationLine) || /underline/.test(cs.textDecorationLine),
                  color: OF.toHex(ics.color), bg: OF.toHex(cs.backgroundColor), h: /center/.test(cs.textAlign) ? 'center' : /right|end/.test(cs.textAlign) ? 'right' : null,
                  sz: parseFloat(ics.fontSize) ? Math.round(parseFloat(ics.fontSize) * 0.75 * 2) / 2 : null,
                };
                if (st.color === '#010203') st.color = null;
                if (st.sz && Math.abs(st.sz - 10) < 0.01 && Math.abs(parseFloat(ics.fontSize) - 13.37) < 0.01) st.sz = null;
                // Excel guarda el formato de número en la regla mso-number-format de la clase o del estilo
                let code = td.getAttribute('data-fmt');
                if (!code) {
                  const styleText = (td.getAttribute('style') || '') + ';' + Array.from(d.querySelectorAll('style')).map((x) => x.textContent).join('\n').split('.' + (td.className || '__')).slice(1).map((x) => x.split('}')[0]).join(';');
                  const m = /mso-number-format:\s*("([^"]*)"|'([^']*)'|([^;]+))/.exec(styleText);
                  if (m) code = (m[2] || m[3] || m[4] || '').replace(/\\([0-9a-fA-F]{4})/g, (x, h) => String.fromCharCode(parseInt(h, 16))).replace(/\\(.)/g, '$1').trim();
                  if (code === 'General' || code === '@' && false) code = code;
                }
                row.push({ text: td.innerText.replace(/ /g, ' ').replace(/\n$/, ''), num: td.getAttribute('x:num'), fmla: td.getAttribute('x:fmla'), code, st, span: +(td.getAttribute('colspan') || 1) });
                for (let k = 1; k < +(td.getAttribute('colspan') || 1); k++) row.push(null);
              }
              out.push(row);
            }
            return out;
          });
        } catch (e) { console.warn(e); grid = null; }
      }
      if (!grid) grid = (text || '').replace(/\r\n?/g, '\n').replace(/\n$/, '').split('\n').map((line) => line.split('\t').map((t) => ({ text: t })));
      if (!grid.length) return;
      const keepSel = { ...this.sel };
      let n = 0;
      grid.forEach((row, dr) => row.forEach((cell, dc) => {
        if (!cell) return;
        const r = start.r + dr, c = start.c + dc;
        if (sh.covered.get(r + ',' + c) === 'x') return;
        this.sel = { r, c };
        let input = cell.num != null && cell.num !== '' ? String(cell.num) : cell.text;
        if (cell.fmla && grid.src) input = '=' + shiftFormula(cell.fmla.replace(/^=/, ''), start.r - grid.src.r, start.c - grid.src.c);
        // una celda numérica copiada como texto con formato se escribe como número
        this.commit(input, true);
        const ch = {};
        const code = this.guessFormat(cell.text, cell.code && cell.code !== 'General' ? cell.code : null);
        if (code && (cell.num != null || /^[-$\d.,% /]+$/.test(cell.text.trim()))) ch.numFmt = code;
        if (cell.st) {
          ch.font = { b: !!cell.st.b, i: !!cell.st.i, u: !!cell.st.u };
          if (cell.st.color) ch.font.color = cell.st.color;
          if (cell.st.sz) ch.font.sz = cell.st.sz;
          if (cell.st.bg && cell.st.bg !== '#FFFFFF') ch.fill = cell.st.bg;
          if (cell.st.h) ch.h = cell.st.h;
        }
        if (Object.keys(ch).length && this.stylesDoc) {
          const el = this.cellEl(sh, r, c, true);
          el.setAttribute('s', this.newXf(+(el.getAttribute('s') || 0), ch));
        }
        n++;
      }));
      this.buildStyleSheet();
      this.undo = [];
      this.markDirty();
      this.sel = keepSel;
      this.showSheetKeep();
      this.anchor = { r: start.r, c: start.c };
      this.select(start.r + grid.length - 1, start.c + Math.max(...grid.map((r) => r.length)) - 1, true);
      OF.toast(`${n} celda${n === 1 ? '' : 's'} pegada${n === 1 ? '' : 's'}`);
    },

    /* -------------------------------- imágenes -------------------------------- */
    async loadPictures(sh) {
      sh.pics = [];
      for (const d of await this.sheetParts(sh, /\/drawing$/)) {
        const rels = await OF.readRels(this.zip, d.path);
        for (const anc of Array.from(d.doc.documentElement.children)) {
          const from = OF.kid(anc, 'from', XDR);
          const pos = (e) => e && { c: +OF.kid(e, 'col', XDR).textContent, co: +OF.kid(e, 'colOff', XDR).textContent / 9525, r: +OF.kid(e, 'row', XDR).textContent, ro: +OF.kid(e, 'rowOff', XDR).textContent / 9525 };
          const item = { from: pos(from), to: pos(OF.kid(anc, 'to', XDR)), anc, d };
          const ext = OF.kid(anc, 'ext', XDR);
          if (ext) { item.w = +ext.getAttribute('cx') / 9525; item.h = +ext.getAttribute('cy') / 9525; }
          if (anc.localName === 'absoluteAnchor') { const o = OF.kid(anc, 'pos', XDR); item.abs = { x: +o.getAttribute('x') / 9525, y: +o.getAttribute('y') / 9525 }; }
          const pic = OF.kid(anc, 'pic', XDR);
          if (pic) {
            const blip = OF.all(pic, 'blip')[0];
            const rel = blip && rels[blip.getAttributeNS(OF.NS.r, 'embed')];
            if (rel && this.zip.file(rel.target) && !/\.(emf|wmf)$/i.test(rel.target)) {
              const b = await this.zip.file(rel.target).async('blob');
              item.url = URL.createObjectURL(new Blob([b], { type: OF.mimeFor(rel.target) }));
            }
          } else if (OF.kid(anc, 'graphicFrame', XDR)) item.chart = true;
          else continue;
          sh.pics.push(item);
        }
      }
    },
    colLeft(sh, c) { let x = 44; for (let i = 0; i < c; i++) { const ci = sh.cols[i]; x += ci && ci.hidden ? 0 : Math.round(((ci && ci.w) || sh.defW) * 7 + 5); } return x; },
    rowTop(sh, r) {
      const tr = this.tbody && this.tbody.querySelector(`tr[data-r="${r}"]`);
      if (tr) return tr.offsetTop;
      let y = 24; for (let i = 0; i < r; i++) { const ri = sh.rowInfo[i] || {}; y += ri.hidden ? 0 : Math.round((ri.ht || sh.defH) * 4 / 3); } return y;
    },
    async drawPictures() {
      const sh = this.sheets[this.cur];
      if (!sh.pics) await this.loadPictures(sh);
      if (this.picLayer) this.picLayer.remove();
      if (!sh.pics.length) return;
      const layer = (this.picLayer = OF.el('div', { class: 'xl-pics' }));
      for (const p of sh.pics) {
        let x, y, w = p.w, h = p.h;
        if (p.abs) { x = 44 + p.abs.x; y = 24 + p.abs.y; }
        else { x = this.colLeft(sh, p.from.c) + p.from.co; y = this.rowTop(sh, p.from.r) + p.from.ro; }
        if (p.to && w == null) { w = this.colLeft(sh, p.to.c) + p.to.co - x; h = this.rowTop(sh, p.to.r) + p.to.ro - y; }
        const el = p.url ? OF.el('img', { src: p.url, alt: '', draggable: 'false' }) : OF.el('div', { class: 'xl-chart' }, p.chart ? 'Gráfico (se conserva)' : 'Objeto');
        Object.assign(el.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
        p.shown = { w, h };
        el.addEventListener('click', (e) => {
          e.stopPropagation();
          layer.querySelectorAll('.sel').forEach((x) => x.classList.remove('sel'));
          el.classList.add('sel');
          OF.imageBar(el, [['Más pequeña', () => this.picAction(p, 0.8)], ['Más grande', () => this.picAction(p, 1.25)], ['Eliminar', () => this.picAction(p, 'del'), true]]);
        });
        layer.append(el);
      }
      this.table.style.position = 'relative';
      this.gridBox.append(layer);
      layer.style.zoom = this.userZoom || 1;
    },
    async picAction(p, a) {
      OF.checkpoint && OF.checkpoint();
      OF.closeImageBar();
      const anc = p.anc;
      if (a === 'del') anc.parentNode.removeChild(anc);
      else {
        const W = Math.round(p.shown.w * a * 9525), H = Math.round(p.shown.h * a * 9525);
        let target = anc;
        if (anc.localName === 'twoCellAnchor') {
          // un anclaje de dos celdas se pasa a anclaje de una celda con tamaño fijo
          target = anc.ownerDocument.createElementNS(XDR, 'xdr:oneCellAnchor');
          target.appendChild(OF.kid(anc, 'from', XDR));
          target.appendChild(anc.ownerDocument.createElementNS(XDR, 'xdr:ext'));
          for (const c of Array.from(anc.children)) if (c.localName !== 'to') target.appendChild(c);
          anc.parentNode.replaceChild(target, anc);
        }
        const ext = OF.kid(target, 'ext', XDR);
        if (ext) { ext.setAttribute('cx', W); ext.setAttribute('cy', H); }
        for (const e of Array.from(target.getElementsByTagNameNS(OF.NS.a, 'ext'))) if (e.hasAttribute('cx')) { e.setAttribute('cx', W); e.setAttribute('cy', H); }
      }
      this.zip.file(p.d.path, OF.serializeXml(p.d.doc));
      this.sheets[this.cur].pics = null;
      this.markDirty();
      await this.drawPictures();
    },
    async insertImage(file) {
      this.stopEdit(true);
      OF.checkpoint && OF.checkpoint();
      const im = await OF.prepareImage(file);
      const sh = this.sheets[this.cur];
      const root = sh.doc.documentElement;
      const media = OF.uniquePart(this.zip, 'xl/media', 'imagen', im.ext);
      this.zip.file(media, im.bytes);
      let drawEl = OF.kid(root, 'drawing', S), dPath, dDoc;
      if (drawEl) {
        dPath = sh.rels[drawEl.getAttributeNS(OF.NS.r, 'id')].target;
        dDoc = OF.parseXml(await this.zip.file(dPath).async('string'));
      } else {
        dPath = OF.uniquePart(this.zip, 'xl/drawings', 'drawing', 'xml');
        dDoc = OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${OF.NS.a}"/>`);
        const rid = await OF.addRel(this.zip, sh.path, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', dPath);
        drawEl = OF.ordered(root, S, root.prefix || '', 'drawing', WS_ORDER);
        if (!root.lookupPrefix(OF.NS.r)) root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:r', OF.NS.r);
        drawEl.setAttributeNS(OF.NS.r, 'r:id', rid);
        await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, dPath, 'application/vnd.openxmlformats-officedocument.drawing+xml'));
        sh.rels = await OF.readRels(this.zip, sh.path);
      }
      this.zip.file(dPath, OF.serializeXml(dDoc));
      const rId = await OF.addRel(this.zip, dPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image', media);
      await OF.contentTypes(this.zip, (ct) => OF.ensureDefaultCT(ct, im.ext, im.mime));
      const ids = OF.all(dDoc, 'cNvPr').map((e) => +e.getAttribute('id') || 0);
      const id = Math.max(1, ...ids) + 1;
      const k = Math.min(1, 360 / im.w);
      const cx = Math.round(im.w * k * 9525), cy = Math.round(im.h * k * 9525);
      const anc = OF.fragment(dDoc, `<xdr:oneCellAnchor xmlns:xdr="${XDR}" xmlns:a="${OF.NS.a}" xmlns:r="${OF.NS.r}"><xdr:from><xdr:col>${this.sel.c}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${this.sel.r}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:ext cx="${cx}" cy="${cy}"/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="${id}" name="Imagen ${id}"/><xdr:cNvPicPr><a:picLocks noChangeAspect="1"/></xdr:cNvPicPr></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>`);
      dDoc.documentElement.appendChild(anc);
      this.zip.file(dPath, OF.serializeXml(dDoc));
      this.zip.file(sh.path, OF.serializeXml(sh.doc));
      sh.pics = null;
      this.markDirty();
      await this.drawPictures();
    },
  });

  ExcelDoc.mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  OF.ExcelDoc = ExcelDoc;
})();
