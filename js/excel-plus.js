/* Excel: fórmulas en español, sugerencias de funciones, bordes, combinar, ancho/alto, ordenar,
   filtrar, autorrellenar, suma rápida, hojas, inmovilizar paneles, gráficos, buscar, deshacer. */
(function () {
  const OF = window.OF;
  const S = OF.NS.s;
  const XDR = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing';
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  const XD = OF.ExcelDoc.prototype;
  const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const colName = OF.colName;
  const colIndex = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const parseRef = (ref) => { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref); return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null; };
  const addr = (r, c) => colName(c) + (r + 1);
  const parseRange = (ref) => { const [a, b] = ref.split(':'); const A = parseRef(a), B = parseRef(b || a); return A && B ? { r1: Math.min(A.r, B.r), c1: Math.min(A.c, B.c), r2: Math.max(A.r, B.r), c2: Math.max(A.c, B.c) } : null; };
  const quoteSheet = (n) => (/^[A-Za-z_][A-Za-z0-9_.]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`);
  function shiftFormula(f, dr, dc) {
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
  OF.shiftFormula = shiftFormula;

  /* ------------------------- fórmulas en español ------------------------- */
  const FN = 'SUM:SUMA AVERAGE:PROMEDIO IF:SI IFS:SI.CONJUNTO COUNT:CONTAR COUNTA:CONTARA COUNTBLANK:CONTAR.BLANCO COUNTIF:CONTAR.SI COUNTIFS:CONTAR.SI.CONJUNTO SUMIF:SUMAR.SI SUMIFS:SUMAR.SI.CONJUNTO AVERAGEIF:PROMEDIO.SI AVERAGEIFS:PROMEDIO.SI.CONJUNTO MAX:MAX MIN:MIN MAXIFS:MAX.SI.CONJUNTO MINIFS:MIN.SI.CONJUNTO ROUND:REDONDEAR ROUNDUP:REDONDEAR.MAS ROUNDDOWN:REDONDEAR.MENOS INT:ENTERO ABS:ABS MOD:RESIDUO POWER:POTENCIA SQRT:RAIZ PRODUCT:PRODUCTO SUMPRODUCT:SUMAPRODUCTO VLOOKUP:BUSCARV HLOOKUP:BUSCARH XLOOKUP:BUSCARX LOOKUP:BUSCAR INDEX:INDICE MATCH:COINCIDIR IFERROR:SI.ERROR IFNA:SI.ND AND:Y OR:O NOT:NO TRUE:VERDADERO FALSE:FALSO LEFT:IZQUIERDA RIGHT:DERECHA MID:EXTRAE LEN:LARGO UPPER:MAYUSC LOWER:MINUSC PROPER:NOMPROPIO TRIM:ESPACIOS CONCATENATE:CONCATENAR CONCAT:CONCAT TEXTJOIN:UNIRCADENAS TEXT:TEXTO VALUE:VALOR FIND:ENCONTRAR SEARCH:HALLAR SUBSTITUTE:SUSTITUIR REPLACE:REEMPLAZAR TODAY:HOY NOW:AHORA DATE:FECHA YEAR:AÑO MONTH:MES DAY:DIA WEEKDAY:DIASEM DATEDIF:SIFECHA EOMONTH:FIN.MES EDATE:FECHA.MES NETWORKDAYS:DIAS.LAB WORKDAY:DIA.LAB MEDIAN:MEDIANA MODE:MODA STDEV:DESVEST LARGE:K.ESIMO.MAYOR SMALL:K.ESIMO.MENOR RANK:JERARQUIA RAND:ALEATORIO RANDBETWEEN:ALEATORIO.ENTRE ISBLANK:ESBLANCO ISNUMBER:ESNUMERO ISTEXT:ESTEXTO ISERROR:ESERROR CHOOSE:ELEGIR PMT:PAGO FV:VF PV:VA NPV:VNA IRR:TIR ROWS:FILAS COLUMNS:COLUMNAS ROW:FILA COLUMN:COLUMNA UNIQUE:UNICOS FILTER:FILTRAR SORT:ORDENAR TRANSPOSE:TRANSPONER'.split(' ').map((x) => x.split(':'));
  const EN2ES = new Map(FN), ES2EN = new Map(FN.map(([a, b]) => [b, a]));
  ES2EN.set('DÍA', 'DAY'); ES2EN.set('ANO', 'YEAR'); ES2EN.set('RAÍZ', 'SQRT'); ES2EN.set('ÍNDICE', 'INDEX');
  const HELP = { SUMA: 'SUMA(A1:A10)', PROMEDIO: 'PROMEDIO(A1:A10)', SI: 'SI(prueba; si_verdadero; si_falso)', CONTAR: 'CONTAR(A1:A10)', CONTARA: 'CONTARA(A1:A10)', 'CONTAR.SI': 'CONTAR.SI(rango; criterio)', 'SUMAR.SI': 'SUMAR.SI(rango; criterio; [rango_suma])', BUSCARV: 'BUSCARV(valor; tabla; columna; FALSO)', BUSCARX: 'BUSCARX(valor; buscar_en; devolver)', REDONDEAR: 'REDONDEAR(número; decimales)', 'SI.ERROR': 'SI.ERROR(valor; si_error)', MAX: 'MAX(A1:A10)', MIN: 'MIN(A1:A10)', HOY: 'HOY()', CONCATENAR: 'CONCATENAR(texto1; texto2)' };
  // reemplaza fuera de textos entre comillas y nombres de hoja entre apóstrofos
  const outside = (f, fn) => f.split(/("(?:[^"]|"")*"|'(?:[^']|'')*')/).map((seg, i) => (i % 2 ? seg : fn(seg))).join('');
  OF.formulaToES = (f) => outside(outside(f, (seg) => seg.replace(/([A-Z][A-Z0-9.]*)(?=\s*\()/gi, (m) => EN2ES.get(m.toUpperCase()) || m).replace(/(?<![A-Za-z0-9_.$])(TRUE|FALSE)(?![A-Za-z0-9_(])/g, (m) => EN2ES.get(m)).replace(/,/g, ';')), (seg) => seg.replace(/(?<![A-Za-z$_\d])(\d+)\.(\d+)/g, '$1,$2'));
  OF.formulaToEN = (f) => {
    let es = false;
    outside(f, (seg) => { if (/;/.test(seg)) es = true; return seg; });
    return outside(f, (seg) => {
      let x = seg.replace(/([A-ZÁÉÍÓÚÑ][A-ZÁÉÍÓÚÑ0-9.]*)(?=\s*\()/gi, (m) => ES2EN.get(m.toUpperCase()) || m.toUpperCase());
      x = x.replace(/(?<![A-Za-z0-9_.$])(VERDADERO|FALSO)(?![A-Za-z0-9_(])/gi, (m) => ES2EN.get(m.toUpperCase()));
      if (es) x = x.replace(/(?<![A-Za-z$_\d])(\d+),(\d+)/g, '$1.$2').replace(/;/g, ',');
      return x;
    });
  };
  const baseInput = XD.inputText, baseParse = XD.parseInput;
  XD.inputText = function (r, c) {
    const t = baseInput.call(this, r, c);
    return this.locale === 'cl' && typeof t === 'string' && t.startsWith('=') ? '=' + OF.formulaToES(t.slice(1)) : t;
  };
  XD.parseInput = function (str) {
    const p = baseParse.call(this, str);
    if (p.type === 'formula') p.f = OF.formulaToEN(p.f);
    return p;
  };
  // sugerencias mientras se escribe una fórmula
  XD.suggest = function (input, anchorEl) {
    if (this._sug) { this._sug.remove(); this._sug = null; }
    const v = input.value.slice(0, input.selectionStart);
    if (!v.startsWith('=')) return;
    const m = /(?:^=|[\s(;,+\-*/^&=<>])([A-Za-zÁÉÍÓÚÑáéíóúñ][A-Za-z0-9.ÁÉÍÓÚÑáéíóúñ]*)$/.exec(v);
    if (!m) return;
    const tok = m[1].toUpperCase();
    const names = (this.locale === 'cl' ? FN.map((x) => x[1]) : FN.map((x) => x[0])).filter((n) => n.startsWith(tok) && n !== tok).slice(0, 7);
    if (!names.length) return;
    const box = (this._sug = OF.el('div', { class: 'xl-sug', role: 'listbox' }));
    const pick = (n) => {
      const before = v.slice(0, v.length - m[1].length) + n + '(';
      input.value = before + input.value.slice(input.selectionStart);
      input.setSelectionRange(before.length, before.length);
      input.focus();
      input.dispatchEvent(new Event('input'));
      if (this._sug) { this._sug.remove(); this._sug = null; }
    };
    for (const n of names) {
      const b = OF.el('button', { type: 'button', role: 'option' }, OF.el('b', null, n), HELP[n] ? OF.el('span', null, HELP[n]) : null);
      b.addEventListener('mousedown', (e) => { e.preventDefault(); pick(n); });
      box.append(b);
    }
    box._pick = () => pick(names[0]);
    document.body.append(box);
    const r = anchorEl.getBoundingClientRect();
    box.style.left = Math.min(r.left, innerWidth - 300) + 'px';
    box.style.top = r.bottom + 2 + 'px';
  };
  const baseStartEdit = XD.startEdit;
  XD.startEdit = function (initial) {
    baseStartEdit.call(this, initial);
    const ed = this.editor;
    if (!ed) return;
    ed.addEventListener('input', () => this.suggest(ed, ed));
    ed.addEventListener('keydown', (e) => { if (e.key === 'Tab' && this._sug) { e.preventDefault(); e.stopImmediatePropagation(); this._sug._pick(); } }, true);
    ed.addEventListener('blur', () => setTimeout(() => { if (this._sug) { this._sug.remove(); this._sug = null; } }, 150));
    this.suggest(ed, ed);
  };

  /* --------------------------- carga y vista --------------------------- */
  const baseLoadSheet = XD.loadSheet;
  XD.loadSheet = async function (sh) {
    await baseLoadSheet.call(this, sh);
    const root = sh.doc.documentElement;
    const af = OF.kid(root, 'autoFilter', S);
    sh.autoFilter = af && af.getAttribute('ref') ? parseRange(af.getAttribute('ref')) : null;
    const pane = OF.all(OF.all(root, 'sheetView', S)[0] || root, 'pane', S)[0];
    sh.freeze = pane && /frozen/.test(pane.getAttribute('state') || '') ? { rows: +(pane.getAttribute('ySplit') || 0), cols: +(pane.getAttribute('xSplit') || 0) } : { rows: 0, cols: 0 };
    sh.pics = null;
  };
  const baseBuildUI = XD.buildUI;
  XD.buildUI = function () {
    baseBuildUI.call(this);
    this.fx.addEventListener('input', () => this.suggest(this.fx, this.fx));
    this.fx.addEventListener('keydown', (e) => { if (e.key === 'Tab' && this._sug) { e.preventDefault(); this._sug._pick(); } });
    this.fx.addEventListener('blur', () => setTimeout(() => { if (this._sug) { this._sug.remove(); this._sug = null; } }, 150));
    const add = OF.el('button', { class: 'xl-tab xl-add', title: 'Agregar hoja', 'aria-label': 'Agregar hoja' }, '+');
    add.addEventListener('click', () => this.sheetOp('add'));
    this.tabs.append(add);
    this.tabs.querySelectorAll('.xl-tab[data-i]').forEach((t) => {
      const i = +t.dataset.i;
      t.addEventListener('contextmenu', (e) => { e.preventDefault(); this.showSheet(i); this.sheetMenu(t, i); });
      t.addEventListener('dblclick', () => this.sheetOp('rename', i));
    });
    // cambiar ancho de columnas y alto de filas arrastrando el borde del encabezado
    this.gridBox.addEventListener('pointerdown', (e) => {
      const th = e.target.closest('th');
      if (!th || !this.table.contains(th)) return;
      const rect = th.getBoundingClientRect();
      const z = this.userZoom || 1;
      const colEdge = th.dataset.hc != null && e.clientX > rect.right - 7 * z;
      const rowEdge = th.parentNode.dataset.r != null && th === th.parentNode.firstChild && e.clientY > rect.bottom - 6 * z;
      if (!colEdge && !rowEdge) return;
      e.preventDefault(); e.stopPropagation();
      const sh = this.sheets[this.cur];
      if (colEdge) {
        const c = +th.dataset.hc, col = this.table.querySelectorAll('colgroup col')[c + 1];
        const start = e.clientX, w0 = parseFloat(col.style.width);
        const mv = (ev) => { col.style.width = Math.max(8, w0 + (ev.clientX - start) / z) + 'px'; };
        const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); const w = parseFloat(col.style.width); if (Math.abs(w - w0) > 1) this.setColWidth([c], (w - 5) / 7); };
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
      } else {
        const r = +th.parentNode.dataset.r, tr = th.parentNode;
        const start = e.clientY, h0 = tr.offsetHeight;
        const mv = (ev) => { tr.style.height = Math.max(8, h0 + (ev.clientY - start) / z) + 'px'; };
        const up = () => { window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up); const h = parseFloat(tr.style.height); if (Math.abs(h - h0) > 1) this.setRowHeight([r], h * 0.75); };
        window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
      }
    }, true);
    this.gridBox.addEventListener('mousemove', (e) => {
      const th = e.target.closest('th');
      if (!th) { this.gridBox.style.cursor = ''; return; }
      const rect = th.getBoundingClientRect();
      this.gridBox.style.cursor = th.dataset.hc != null && e.clientX > rect.right - 7 ? 'col-resize' : th.parentNode.dataset.r != null && e.clientY > rect.bottom - 6 ? 'row-resize' : '';
    });
    this.gridBox.addEventListener('dblclick', (e) => {
      const th = e.target.closest('thead th[data-hc]');
      if (th && e.clientX > th.getBoundingClientRect().right - 8) this.autofit([+th.dataset.hc]);
    });
    // atajos de Excel
    this.gridBox.addEventListener('keydown', (e) => {
      if (this.editor) return;
      const k = e.key.toLowerCase();
      if ((e.ctrlKey || e.metaKey) && (k === 'n' || k === 'k' || k === 'i' || k === 'u')) { e.preventDefault(); this.toggle(k === 'n' ? 'b' : k === 'u' ? 'u' : 'i'); this.onselect && this.onselect(); }
      if ((e.ctrlKey || e.metaKey) && k === 'd') { e.preventDefault(); this.fillDir('down'); }
      if ((e.ctrlKey || e.metaKey) && k === 'r') { e.preventDefault(); this.fillDir('right'); }
    }, true);
  };
  const baseShowSheet = XD.showSheet;
  XD.showSheet = function (i) {
    baseShowSheet.call(this, i);
    this.applyFreeze();
    this.onselect && this.onselect();
  };
  const baseRenderRows = XD.renderRows;
  XD.renderRows = function (upto) { const from = this.rendered; baseRenderRows.call(this, upto); if (from) this.applyFreeze(); };
  const baseFillCell = XD.fillCell;
  XD.fillCell = function (td, sh, r, c) {
    baseFillCell.call(this, td, sh, r, c);
    const af = sh.autoFilter;
    if (af && r === af.r1 && c >= af.c1 && c <= af.c2) {
      const active = sh.filters && sh.filters[c] != null;
      const b = OF.el('button', { class: 'xl-fbtn' + (active ? ' on' : ''), type: 'button', title: 'Filtrar', 'aria-label': 'Filtrar columna ' + colName(c) }, active ? '⏷' : '▾');
      b.addEventListener('click', (e) => { e.stopPropagation(); this.filterMenu(b, c); });
      td.append(b);
      td.classList.add('has-fbtn');
    }
  };
  const baseSelect = XD.select;
  XD.select = function (r, c, extend) {
    baseSelect.call(this, r, c, extend);
    this.placeHandle();
  };

  /* -------------------------- inmovilizar paneles -------------------------- */
  XD.applyFreeze = function () {
    const sh = this.sheets[this.cur];
    if (!sh || !this.table) return;
    const { rows, cols } = sh.freeze || { rows: 0, cols: 0 };
    this.table.classList.toggle('frozen', !!(rows || cols));
    if (!rows && !cols) return;
    const head = this.table.tHead.offsetHeight;
    const lefts = []; let x = 44;
    const colsEl = this.table.querySelectorAll('colgroup col');
    for (let c = 0; c < cols; c++) { lefts[c] = x; x += parseFloat(colsEl[c + 1].style.width) || 0; }
    for (let r = 0; r < Math.min(rows, this.rendered); r++) {
      const tr = this.tbody.querySelector(`tr[data-r="${r}"]`);
      if (!tr) continue;
      const top = tr.offsetTop;
      tr.querySelectorAll('td, th').forEach((cell) => { cell.style.position = 'sticky'; cell.style.top = top + 'px'; cell.classList.add('frz'); });
      if (r === rows - 1) tr.classList.add('frz-last');
      tr.firstChild.style.zIndex = 5;
      void head;
    }
    if (cols) for (const tr of this.tbody.children) {
      for (const td of tr.querySelectorAll('td')) {
        const c = +td.dataset.c;
        if (c >= cols) break;
        td.style.position = 'sticky'; td.style.left = lefts[c] + 'px'; td.classList.add('frz');
        td.style.zIndex = +tr.dataset.r < rows ? 4 : 3;
        if (c + (td.colSpan || 1) === cols) td.classList.add('frz-lastc');
      }
    }
    for (let r = 0; r < Math.min(rows, this.rendered); r++) { const tr = this.tbody.querySelector(`tr[data-r="${r}"]`); if (tr) tr.querySelectorAll('td').forEach((td) => { if (!td.style.zIndex) td.style.zIndex = 3; }); }
  };
  XD.freeze = function (mode) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur];
    const rows = mode === 'row' ? 1 : mode === 'col' ? 0 : mode === 'cell' ? this.sel.r : 0;
    const cols = mode === 'col' ? 1 : mode === 'row' ? 0 : mode === 'cell' ? this.sel.c : 0;
    const root = sh.doc.documentElement;
    let views = OF.kid(root, 'sheetViews', S);
    if (!views) { views = OF.ordered(root, S, root.prefix || '', 'sheetViews', WS_ORDER); }
    let sv = OF.kid(views, 'sheetView', S);
    if (!sv) { sv = sh.doc.createElementNS(S, 'sheetView'); sv.setAttribute('workbookViewId', '0'); views.appendChild(sv); }
    OF.removeKids(sv, ['pane'], S);
    for (const sel of OF.kids(sv, 'selection', S)) if (sel.getAttribute('pane')) sv.removeChild(sel);
    if (rows || cols) {
      const pane = sh.doc.createElementNS(S, 'pane');
      if (cols) pane.setAttribute('xSplit', cols);
      if (rows) pane.setAttribute('ySplit', rows);
      pane.setAttribute('topLeftCell', addr(rows, cols));
      pane.setAttribute('activePane', rows && cols ? 'bottomRight' : rows ? 'bottomLeft' : 'topRight');
      pane.setAttribute('state', 'frozen');
      sv.insertBefore(pane, sv.firstChild);
    }
    sh.freeze = { rows, cols };
    this.markDirty();
    this.showSheetKeep();
    OF.toast(rows || cols ? 'Paneles inmovilizados' : 'Paneles movilizados');
  };

  /* -------------------------- ancho y alto -------------------------- */
  XD.setColWidth = function (cs, width) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], root = sh.doc.documentElement;
    let cols = OF.kid(root, 'cols', S);
    if (!cols) cols = OF.ordered(root, S, root.prefix || '', 'cols', WS_ORDER);
    for (const c of cs) {
      // separar rangos <col> que incluyan la columna
      let style = null;
      for (const col of OF.kids(cols, 'col', S)) {
        const a = +col.getAttribute('min') - 1, b = +col.getAttribute('max') - 1;
        if (c < a || c > b) continue;
        style = col.getAttribute('style');
        if (a < c) { const l = col.cloneNode(true); l.setAttribute('max', c); cols.insertBefore(l, col); }
        if (b > c) { const rr = col.cloneNode(true); rr.setAttribute('min', c + 2); cols.insertBefore(rr, col.nextSibling); }
        cols.removeChild(col);
      }
      const e = sh.doc.createElementNS(S, 'col');
      e.setAttribute('min', c + 1); e.setAttribute('max', c + 1);
      e.setAttribute('width', Math.max(0.5, Math.round(width * 100) / 100)); e.setAttribute('customWidth', '1');
      if (style) e.setAttribute('style', style);
      let before = null;
      for (const col of OF.kids(cols, 'col', S)) if (+col.getAttribute('min') > c + 1) { before = col; break; }
      cols.insertBefore(e, before);
      sh.cols[c] = { ...(sh.cols[c] || {}), w: width, hidden: false };
    }
    this.markDirty();
    this.showSheetKeep();
  };
  XD.setRowHeight = function (rs, pts) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur];
    for (const r of rs) {
      let row = sh.rowEls.get(r);
      if (!row) { this.cellEl(sh, r, 0, true); row = sh.rowEls.get(r); const c0 = sh.cells.get(r + ',0'); if (c0 && !c0.firstElementChild && !c0.getAttribute('s')) { row.removeChild(c0); sh.cells.delete(r + ',0'); } }
      row.setAttribute('ht', Math.max(1, Math.round(pts * 4) / 4)); row.setAttribute('customHeight', '1');
      sh.rowInfo[r] = { ...(sh.rowInfo[r] || {}), ht: pts };
    }
    this.markDirty();
    this.showSheetKeep();
  };
  XD.autofit = function (cs) {
    const sh = this.sheets[this.cur];
    const cv = document.createElement('canvas').getContext('2d');
    const out = [];
    for (const c of cs) {
      let w = 0;
      for (let r = 0; r <= sh.maxR; r++) {
        const x = this.value(this.cur, r, c);
        const t = this.display(x);
        if (!t) continue;
        const td = this.tdAt(r, c);
        if (td && td.colSpan > 1) continue;
        cv.font = td ? getComputedStyle(td).font : '11pt Calibri';
        w = Math.max(w, cv.measureText(t).width);
      }
      out.push([c, Math.max(4, (w + 12) / 7)]);
    }
    OF.checkpoint && OF.checkpoint();
    for (const [c, w] of out) this.setColWidth([c], w);
  };
  XD.promptSize = function (kind) {
    const g = this.workRect();
    if (kind === 'col') {
      const cur = ((this.sheets[this.cur].cols[g.c1] || {}).w || this.sheets[this.cur].defW).toFixed(1);
      const v = prompt('Ancho de columna (caracteres):', cur.replace('.', ','));
      if (v && !isNaN(parseFloat(v.replace(',', '.')))) { const cs = []; for (let c = g.c1; c <= g.c2; c++) cs.push(c); this.setColWidth(cs, parseFloat(v.replace(',', '.'))); }
    } else {
      const cur = ((this.sheets[this.cur].rowInfo[g.r1] || {}).ht || this.sheets[this.cur].defH).toFixed(1);
      const v = prompt('Alto de fila (puntos):', cur.replace('.', ','));
      if (v && !isNaN(parseFloat(v.replace(',', '.')))) { const rs = []; for (let r = g.r1; r <= Math.min(g.r2, this.sheets[this.cur].maxR + 1); r++) rs.push(r); this.setRowHeight(rs, parseFloat(v.replace(',', '.'))); }
    }
  };
  XD.autofitSelection = function () { const g = this.workRect(); const cs = []; for (let c = g.c1; c <= g.c2; c++) cs.push(c); this.autofit(cs); };

  /* -------------------------- bordes y combinar -------------------------- */
  XD.setBorders = function (preset) {
    if (!this.stylesDoc) return;
    this.stopEdit(true);
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], g = this.workRect();
    const st = preset === 'thick' ? 'medium' : 'thin';
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      if (sh.covered.get(r + ',' + c) === 'x') continue;
      const m = sh.covered.get(r + ',' + c);
      const r2 = m ? m.r2 : r, c2 = m ? m.c2 : c;
      let sides;
      if (preset === 'all') sides = { left: 'thin', right: 'thin', top: 'thin', bottom: 'thin' };
      else if (preset === 'none') sides = { left: null, right: null, top: null, bottom: null };
      else if (preset === 'bottom') { if (r2 !== g.r2) continue; sides = { bottom: 'thin' }; }
      else if (preset === 'doubleBottom') { if (r2 !== g.r2) continue; sides = { bottom: 'double' }; }
      else { sides = {}; if (r === g.r1) sides.top = st; if (r2 === g.r2) sides.bottom = st; if (c === g.c1) sides.left = st; if (c2 === g.c2) sides.right = st; if (!Object.keys(sides).length) continue; }
      const el = this.cellEl(sh, r, c, true);
      el.setAttribute('s', this.newXf(+(el.getAttribute('s') || 0), { sides }));
      // la celda vecina comparte el borde: quitarlo también allí al borrar
      if (preset === 'none') {
        for (const [rr, cc, side] of [[r - 1, c, 'bottom'], [r, c - 1, 'right'], [r2 + 1, c, 'top'], [r, c2 + 1, 'left']]) {
          if (rr < 0 || cc < 0 || (rr >= g.r1 && rr <= g.r2 && cc >= g.c1 && cc <= g.c2)) continue;
          const ne = sh.cells.get(rr + ',' + cc);
          if (ne && (this.borders[(this.xfs[+(ne.getAttribute('s') || 0)] || {}).borderId] || {})[side]) ne.setAttribute('s', this.newXf(+(ne.getAttribute('s') || 0), { sides: { [side]: null } }));
        }
      }
    }
    this.buildStyleSheet();
    this.repaintRect(g);
    this.markDirty();
  };
  XD.merge = async function (mode) {
    this.stopEdit(true);
    const sh = this.sheets[this.cur], g = this.workRect();
    const root = sh.doc.documentElement;
    let mc = OF.kid(root, 'mergeCells', S);
    const overlaps = (m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2);
    if (mode === 'split') {
      if (!mc || !sh.merges.some(overlaps)) { OF.toast('No hay celdas combinadas en la selección.'); return; }
      OF.checkpoint && OF.checkpoint();
      for (const e of OF.kids(mc, 'mergeCell', S)) { const m = parseRange(e.getAttribute('ref')); if (m && overlaps(m)) mc.removeChild(e); }
    } else {
      if (g.r1 === g.r2 && g.c1 === g.c2) { OF.toast('Selecciona varias celdas para combinarlas.'); return; }
      let lost = 0;
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) if ((r !== g.r1 || c !== g.c1) && this.inputText(r, c) !== '') lost++;
      if (lost && !confirm('Al combinar solo se conserva el valor de la celda superior izquierda. ¿Continuar?')) return;
      OF.checkpoint && OF.checkpoint();
      if (!mc) mc = OF.ordered(root, S, root.prefix || '', 'mergeCells', WS_ORDER);
      for (const e of OF.kids(mc, 'mergeCell', S)) { const m = parseRange(e.getAttribute('ref')); if (m && overlaps(m)) mc.removeChild(e); }
      for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
        if (r === g.r1 && c === g.c1) continue;
        const el = sh.cells.get(r + ',' + c);
        if (el) { for (const ch of Array.from(el.children)) if (['f', 'v', 'is'].includes(ch.localName)) el.removeChild(ch); el.removeAttribute('t'); }
      }
      const e = sh.doc.createElementNS(S, 'mergeCell');
      e.setAttribute('ref', addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2));
      mc.appendChild(e);
      if (mode === 'center') { const el = this.cellEl(sh, g.r1, g.c1, true); el.setAttribute('s', this.newXf(+(el.getAttribute('s') || 0), { h: 'center' })); this.buildStyleSheet(); }
    }
    const n = OF.kids(mc, 'mergeCell', S).length;
    if (n) mc.setAttribute('count', n); else root.removeChild(mc);
    await this.reloadSheet();
  };
  XD.reloadSheet = async function () {
    const sh = this.sheets[this.cur];
    this.zip.file(sh.path, OF.serializeXml(sh.doc));
    await this.loadSheet(sh);
    this.buildEngine();
    this.markDirty();
    this.showSheetKeep();
  };

  /* ----------------------------- ordenar ----------------------------- */
  // Región de datos alrededor de la celda (como Ctrl+A en Excel)
  XD.dataRegion = function () {
    if (this.isRange()) return this.workRect();
    const sh = this.sheets[this.cur];
    const has = (r, c) => r >= 0 && c >= 0 && this.inputText(r, c) !== '';
    let g = { r1: this.sel.r, r2: this.sel.r, c1: this.sel.c, c2: this.sel.c };
    for (let grew = true; grew;) {
      grew = false;
      const tryDir = (dr, dc) => {
        const n = { ...g };
        if (dr < 0) n.r1--; if (dr > 0) n.r2++; if (dc < 0) n.c1--; if (dc > 0) n.c2++;
        const rr = dr ? (dr < 0 ? n.r1 : n.r2) : null, cc = dc ? (dc < 0 ? n.c1 : n.c2) : null;
        let any = false;
        if (rr != null) for (let c = n.c1; c <= n.c2; c++) if (has(rr, c)) any = true;
        if (cc != null) for (let r = n.r1; r <= n.r2; r++) if (has(r, cc)) any = true;
        if (any) { g = n; grew = true; }
      };
      tryDir(-1, 0); tryDir(1, 0); tryDir(0, -1); tryDir(0, 1);
      if (g.r2 > sh.maxR + 1 || g.c2 > sh.maxC + 1) break;
    }
    return g;
  };
  XD.hasHeader = function (g) {
    if (g.r1 === g.r2) return false;
    const sh = this.sheets[this.cur];
    let textTop = 0, numBelow = 0, bold = false;
    for (let c = g.c1; c <= g.c2; c++) {
      const a = this.raw(sh, g.r1, c), b = this.raw(sh, g.r1 + 1, c);
      if (typeof a.v === 'string' && a.f == null) textTop++;
      if (typeof b.v === 'number' || b.f != null) numBelow++;
      if ((this.fonts[(this.xfs[a.s] || {}).fontId] || {}).b) bold = true;
    }
    return bold || (textTop === g.c2 - g.c1 + 1 && numBelow > 0) || (textTop === g.c2 - g.c1 + 1 && g.r2 - g.r1 > 1 && (() => { for (let c = g.c1; c <= g.c2; c++) { const a = this.inputText(g.r1, c), b = this.inputText(g.r1 + 1, c); if (a && b && a === b) return false; } return true; })());
  };
  XD.sortRange = async function (dir) {
    this.stopEdit(true);
    const sh = this.sheets[this.cur];
    const region = this.dataRegion();
    const g = { ...region };
    if (sh.merges.some((m) => !(m.r2 < g.r1 || m.r1 > g.r2 || m.c2 < g.c1 || m.c1 > g.c2))) { OF.toast('No se puede ordenar un rango con celdas combinadas.', 4000); return; }
    const header = sh.autoFilter && sh.autoFilter.r1 === g.r1 ? true : this.hasHeader(g);
    if (header) g.r1++;
    if (g.r2 <= g.r1) { OF.toast('No hay suficientes filas para ordenar.'); return; }
    const key = Math.min(Math.max(this.sel.c, g.c1), g.c2);
    OF.checkpoint && OF.checkpoint();
    for (const si of Object.keys(sh.shared)) this.expandShared(sh, si);
    const rows = [];
    for (let r = g.r1; r <= g.r2; r++) {
      const cells = [];
      for (let c = g.c1; c <= g.c2; c++) { const el = sh.cells.get(r + ',' + c); cells.push(el || null); }
      const v = this.value(this.cur, r, key).v;
      rows.push({ r, cells, v, hidden: sh.rowInfo[r] && sh.rowInfo[r].hidden });
    }
    const rank = (v) => (v == null || v === '' ? 4 : typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : typeof v === 'boolean' ? 2 : 3);
    rows.sort((a, b) => {
      const ra = rank(a.v), rb = rank(b.v);
      if (ra === 4 || rb === 4) return ra - rb; // vacíos siempre al final
      if (ra !== rb) return (ra - rb) * (dir === 'asc' ? 1 : -1);
      const cmp = ra === 0 ? a.v - b.v : ra === 1 ? String(a.v).localeCompare(String(b.v), 'es', { sensitivity: 'base', numeric: true }) : ra === 2 ? (a.v === b.v ? 0 : a.v ? 1 : -1) : 0;
      return (dir === 'asc' ? cmp : -cmp) || a.r - b.r;
    });
    // quitar celdas y volver a ubicarlas
    for (const row of rows) for (const el of row.cells) if (el) { el.parentNode.removeChild(el); }
    rows.forEach((row, i) => {
      const nr = g.r1 + i, dr = nr - row.r;
      row.cells.forEach((el, k) => {
        if (!el) return;
        const c = g.c1 + k;
        el.setAttribute('r', addr(nr, c));
        const f = OF.kid(el, 'f', S);
        if (f && f.textContent && dr) f.textContent = shiftFormula(f.textContent, dr, 0);
        this.placeCell(sh, nr, c, el);
      });
    });
    await this.reloadSheet();
    OF.toast(`Ordenado ${dir === 'asc' ? 'de A a Z' : 'de Z a A'} por la columna ${colName(key)}${header ? ' (con encabezado)' : ''}`);
  };
  // inserta un elemento <c> en la fila correcta y en orden de columna
  XD.placeCell = function (sh, r, c, el) {
    let row = sh.rowEls.get(r);
    if (!row) {
      row = sh.doc.createElementNS(S, 'row');
      row.setAttribute('r', r + 1);
      let next = null;
      for (const [rr, re] of sh.rowEls) if (rr > r && (!next || rr < next[0])) next = [rr, re];
      sh.data.insertBefore(row, next ? next[1] : null);
      sh.rowEls.set(r, row);
    }
    row.removeAttribute('spans');
    let next = null;
    for (const ce of OF.kids(row, 'c', S)) { const p = parseRef(ce.getAttribute('r')); if (p && p.c > c) { next = ce; break; } }
    row.insertBefore(el, next);
  };

  /* ----------------------------- filtrar ----------------------------- */
  XD.toggleFilter = async function () {
    this.stopEdit(true);
    const sh = this.sheets[this.cur], root = sh.doc.documentElement;
    OF.checkpoint && OF.checkpoint();
    const af = OF.kid(root, 'autoFilter', S);
    if (af) {
      const g = sh.autoFilter;
      root.removeChild(af);
      if (g) for (let r = g.r1 + 1; r <= g.r2; r++) { const row = sh.rowEls.get(r); if (row && row.getAttribute('hidden') === '1') row.removeAttribute('hidden'); }
      sh.filters = null;
      await this.reloadSheet();
      OF.toast('Filtro quitado');
      return;
    }
    const g = this.dataRegion();
    if (g.r2 === g.r1) { OF.toast('Selecciona una tabla con encabezados para filtrar.'); return; }
    const e = OF.ordered(root, S, root.prefix || '', 'autoFilter', WS_ORDER);
    e.setAttribute('ref', addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2));
    sh.filters = {};
    await this.reloadSheet();
    OF.toast('Filtro activado: usa ▾ en los encabezados');
  };
  XD.filterMenu = function (btn, c) {
    const sh = this.sheets[this.cur], g = sh.autoFilter;
    const vals = new Map();
    for (let r = g.r1 + 1; r <= g.r2; r++) { const t = this.display(this.value(this.cur, r, c)); vals.set(t, (vals.get(t) || 0) + 1); }
    const list = Array.from(vals.keys()).sort((a, b) => (a === '' ? 1 : b === '' ? -1 : a.localeCompare(b, 'es', { numeric: true })));
    const cur = sh.filters && sh.filters[c];
    OF.popup(btn, (p, close) => {
      p.classList.add('flt');
      p.append(OF.el('button', { class: 'pop-item', onclick: () => { close(); this.sel = { r: g.r1 + 1, c }; this.anchor = null; this.sortRange('asc'); } }, 'Ordenar de A a Z'));
      p.append(OF.el('button', { class: 'pop-item', onclick: () => { close(); this.sel = { r: g.r1 + 1, c }; this.anchor = null; this.sortRange('desc'); } }, 'Ordenar de Z a A'));
      p.append(OF.el('hr'));
      const search = OF.el('input', { class: 'flt-q', placeholder: 'Buscar', type: 'search' });
      const boxes = OF.el('div', { class: 'flt-list' });
      const all = OF.el('input', { type: 'checkbox' });
      all.checked = !cur;
      const items = list.map((v) => { const cb = OF.el('input', { type: 'checkbox' }); cb.checked = !cur || cur.includes(v); return { v, cb, row: OF.el('label', null, cb, OF.el('span', null, v === '' ? '(Vacías)' : v)) }; });
      all.addEventListener('change', () => items.forEach((i) => { if (i.row.style.display !== 'none') i.cb.checked = all.checked; }));
      search.addEventListener('input', () => { const q = search.value.toLowerCase(); items.forEach((i) => { i.row.style.display = i.v.toLowerCase().includes(q) ? '' : 'none'; }); });
      boxes.append(OF.el('label', { class: 'flt-all' }, all, OF.el('span', null, '(Seleccionar todo)')), ...items.map((i) => i.row));
      p.append(search, boxes);
      const ok = OF.el('button', { class: 'btn primary', onclick: () => { close(); const sel = items.filter((i) => i.cb.checked).map((i) => i.v); this.applyFilter(c, sel.length === items.length ? null : sel); } }, 'Aceptar');
      const clear = OF.el('button', { class: 'btn', onclick: () => { close(); this.applyFilter(c, null); } }, 'Borrar filtro');
      p.append(OF.el('div', { class: 'flt-btns' }, clear, ok));
    });
  };
  XD.applyFilter = async function (c, allowed) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], g = sh.autoFilter, root = sh.doc.documentElement;
    sh.filters = sh.filters || {};
    if (allowed) sh.filters[c] = allowed; else delete sh.filters[c];
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      let show = true;
      for (const cc in sh.filters) if (!sh.filters[cc].includes(this.display(this.value(this.cur, r, +cc)))) show = false;
      let row = sh.rowEls.get(r);
      if (!show) { if (!row) { this.cellEl(sh, r, g.c1, true); row = sh.rowEls.get(r); } row.setAttribute('hidden', '1'); }
      else if (row) row.removeAttribute('hidden');
    }
    const af = OF.kid(root, 'autoFilter', S);
    for (const fc of OF.kids(af, 'filterColumn', S)) af.removeChild(fc);
    for (const cc of Object.keys(sh.filters).map(Number).sort((a, b) => a - b)) {
      const fc = sh.doc.createElementNS(S, 'filterColumn');
      fc.setAttribute('colId', cc - g.c1);
      const fs = sh.doc.createElementNS(S, 'filters');
      for (const v of sh.filters[cc]) { if (v === '') { fs.setAttribute('blank', '1'); continue; } const f = sh.doc.createElementNS(S, 'filter'); f.setAttribute('val', v); fs.appendChild(f); }
      fc.appendChild(fs);
      af.appendChild(fc);
    }
    const keep = sh.filters;
    await this.reloadSheet();
    sh.filters = keep;
    this.showSheetKeep();
    let shown = 0; for (let r = g.r1 + 1; r <= g.r2; r++) if (!(sh.rowInfo[r] || {}).hidden) shown++;
    OF.toast(`${shown} de ${g.r2 - g.r1} filas visibles`);
  };

  /* --------------------------- autorrellenar --------------------------- */
  const DAYS = ['lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado', 'domingo'];
  const MONTHS = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
  const MONTHS3 = MONTHS.map((m) => m.slice(0, 3));
  const caseLike = (src, s) => (src === src.toUpperCase() ? s.toUpperCase() : src[0] === src[0].toUpperCase() ? s[0].toUpperCase() + s.slice(1) : s);
  // valor en la posición k (k = 0,1,2... desde el inicio del patrón) a partir de los valores de origen
  XD.seriesValue = function (src, k) {
    const n = src.length;
    const all = src.every((x) => x.type === 'num');
    if (all && n >= 2) {
      const step = (src[n - 1].v - src[0].v) / (n - 1);
      return { type: 'num', v: +(src[0].v + step * k).toPrecision(15) };
    }
    const base = src[k % n], cycle = Math.floor(k / n);
    if (base.type === 'formula') return { type: 'formula', f: base.f, shift: k - (k % n) + 0 };
    if (base.type === 'num' && base.date && n === 1) return { type: 'num', v: base.v + k };
    if (base.type === 'str') {
      const t = base.v, lo = t.toLowerCase();
      let i = DAYS.indexOf(lo); if (i >= 0) return { type: 'str', v: caseLike(t, DAYS[(i + k) % 7]) };
      i = MONTHS.indexOf(lo); if (i >= 0) return { type: 'str', v: caseLike(t, MONTHS[(i + k) % 12]) };
      i = MONTHS3.indexOf(lo); if (i >= 0) return { type: 'str', v: caseLike(t, MONTHS3[(i + k) % 12]) };
      const m = /^(.*?)(\d+)(\D*)$/.exec(t);
      if (m) { const step = n === 1 ? k : cycle * n; return { type: 'str', v: m[1] + String(+m[2] + step).padStart(m[2].length, '0') + m[3] }; }
    }
    return { ...base };
  };
  XD.readSrc = function (sh, r, c) {
    const x = this.raw(sh, r, c);
    if (x.f != null) return { type: 'formula', f: x.f, r, c, s: x.s };
    if (x.v == null || x.v === '') return { type: 'empty', s: x.s };
    if (typeof x.v === 'number') return { type: 'num', v: x.v, date: SSF.is_date(this.fmtCode(x.s)), s: x.s };
    if (typeof x.v === 'boolean') return { type: 'bool', v: x.v, s: x.s };
    if (typeof x.v === 'object') return { type: 'empty', s: x.s };
    return { type: 'str', v: x.v, s: x.s };
  };
  XD.writeCell = function (sh, r, c, p, style) {
    const el = this.cellEl(sh, r, c, p.type !== 'empty' || style != null);
    if (!el) return;
    const oldF = OF.kid(el, 'f', S);
    if (oldF && oldF.getAttribute('t') === 'shared' && oldF.getAttribute('ref')) this.expandShared(sh, oldF.getAttribute('si'));
    for (const ch of Array.from(el.children)) if (['f', 'v', 'is'].includes(ch.localName)) el.removeChild(ch);
    el.removeAttribute('t');
    const add = (name, text) => { const e = sh.doc.createElementNS(S, name); e.textContent = text; el.appendChild(e); return e; };
    if (p.type === 'formula') add('f', p.f);
    else if (p.type === 'num') add('v', String(p.v));
    else if (p.type === 'bool') { el.setAttribute('t', 'b'); add('v', p.v ? '1' : '0'); }
    else if (p.type === 'str') { const idx = this.addString(p.v); if (idx != null) { el.setAttribute('t', 's'); add('v', String(idx)); } else { el.setAttribute('t', 'inlineStr'); const is = add('is', ''); const t = sh.doc.createElementNS(S, 't'); OF.setText(t, p.v); is.appendChild(t); } }
    if (style != null) { if (+style) el.setAttribute('s', style); else el.removeAttribute('s'); }
    if (r > sh.maxR) sh.maxR = r; if (c > sh.maxC) sh.maxC = c;
  };
  XD.fill = async function (src, dst) {
    // src y dst: rectángulos; dst incluye src y se extiende hacia abajo, arriba, derecha o izquierda
    const sh = this.sheets[this.cur];
    OF.checkpoint && OF.checkpoint();
    for (const si of Object.keys(sh.shared)) this.expandShared(sh, si);
    const vertical = dst.r1 !== src.r1 || dst.r2 !== src.r2;
    const lines = vertical ? [src.c1, src.c2] : [src.r1, src.r2];
    for (let L = lines[0]; L <= lines[1]; L++) {
      const srcCells = [];
      const len = vertical ? src.r2 - src.r1 + 1 : src.c2 - src.c1 + 1;
      for (let i = 0; i < len; i++) srcCells.push(vertical ? this.readSrc(sh, src.r1 + i, L) : this.readSrc(sh, L, src.c1 + i));
      const forward = vertical ? dst.r2 > src.r2 : dst.c2 > src.c2;
      const count = vertical ? (forward ? dst.r2 - src.r2 : src.r1 - dst.r1) : (forward ? dst.c2 - src.c2 : src.c1 - dst.c1);
      const seq = forward ? srcCells : srcCells.slice().reverse();
      const nonEmpty = seq.filter((x) => x.type !== 'empty');
      for (let k = 1; k <= count; k++) {
        const r = vertical ? (forward ? src.r2 + k : src.r1 - k) : L;
        const c = vertical ? L : (forward ? src.c2 + k : src.c1 - k);
        const base = seq[(k - 1 + len) % len];
        let p;
        if (base.type === 'formula') {
          const dr = r - base.r, dc = c - base.c;
          p = { type: 'formula', f: shiftFormula(base.f, dr, dc) };
        } else if (base.type === 'empty') p = { type: 'empty' };
        else {
          const v = this.seriesValue(nonEmpty.length === seq.length ? seq : [base], nonEmpty.length === seq.length ? len - 1 + k : k);
          p = forward ? v : (() => { const t = this.seriesValue(seq, -k); return t; })();
          if (!forward && seq.every((x) => x.type === 'num') && len >= 2) { const step = (seq[len - 1].v - seq[0].v) / (len - 1); p = { type: 'num', v: +(seq[0].v - step * k).toPrecision(15) }; }
          else if (!forward) p = { ...base };
        }
        this.writeCell(sh, r, c, p, base.s);
      }
    }
    await this.reloadSheet();
    this.anchor = { r: dst.r1, c: dst.c1 };
    this.select(dst.r2, dst.c2, true);
  };
  XD.fillDir = function (dir) {
    const g = this.workRect();
    if (dir === 'down') { if (g.r2 === g.r1) { if (g.r1 === 0) return; this.fill({ ...g, r1: g.r1 - 1, r2: g.r1 - 1 }, { ...g, r1: g.r1 - 1 }); } else this.fill({ ...g, r2: g.r1 }, g); }
    else { if (g.c2 === g.c1) { if (g.c1 === 0) return; this.fill({ ...g, c1: g.c1 - 1, c2: g.c1 - 1 }, { ...g, c1: g.c1 - 1 }); } else this.fill({ ...g, c2: g.c1 }, g); }
  };
  // cuadrito de relleno en la esquina de la selección
  XD.placeHandle = function () {
    if (!this.handle) {
      this.handle = OF.el('div', { class: 'xl-handle', title: 'Arrastra para rellenar', 'aria-hidden': 'true' });
      this.handle.addEventListener('pointerdown', (e) => this.dragFill(e));
    }
    if (this.handle.parentNode !== this.gridBox) this.gridBox.append(this.handle);
    const g = this.rangeRect();
    const sh = this.sheets[this.cur];
    const m = sh.covered.get(g.r2 + ',' + g.c2);
    const td = this.tdAt(g.r2, g.c2) || (m === 'x' ? null : null);
    if (!td || this.whole) { this.handle.style.display = 'none'; return; }
    const box = this.gridBox.getBoundingClientRect(), r = td.getBoundingClientRect();
    this.handle.style.display = '';
    this.handle.style.left = r.right - box.left + this.gridBox.scrollLeft - 5 + 'px';
    this.handle.style.top = r.bottom - box.top + this.gridBox.scrollTop - 5 + 'px';
  };
  XD.dragFill = function (e) {
    e.preventDefault(); e.stopPropagation();
    const src = this.rangeRect();
    let dst = { ...src };
    const ghost = OF.el('div', { class: 'xl-fillghost' });
    this.gridBox.append(ghost);
    const cellAt = (x, y) => { const el = document.elementFromPoint(x, y); const td = el && el.closest && el.closest('td[data-c]'); return td && this.table.contains(td) ? { r: +td.parentNode.dataset.r, c: +td.dataset.c } : null; };
    const paint = () => {
      const a = this.tdAt(dst.r1, dst.c1), b = this.tdAt(dst.r2, dst.c2);
      if (!a || !b) return;
      const box = this.gridBox.getBoundingClientRect(), ra = a.getBoundingClientRect(), rb = b.getBoundingClientRect();
      Object.assign(ghost.style, { left: ra.left - box.left + this.gridBox.scrollLeft + 'px', top: ra.top - box.top + this.gridBox.scrollTop + 'px', width: rb.right - ra.left + 'px', height: rb.bottom - ra.top + 'px' });
    };
    const mv = (ev) => {
      const p = cellAt(ev.clientX, ev.clientY);
      if (!p) return;
      const down = p.r - src.r2, up = src.r1 - p.r, right = p.c - src.c2, left = src.c1 - p.c;
      const best = Math.max(down, up, right, left, 0);
      dst = { ...src };
      if (best > 0) { if (best === down) dst.r2 = p.r; else if (best === up) dst.r1 = p.r; else if (best === right) dst.c2 = p.c; else dst.c1 = p.c; }
      paint();
      const box = this.gridBox.getBoundingClientRect();
      if (ev.clientY > box.bottom - 30) this.gridBox.scrollTop += 20;
      if (ev.clientX > box.right - 30) this.gridBox.scrollLeft += 20;
    };
    const up = () => {
      window.removeEventListener('pointermove', mv); window.removeEventListener('pointerup', up);
      ghost.remove();
      if (dst.r1 !== src.r1 || dst.r2 !== src.r2 || dst.c1 !== src.c1 || dst.c2 !== src.c2) this.fill(src, dst);
    };
    window.addEventListener('pointermove', mv); window.addEventListener('pointerup', up);
    paint();
  };

  /* --------------------------- autosuma --------------------------- */
  XD.autoSum = async function (fn) {
    this.stopEdit(true);
    const sh = this.sheets[this.cur];
    const isNum = (r, c) => { const x = this.raw(sh, r, c); return typeof x.v === 'number' || x.f != null; };
    const name = this.locale === 'cl' ? OF.formulaToES(fn + '(').slice(0, -1) : fn;
    if (this.isRange()) {
      const g = this.workRect();
      OF.checkpoint && OF.checkpoint();
      const vertical = g.r2 > g.r1 || g.c1 === g.c2;
      if (vertical) for (let c = g.c1; c <= g.c2; c++) this.writeCell(sh, g.r2 + 1, c, { type: 'formula', f: `${fn}(${addr(g.r1, c)}:${addr(g.r2, c)})` }, null);
      else for (let r = g.r1; r <= g.r2; r++) this.writeCell(sh, r, g.c2 + 1, { type: 'formula', f: `${fn}(${addr(r, g.c1)}:${addr(r, g.c2)})` }, null);
      await this.reloadSheet();
      if (vertical) { this.anchor = { r: g.r2 + 1, c: g.c1 }; this.select(g.r2 + 1, g.c2, true); } else { this.anchor = { r: g.r1, c: g.c2 + 1 }; this.select(g.r2, g.c2 + 1, true); }
      return;
    }
    const { r, c } = this.sel;
    let r1 = r - 1; while (r1 >= 0 && isNum(r1, c)) r1--;
    if (r1 < r - 1) { this.commit(`=${name}(${addr(r1 + 1, c)}:${addr(r - 1, c)})`); return; }
    let c1 = c - 1; while (c1 >= 0 && isNum(r, c1)) c1--;
    if (c1 < c - 1) { this.commit(`=${name}(${addr(r, c1 + 1)}:${addr(r, c - 1)})`); return; }
    this.startEdit(`=${name}(`);
  };

  /* --------------------------- hojas --------------------------- */
  XD.flush = function () {
    for (const sh of this.sheets) this.zip.file(sh.path, OF.serializeXml(sh.doc));
    if (this.ssDoc && this.ssRel) this.zip.file(this.ssRel.target, OF.serializeXml(this.ssDoc));
    if (this.stylesDoc) this.zip.file(this.stylesPath, OF.serializeXml(this.stylesDoc));
    this.zip.file(this.wbPath, OF.serializeXml(this.wb));
  };
  XD.reopen = async function (sheetIdx, sel) {
    const bytes = await this.zip.generateAsync({ type: 'arraybuffer', compression: 'STORE' });
    const keep = { scroll: [this.gridBox.scrollTop, this.gridBox.scrollLeft], zoom: this.userZoom };
    const wasDirty = this.dirty;
    await this.open(bytes, this.host);
    this.dirty = wasDirty;
    if (sheetIdx != null && this.sheets[sheetIdx]) this.showSheet(sheetIdx);
    if (sel) this.select(sel.r, sel.c);
    if (keep.zoom) { this.userZoom = keep.zoom; this.zoom(1); }
  };
  XD.sheetMenu = function (tab, i) {
    OF.popup(tab, (p, close) => {
      const it = (label, fn, danger) => p.append(OF.el('button', { class: 'pop-item' + (danger ? ' danger' : ''), onclick: () => { close(); fn(); } }, label));
      it('Cambiar nombre', () => this.sheetOp('rename', i));
      it('Duplicar hoja', () => this.sheetOp('dup', i));
      it('Agregar hoja', () => this.sheetOp('add'));
      it('Eliminar hoja', () => this.sheetOp('delete', i), true);
    }, true);
  };
  XD.sheetOp = async function (op, i) {
    this.stopEdit(true);
    const wbRoot = this.wb.documentElement;
    const sheetEls = OF.all(this.wb, 'sheet');
    const names = sheetEls.map((e) => e.getAttribute('name'));
    const valid = (n) => n && n.length <= 31 && !/[\[\]:*?/\\]/.test(n) && !/^'|'$/.test(n) && !names.some((x) => x.toLowerCase() === n.toLowerCase());
    const relPath = OF.relsPathFor(this.wbPath);
    if (op === 'rename') {
      const old = this.sheets[i].name;
      const n = (prompt('Nombre de la hoja:', old) || '').trim();
      if (!n || n === old) return;
      if (!valid(n)) { OF.toast('Nombre no válido o repetido (máx. 31 caracteres, sin [ ] : * ? / \\).', 4000); return; }
      OF.checkpoint && OF.checkpoint();
      sheetEls.find((e) => e.getAttribute('name') === old).setAttribute('name', n);
      const re = new RegExp(`(^|[^A-Za-z0-9_.'])(${old.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}|'${old.replace(/'/g, "''").replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}')!`, 'g');
      const fix = (t) => t.replace(re, (m, pre) => pre + quoteSheet(n) + '!');
      for (const sh of this.sheets) for (const f of OF.all(sh.doc, 'f', S)) f.textContent = fix(f.textContent);
      for (const dn of OF.all(this.wb, 'definedName')) dn.textContent = fix(dn.textContent);
      for (const path of Object.keys(this.zip.files).filter((x) => /^xl\/charts\/chart\d*\.xml$/.test(x))) {
        const doc = OF.parseXml(await this.zip.file(path).async('string'));
        for (const f of OF.all(doc, 'f', C)) f.textContent = fix(f.textContent);
        this.zip.file(path, OF.serializeXml(doc));
      }
      this.flush();
      this.markDirty();
      await this.reopen(i);
      return;
    }
    if (op === 'add' || op === 'dup') {
      let n = op === 'dup' ? this.sheets[i].name + ' (2)' : 'Hoja' + (sheetEls.length + 1);
      for (let k = 2; !valid(n); k++) n = op === 'dup' ? `${this.sheets[i].name} (${k})` : 'Hoja' + (sheetEls.length + k);
      OF.checkpoint && OF.checkpoint();
      this.flush();
      const path = OF.uniquePart(this.zip, 'xl/worksheets', 'sheet', 'xml');
      if (op === 'dup') {
        const src = this.sheets[i];
        const doc = OF.parseXml(OF.serializeXml(src.doc));
        // dibujos, tablas y comentarios no se duplican (quedarían compartidos)
        for (const name of ['drawing', 'legacyDrawing', 'tableParts', 'picture', 'oleObjects', 'controls']) OF.removeKids(doc.documentElement, [name], S);
        for (const sv of OF.all(doc, 'sheetView', S)) sv.removeAttribute('tabSelected');
        this.zip.file(path, OF.serializeXml(doc));
      } else {
        this.zip.file(path, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="${S}" xmlns:r="${OF.NS.r}"><sheetViews><sheetView workbookViewId="0"/></sheetViews><sheetFormatPr defaultRowHeight="15"/><sheetData/><pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>`);
      }
      const rid = await OF.addRel(this.zip, this.wbPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet', path);
      await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, path, 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'));
      const sheetsEl = OF.kid(wbRoot, 'sheets', S);
      const e = this.wb.createElementNS(S, 'sheet');
      e.setAttribute('name', n);
      e.setAttribute('sheetId', Math.max(0, ...sheetEls.map((x) => +x.getAttribute('sheetId'))) + 1);
      e.setAttributeNS(OF.NS.r, 'r:id', rid);
      const after = op === 'dup' ? sheetEls[OF.all(this.wb, 'sheet').findIndex((x) => x.getAttribute('name') === this.sheets[i].name)] : null;
      sheetsEl.insertBefore(e, after ? after.nextSibling : null);
      // las hojas definidas por posición (localSheetId) se corren si se inserta en medio
      const pos = OF.all(this.wb, 'sheet').indexOf(e);
      for (const dn of OF.all(this.wb, 'definedName')) { const l = dn.getAttribute('localSheetId'); if (l != null && +l >= pos) dn.setAttribute('localSheetId', +l + 1); }
      this.zip.file(this.wbPath, OF.serializeXml(this.wb));
      this.markDirty();
      await this.reopen(null);
      const idx = this.sheets.findIndex((s) => s.name === n);
      this.showSheet(idx);
      OF.toast(op === 'dup' ? `Hoja duplicada: «${n}»` : `Hoja agregada: «${n}»`);
      return;
    }
    if (op === 'delete') {
      if (this.sheets.filter((s) => !s.hidden).length <= 1) { OF.toast('El libro debe tener al menos una hoja visible.'); return; }
      const name = this.sheets[i].name;
      if (!confirm(`¿Eliminar la hoja «${name}»? Las fórmulas que la usan mostrarán #REF!.`)) return;
      OF.checkpoint && OF.checkpoint();
      this.flush();
      const pos = sheetEls.findIndex((e) => e.getAttribute('name') === name);
      const el = sheetEls[pos];
      const rid = el.getAttributeNS(OF.NS.r, 'id');
      el.parentNode.removeChild(el);
      const target = this.sheets[i].path;
      const relDoc = OF.parseXml(await this.zip.file(relPath).async('string'));
      for (const r of OF.all(relDoc, 'Relationship')) if (r.getAttribute('Id') === rid) r.parentNode.removeChild(r);
      this.zip.file(relPath, OF.serializeXml(relDoc));
      this.zip.remove(target); this.zip.remove(OF.relsPathFor(target));
      await OF.contentTypes(this.zip, (ct) => { for (const o of OF.all(ct, 'Override')) if (o.getAttribute('PartName') === '/' + target) o.parentNode.removeChild(o); });
      for (const dn of Array.from(OF.all(this.wb, 'definedName'))) { const l = dn.getAttribute('localSheetId'); if (l == null) continue; if (+l === pos) dn.parentNode.removeChild(dn); else if (+l > pos) dn.setAttribute('localSheetId', +l - 1); }
      const re = new RegExp(`(^|[^A-Za-z0-9_.'])(${name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}|'${name.replace(/'/g, "''").replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}')!\\$?[A-Z]{1,3}\\$?\\d+(:\\$?[A-Z]{1,3}\\$?\\d+)?`, 'g');
      for (const sh of this.sheets) if (sh.path !== target) { for (const f of OF.all(sh.doc, 'f', S)) f.textContent = f.textContent.replace(re, (m, pre) => pre + '#REF!'); this.zip.file(sh.path, OF.serializeXml(sh.doc)); }
      for (const dn of OF.all(this.wb, 'definedName')) dn.textContent = dn.textContent.replace(re, (m, pre) => pre + '#REF!');
      const bv = OF.all(this.wb, 'workbookView')[0];
      if (bv) { bv.setAttribute('activeTab', '0'); bv.removeAttribute('firstSheet'); }
      this.zip.file(this.wbPath, OF.serializeXml(this.wb));
      this.markDirty();
      await this.reopen(0);
      OF.toast(`Hoja «${name}» eliminada`);
    }
  };

  /* ----------------------------- gráficos ----------------------------- */
  XD.resolveRef = function (ref) {
    const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(ref.trim());
    if (!m) return null;
    const name = (m[1] || m[2]).replace(/''/g, "'");
    const si = this.sheets.findIndex((s) => s.name === name);
    if (si < 0) return null;
    const g = parseRange(m[3].replace(/\$/g, ''));
    if (!g) return null;
    const out = [];
    for (let r = g.r1; r <= g.r2; r++) for (let c = g.c1; c <= g.c2; c++) {
      const x = this.value(si, r, c);
      out.push(typeof x.v === 'number' ? x.v : x.v == null ? null : typeof x.v === 'object' ? null : this.display(x));
    }
    return out;
  };
  XD.loadPictures = async function (sh) {
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
        const gf = OF.kid(anc, 'graphicFrame', XDR);
        if (pic) {
          const blip = OF.all(pic, 'blip')[0];
          const rel = blip && rels[blip.getAttributeNS(OF.NS.r, 'embed')];
          if (rel && this.zip.file(rel.target) && !/\.(emf|wmf)$/i.test(rel.target)) {
            const b = await this.zip.file(rel.target).async('blob');
            item.url = URL.createObjectURL(new Blob([b], { type: OF.mimeFor(rel.target) }));
          }
        } else if (gf) {
          item.chart = true;
          const ch = OF.all(gf, 'chart', C)[0];
          const rel = ch && rels[ch.getAttributeNS(OF.NS.r, 'id')];
          if (rel && this.zip.file(rel.target)) { try { item.spec = OF.parseChart(OF.parseXml(await this.zip.file(rel.target).async('string'))); } catch (e) {} }
        } else continue;
        sh.pics.push(item);
      }
    }
  };
  XD.drawPictures = async function () {
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
      let el;
      if (p.url) el = OF.el('img', { src: p.url, alt: '', draggable: 'false' });
      else if (p.spec) { el = OF.el('div', { class: 'xl-chartbox' }); el.append(OF.drawChart(p.spec, Math.max(40, w), Math.max(40, h), (ref) => this.resolveRef(ref))); }
      else el = OF.el('div', { class: 'xl-chart' }, p.chart ? 'Gráfico (se conserva)' : 'Objeto');
      Object.assign(el.style, { left: x + 'px', top: y + 'px', width: w + 'px', height: h + 'px' });
      p.shown = { w, h };
      p.el = el;
      el.addEventListener('click', (e) => {
        e.stopPropagation();
        layer.querySelectorAll('.sel').forEach((q) => q.classList.remove('sel'));
        el.classList.add('sel');
        OF.imageBar(el, [['Más pequeño', () => this.picAction(p, 0.8)], ['Más grande', () => this.picAction(p, 1.25)], ['Eliminar', () => this.picAction(p, 'del'), true]]);
      });
      layer.append(el);
    }
    this.table.style.position = 'relative';
    this.gridBox.append(layer);
    layer.style.zoom = this.userZoom || 1;
  };
  XD.redrawCharts = function () {
    const sh = this.sheets[this.cur];
    if (!sh || !sh.pics) return;
    for (const p of sh.pics) if (p.spec && p.el && p.el.isConnected) { p.el.innerHTML = ''; p.el.append(OF.drawChart(p.spec, Math.max(40, p.shown.w), Math.max(40, p.shown.h), (ref) => this.resolveRef(ref))); }
  };
  const baseCommit = XD.commit;
  XD.commit = function (input, noHist) { baseCommit.call(this, input, noHist); this.redrawCharts(); };
  XD.ensureDrawing = async function (sh) {
    const root = sh.doc.documentElement;
    let drawEl = OF.kid(root, 'drawing', S);
    if (drawEl) { const dPath = sh.rels[drawEl.getAttributeNS(OF.NS.r, 'id')].target; return { dPath, dDoc: OF.parseXml(await this.zip.file(dPath).async('string')) }; }
    const dPath = OF.uniquePart(this.zip, 'xl/drawings', 'drawing', 'xml');
    const dDoc = OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><xdr:wsDr xmlns:xdr="${XDR}" xmlns:a="${OF.NS.a}"/>`);
    this.zip.file(dPath, OF.serializeXml(dDoc));
    const rid = await OF.addRel(this.zip, sh.path, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing', dPath);
    drawEl = OF.ordered(root, S, root.prefix || '', 'drawing', WS_ORDER);
    if (!root.lookupPrefix(OF.NS.r)) root.setAttributeNS('http://www.w3.org/2000/xmlns/', 'xmlns:r', OF.NS.r);
    drawEl.setAttributeNS(OF.NS.r, 'r:id', rid);
    await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, dPath, 'application/vnd.openxmlformats-officedocument.drawing+xml'));
    sh.rels = await OF.readRels(this.zip, sh.path);
    return { dPath, dDoc };
  };
  XD.insertChart = async function (type) {
    this.stopEdit(true);
    const sh = this.sheets[this.cur];
    const g = this.dataRegion();
    if (g.r2 - g.r1 < 1 && g.c2 - g.c1 < 1) { OF.toast('Selecciona los datos del gráfico (con títulos en la primera fila).', 4000); return; }
    const isText = (r, c) => { const x = this.value(this.cur, r, c); return x.v != null && typeof x.v === 'string'; };
    let headerRow = false; for (let c = g.c1; c <= g.c2; c++) if (isText(g.r1, c)) headerRow = true;
    let labelCol = true; for (let r = g.r1 + (headerRow ? 1 : 0); r <= g.r2; r++) { const x = this.value(this.cur, r, g.c1); if (typeof x.v === 'number' && !SSF.is_date(this.fmtCode(x.s))) labelCol = false; }
    if (g.c1 === g.c2) labelCol = false;
    const r0 = g.r1 + (headerRow ? 1 : 0), c0 = g.c1 + (labelCol ? 1 : 0);
    if (r0 > g.r2 || c0 > g.c2) { OF.toast('No encontré números para graficar.'); return; }
    OF.checkpoint && OF.checkpoint();
    const sn = quoteSheet(sh.name);
    const abs = (r, c) => `$${colName(c)}$${r + 1}`;
    const cats = [];
    for (let r = r0; r <= g.r2; r++) cats.push(labelCol ? this.display(this.value(this.cur, r, g.c1)) : String(r - r0 + 1));
    const catRef = labelCol ? `${sn}!${abs(r0, g.c1)}:${abs(g.r2, g.c1)}` : `${sn}!${abs(r0, c0)}:${abs(g.r2, c0)}`;
    const series = [];
    for (let c = c0; c <= g.c2; c++) {
      const vals = []; for (let r = r0; r <= g.r2; r++) { const x = this.value(this.cur, r, c); vals.push(typeof x.v === 'number' ? x.v : null); }
      series.push({ name: headerRow ? this.display(this.value(this.cur, g.r1, c)) : 'Serie ' + (c - c0 + 1), nameRef: headerRow ? `${sn}!${abs(g.r1, c)}` : null, cats, catRef, vals, valRef: `${sn}!${abs(r0, c)}:${abs(g.r2, c)}` });
    }
    if (type === 'pie') series.length = 1;
    const title = type === 'pie' || series.length === 1 ? series[0].name : '';
    const chartPath = OF.uniquePart(this.zip, 'xl/charts', 'chart', 'xml');
    this.zip.file(chartPath, OF.chartXml(type, title, series));
    await OF.contentTypes(this.zip, (ct) => OF.addOverrideCT(ct, chartPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'));
    const { dPath, dDoc } = await this.ensureDrawing(sh);
    const rId = await OF.addRel(this.zip, dPath, 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart', chartPath);
    const ids = OF.all(dDoc, 'cNvPr').map((e) => +e.getAttribute('id') || 0);
    const id = Math.max(1, ...ids) + 1;
    const fc = g.c2 + 2, fr = g.r1;
    const anc = OF.fragment(dDoc, `<xdr:twoCellAnchor xmlns:xdr="${XDR}" xmlns:a="${OF.NS.a}" xmlns:r="${OF.NS.r}" xmlns:c="${C}" editAs="oneCell"><xdr:from><xdr:col>${fc}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fr}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>${fc + 7}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${fr + 15}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to><xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${id}" name="Gráfico ${id}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr><xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="${rId}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`);
    dDoc.documentElement.appendChild(anc);
    this.zip.file(dPath, OF.serializeXml(dDoc));
    this.zip.file(sh.path, OF.serializeXml(sh.doc));
    sh.pics = null;
    this.markDirty();
    await this.drawPictures();
    const p = sh.pics[sh.pics.length - 1];
    if (p && p.el) p.el.scrollIntoView({ block: 'nearest', inline: 'nearest' });
    OF.toast('Gráfico insertado');
  };

  /* ------------------------ estado, búsqueda, deshacer ------------------------ */
  XD.formatState = function () {
    if (!this.sel) return null;
    const x = this.raw(this.sheets[this.cur], this.sel.r, this.sel.c);
    const xf = this.xfs[x.s] || this.xfs[0], f = this.fonts[xf.fontId] || {};
    return { b: !!f.b, i: !!f.i, u: !!f.u, font: f.name || 'Calibri', size: f.sz || 11, align: xf.h === 'center' || xf.h === 'right' || xf.h === 'left' ? xf.h : null, wrap: !!xf.wrap };
  };
  XD.statusText = function () {
    if (!this.sel) return '';
    const sh = this.sheets[this.cur];
    if (!this.isRange()) return sh.freeze && (sh.freeze.rows || sh.freeze.cols) ? 'Paneles inmovilizados' : '';
    const g = this.workRect();
    let n = 0, cnt = 0, sum = 0;
    for (let r = g.r1; r <= Math.min(g.r2, sh.maxR); r++) {
      if ((sh.rowInfo[r] || {}).hidden) continue;
      for (let c = g.c1; c <= Math.min(g.c2, sh.maxC); c++) {
        const x = this.value(this.cur, r, c);
        if (x.v == null || x.v === '') continue;
        cnt++;
        if (typeof x.v === 'number') { n++; sum += x.v; }
      }
    }
    const f = (v) => (this.locale === 'cl' ? v.toLocaleString('es-CL', { maximumFractionDigits: 2 }) : v.toLocaleString('en-US', { maximumFractionDigits: 2 }));
    return n ? `Promedio: ${f(sum / n)} · Recuento: ${cnt} · Suma: ${f(sum)}` : cnt ? `Recuento: ${cnt}` : '';
  };
  XD.findMatches = function (q) {
    const sh = this.sheets[this.cur], out = [], lq = q.toLowerCase();
    for (let r = 0; r <= sh.maxR; r++) for (let c = 0; c <= sh.maxC; c++) {
      const t = this.display(this.value(this.cur, r, c)), raw = this.inputText(r, c);
      if (t.toLowerCase().includes(lq) || (raw.startsWith('=') && raw.toLowerCase().includes(lq))) out.push({ r, c });
    }
    return out;
  };
  XD.find = function (q, dir) {
    if (!q) return { i: -1, n: 0 };
    const m = this.findMatches(q);
    if (!m.length) return { i: -1, n: 0 };
    const cur = this.sel || { r: 0, c: 0 };
    const pos = (x) => x.r * 100000 + x.c;
    let i;
    if (dir < 0) { i = m.length - 1; for (let k = m.length - 1; k >= 0; k--) if (pos(m[k]) < pos(cur)) { i = k; break; } }
    else { i = 0; for (let k = 0; k < m.length; k++) if (pos(m[k]) > pos(cur) || (this._fresh && pos(m[k]) === pos(cur))) { i = k; break; } }
    this._fresh = false;
    this.select(m[i].r, m[i].c);
    return { i, n: m.length };
  };
  XD.replaceText = function (text, q, rep) { return text.replace(new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'gi'), () => rep); };
  XD.replaceOne = function (q, rep) {
    if (!q || !this.sel) return this.find(q, 1);
    const t = this.inputText(this.sel.r, this.sel.c);
    if (t.toLowerCase().includes(q.toLowerCase())) this.commit(this.replaceText(t, q, rep));
    return this.find(q, 1);
  };
  XD.replaceAll = function (q, rep) {
    if (!q) return 0;
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], keep = this.sel;
    let n = 0;
    for (const m of this.findMatches(q)) {
      const t = this.inputText(m.r, m.c);
      if (!t.toLowerCase().includes(q.toLowerCase())) continue;
      this.sel = { r: m.r, c: m.c };
      this.commit(this.replaceText(t, q, rep), true);
      n++;
    }
    void sh;
    if (keep) this.select(keep.r, keep.c);
    return n;
  };
  XD.snapshot = function () {
    this.flush();
    return { files: Object.assign({}, this.zip.files), cur: this.cur, sel: this.sel && { ...this.sel }, scroll: [this.gridBox.scrollTop, this.gridBox.scrollLeft] };
  };
  XD.restore = async function (s) {
    const z = new JSZip();
    for (const k in s.files) z.files[k] = s.files[k];
    this.zip = z;
    await this.reopen(s.cur, s.sel);
    this.gridBox.scrollTop = s.scroll[0]; this.gridBox.scrollLeft = s.scroll[1];
  };
  XD.printSetup = function () {
    this.stopEdit(true);
    const sh = this.sheets[this.cur];
    // solo la zona con datos (y lo que cubren imágenes y gráficos)
    let mr = sh.maxR, mc = sh.maxC;
    for (const p of sh.pics || []) { if (p.to) { mr = Math.max(mr, p.to.r); mc = Math.max(mc, p.to.c); } else if (p.from) { mr = Math.max(mr, p.from.r + Math.ceil((p.h || 0) / 20)); mc = Math.max(mc, p.from.c + Math.ceil((p.w || 0) / 64)); } }
    while (this.rendered <= mr && this.rendered < sh.nR) this.renderRows(Math.min(sh.nR, this.rendered + 200));
    this.tbody.querySelectorAll('tr').forEach((tr) => tr.classList.toggle('xl-noprint', +tr.dataset.r > mr));
    this.table.querySelectorAll('td[data-c]').forEach((td) => td.classList.toggle('xl-noprint', +td.dataset.c > mc));
    this.table.querySelectorAll('colgroup col').forEach((col, i) => col.classList.toggle('xl-noprint', i - 1 > mc));
    return '@page { size: landscape; margin: 12mm }';
  };
  XD.fit = function () { if (this.table) this.table.querySelectorAll('.xl-noprint').forEach((e) => e.classList.remove('xl-noprint')); };
})();
