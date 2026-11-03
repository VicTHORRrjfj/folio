/* Excel: formato condicional (dibujar y crear) y filtros por condición. */
(function () {
  const OF = window.OF;
  const S = OF.NS.s;
  const XD = OF.ExcelDoc.prototype;
  const WS_ORDER = ['sheetPr', 'dimension', 'sheetViews', 'sheetFormatPr', 'cols', 'sheetData', 'sheetCalcPr', 'sheetProtection', 'protectedRanges', 'scenarios', 'autoFilter', 'sortState', 'dataConsolidate', 'customSheetViews', 'mergeCells', 'phoneticPr', 'conditionalFormatting', 'dataValidations', 'hyperlinks', 'printOptions', 'pageMargins', 'pageSetup', 'headerFooter', 'rowBreaks', 'colBreaks', 'customProperties', 'cellWatches', 'ignoredErrors', 'smartTags', 'drawing', 'legacyDrawing', 'legacyDrawingHF', 'drawingHF', 'picture', 'oleObjects', 'controls', 'webPublishItems', 'tableParts', 'extLst'];
  const STYLE_ORDER = ['numFmts', 'fonts', 'fills', 'borders', 'cellStyleXfs', 'cellXfs', 'cellStyles', 'dxfs', 'tableStyles', 'colors', 'extLst'];
  const colName = OF.colName;
  const colIndex = (s) => { let n = 0; for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64); return n - 1; };
  const parseRef = (ref) => { const m = /^\$?([A-Z]+)\$?(\d+)$/i.exec(ref); return m ? { r: +m[2] - 1, c: colIndex(m[1]) } : null; };
  const addr = (r, c) => colName(c) + (r + 1);
  const parseRange = (ref) => { const [a, b] = ref.split(':'); const A = parseRef(a), B = parseRef(b || a); return A && B ? { r1: Math.min(A.r, B.r), c1: Math.min(A.c, B.c), r2: Math.max(A.r, B.r), c2: Math.max(A.c, B.c) } : null; };
  const rangeStr = (g) => (g.r1 === g.r2 && g.c1 === g.c2 ? addr(g.r1, g.c1) : addr(g.r1, g.c1) + ':' + addr(g.r2, g.c2));
  function shiftFormula(f, dr, dc) {
    return f.split(/("(?:[^"]|"")*")/).map((seg, i) => {
      if (i % 2) return seg;
      return seg.replace(/(?<![A-Za-z0-9_.!])(\$?)([A-Z]{1,3})(\$?)(\d+)(?![A-Za-z0-9_(])/g, (m, d1, col, d2, row) => {
        const c = d1 ? colIndex(col) : colIndex(col) + dc;
        const r = d2 ? +row : +row + dr;
        if (c < 0 || r < 1) return '#REF!';
        return d1 + colName(c) + d2 + r;
      });
    }).join('');
  }
  const hex = (c) => { c = c.replace('#', ''); return [0, 2, 4].map((i) => parseInt(c.slice(i, i + 2), 16)); };
  const mix = (a, b, t) => { const A = hex(a), B = hex(b); return '#' + A.map((v, i) => Math.round(v + (B[i] - v) * t).toString(16).padStart(2, '0')).join(''); };

  /* ------------------------------- lectura ------------------------------- */
  XD.loadDxfs = function () {
    this.dxfs = [];
    if (!this.stylesDoc) return;
    const list = OF.kid(this.stylesDoc.documentElement, 'dxfs', S);
    for (const d of OF.kids(list, 'dxf', S)) {
      const out = {};
      const f = OF.kid(d, 'font', S);
      if (f) {
        const on = (n) => { const e = OF.kid(f, n, S); return e ? e.getAttribute('val') !== '0' && e.getAttribute('val') !== 'false' : null; };
        if (on('b') != null) out.b = on('b');
        if (on('i') != null) out.i = on('i');
        if (on('strike') != null) out.strike = on('strike');
        const u = OF.kid(f, 'u', S); if (u) out.u = u.getAttribute('val') !== 'none';
        const c = this.color(OF.kid(f, 'color', S)); if (c) out.color = c;
      }
      const pf = OF.kid(OF.kid(d, 'fill', S), 'patternFill', S);
      if (pf && pf.getAttribute('patternType') !== 'none') { const c = this.color(OF.kid(pf, 'bgColor', S)) || this.color(OF.kid(pf, 'fgColor', S)); if (c) out.bg = c; }
      const bd = OF.kid(d, 'border', S);
      if (bd) for (const side of ['left', 'right', 'top', 'bottom']) { const e = OF.kid(bd, side, S); if (e && e.getAttribute('style')) out['b' + side] = this.color(OF.kid(e, 'color', S)) || '#000'; }
      this.dxfs.push(out);
    }
  };
  XD.parseCF = function (sh) {
    const root = sh.doc.documentElement;
    sh.cf = [];
    for (const cf of OF.kids(root, 'conditionalFormatting', S)) {
      const ranges = (cf.getAttribute('sqref') || '').split(/\s+/).map(parseRange).filter(Boolean);
      if (!ranges.length) continue;
      for (const rule of OF.kids(cf, 'cfRule', S)) {
        const R = { el: rule, ranges, type: rule.getAttribute('type'), priority: +(rule.getAttribute('priority') || 999), dxf: rule.hasAttribute('dxfId') ? +rule.getAttribute('dxfId') : null, op: rule.getAttribute('operator'), text: rule.getAttribute('text'), stop: rule.getAttribute('stopIfTrue') === '1', formulas: OF.kids(rule, 'formula', S).map((f) => f.textContent) };
        const cs = OF.kid(rule, 'colorScale', S), db = OF.kid(rule, 'dataBar', S), is = OF.kid(rule, 'iconSet', S);
        const cfvos = (el) => OF.kids(el, 'cfvo', S).map((v) => ({ type: v.getAttribute('type'), val: v.getAttribute('val'), gte: v.getAttribute('gte') !== '0' }));
        if (cs) { R.cfvo = cfvos(cs); R.colors = OF.kids(cs, 'color', S).map((c) => this.color(c) || '#ffffff'); }
        if (db) { R.cfvo = cfvos(db); R.barColor = this.color(OF.kid(db, 'color', S)) || '#638ec6'; R.showValue = db.getAttribute('showValue') !== '0'; }
        if (is) { R.cfvo = cfvos(is); R.iconSet = is.getAttribute('iconSet') || '3TrafficLights1'; R.reverse = is.getAttribute('reverse') === '1'; R.showValue = is.getAttribute('showValue') !== '0'; }
        R.rank = +(rule.getAttribute('rank') || 10); R.bottom = rule.getAttribute('bottom') === '1'; R.percent = rule.getAttribute('percent') === '1';
        R.above = rule.getAttribute('aboveAverage') !== '0'; R.equalAvg = rule.getAttribute('equalAverage') === '1';
        sh.cf.push(R);
      }
    }
    sh.cf.sort((a, b) => a.priority - b.priority);
  };
  const baseLoadSheet = XD.loadSheet;
  XD.loadSheet = async function (sh) {
    await baseLoadSheet.call(this, sh);
    this.parseCF(sh);
    this.parseFilters(sh);
  };

  /* ------------------------------- evaluación ------------------------------- */
  XD.cfCells = function (R, sh) {
    const out = [];
    for (const g of R.ranges) {
      const r2 = Math.min(g.r2, Math.max(sh.maxR, g.r1)), c2 = Math.min(g.c2, Math.max(sh.maxC, g.c1));
      for (let r = g.r1; r <= r2; r++) for (let c = g.c1; c <= c2; c++) { out.push([r, c]); if (out.length > 30000) return out; }
    }
    return out;
  };
  XD.calc = function (sh, f) {
    if (!this.hf || sh.hfId == null) return null;
    try { const v = this.hf.calculateFormula('=' + f, sh.hfId); return v && typeof v === 'object' && 'value' in v ? { error: v.value } : v; } catch (e) { return null; }
  };
  XD.constOrCalc = function (sh, f, dr, dc) {
    const t = (f || '').trim();
    if (/^-?\d+(\.\d+)?(E[+-]?\d+)?$/i.test(t)) return +t;
    if (/^"(?:[^"]|"")*"$/.test(t)) return t.slice(1, -1).replace(/""/g, '"');
    return this.calc(sh, dr || dc ? shiftFormula(t, dr, dc) : t);
  };
  XD.computeCF = function (sh) {
    const map = new Map();
    if (!sh.cf || !sh.cf.length) { sh.cfMap = map; return map; }
    if (!this.dxfs) this.loadDxfs();
    const si = this.sheets.indexOf(sh);
    const val = (r, c) => this.value(si, r, c);
    const set = (k, props) => { let o = map.get(k); if (!o) map.set(k, (o = {})); for (const p in props) if (o[p] === undefined && props[p] !== undefined) o[p] = props[p]; };
    const stopped = new Set();
    for (const R of sh.cf) {
      const cells = this.cfCells(R, sh);
      const nums = [];
      const xs = cells.map(([r, c]) => { const x = val(r, c); if (typeof x.v === 'number') nums.push(x.v); return x; });
      const sorted = nums.slice().sort((a, b) => a - b);
      const min = sorted[0], max = sorted[sorted.length - 1];
      const resolve = (v, fallback) => {
        if (!v) return fallback;
        const n = parseFloat(v.val);
        switch (v.type) {
          case 'min': return min;
          case 'max': return max;
          case 'num': return isNaN(n) ? (typeof this.constOrCalc(sh, v.val) === 'number' ? this.constOrCalc(sh, v.val) : fallback) : n;
          case 'percent': return min + ((max - min) * n) / 100;
          case 'percentile': { if (!sorted.length) return fallback; const pos = ((sorted.length - 1) * n) / 100, lo = Math.floor(pos); return sorted[lo] + (sorted[Math.min(lo + 1, sorted.length - 1)] - sorted[lo]) * (pos - lo); }
          case 'formula': { const x = this.constOrCalc(sh, v.val); return typeof x === 'number' ? x : fallback; }
          case 'autoMin': return Math.min(0, min);
          case 'autoMax': return Math.max(0, max);
          default: return fallback;
        }
      };
      const style = R.dxf != null ? this.dxfs[R.dxf] || {} : {};
      const hit = (k) => { if (stopped.has(k)) return; set(k, style); if (R.stop) stopped.add(k); };
      const first = R.ranges[0];
      if (R.type === 'colorScale' && R.colors && sorted.length) {
        const stops = R.cfvo.map((v, i) => resolve(v, i === 0 ? min : max));
        cells.forEach(([r, c], i) => {
          const v = xs[i].v; if (typeof v !== 'number') return;
          const k = r + ',' + c; if (stopped.has(k)) return;
          let col;
          if (v <= stops[0]) col = R.colors[0];
          else if (v >= stops[stops.length - 1]) col = R.colors[R.colors.length - 1];
          else for (let j = 0; j < stops.length - 1; j++) if (v <= stops[j + 1]) { col = mix(R.colors[j], R.colors[j + 1], stops[j + 1] === stops[j] ? 1 : (v - stops[j]) / (stops[j + 1] - stops[j])); break; }
          set(k, { bg: col });
        });
      } else if (R.type === 'dataBar' && sorted.length) {
        const lo = resolve(R.cfvo[0], min), hi = resolve(R.cfvo[1], max);
        cells.forEach(([r, c], i) => {
          const v = xs[i].v; if (typeof v !== 'number') return;
          const k = r + ',' + c; if (stopped.has(k)) return;
          const pct = hi === lo ? 100 : Math.max(0, Math.min(1, (v - lo) / (hi - lo))) * 100;
          set(k, { bar: { pct: Math.max(pct, 4), color: R.barColor }, hideValue: !R.showValue });
        });
      } else if (R.type === 'iconSet' && sorted.length) {
        const th = R.cfvo.map((v, i) => resolve(v, i ? max : min));
        const n = th.length;
        cells.forEach(([r, c], i) => {
          const v = xs[i].v; if (typeof v !== 'number') return;
          const k = r + ',' + c; if (stopped.has(k)) return;
          let idx = 0;
          for (let j = n - 1; j > 0; j--) if (R.cfvo[j].gte ? v >= th[j] : v > th[j]) { idx = j; break; }
          if (R.reverse) idx = n - 1 - idx;
          set(k, { icon: { set: R.iconSet, idx, n }, hideValue: !R.showValue });
        });
      } else if (R.type === 'top10') {
        const cnt = R.percent ? Math.max(1, Math.floor((sorted.length * R.rank) / 100)) : R.rank;
        const lim = R.bottom ? sorted[Math.min(cnt, sorted.length) - 1] : sorted[Math.max(0, sorted.length - cnt)];
        cells.forEach(([r, c], i) => { const v = xs[i].v; if (typeof v === 'number' && (R.bottom ? v <= lim : v >= lim)) hit(r + ',' + c); });
      } else if (R.type === 'aboveAverage') {
        const avg = nums.reduce((a, b) => a + b, 0) / (nums.length || 1);
        cells.forEach(([r, c], i) => { const v = xs[i].v; if (typeof v !== 'number') return; if (R.above ? v > avg || (R.equalAvg && v === avg) : v < avg || (R.equalAvg && v === avg)) hit(r + ',' + c); });
      } else if (R.type === 'duplicateValues' || R.type === 'uniqueValues') {
        const cnt = new Map();
        const keyOf = (x) => (x.v == null || x.v === '' ? null : String(typeof x.v === 'string' ? x.v.toLowerCase() : x.v));
        xs.forEach((x) => { const k = keyOf(x); if (k != null) cnt.set(k, (cnt.get(k) || 0) + 1); });
        cells.forEach(([r, c], i) => { const k = keyOf(xs[i]); if (k != null && (R.type === 'duplicateValues' ? cnt.get(k) > 1 : cnt.get(k) === 1)) hit(r + ',' + c); });
      } else {
        cells.forEach(([r, c], i) => {
          const x = xs[i], v = x.v, dr = r - first.r1, dc = c - first.c1;
          let ok = false;
          const txt = this.display(x).toLowerCase(), q = (R.text || '').toLowerCase();
          switch (R.type) {
            case 'cellIs': {
              const a = this.constOrCalc(sh, R.formulas[0], dr, dc), b = this.constOrCalc(sh, R.formulas[1], dr, dc);
              if (v && typeof v === 'object') break;
              ok = cmp(R.op, v == null ? 0 : v, a, b);
              break;
            }
            case 'containsText': ok = txt.includes(q); break;
            case 'notContainsText': ok = !txt.includes(q); break;
            case 'beginsWith': ok = txt.startsWith(q); break;
            case 'endsWith': ok = txt.endsWith(q); break;
            case 'containsBlanks': ok = v == null || String(v).trim() === ''; break;
            case 'notContainsBlanks': ok = !(v == null || String(v).trim() === ''); break;
            case 'containsErrors': ok = !!(v && typeof v === 'object'); break;
            case 'notContainsErrors': ok = !(v && typeof v === 'object'); break;
            case 'expression': { const res = R.formulas[0] ? this.calc(sh, shiftFormula(R.formulas[0], dr, dc)) : null; ok = res === true || (typeof res === 'number' && res !== 0); break; }
            default: ok = false;
          }
          if (ok) hit(r + ',' + c);
        });
      }
    }
    sh.cfMap = map;
    return map;
  };
  function cmp(op, v, a, b) {
    const norm = (x) => (typeof x === 'string' ? x.toLowerCase() : x);
    const V = norm(v), A = norm(a), B = norm(b);
    const lt = (x, y) => (typeof x === typeof y ? x < y : typeof x === 'number');
    switch (op) {
      case 'equal': return V === A || (typeof V === 'number' && +A === V && A !== '' && A != null);
      case 'notEqual': return !(V === A || (typeof V === 'number' && +A === V));
      case 'greaterThan': return lt(A, V);
      case 'lessThan': return lt(V, A);
      case 'greaterThanOrEqual': return !lt(V, A);
      case 'lessThanOrEqual': return !lt(A, V);
      case 'between': { const lo = lt(B, A) ? B : A, hi = lo === A ? B : A; return !lt(V, lo) && !lt(hi, V); }
      case 'notBetween': { const lo = lt(B, A) ? B : A, hi = lo === A ? B : A; return lt(V, lo) || lt(hi, V); }
      default: return false;
    }
  }

  /* ------------------------------- dibujo ------------------------------- */
  const ICONS = {
    arrows: ['▼', '►', '▲'], lights: ['●', '●', '●'], symbols: ['✖', '!', '✔'], flags: ['⚑', '⚑', '⚑'], stars: ['☆', '★', '★'], rating: ['▁', '▃', '▅', '▇'],
  };
  const ICON_COLORS = ['#c0392b', '#e6a118', '#2e9e4f'];
  function iconFor(ic) {
    const s = ic.set;
    const fam = /Arrow/i.test(s) ? 'arrows' : /Symbol/i.test(s) ? 'symbols' : /Flag/i.test(s) ? 'flags' : /Star/i.test(s) ? 'stars' : /Rating|Quarters|Boxes/i.test(s) ? 'rating' : 'lights';
    const t = ic.n > 1 ? ic.idx / (ic.n - 1) : 1;
    const color = t < 0.34 ? ICON_COLORS[0] : t < 0.67 ? ICON_COLORS[1] : ICON_COLORS[2];
    const glyphs = ICONS[fam];
    let g = glyphs[Math.min(glyphs.length - 1, Math.round(t * (glyphs.length - 1)))];
    if (fam === 'arrows' && ic.n >= 4 && t > 0.2 && t < 0.8) g = t < 0.5 ? '◢' : '◥';
    return { g, color: /Gray/i.test(s) ? '#6b7480' : color };
  }
  const baseFillCell = XD.fillCell;
  XD.fillCell = function (td, sh, r, c) {
    baseFillCell.call(this, td, sh, r, c);
    if (td.dataset.cf) { td.style.backgroundColor = td.style.color = td.style.fontWeight = td.style.fontStyle = td.style.textDecoration = td.style.boxShadow = ''; delete td.dataset.cf; td.classList.remove('cf-bar'); }
    const st = sh.cfMap && sh.cfMap.get(r + ',' + c);
    if (!st) return;
    td.dataset.cf = '1';
    if (st.bg) { td.style.backgroundColor = st.bg; td.style.backgroundImage = 'none'; }
    if (st.color) td.style.color = st.color;
    if (st.b != null) td.style.fontWeight = st.b ? '700' : '400';
    if (st.i != null) td.style.fontStyle = st.i ? 'italic' : 'normal';
    if (st.strike || st.u) td.style.textDecoration = [st.u && 'underline', st.strike && 'line-through'].filter(Boolean).join(' ');
    const sh2 = [st.btop && `inset 0 1px 0 ${st.btop}`, st.bbottom && `inset 0 -1px 0 ${st.bbottom}`, st.bleft && `inset 1px 0 0 ${st.bleft}`, st.bright && `inset -1px 0 0 ${st.bright}`].filter(Boolean);
    if (sh2.length) td.style.boxShadow = sh2.join(',');
    const v = td.querySelector('.xl-v');
    if (st.hideValue && v) v.style.visibility = 'hidden';
    if (st.bar) {
      td.classList.add('cf-bar');
      td.prepend(OF.el('i', { class: 'xl-cfbar', style: { width: `calc(${st.bar.pct}% - 4px)`, background: `linear-gradient(90deg, ${st.bar.color}, ${mix(st.bar.color, '#ffffff', 0.75)})`, borderColor: st.bar.color } }));
    }
    if (st.icon) { const ic = iconFor(st.icon); td.prepend(OF.el('i', { class: 'xl-cfic', style: { color: ic.color } }, ic.g)); }
  };
  const baseShowSheet = XD.showSheet;
  XD.showSheet = function (i) {
    if (this.sheets[i]) this.computeCF(this.sheets[i]);
    baseShowSheet.call(this, i);
  };
  XD.refreshCF = function () {
    const sh = this.sheets[this.cur];
    if (!sh) return;
    if ((!sh.cf || !sh.cf.length) && !(sh.cfMap && sh.cfMap.size)) { sh.cfMap = new Map(); return; }
    const old = sh.cfMap || new Map();
    const now = this.computeCF(sh);
    const keys = new Set([...old.keys(), ...now.keys()]);
    for (const k of keys) {
      if (JSON.stringify(old.get(k)) === JSON.stringify(now.get(k))) continue;
      const [r, c] = k.split(',').map(Number);
      const td = this.tdAt(r, c);
      if (td) this.fillCell(td, sh, r, c);
    }
  };
  const baseCommit = XD.commit;
  XD.commit = function (input, noHist) { const res = baseCommit.call(this, input, noHist); this.refreshCF(); return res; };

  /* ------------------------------- crear reglas ------------------------------- */
  const PRESETS = {
    red: { label: 'Relleno rojo claro con texto rojo oscuro', bg: 'FFC7CE', color: '9C0006' },
    yellow: { label: 'Relleno amarillo con texto amarillo oscuro', bg: 'FFEB9C', color: '9C5700' },
    green: { label: 'Relleno verde con texto verde oscuro', bg: 'C6EFCE', color: '006100' },
    redFill: { label: 'Relleno rojo claro', bg: 'FFC7CE' },
    redText: { label: 'Texto rojo', color: 'FF0000' },
    bold: { label: 'Negrita', b: true },
  };
  XD.addDxf = function (p) {
    const doc = this.stylesDoc, root = doc.documentElement;
    const mk = (n) => doc.createElementNS(S, root.prefix ? root.prefix + ':' + n : n);
    const list = OF.ordered(root, S, root.prefix || '', 'dxfs', STYLE_ORDER);
    const d = mk('dxf');
    if (p.color || p.b) {
      const f = mk('font');
      if (p.b) f.appendChild(mk('b'));
      if (p.color) { const c = mk('color'); c.setAttribute('rgb', 'FF' + p.color); f.appendChild(c); }
      d.appendChild(f);
    }
    if (p.bg) { const fl = mk('fill'), pf = mk('patternFill'), bg = mk('bgColor'); bg.setAttribute('rgb', 'FF' + p.bg); pf.appendChild(bg); fl.appendChild(pf); d.appendChild(fl); }
    list.appendChild(d);
    const n = OF.kids(list, 'dxf', S).length;
    list.setAttribute('count', n);
    this.stylesDirty = true;
    this.dxfs = null;
    return n - 1;
  };
  XD.cfTarget = function () {
    const g = this.workRect();
    if (g.r1 !== g.r2 || g.c1 !== g.c2) return g;
    // una sola celda: la columna de la tabla en la que está (sin el encabezado)
    const d = this.dataRegion();
    if (d.r2 > d.r1 && g.c1 >= d.c1 && g.c1 <= d.c2 && g.r1 >= d.r1 && g.r1 <= d.r2) {
      const head = this.hasHeader(d) ? 1 : 0;
      return { r1: d.r1 + head, r2: d.r2, c1: g.c1, c2: g.c1 };
    }
    return g;
  };
  XD.addCF = function (g, build) {
    if (!this.stylesDoc) { OF.toast('Este archivo no tiene hoja de estilos.'); return; }
    this.stopEdit && this.stopEdit(true);
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], doc = sh.doc, root = doc.documentElement;
    const mk = (n) => doc.createElementNS(S, root.prefix ? root.prefix + ':' + n : n);
    for (const r of OF.all(root, 'cfRule', S)) r.setAttribute('priority', +(r.getAttribute('priority') || 1) + 1);
    const cf = mk('conditionalFormatting');
    cf.setAttribute('sqref', rangeStr(g));
    const rule = mk('cfRule');
    rule.setAttribute('priority', '1');
    build(rule, mk, addr(g.r1, g.c1));
    cf.appendChild(rule);
    const existing = OF.kids(root, 'conditionalFormatting', S);
    if (existing.length) existing[existing.length - 1].after(cf);
    else {
      const idx = WS_ORDER.indexOf('conditionalFormatting');
      let before = null;
      for (let c = root.firstElementChild; c; c = c.nextElementSibling) if (WS_ORDER.indexOf(c.localName) > idx) { before = c; break; }
      root.insertBefore(cf, before);
    }
    this.parseCF(sh);
    this.markDirty();
    this.refreshCF();
  };
  const numText = (doc, s) => {
    let t = String(s).trim();
    if (doc.locale === 'cl') { if (/^-?[\d.]*,?\d+$/.test(t) || /^-?\d{1,3}(\.\d{3})+$/.test(t)) t = t.replace(/\./g, '').replace(',', '.'); }
    else t = t.replace(/,/g, '');
    return t !== '' && isFinite(+t) ? String(+t) : '"' + String(s).replace(/"/g, '""') + '"';
  };
  const OPS = { gt: ['greaterThan', 'Mayor que'], lt: ['lessThan', 'Menor que'], between: ['between', 'Entre'], eq: ['equal', 'Igual a'], text: ['containsText', 'Texto que contiene'] };
  XD.cfDialog = function (kind) {
    const anchor = document.querySelector('.ribbon button[title="Formato condicional"]') || document.querySelector('.ribbon');
    const g = this.cfTarget();
    OF.popup(anchor, (p, close) => {
      p.classList.add('cfd');
      p.append(OF.el('div', { class: 'pop-h' }, `${OPS[kind][1]}… · ${rangeStr(g)}`));
      const in1 = OF.el('input', { class: 'flt-q', placeholder: kind === 'text' ? 'Texto' : 'Valor' });
      const in2 = OF.el('input', { class: 'flt-q', placeholder: 'y' });
      p.append(in1);
      if (kind === 'between') p.append(in2);
      const selP = OF.el('select', { class: 'flt-q' }, ...Object.entries(PRESETS).map(([k, v]) => OF.el('option', { value: k }, v.label)));
      p.append(OF.el('div', { class: 'cfd-l' }, 'con'), selP);
      const ok = OF.el('button', { class: 'btn primary', onclick: () => {
        if (!in1.value.trim()) { in1.focus(); return; }
        close();
        const dx = this.addDxf(PRESETS[selP.value]);
        this.addCF(g, (rule, mk, tl) => {
          rule.setAttribute('dxfId', dx);
          if (kind === 'text') {
            const q = in1.value.trim();
            rule.setAttribute('type', 'containsText'); rule.setAttribute('operator', 'containsText'); rule.setAttribute('text', q);
            const f = mk('formula'); f.textContent = `NOT(ISERROR(SEARCH("${q.replace(/"/g, '""')}",${tl})))`; rule.appendChild(f);
          } else {
            rule.setAttribute('type', 'cellIs'); rule.setAttribute('operator', OPS[kind][0]);
            for (const inp of kind === 'between' ? [in1, in2] : [in1]) { const f = mk('formula'); f.textContent = numText(this, inp.value); rule.appendChild(f); }
          }
        });
        OF.toast('Regla de formato condicional aplicada a ' + rangeStr(g));
      } }, 'Aceptar');
      p.append(OF.el('div', { class: 'flt-btns' }, OF.el('button', { class: 'btn', onclick: close }, 'Cancelar'), ok));
      p.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok.click(); } });
    }, false);
  };
  XD.cfQuick = function (kind, opt) {
    const g = this.cfTarget();
    const done = (msg) => OF.toast((msg || 'Formato condicional aplicado a ') + rangeStr(g));
    if (kind === 'dup' || kind === 'top' || kind === 'bottom' || kind === 'above' || kind === 'below') {
      const dx = this.addDxf(PRESETS[kind === 'below' || kind === 'bottom' ? 'red' : kind === 'dup' ? 'red' : 'green']);
      this.addCF(g, (rule) => {
        rule.setAttribute('dxfId', dx);
        if (kind === 'dup') rule.setAttribute('type', 'duplicateValues');
        else if (kind === 'top' || kind === 'bottom') { rule.setAttribute('type', 'top10'); rule.setAttribute('rank', '10'); if (kind === 'bottom') rule.setAttribute('bottom', '1'); }
        else { rule.setAttribute('type', 'aboveAverage'); if (kind === 'below') rule.setAttribute('aboveAverage', '0'); }
      });
      return done();
    }
    if (kind === 'bar') {
      this.addCF(g, (rule, mk) => {
        rule.setAttribute('type', 'dataBar');
        const db = mk('dataBar');
        for (const t of ['min', 'max']) { const v = mk('cfvo'); v.setAttribute('type', t); db.appendChild(v); }
        const c = mk('color'); c.setAttribute('rgb', 'FF' + opt); db.appendChild(c);
        rule.appendChild(db);
      });
      return done();
    }
    if (kind === 'scale') {
      this.addCF(g, (rule, mk) => {
        rule.setAttribute('type', 'colorScale');
        const cs = mk('colorScale');
        const three = opt.length === 3;
        const vs = three ? [['min'], ['percentile', '50'], ['max']] : [['min'], ['max']];
        for (const [t, val] of vs) { const v = mk('cfvo'); v.setAttribute('type', t); if (val) v.setAttribute('val', val); cs.appendChild(v); }
        for (const col of opt) { const c = mk('color'); c.setAttribute('rgb', 'FF' + col); cs.appendChild(c); }
        rule.appendChild(cs);
      });
      return done();
    }
    if (kind === 'icons') {
      this.addCF(g, (rule, mk) => {
        rule.setAttribute('type', 'iconSet');
        const is = mk('iconSet'); is.setAttribute('iconSet', opt);
        for (const val of ['0', '33', '67']) { const v = mk('cfvo'); v.setAttribute('type', 'percent'); v.setAttribute('val', val); is.appendChild(v); }
        rule.appendChild(is);
      });
      return done();
    }
  };
  XD.clearCF = function (scope) {
    const sh = this.sheets[this.cur], root = sh.doc.documentElement;
    const list = OF.kids(root, 'conditionalFormatting', S);
    if (!list.length) { OF.toast('Esta hoja no tiene reglas de formato condicional.'); return; }
    OF.checkpoint && OF.checkpoint();
    const g = this.workRect();
    let n = 0;
    for (const cf of list) {
      if (scope === 'sheet') { root.removeChild(cf); n++; continue; }
      const ranges = (cf.getAttribute('sqref') || '').split(/\s+/).filter(Boolean);
      const keep = ranges.filter((s) => { const q = parseRange(s); return !(q && q.r1 <= g.r2 && q.r2 >= g.r1 && q.c1 <= g.c2 && q.c2 >= g.c1); });
      if (keep.length === ranges.length) continue;
      n++;
      if (keep.length) cf.setAttribute('sqref', keep.join(' ')); else root.removeChild(cf);
    }
    this.parseCF(sh);
    this.markDirty();
    this.refreshCF();
    OF.toast(n ? 'Reglas borradas' : 'Las celdas seleccionadas no tienen reglas.');
  };

  /* ------------------------------- filtros por condición ------------------------------- */
  // sh.filters[c] = ['a','b'] (lista de valores) | { type:'custom', and, conds:[{op,val}] } | { type:'avg', above } | { type:'top', n, bottom }
  XD.parseFilters = function (sh) {
    const af = OF.kid(sh.doc.documentElement, 'autoFilter', S);
    sh.filters = af ? {} : null;
    if (!af || !sh.autoFilter) return;
    for (const fc of OF.kids(af, 'filterColumn', S)) {
      const c = sh.autoFilter.c1 + +(fc.getAttribute('colId') || 0);
      const fs = OF.kid(fc, 'filters', S), cfs = OF.kid(fc, 'customFilters', S), dy = OF.kid(fc, 'dynamicFilter', S), t10 = OF.kid(fc, 'top10', S);
      if (fs) { const l = OF.kids(fs, 'filter', S).map((f) => f.getAttribute('val')); if (fs.getAttribute('blank') === '1') l.push(''); sh.filters[c] = l; }
      else if (cfs) sh.filters[c] = { type: 'custom', and: cfs.getAttribute('and') === '1', conds: OF.kids(cfs, 'customFilter', S).map((f) => ({ op: f.getAttribute('operator') || 'equal', val: f.getAttribute('val') || '' })) };
      else if (dy && /Average/.test(dy.getAttribute('type') || '')) sh.filters[c] = { type: 'avg', above: dy.getAttribute('type') === 'aboveAverage' };
      else if (t10) sh.filters[c] = { type: 'top', n: +(t10.getAttribute('val') || 10), bottom: t10.getAttribute('top') === '0', percent: t10.getAttribute('percent') === '1' };
    }
  };
  const wild = (pat) => new RegExp('^' + pat.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/~\*/g, '\u0001').replace(/~\?/g, '\u0002').replace(/\*/g, '.*').replace(/\?/g, '.').replace(/\u0001/g, '\\*').replace(/\u0002/g, '\\?') + '$', 'i');
  function condMatch(cond, x, txt) {
    const v = x.v;
    const num = cond.val !== '' && isFinite(+cond.val) && !/[*?]/.test(cond.val) ? +cond.val : null;
    const isNum = typeof v === 'number';
    switch (cond.op) {
      case 'equal': return num != null && isNum ? v === num : wild(cond.val).test(txt);
      case 'notEqual': return !(num != null && isNum ? v === num : wild(cond.val).test(txt));
      case 'greaterThan': return num != null ? isNum && v > num : txt.toLowerCase() > cond.val.toLowerCase();
      case 'greaterThanOrEqual': return num != null ? isNum && v >= num : txt.toLowerCase() >= cond.val.toLowerCase();
      case 'lessThan': return num != null ? isNum && v < num : txt.toLowerCase() < cond.val.toLowerCase();
      case 'lessThanOrEqual': return num != null ? isNum && v <= num : txt.toLowerCase() <= cond.val.toLowerCase();
      default: return true;
    }
  }
  XD.filterPass = function (sh, spec, r, c, stats) {
    const x = this.value(this.cur, r, c), txt = this.display(x);
    if (Array.isArray(spec)) return spec.includes(txt);
    if (spec.type === 'custom') { const res = spec.conds.map((cd) => condMatch(cd, x, txt)); return spec.and ? res.every(Boolean) : res.some(Boolean); }
    if (typeof x.v !== 'number') return false;
    if (spec.type === 'avg') return spec.above ? x.v > stats.avg : x.v < stats.avg;
    if (spec.type === 'top') return spec.bottom ? x.v <= stats.lim : x.v >= stats.lim;
    return true;
  };
  XD.applyFilter = async function (c, spec) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur], g = sh.autoFilter, root = sh.doc.documentElement;
    sh.filters = sh.filters || {};
    if (spec) sh.filters[c] = spec; else delete sh.filters[c];
    const stats = {};
    for (const cc in sh.filters) {
      const sp = sh.filters[cc];
      if (Array.isArray(sp) || sp.type === 'custom') continue;
      const nums = []; for (let r = g.r1 + 1; r <= g.r2; r++) { const v = this.value(this.cur, r, +cc).v; if (typeof v === 'number') nums.push(v); }
      nums.sort((a, b) => a - b);
      const n = sp.percent ? Math.max(1, Math.floor((nums.length * sp.n) / 100)) : sp.n;
      stats[cc] = { avg: nums.reduce((a, b) => a + b, 0) / (nums.length || 1), lim: sp.bottom ? nums[Math.min(n, nums.length) - 1] : nums[Math.max(0, nums.length - n)] };
    }
    for (let r = g.r1 + 1; r <= g.r2; r++) {
      let show = true;
      for (const cc in sh.filters) if (!this.filterPass(sh, sh.filters[cc], r, +cc, stats[cc])) { show = false; break; }
      let row = sh.rowEls.get(r);
      if (!show) { if (!row) { this.cellEl(sh, r, g.c1, true); row = sh.rowEls.get(r); } row.setAttribute('hidden', '1'); }
      else if (row) row.removeAttribute('hidden');
    }
    const af = OF.kid(root, 'autoFilter', S);
    const mk = (n) => sh.doc.createElementNS(S, root.prefix ? root.prefix + ':' + n : n);
    for (const fc of OF.kids(af, 'filterColumn', S)) af.removeChild(fc);
    for (const cc of Object.keys(sh.filters).map(Number).sort((a, b) => a - b)) {
      const sp = sh.filters[cc];
      const fc = mk('filterColumn');
      fc.setAttribute('colId', cc - g.c1);
      if (Array.isArray(sp)) {
        const fs = mk('filters');
        for (const v of sp) { if (v === '') { fs.setAttribute('blank', '1'); continue; } const f = mk('filter'); f.setAttribute('val', v); fs.appendChild(f); }
        fc.appendChild(fs);
      } else if (sp.type === 'custom') {
        const cfs = mk('customFilters');
        if (sp.and) cfs.setAttribute('and', '1');
        for (const cd of sp.conds) { const f = mk('customFilter'); if (cd.op !== 'equal') f.setAttribute('operator', cd.op); f.setAttribute('val', cd.val); cfs.appendChild(f); }
        fc.appendChild(cfs);
      } else if (sp.type === 'avg') {
        const d = mk('dynamicFilter'); d.setAttribute('type', sp.above ? 'aboveAverage' : 'belowAverage'); d.setAttribute('val', String(stats[cc].avg)); fc.appendChild(d);
      } else if (sp.type === 'top') {
        const t = mk('top10'); if (sp.bottom) t.setAttribute('top', '0'); if (sp.percent) t.setAttribute('percent', '1'); t.setAttribute('val', sp.n); if (stats[cc].lim != null) t.setAttribute('filterVal', stats[cc].lim); fc.appendChild(t);
      }
      af.appendChild(fc);
    }
    const keep = sh.filters;
    await this.reloadSheet();
    sh.filters = keep;
    this.showSheetKeep();
    let shown = 0; for (let r = g.r1 + 1; r <= g.r2; r++) if (!(sh.rowInfo[r] || {}).hidden) shown++;
    OF.toast(`${shown} de ${g.r2 - g.r1} filas visibles`);
  };
  const baseFilterMenu = XD.filterMenu;
  XD.filterMenu = function (btn, c) {
    baseFilterMenu.call(this, btn, c);
    const p = document.querySelector('.pop.flt');
    if (!p) return;
    const sh = this.sheets[this.cur], g = sh.autoFilter;
    let n = 0, t = 0;
    for (let r = g.r1 + 1; r <= g.r2; r++) { const v = this.value(this.cur, r, c).v; if (typeof v === 'number') n++; else if (v != null && v !== '') t++; }
    const numeric = n >= t;
    const cur = sh.filters && sh.filters[c];
    const b = OF.el('button', { class: 'pop-item' + (cur && !Array.isArray(cur) ? ' on' : ''), onclick: () => this.condFilterMenu(btn, c, numeric) }, (numeric ? 'Filtros de número' : 'Filtros de texto') + '  ›');
    const hr = p.querySelector('hr');
    hr.after(b, OF.el('hr'));
    // si hay un filtro por condición, la lista de valores aparece marcada completa
    if (cur && !Array.isArray(cur)) p.querySelectorAll('.flt-list input').forEach((cb) => { cb.checked = true; });
  };
  const NUM_OPS = [['equal', 'Es igual a'], ['notEqual', 'No es igual a'], ['greaterThan', 'Mayor que'], ['greaterThanOrEqual', 'Mayor o igual que'], ['lessThan', 'Menor que'], ['lessThanOrEqual', 'Menor o igual que'], ['between', 'Entre…']];
  const TXT_OPS = [['equal', 'Es igual a'], ['notEqual', 'No es igual a'], ['beginsWith', 'Comienza por'], ['endsWith', 'Termina con'], ['contains', 'Contiene'], ['notContains', 'No contiene']];
  const toXml = (op, v) => (op === 'beginsWith' ? { op: 'equal', val: v + '*' } : op === 'endsWith' ? { op: 'equal', val: '*' + v } : op === 'contains' ? { op: 'equal', val: '*' + v + '*' } : op === 'notContains' ? { op: 'notEqual', val: '*' + v + '*' } : { op, val: v });
  const fromXml = (cd) => {
    const m = /^\*(.*)\*$/.exec(cd.val);
    if (m && cd.op === 'equal') return ['contains', m[1]];
    if (m && cd.op === 'notEqual') return ['notContains', m[1]];
    if (cd.op === 'equal' && /^[^*]+\*$/.test(cd.val)) return ['beginsWith', cd.val.slice(0, -1)];
    if (cd.op === 'equal' && /^\*[^*]+$/.test(cd.val)) return ['endsWith', cd.val.slice(1)];
    return [cd.op, cd.val];
  };
  XD.condFilterMenu = function (btn, c, numeric) {
    const sh = this.sheets[this.cur];
    const cur = sh.filters && sh.filters[c];
    const ops = numeric ? NUM_OPS : TXT_OPS;
    const toLocal = (v) => (numeric && this.locale === 'cl' && isFinite(+v) && v !== '' ? String(v).replace('.', ',') : v);
    const fromLocal = (v) => { v = v.trim(); if (numeric && this.locale === 'cl' && /^-?[\d.]*,?\d+$/.test(v)) return v.replace(/\./g, '').replace(',', '.'); return v; };
    OF.popup(btn, (p, close) => {
      p.classList.add('flt', 'cfd');
      p.append(OF.el('div', { class: 'pop-h' }, (numeric ? 'Filtro de número · ' : 'Filtro de texto · ') + 'columna ' + colName(c)));
      const mkSel = () => OF.el('select', { class: 'flt-q' }, OF.el('option', { value: '' }, '—'), ...ops.map(([k, l]) => OF.el('option', { value: k }, l)));
      const s1 = mkSel(), v1 = OF.el('input', { class: 'flt-q', placeholder: 'Valor' }), s2 = mkSel(), v2 = OF.el('input', { class: 'flt-q', placeholder: 'Valor' });
      const and = OF.el('input', { type: 'radio', name: 'fl-andor', checked: true }), or = OF.el('input', { type: 'radio', name: 'fl-andor' });
      s1.value = ops[numeric ? 2 : 4][0];
      if (cur && cur.type === 'custom') {
        const [o1, x1] = fromXml(cur.conds[0]); s1.value = o1; v1.value = toLocal(x1);
        if (cur.conds[1]) { const [o2, x2] = fromXml(cur.conds[1]); s2.value = o2; v2.value = toLocal(x2); }
        (cur.and ? and : or).checked = true;
        // "entre" guardado como >= y <=
        if (cur.and && cur.conds.length === 2 && cur.conds[0].op === 'greaterThanOrEqual' && cur.conds[1].op === 'lessThanOrEqual' && numeric) { s1.value = 'between'; v1.value = toLocal(cur.conds[0].val); v2.value = toLocal(cur.conds[1].val); s2.value = ''; }
      }
      const second = OF.el('div', { class: 'cfd-2' }, OF.el('div', { class: 'cfd-andor' }, OF.el('label', null, and, ' Y'), OF.el('label', null, or, ' O')), s2, v2);
      const sync = () => { const btw = s1.value === 'between'; second.style.display = btw ? 'none' : ''; v2.placeholder = 'Valor'; if (btw) { betweenTo.style.display = ''; } else betweenTo.style.display = 'none'; };
      const betweenTo = OF.el('input', { class: 'flt-q', placeholder: 'y' });
      if (s1.value === 'between') betweenTo.value = v2.value;
      s1.addEventListener('change', sync);
      p.append(s1, v1, betweenTo, second);
      sync();
      if (numeric) {
        const quick = (label, spec) => OF.el('button', { class: 'pop-item', onclick: () => { close(); this.applyFilter(c, spec); } }, label);
        p.append(OF.el('hr'), quick('Por encima del promedio', { type: 'avg', above: true }), quick('Por debajo del promedio', { type: 'avg', above: false }), quick('10 mejores', { type: 'top', n: 10, bottom: false }), quick('10 inferiores', { type: 'top', n: 10, bottom: true }));
      }
      const ok = OF.el('button', { class: 'btn primary', onclick: () => {
        if (!s1.value || v1.value.trim() === '') { v1.focus(); return; }
        let conds;
        if (s1.value === 'between') {
          if (betweenTo.value.trim() === '') { betweenTo.focus(); return; }
          conds = [{ op: 'greaterThanOrEqual', val: fromLocal(v1.value) }, { op: 'lessThanOrEqual', val: fromLocal(betweenTo.value) }];
          close(); this.applyFilter(c, { type: 'custom', and: true, conds }); return;
        }
        conds = [toXml(s1.value, fromLocal(v1.value))];
        if (s2.value && v2.value.trim() !== '') conds.push(toXml(s2.value, fromLocal(v2.value)));
        close();
        this.applyFilter(c, { type: 'custom', and: and.checked, conds });
      } }, 'Aceptar');
      p.append(OF.el('div', { class: 'flt-btns' }, OF.el('button', { class: 'btn', onclick: () => { close(); this.applyFilter(c, null); } }, 'Borrar filtro'), ok));
      p.addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); ok.click(); } });
    }, false);
  };
  /* ------------------------------- administrar reglas ------------------------------- */
  const OPL = { greaterThan: 'mayor que', lessThan: 'menor que', greaterThanOrEqual: 'mayor o igual que', lessThanOrEqual: 'menor o igual que', equal: 'igual a', notEqual: 'distinto de', between: 'entre', notBetween: 'no está entre' };
  XD.cfDescribe = function (R) {
    const es = (f) => (this.locale === 'cl' && OF.formulaToES ? OF.formulaToES(f) : f);
    switch (R.type) {
      case 'cellIs': return `Valor ${OPL[R.op] || R.op} ${R.formulas.map(es).join(R.op && /etween/.test(R.op) ? ' y ' : ', ')}`;
      case 'containsText': return `Contiene «${R.text || ''}»`;
      case 'notContainsText': return `No contiene «${R.text || ''}»`;
      case 'beginsWith': return `Comienza por «${R.text || ''}»`;
      case 'endsWith': return `Termina con «${R.text || ''}»`;
      case 'expression': return `Fórmula: =${es(R.formulas[0] || '')}`;
      case 'colorScale': return 'Escala de colores';
      case 'dataBar': return 'Barras de datos';
      case 'iconSet': return 'Conjunto de iconos';
      case 'top10': return `${R.bottom ? 'Inferiores' : 'Superiores'} ${R.rank}${R.percent ? '%' : ''}`;
      case 'aboveAverage': return R.above ? 'Por encima del promedio' : 'Por debajo del promedio';
      case 'duplicateValues': return 'Valores duplicados';
      case 'uniqueValues': return 'Valores únicos';
      case 'containsBlanks': return 'Celdas en blanco';
      case 'notContainsBlanks': return 'Celdas con datos';
      case 'containsErrors': return 'Celdas con errores';
      default: return R.type;
    }
  };
  XD.cfPreview = function (R) {
    if (!this.dxfs) this.loadDxfs();
    const st = R.dxf != null ? this.dxfs[R.dxf] || {} : {};
    const sw = OF.el('span', { class: 'cfm-sw' }, 'AaBb');
    if (R.colors) sw.style.background = `linear-gradient(90deg, ${R.colors.join(',')})`;
    else if (R.barColor) sw.style.background = `linear-gradient(90deg, ${R.barColor} 60%, transparent 60%)`;
    else if (R.iconSet) sw.textContent = '▲ ► ▼';
    if (st.bg) sw.style.background = st.bg;
    if (st.color) sw.style.color = st.color;
    if (st.b) sw.style.fontWeight = '700';
    return sw;
  };
  XD.cfEdit = function (fn) {
    OF.checkpoint && OF.checkpoint();
    const sh = this.sheets[this.cur];
    fn(sh, sh.doc.documentElement);
    for (const cf of OF.kids(sh.doc.documentElement, 'conditionalFormatting', S)) if (!OF.kids(cf, 'cfRule', S).length) cf.parentNode.removeChild(cf);
    this.parseCF(sh);
    this.markDirty();
    this.refreshCF();
  };
  XD.cfManager = function (anchor) {
    anchor = anchor || document.querySelector('.ribbon button[title="Formato condicional"]') || document.querySelector('.ribbon');
    const sh = this.sheets[this.cur];
    OF.popup(anchor, (p, close) => {
      p.classList.add('cfd', 'cfm');
      p.append(OF.el('div', { class: 'pop-h' }, `Reglas de formato condicional · ${sh.name}`));
      const rules = sh.cf || [];
      if (!rules.length) p.append(OF.el('div', { class: 'cfm-empty' }, 'Esta hoja no tiene reglas.'));
      rules.forEach((R, i) => {
        const range = OF.el('input', { class: 'cfm-rng', value: R.ranges.map(rangeStr).join(' '), title: 'Se aplica a', 'aria-label': 'Se aplica a' });
        range.addEventListener('change', () => {
          const parts = range.value.trim().toUpperCase().split(/[\s;,]+/).filter(Boolean);
          if (!parts.length || !parts.every((x) => parseRange(x))) { OF.toast('Rango no válido. Ejemplo: B2:B20'); range.value = R.ranges.map(rangeStr).join(' '); return; }
          this.cfEdit((s2, root) => {
            const cf = R.el.parentNode;
            if (OF.kids(cf, 'cfRule', S).length > 1) {
              const nc = cf.cloneNode(false);
              cf.after(nc); nc.appendChild(R.el);
              nc.setAttribute('sqref', parts.join(' '));
            } else cf.setAttribute('sqref', parts.join(' '));
            void root;
          });
          this.cfManager(anchor);
        });
        const mv = (dir) => () => {
          const other = rules[i + dir];
          if (!other) return;
          this.cfEdit(() => { const a = R.el.getAttribute('priority'), b = other.el.getAttribute('priority'); R.el.setAttribute('priority', b); other.el.setAttribute('priority', a); });
          this.cfManager(anchor);
        };
        const del = () => { this.cfEdit(() => R.el.parentNode.removeChild(R.el)); this.cfManager(anchor); };
        const stop = OF.el('input', { type: 'checkbox', title: 'Detener si es verdad' });
        stop.checked = R.stop;
        stop.addEventListener('change', () => { this.cfEdit(() => { if (stop.checked) R.el.setAttribute('stopIfTrue', '1'); else R.el.removeAttribute('stopIfTrue'); }); });
        p.append(OF.el('div', { class: 'cfm-row' },
          this.cfPreview(R),
          OF.el('div', { class: 'cfm-main' }, OF.el('div', { class: 'cfm-d' }, this.cfDescribe(R)), range),
          OF.el('label', { class: 'cfm-stop', title: 'Detener si es verdad: las reglas de abajo no se aplican a esas celdas' }, stop, 'Detener'),
          OF.el('button', { class: 'cfm-b', title: 'Subir prioridad', disabled: i === 0, onclick: mv(-1) }, '↑'),
          OF.el('button', { class: 'cfm-b', title: 'Bajar prioridad', disabled: i === rules.length - 1, onclick: mv(1) }, '↓'),
          OF.el('button', { class: 'cfm-b danger', title: 'Eliminar regla', onclick: del }, '🗑')));
      });
      // nueva regla con fórmula
      p.append(OF.el('hr'), OF.el('div', { class: 'pop-h' }, 'Nueva regla con fórmula (se aplica a la selección)'));
      const g = this.cfTarget();
      const fIn = OF.el('input', { class: 'flt-q', placeholder: `Ej.: =${addr(g.r1, g.c1)}>${addr(g.r1, g.c1 + 1)}` });
      const selP = OF.el('select', { class: 'flt-q' }, ...Object.entries(PRESETS).map(([k, v]) => OF.el('option', { value: k }, v.label)));
      const add = OF.el('button', { class: 'btn primary', onclick: () => {
        let f = fIn.value.trim().replace(/^=/, '');
        if (!f) { fIn.focus(); return; }
        if (this.locale === 'cl' && OF.formulaToEN) f = OF.formulaToEN(f);
        const dx = this.addDxf(PRESETS[selP.value]);
        this.addCF(g, (rule, mk) => { rule.setAttribute('type', 'expression'); rule.setAttribute('dxfId', dx); const fe = mk('formula'); fe.textContent = f; rule.appendChild(fe); });
        this.cfManager(anchor);
      } }, 'Agregar');
      p.append(OF.el('div', { class: 'cfd-l' }, `Rango: ${rangeStr(g)}. La fórmula se escribe para la primera celda.`), fIn, selP, OF.el('div', { class: 'flt-btns' }, OF.el('button', { class: 'btn', onclick: close }, 'Cerrar'), add));
    }, false);
  };
})();
