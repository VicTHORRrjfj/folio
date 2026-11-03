/* Gráficos de Office (DrawingML chart): lectura del XML y dibujo en SVG.
   Se usa en Excel (con datos en vivo desde las celdas) y en PowerPoint (con los valores guardados). */
(function () {
  const OF = window.OF;
  const C = 'http://schemas.openxmlformats.org/drawingml/2006/chart';
  OF.NS.c = C;
  const PALETTE = ['#4472C4', '#ED7D31', '#A5A5A5', '#FFC000', '#5B9BD5', '#70AD47', '#264478', '#9E480E', '#636363', '#997300'];

  const k = (el, n) => OF.kid(el, n, C);
  const cacheVals = (ref) => {
    if (!ref) return null;
    let cache = k(ref, 'numCache') || k(ref, 'strCache') || k(ref, 'multiLvlStrCache');
    if (!cache) return null;
    const n = +(k(cache, 'ptCount') ? k(cache, 'ptCount').getAttribute('val') : 0);
    // categorías de varios niveles: se usa el primer nivel
    const ptsFrom = k(cache, 'lvl') || cache;
    const out = new Array(n).fill(null);
    for (const pt of OF.kids(ptsFrom, 'pt', C)) { const v = k(pt, 'v'); out[+pt.getAttribute('idx')] = v ? v.textContent : null; }
    return out;
  };
  // lee un elemento cat/val/tx: { ref: 'Hoja1!$B$2:$B$5', cache: [...] }
  const source = (el) => {
    if (!el) return null;
    const r = k(el, 'numRef') || k(el, 'strRef') || k(el, 'multiLvlStrRef');
    if (r) { const f = k(r, 'f'); return { ref: f ? f.textContent : null, cache: cacheVals(r) }; }
    const lit = k(el, 'numLit') || k(el, 'strLit');
    if (lit) { const out = []; for (const pt of OF.kids(lit, 'pt', C)) { const v = k(pt, 'v'); out[+pt.getAttribute('idx')] = v ? v.textContent : null; } return { ref: null, cache: out }; }
    const v = k(el, 'v');
    if (v) return { ref: null, cache: [v.textContent] };
    return null;
  };

  // XML del gráfico -> especificación simple
  OF.parseChart = function (doc) {
    const plot = OF.all(doc, 'plotArea', C)[0];
    if (!plot) return null;
    const groups = [];
    for (const g of Array.from(plot.children)) {
      const n = g.localName;
      if (!/Chart$/.test(n)) continue;
      let type = n.replace(/3D/, '').replace('Chart', '');
      if (type === 'bar') type = (k(g, 'barDir') && k(g, 'barDir').getAttribute('val') === 'bar') ? 'bar' : 'col';
      if (type === 'doughnut' || type === 'ofPie') type = type === 'doughnut' ? 'doughnut' : 'pie';
      if (type === 'scatter' || type === 'bubble' || type === 'radar' || type === 'stock' || type === 'surface') type = 'line';
      const grouping = k(g, 'grouping') ? k(g, 'grouping').getAttribute('val') : 'clustered';
      const series = OF.kids(g, 'ser', C).map((s, i) => {
        const sp = k(s, 'spPr');
        const fill = sp && OF.all(sp, 'srgbClr')[0];
        return { name: source(k(s, 'tx')), cat: source(k(s, 'cat') || k(s, 'xVal')), val: source(k(s, 'val') || k(s, 'yVal')), color: fill ? '#' + fill.getAttribute('val') : null, idx: i };
      });
      groups.push({ type, grouping, series });
    }
    const t = OF.all(doc, 'title', C)[0];
    let title = null;
    if (t) {
      const txt = OF.all(t, 't').map((x) => x.textContent).join('');
      if (txt) title = txt;
      else { const s = source(k(t, 'tx')); if (s && s.cache) title = s.cache[0]; }
    }
    const auto = OF.all(doc, 'autoTitleDeleted', C)[0];
    if (!title && !(auto && auto.getAttribute('val') === '1') && groups[0] && groups[0].series.length === 1) title = '__serie__';
    const legend = !!OF.all(doc, 'legend', C)[0];
    return { groups, title, legend };
  };

  // resolve(ref) -> array de valores (Excel lo usa para leer las celdas en vivo)
  const values = (src, resolve) => {
    if (!src) return [];
    if (src.ref && resolve) { const v = resolve(src.ref); if (v) return v; }
    return src.cache || [];
  };

  const fmtNum = (v) => {
    if (v == null || isNaN(v)) return '';
    const a = Math.abs(v);
    if (a >= 1e6) return (v / 1e6).toLocaleString('es-CL', { maximumFractionDigits: 1 }) + ' M';
    if (a >= 1e4) return (v / 1e3).toLocaleString('es-CL', { maximumFractionDigits: 0 }) + ' mil';
    return v.toLocaleString('es-CL', { maximumFractionDigits: 2 });
  };
  const niceStep = (range) => {
    const raw = range / 5, p = Math.pow(10, Math.floor(Math.log10(raw || 1)));
    const m = raw / p;
    return (m <= 1 ? 1 : m <= 2 ? 2 : m <= 2.5 ? 2.5 : m <= 5 ? 5 : 10) * p;
  };
  const SVGNS = 'http://www.w3.org/2000/svg';
  const S = (tag, attrs, text) => { const e = document.createElementNS(SVGNS, tag); for (const a in attrs) e.setAttribute(a, attrs[a]); if (text != null) e.textContent = text; return e; };

  OF.drawChart = function (spec, w, h, resolve, palette) {
    const pal = palette || PALETTE;
    const svg = S('svg', { width: w, height: h, viewBox: `0 0 ${w} ${h}`, class: 'of-chart' });
    svg.append(S('rect', { x: 0, y: 0, width: w, height: h, fill: '#fff' }));
    if (!spec || !spec.groups.length) { svg.append(S('text', { x: w / 2, y: h / 2, 'text-anchor': 'middle', 'font-size': 13, fill: '#888' }, 'Gráfico')); return svg; }
    const fs = Math.max(9, Math.min(13, Math.round(Math.min(w, h) / 28)));
    const font = 'Calibri, Carlito, Arial, sans-serif';
    let top = 8;
    const g0 = spec.groups[0];
    const series = [];
    spec.groups.forEach((g) => g.series.forEach((s) => {
      const nm = values(s.name, resolve);
      series.push({ type: g.type, grouping: g.grouping, name: nm && nm[0] != null ? String(nm[0]) : 'Serie ' + (series.length + 1), cats: values(s.cat, resolve).map((x) => (x == null ? '' : String(x))), vals: values(s.val, resolve).map((x) => (x == null || x === '' || isNaN(+x) ? null : +x)), color: s.color || pal[series.length % pal.length] });
    }));
    let title = spec.title === '__serie__' ? (series[0] && series[0].name) : spec.title;
    if (title) { svg.append(S('text', { x: w / 2, y: top + fs + 3, 'text-anchor': 'middle', 'font-size': fs + 3, 'font-family': font, fill: '#404040' }, title)); top += fs + 12; }
    const isPie = g0.type === 'pie' || g0.type === 'doughnut';
    const cats = (series[0] && series[0].cats.length ? series[0].cats : (series[0] ? series[0].vals.map((_, i) => String(i + 1)) : []));
    // leyenda
    let bottom = h - 6;
    const legendItems = isPie ? cats.map((c, i) => ({ name: c, color: pal[i % pal.length] })) : series.map((s) => ({ name: s.name, color: s.color }));
    if (spec.legend || isPie || series.length > 1) {
      const lg = S('g', {});
      let x = 0; const items = [];
      for (const it of legendItems) { const width = 14 + it.name.length * fs * 0.55 + 14; items.push({ ...it, x, width }); x += width; }
      const startX = Math.max(6, (w - x) / 2);
      for (const it of items) {
        if (startX + it.x + it.width > w) break;
        lg.append(S('rect', { x: startX + it.x, y: bottom - fs, width: 10, height: 10, fill: it.color }));
        lg.append(S('text', { x: startX + it.x + 14, y: bottom - 1, 'font-size': fs, 'font-family': font, fill: '#595959' }, it.name));
      }
      svg.append(lg);
      bottom -= fs + 10;
    }
    if (isPie) {
      const vals = series[0].vals.map((v) => Math.max(0, v || 0));
      const total = vals.reduce((a, b) => a + b, 0) || 1;
      const cx = w / 2, cy = (top + bottom) / 2, r = Math.max(10, Math.min(w, bottom - top) / 2 - 6);
      let a0 = -Math.PI / 2;
      vals.forEach((v, i) => {
        const a1 = a0 + (v / total) * Math.PI * 2;
        const large = a1 - a0 > Math.PI ? 1 : 0;
        const p = v >= total ? `M${cx - r},${cy}a${r},${r} 0 1,0 ${2 * r},0a${r},${r} 0 1,0 ${-2 * r},0` : `M${cx},${cy}L${cx + r * Math.cos(a0)},${cy + r * Math.sin(a0)}A${r},${r} 0 ${large} 1 ${cx + r * Math.cos(a1)},${cy + r * Math.sin(a1)}Z`;
        svg.append(S('path', { d: p, fill: pal[i % pal.length], stroke: '#fff', 'stroke-width': 1.5 }));
        if (v / total > 0.05) { const am = (a0 + a1) / 2; svg.append(S('text', { x: cx + r * 0.65 * Math.cos(am), y: cy + r * 0.65 * Math.sin(am) + fs / 3, 'text-anchor': 'middle', 'font-size': fs, 'font-family': font, fill: '#fff' }, Math.round((v / total) * 100) + '%')); }
        a0 = a1;
      });
      if (g0.type === 'doughnut') svg.append(S('circle', { cx, cy, r: r * 0.5, fill: '#fff' }));
      return svg;
    }
    // ejes
    const horiz = g0.type === 'bar';
    const stacked = /stacked/i.test(g0.grouping), pct = /percent/i.test(g0.grouping);
    const n = Math.max(cats.length, ...series.map((s) => s.vals.length));
    let lo = 0, hi = 0;
    if (stacked || pct) for (let i = 0; i < n; i++) { let p = 0, q = 0; for (const s of series) { const v = s.vals[i] || 0; if (v >= 0) p += v; else q += v; } hi = Math.max(hi, p); lo = Math.min(lo, q); }
    else for (const s of series) for (const v of s.vals) if (v != null) { hi = Math.max(hi, v); lo = Math.min(lo, v); }
    if (pct) { hi = 1; lo = Math.min(lo, 0) ? -1 : 0; }
    if (hi === lo) hi = lo + 1;
    const step = niceStep(hi - lo);
    hi = Math.ceil(hi / step) * step; lo = Math.floor(lo / step) * step;
    const labelW = Math.max(...[hi, lo].map((v) => (pct ? Math.round(v * 100) + '%' : fmtNum(v)).length)) * fs * 0.55 + 8;
    const catW = horiz ? Math.min(w * 0.3, Math.max(...cats.map((c) => c.length)) * fs * 0.55 + 8) : 0;
    const L = 8 + (horiz ? catW : labelW), R = w - 12, T = top + 4, B = bottom - (horiz ? fs + 6 : fs + 8);
    const scale = (v) => (horiz ? L + ((v - lo) / (hi - lo)) * (R - L) : B - ((v - lo) / (hi - lo)) * (B - T));
    for (let v = lo; v <= hi + step / 2; v += step) {
      const p = scale(v);
      if (horiz) { svg.append(S('line', { x1: p, x2: p, y1: T, y2: B, stroke: '#d9d9d9' })); svg.append(S('text', { x: p, y: B + fs + 2, 'text-anchor': 'middle', 'font-size': fs, 'font-family': font, fill: '#595959' }, pct ? Math.round(v * 100) + '%' : fmtNum(v))); }
      else { svg.append(S('line', { x1: L, x2: R, y1: p, y2: p, stroke: '#d9d9d9' })); svg.append(S('text', { x: L - 4, y: p + fs / 3, 'text-anchor': 'end', 'font-size': fs, 'font-family': font, fill: '#595959' }, pct ? Math.round(v * 100) + '%' : fmtNum(v))); }
    }
    const band = (horiz ? B - T : R - L) / Math.max(1, n);
    const every = Math.ceil((cats.length * fs * 4) / ((horiz ? B - T : R - L) * 1.2)) || 1;
    cats.forEach((c, i) => {
      if (i % every) return;
      const mid = (horiz ? T : L) + band * (i + 0.5);
      const label = c.length > 14 ? c.slice(0, 13) + '…' : c;
      if (horiz) svg.append(S('text', { x: L - 4, y: mid + fs / 3, 'text-anchor': 'end', 'font-size': fs, 'font-family': font, fill: '#595959' }, label));
      else svg.append(S('text', { x: mid, y: B + fs + 4, 'text-anchor': 'middle', 'font-size': fs, 'font-family': font, fill: '#595959' }, label));
    });
    const bars = series.filter((s) => s.type === 'col' || s.type === 'bar' || s.type === 'area');
    const lines = series.filter((s) => s.type === 'line');
    const nb = stacked || pct ? 1 : bars.length;
    const totals = [];
    if (pct) for (let i = 0; i < n; i++) totals[i] = series.reduce((a, s) => a + Math.abs(s.vals[i] || 0), 0) || 1;
    const acc = []; const accN = [];
    bars.forEach((s, si) => {
      if (s.type === 'area') {
        let d = '';
        s.vals.forEach((v, i) => { const x = L + band * (i + 0.5), y = scale(v || 0); d += (i ? 'L' : 'M') + x + ',' + y; });
        d += `L${L + band * (s.vals.length - 0.5)},${scale(0)}L${L + band * 0.5},${scale(0)}Z`;
        svg.append(S('path', { d, fill: s.color, 'fill-opacity': 0.75 }));
        return;
      }
      s.vals.forEach((v0, i) => {
        if (v0 == null) return;
        let v = pct ? v0 / totals[i] : v0;
        let base = 0;
        if (stacked || pct) { if (v >= 0) { base = acc[i] || 0; acc[i] = base + v; } else { base = accN[i] || 0; accN[i] = base + v; } }
        const gap = band * 0.25, bw = (band - gap) / nb;
        const off = (horiz ? T : L) + band * i + gap / 2 + (stacked || pct ? 0 : bw * si);
        const a = scale(base), b = scale(base + v);
        if (horiz) svg.append(S('rect', { x: Math.min(a, b), y: off, width: Math.abs(b - a), height: bw, fill: s.color }));
        else svg.append(S('rect', { x: off, y: Math.min(a, b), width: bw, height: Math.abs(b - a), fill: s.color }));
      });
    });
    lines.forEach((s) => {
      let d = '';
      s.vals.forEach((v, i) => { if (v == null) return; const x = L + band * (i + 0.5), y = scale(v); d += (d ? 'L' : 'M') + x + ',' + y; });
      svg.append(S('path', { d, fill: 'none', stroke: s.color, 'stroke-width': 2.25, 'stroke-linejoin': 'round' }));
      s.vals.forEach((v, i) => { if (v != null) svg.append(S('circle', { cx: L + band * (i + 0.5), cy: scale(v), r: 3, fill: s.color })); });
    });
    svg.append(S('line', { x1: horiz ? scale(0) : L, x2: horiz ? scale(0) : R, y1: horiz ? T : scale(0), y2: horiz ? B : scale(0), stroke: '#bfbfbf' }));
    return svg;
  };

  /* ----------------------- crear un gráfico nuevo (XML) ----------------------- */
  const esc = (s) => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const strCache = (arr) => `<c:strCache><c:ptCount val="${arr.length}"/>${arr.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(v)}</c:v></c:pt>`).join('')}</c:strCache>`;
  const numCache = (arr) => `<c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${arr.length}"/>${arr.map((v, i) => (v == null || isNaN(v) ? '' : `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>`)).join('')}</c:numCache>`;
  // series: [{name, nameRef, catRef, cats, valRef, vals}]
  OF.chartXml = function (type, title, series) {
    const kind = type === 'line' ? 'lineChart' : type === 'pie' ? 'pieChart' : 'barChart';
    const ser = series.map((s, i) => `<c:ser><c:idx val="${i}"/><c:order val="${i}"/><c:tx>${s.nameRef ? `<c:strRef><c:f>${esc(s.nameRef)}</c:f>${strCache([s.name])}</c:strRef>` : `<c:v>${esc(s.name)}</c:v>`}</c:tx>` +
      (type === 'line' ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${PALETTE[i % PALETTE.length].slice(1)}"/></a:solidFill><a:round/></a:ln></c:spPr><c:marker><c:symbol val="circle"/><c:size val="5"/></c:marker>` : type === 'pie' ? '' : `<c:spPr><a:solidFill><a:srgbClr val="${PALETTE[i % PALETTE.length].slice(1)}"/></a:solidFill></c:spPr><c:invertIfNegative val="0"/>`) +
      (type === 'pie' ? series[0].cats.map((_, j) => `<c:dPt><c:idx val="${j}"/><c:bubble3D val="0"/><c:spPr><a:solidFill><a:srgbClr val="${PALETTE[j % PALETTE.length].slice(1)}"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FFFFFF"/></a:solidFill></a:ln></c:spPr></c:dPt>`).join('') : '') +
      `<c:cat><c:strRef><c:f>${esc(s.catRef)}</c:f>${strCache(s.cats)}</c:strRef></c:cat><c:val><c:numRef><c:f>${esc(s.valRef)}</c:f>${numCache(s.vals)}</c:numRef></c:val>` +
      (type === 'line' ? '<c:smooth val="0"/>' : '') + `</c:ser>`).join('');
    const head = kind === 'barChart' ? `<c:barDir val="${type === 'bar' ? 'bar' : 'col'}"/><c:grouping val="clustered"/><c:varyColors val="0"/>` : kind === 'lineChart' ? '<c:grouping val="standard"/><c:varyColors val="0"/>' : '<c:varyColors val="1"/>';
    const tail = kind === 'barChart' ? '<c:gapWidth val="150"/><c:axId val="500000001"/><c:axId val="500000002"/>' : kind === 'lineChart' ? '<c:marker val="1"/><c:axId val="500000001"/><c:axId val="500000002"/>' : '<c:firstSliceAng val="0"/>';
    const axes = kind === 'pieChart' ? '' : `<c:catAx><c:axId val="500000001"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'l' : 'b'}"/><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000002"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/><c:noMultiLvlLbl val="0"/></c:catAx><c:valAx><c:axId val="500000002"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${type === 'bar' ? 'b' : 'l'}"/><c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines><c:numFmt formatCode="General" sourceLinked="1"/><c:majorTickMark val="none"/><c:minorTickMark val="none"/><c:tickLblPos val="nextTo"/><c:crossAx val="500000001"/><c:crosses val="autoZero"/><c:crossBetween val="between"/></c:valAx>`;
    const titleXml = title ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"/></a:pPr><a:r><a:rPr lang="es-CL" sz="1400" b="0"/><a:t>${esc(title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>` : '<c:autoTitleDeleted val="1"/>';
    return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><c:chartSpace xmlns:c="${C}" xmlns:a="${OF.NS.a}" xmlns:r="${OF.NS.r}"><c:roundedCorners val="0"/><c:chart>${titleXml}<c:plotArea><c:layout/><c:${kind}>${head}${ser}${tail}</c:${kind}>${axes}</c:plotArea><c:legend><c:legendPos val="${type === 'pie' ? 'r' : 'b'}"/><c:overlay val="0"/></c:legend><c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart><c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1000"/></a:pPr><a:endParaRPr lang="es-CL"/></a:p></c:txPr></c:chartSpace>`;
  };
})();
