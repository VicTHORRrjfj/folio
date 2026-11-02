/* Folio: interfaz, archivos recientes (IndexedDB), abrir/guardar/compartir, modo sin conexión. */
(function () {
  const OF = window.OF;
  const $ = (id) => document.getElementById(id);
  const KINDS = {
    word: { ext: 'docx', cls: () => OF.WordDoc, label: 'DOC', name: 'Documento' },
    excel: { ext: 'xlsx', cls: () => OF.ExcelDoc, label: 'XLS', name: 'Planilla' },
    ppt: { ext: 'pptx', cls: () => OF.PptDoc, label: 'PPT', name: 'Presentación' },
  };
  const state = { doc: null, name: '', kind: null, handle: null, recentId: null };

  /* ----------------------------- recientes ----------------------------- */
  const db = new Promise((res, rej) => {
    const r = indexedDB.open('folio', 1);
    r.onupgradeneeded = () => r.result.createObjectStore('files', { keyPath: 'id' });
    r.onsuccess = () => res(r.result);
    r.onerror = () => rej(r.error);
  });
  async function tx(mode, fn) {
    const d = await db;
    return new Promise((res, rej) => {
      const t = d.transaction('files', mode);
      const out = fn(t.objectStore('files'));
      t.oncomplete = () => res(out && out.result !== undefined ? out.result : out);
      t.onerror = () => rej(t.error);
    });
  }
  async function listRecent() {
    try { const all = await tx('readonly', (s) => s.getAll()); return all.sort((a, b) => b.at - a.at); } catch (e) { return []; }
  }
  async function putRecent(rec) {
    try {
      await tx('readwrite', (s) => s.put(rec));
      const all = await listRecent();
      for (const old of all.slice(20)) await tx('readwrite', (s) => s.delete(old.id));
    } catch (e) { console.warn('No se pudo guardar en recientes', e); }
  }
  async function renderRecent() {
    const list = $('recent');
    const items = await listRecent();
    list.innerHTML = '';
    $('recentEmpty').classList.toggle('hidden', items.length > 0);
    list.classList.toggle('hidden', !items.length);
    for (const it of items) {
      const when = new Date(it.at).toLocaleString('es-CL', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
      const li = OF.el('li', null,
        OF.el('button', { class: 'open', onclick: () => openBuffer(it.data, it.name, null, it.id) },
          OF.el('span', { class: 'tag ' + it.kind }, KINDS[it.kind].label),
          OF.el('span', { style: { minWidth: 0 } }, OF.el('div', { class: 'nm' }, it.name), OF.el('div', { class: 'dt' }, (it.saved ? 'Guardado ' : 'Abierto ') + when))),
        OF.el('button', { class: 'icon-btn', title: 'Quitar de recientes', 'aria-label': 'Quitar ' + it.name + ' de recientes', onclick: async () => { await tx('readwrite', (s) => s.delete(it.id)); renderRecent(); } },
          svgIcon('M6 6l12 12M18 6L6 18')));
      list.append(li);
    }
  }
  function svgIcon(d) {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('fill', 'none'); s.setAttribute('stroke', 'currentColor'); s.setAttribute('stroke-width', '2'); s.setAttribute('stroke-linecap', 'round');
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); s.append(p);
    return s;
  }

  /* ------------------------------- abrir ------------------------------- */
  function kindOf(name) {
    const ext = name.split('.').pop().toLowerCase();
    if (ext === 'docx' || ext === 'docm') return 'word';
    if (ext === 'xlsx' || ext === 'xlsm') return 'excel';
    if (ext === 'pptx' || ext === 'pptm') return 'ppt';
    if (['doc', 'xls', 'ppt'].includes(ext)) throw new Error(`Los archivos .${ext} usan el formato antiguo de Office. Ábrelo en Office y guárdalo como .${ext}x para editarlo aquí.`);
    throw new Error('Folio abre archivos .docx, .xlsx y .pptx.');
  }

  async function openBuffer(buffer, name, handle, recentId) {
    if (!(await confirmLeave())) return;
    let kind;
    try { kind = kindOf(name); } catch (e) { return toast(e.message, 6000); }
    if (kind === 'excel' && !window.HyperFormula) await waitFor(() => window.HyperFormula, 3000);
    showEditor(kind, name);
    const view = $('view');
    view.innerHTML = '<div class="loading">Abriendo…</div>';
    const doc = new (KINDS[kind].cls())();
    try {
      await doc.open(buffer.slice(0), view);
    } catch (e) {
      console.error(e);
      showHome();
      return toast('No se pudo abrir el archivo: ' + (e.message || e), 6000);
    }
    doc.onchange = updateState;
    Object.assign(state, { doc, name, kind, handle: handle || null, recentId: recentId || cryptoId() });
    updateState();
    putRecent({ id: state.recentId, name, kind, data: buffer, at: Date.now(), saved: false });
    if (kind === 'excel') setTimeout(() => doc.gridBox && doc.gridBox.focus({ preventScroll: true }), 50);
  }
  function waitFor(cond, ms) { return new Promise((res) => { const t0 = Date.now(); (function loop() { if (cond() || Date.now() - t0 > ms) res(); else setTimeout(loop, 50); })(); }); }
  function cryptoId() { return (crypto.randomUUID && crypto.randomUUID()) || String(Date.now() + Math.random()); }

  async function pickFile() {
    if (window.showOpenFilePicker && !/Android|iPhone|iPad/i.test(navigator.userAgent)) {
      try {
        const [h] = await showOpenFilePicker({ types: [{ description: 'Documentos de Office', accept: { 'application/vnd.openxmlformats-officedocument.wordprocessingml.document': ['.docx'], 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx', '.xlsm'], 'application/vnd.openxmlformats-officedocument.presentationml.presentation': ['.pptx'] } }] });
        const f = await h.getFile();
        return openBuffer(await f.arrayBuffer(), f.name, h);
      } catch (e) { if (e.name !== 'AbortError') toast('No se pudo abrir: ' + e.message); return; }
    }
    $('fileInput').click();
  }
  $('fileInput').addEventListener('change', async (e) => {
    const f = e.target.files[0];
    e.target.value = '';
    if (f) openBuffer(await f.arrayBuffer(), f.name);
  });

  async function newFile(kind) {
    try {
      const res = await fetch(`assets/nuevo.${KINDS[kind].ext}`);
      const buf = await res.arrayBuffer();
      const stamp = new Date().toLocaleDateString('es-CL').replace(/\//g, '-');
      openBuffer(buf, `${KINDS[kind].name} ${stamp}.${KINDS[kind].ext}`);
    } catch (e) { toast('No se pudo crear el archivo.'); }
  }

  /* ------------------------------ guardar ------------------------------ */
  async function save(asCopy) {
    const { doc } = state;
    if (!doc) return;
    let blob;
    try { blob = await doc.save(); } catch (e) { console.error(e); return toast('No se pudo guardar: ' + e.message, 6000); }
    let name = state.name;
    if (asCopy) {
      const base = name.replace(/\.[^.]+$/, ''), ext = name.split('.').pop();
      name = `${base} (copia).${ext}`;
    }
    let where = 'Descargas';
    if (!asCopy && state.handle && state.handle.createWritable) {
      try {
        if ((await state.handle.queryPermission({ mode: 'readwrite' })) !== 'granted') await state.handle.requestPermission({ mode: 'readwrite' });
        const w = await state.handle.createWritable();
        await w.write(blob); await w.close();
        where = null;
      } catch (e) { console.warn(e); downloadBlob(blob, name); }
    } else if (window.showSaveFilePicker && !/Android|iPhone|iPad/i.test(navigator.userAgent)) {
      try {
        const h = await showSaveFilePicker({ suggestedName: name });
        const w = await h.createWritable(); await w.write(blob); await w.close();
        if (!asCopy) state.handle = h;
        where = null;
      } catch (e) { if (e.name === 'AbortError') return; downloadBlob(blob, name); }
    } else downloadBlob(blob, name);
    await putRecent({ id: state.recentId, name: state.name, kind: state.kind, data: await blob.arrayBuffer(), at: Date.now(), saved: true });
    updateState(true);
    toast(where ? `Guardado en ${where} como «${name}»` : `Guardado «${name}»`);
  }
  function downloadBlob(blob, name) {
    const a = OF.el('a', { href: URL.createObjectURL(blob), download: name });
    document.body.append(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  async function share() {
    const { doc } = state;
    if (!doc) return;
    const blob = await doc.save();
    updateState(true);
    const file = new File([blob], state.name, { type: blob.type });
    if (navigator.canShare && navigator.canShare({ files: [file] })) {
      try { await navigator.share({ files: [file], title: state.name }); } catch (e) { if (e.name !== 'AbortError') toast('No se pudo compartir.'); }
    } else { downloadBlob(blob, state.name); toast('Este navegador no permite compartir archivos; se descargó una copia.'); }
  }

  /* ------------------------------ barra de herramientas ------------------------------ */
  const ICON = {
    left: 'M4 6h16M4 10h10M4 14h16M4 18h10', center: 'M4 6h16M7 10h10M4 14h16M7 18h10', right: 'M4 6h16M10 10h10M4 14h16M10 18h10', justify: 'M4 6h16M4 10h16M4 14h16M4 18h16',
    image: 'M4 5h16v14H4zM4 16l5-5 4 4 3-3 4 4M15 9.5a1 1 0 1 0 0-.01', rows: 'M3 6h18M3 12h18M3 18h18M12 3v18', cols: 'M6 3v18M12 3v18M18 3v18M3 12h18',
    slide: 'M3 5h18v12H3zM8 21h8M12 17v4', table: 'M3 5h18v14H3zM3 10h18M3 15h18M9 5v14M15 5v14', range: 'M4 4h7v7H4zM13 13h7v7h-7zM11 7.5h5.5V13', wrap: 'M4 6h16M4 12h13a3 3 0 0 1 0 6h-4m0 0l2-2m-2 2l2 2M4 18h5', num: 'M5 9h14M5 15h14M10 4l-2 16M16 4l-2 16',
  };
  const COLORS = ['#000000', '#595959', '#A6A6A6', '#FFFFFF', '#C00000', '#FF0000', '#ED7D31', '#FFC000', '#FFFF00', '#70AD47', '#00B050', '#00B0F0', '#4472C4', '#002060', '#7030A0', '#1F4E79'];
  const FILLS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#EDEDED', '#FFD966', '#A9D08E', '#9BC2E6', '#F4B084', '#BFBFBF', '#FFC000', '#70AD47', '#4472C4', '#ED7D31', '#C00000', '#1F4E79'];
  const SIZES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 48, 60, 72];
  const HIGHLIGHTS = [['yellow', '#FFFF00'], ['green', '#00FF00'], ['cyan', '#00FFFF'], ['magenta', '#FF00FF'], ['red', '#FF0000'], ['lightGray', '#C0C0C0']];

  function icon(d) {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('fill', 'none'); s.setAttribute('stroke', 'currentColor'); s.setAttribute('stroke-width', '1.8'); s.setAttribute('stroke-linecap', 'round'); s.setAttribute('stroke-linejoin', 'round');
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); s.append(p);
    return s;
  }
  async function run(fn) {
    try { await fn(state.doc); } catch (e) { console.error(e); toast('No se pudo completar: ' + (e.message || e), 5000); }
  }
  let pop = null;
  function closePop() { if (pop) { pop.remove(); pop = null; } }
  function openPop(btn, build) {
    const was = pop && pop.dataset.for === btn.id;
    closePop();
    if (was) return;
    pop = OF.el('div', { class: 'pop', 'data-for': btn.id, role: 'menu' });
    build(pop);
    document.body.append(pop);
    const r = btn.getBoundingClientRect();
    const w = pop.offsetWidth;
    pop.style.left = Math.max(8, Math.min(innerWidth - w - 8, r.left)) + 'px';
    pop.style.top = r.bottom + 6 + 'px';
  }
  function swatches(list, onPick, noneLabel) {
    const wrap = OF.el('div', { class: 'sw-grid' });
    for (const c of list) wrap.append(OF.el('button', { class: 'sw', title: c, 'aria-label': c, style: { background: c }, onclick: () => { closePop(); onPick(c); } }));
    return [OF.el('button', { class: 'pop-item', onclick: () => { closePop(); onPick(null); } }, noneLabel), wrap];
  }
  function menuItems(items) {
    return items.map(([label, fn, danger]) => OF.el('button', { class: 'pop-item' + (danger ? ' danger' : ''), onclick: () => { closePop(); run(fn); } }, label));
  }
  let rbId = 0;
  function rb(label, title, onclick, opts = {}) {
    const b = OF.el('button', { class: 'rb' + (opts.cls ? ' ' + opts.cls : ''), id: 'rb' + rbId++, title, 'aria-label': title, type: 'button' });
    if (typeof label === 'string') b.append(OF.el('span', { class: 'rb-t' }, label)); else b.append(label);
    if (opts.text) b.append(OF.el('span', { class: 'rb-l' }, opts.text));
    if (opts.caret) b.append(OF.el('span', { class: 'rb-c', 'aria-hidden': 'true' }, '▾'));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', (e) => { e.stopPropagation(); onclick(b); });
    return b;
  }
  const sep = () => OF.el('span', { class: 'rb-sep' });

  function buildRibbon(kind) {
    const bar = $('ribbon');
    bar.innerHTML = '';
    const items = [];
    items.push(rb('B', 'Negrita', () => run((d) => d.toggle('b')), { cls: 'b' }));
    items.push(rb('I', 'Cursiva', () => run((d) => d.toggle('i')), { cls: 'i' }));
    items.push(rb('U', 'Subrayado', () => run((d) => d.toggle('u')), { cls: 'u' }));
    const colorBtn = rb(OF.el('span', { class: 'rb-a' }, 'A'), 'Color de texto', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Color de texto'), ...swatches(COLORS, (c) => run((d) => d.setColor(c)), 'Automático'))), { caret: true });
    items.push(colorBtn);
    if (kind === 'excel') items.push(rb(OF.el('span', { class: 'rb-fill' }), 'Color de relleno', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Relleno de celda'), ...swatches(FILLS, (c) => run((d) => d.setFill(c)), 'Sin relleno'))), { caret: true }));
    if (kind === 'word') items.push(rb(OF.el('span', { class: 'rb-hl' }, 'ab'), 'Resaltar', (b) => openPop(b, (p) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Resaltado'), OF.el('button', { class: 'pop-item', onclick: () => { closePop(); run((d) => d.setHighlight(null)); } }, 'Sin resaltado'));
      const g = OF.el('div', { class: 'sw-grid' });
      for (const [n, c] of HIGHLIGHTS) g.append(OF.el('button', { class: 'sw', title: n, 'aria-label': n, style: { background: c }, onclick: () => { closePop(); run((d) => d.setHighlight(n, c)); } }));
      p.append(g);
    }), { caret: true }));
    items.push(rb('Aa', 'Tamaño de letra', (b) => openPop(b, (p) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Tamaño'));
      const g = OF.el('div', { class: 'size-grid' });
      for (const s of SIZES) g.append(OF.el('button', { class: 'pop-item', onclick: () => { closePop(); run((d) => d.setSize(s)); } }, String(s)));
      p.append(g);
    }), { caret: true }));
    items.push(sep());
    for (const [a, t] of [['left', 'Alinear a la izquierda'], ['center', 'Centrar'], ['right', 'Alinear a la derecha']]) items.push(rb(icon(ICON[a]), t, () => run((d) => d.align(a))));
    if (kind !== 'excel') items.push(rb(icon(ICON.justify), 'Justificar', () => run((d) => d.align('justify'))));
    if (kind === 'excel') {
      items.push(rb(icon(ICON.wrap), 'Ajustar texto en la celda', () => run((d) => d.wrapText())));
      items.push(rb(icon(ICON.num), 'Formato de número', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Formato de número'), ...menuItems([
        ['General', (d) => d.numFormat('General')], ['Número  1.234,56', (d) => d.numFormat('#,##0.00')], ['Moneda  $1.234', (d) => d.numFormat('"$"#,##0')],
        ['Porcentaje  12%', (d) => d.numFormat('0%')], ['Fecha  15-03-2026', (d) => d.numFormat('dd-mm-yyyy')], ['Texto', (d) => d.numFormat('@')],
      ]))), { caret: true }));
    }
    items.push(sep());
    items.push(rb(icon(ICON.image), 'Insertar imagen', async () => { const f = await OF.pickImage(); if (f) run((d) => d.insertImage(f)); }, { text: 'Imagen' }));
    if (kind === 'excel') {
      items.push(rb(icon(ICON.rows), 'Filas', (b) => openPop(b, (p) => p.append(...menuItems([['Insertar fila arriba', (d) => d.rowOp('above')], ['Insertar fila abajo', (d) => d.rowOp('below')], ['Eliminar fila', (d) => d.rowOp('del'), true]]))), { text: 'Filas', caret: true }));
      items.push(rb(icon(ICON.cols), 'Columnas', (b) => openPop(b, (p) => p.append(...menuItems([['Insertar columna a la izquierda', (d) => d.colOp('left')], ['Insertar columna a la derecha', (d) => d.colOp('right')], ['Eliminar columna', (d) => d.colOp('del'), true]]))), { text: 'Columnas', caret: true }));
      const rg = rb(icon(ICON.range), 'Seleccionar rango: toca la primera y la última celda', (b) => { const d = state.doc; d.rangeMode = !d.rangeMode; b.classList.toggle('on', d.rangeMode); toast(d.rangeMode ? 'Toca la última celda del rango' : 'Selección de una celda'); }, { text: 'Rango' });
      items.push(rg);
    } else {
      items.push(rb(icon(ICON.table), 'Filas de tabla', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Toca antes una celda de la tabla'), ...menuItems([['Insertar fila arriba', (d) => d.tableRow('above')], ['Insertar fila abajo', (d) => d.tableRow('below')], ['Eliminar fila', (d) => d.tableRow('del'), true]]))), { text: 'Tabla', caret: true }));
    }
    if (kind === 'ppt') {
      items.push(rb(icon(ICON.slide), 'Diapositivas', (b) => openPop(b, (p) => p.append(...menuItems([
        ['Nueva diapositiva', (d) => d.slideOp('new')], ['Duplicar diapositiva', (d) => d.slideOp('dup')],
        ['Mover hacia arriba', (d) => d.slideOp('up')], ['Mover hacia abajo', (d) => d.slideOp('down')], ['Eliminar diapositiva', (d) => d.slideOp('delete'), true],
      ]))), { text: 'Diapositiva', caret: true }));
    }
    bar.append(...items);
  }
  document.addEventListener('click', (e) => { if (pop && !e.target.closest('.pop')) closePop(); });
  window.addEventListener('resize', closePop);

  /* ------------------------------- vistas ------------------------------- */
  function showEditor(kind, name) {
    $('home').classList.add('hidden');
    $('editor').classList.remove('hidden');
    $('editor').dataset.kind = kind;
    $('fname').textContent = name;
    $('presentBtn').classList.toggle('hidden', kind !== 'ppt');
    $('localeBtn').classList.toggle('hidden', kind !== 'excel');
    const hints = {
      word: 'Escribe directamente sobre la página. Para dar formato, selecciona el texto y usa la barra. Encabezados y pies se conservan tal cual.',
      excel: 'Toca una celda para seleccionarla y vuelve a tocar para editar. Para un rango usa «Rango» o toca los encabezados de fila/columna. Las fórmulas van en inglés (SUM, IF).',
      ppt: 'Toca un cuadro de texto para editarlo. Para dar formato, selecciona el texto y usa la barra. Las diapositivas se agregan desde «Diapositiva».',
    };
    $('menuHint').textContent = hints[kind];
    buildRibbon(kind);
    history.pushState({ editor: true }, '');
  }
  async function showHome(fromPop) {
    if (state.doc && state.doc.dirty && !(await confirmLeave())) { if (fromPop) history.pushState({ editor: true }, ''); return; }
    state.doc = null;
    $('editor').classList.add('hidden');
    $('home').classList.remove('hidden');
    $('view').innerHTML = '';
    closeMenu();
    renderRecent();
  }
  async function confirmLeave() {
    if (!state.doc || !state.doc.dirty) return true;
    return confirm('Tienes cambios sin guardar. ¿Salir sin guardar?');
  }
  function updateState(justSaved) {
    const el = $('fstate');
    const dirty = state.doc && state.doc.dirty;
    el.textContent = dirty ? 'Cambios sin guardar' : justSaved === true ? 'Guardado' : 'Sin cambios';
    el.classList.toggle('dirty', !!dirty);
    $('fname').textContent = state.name;
  }
  let toastT;
  function toast(msg, ms) {
    let t = document.querySelector('.toast');
    if (!t) { t = OF.el('div', { class: 'toast', role: 'status' }); document.body.append(t); }
    t.textContent = msg;
    clearTimeout(toastT);
    toastT = setTimeout(() => t.remove(), ms || 3200);
  }
  OF.toast = toast;

  function closeMenu() { $('menu').classList.add('hidden'); $('moreBtn').setAttribute('aria-expanded', 'false'); }

  /* ------------------------------- eventos ------------------------------- */
  $('openBtn').addEventListener('click', pickFile);
  document.querySelectorAll('[data-new]').forEach((b) => b.addEventListener('click', () => newFile(b.dataset.new)));
  $('backBtn').addEventListener('click', () => history.back());
  window.addEventListener('popstate', () => { if (!$('editor').classList.contains('hidden')) showHome(true); });
  $('save').addEventListener('click', () => save(false));
  $('saveCopy').addEventListener('click', () => { closeMenu(); save(true); });
  $('shareBtn').addEventListener('click', () => { closeMenu(); share(); });
  $('zoomIn').addEventListener('click', () => state.doc && state.doc.zoom(1.15));
  $('zoomOut').addEventListener('click', () => state.doc && state.doc.zoom(1 / 1.15));
  $('presentBtn').addEventListener('click', () => state.doc && state.doc.present && state.doc.present(0));
  $('moreBtn').addEventListener('click', (e) => {
    e.stopPropagation();
    const m = $('menu');
    m.classList.toggle('hidden');
    $('moreBtn').setAttribute('aria-expanded', String(!m.classList.contains('hidden')));
  });
  document.addEventListener('click', (e) => { if (!e.target.closest('#menu')) closeMenu(); });
  $('localeBtn').addEventListener('click', () => {
    const d = state.doc;
    if (!d || d.kind !== 'excel') return;
    d.locale = d.locale === 'cl' ? 'us' : 'cl';
    $('localeLbl').textContent = d.locale === 'cl' ? 'Números: 1.234,56' : 'Números: 1,234.56';
    try { localStorage.setItem('folio-locale', d.locale); } catch (e) {}
    d.showSheet(d.cur);
  });
  document.addEventListener('keydown', (e) => {
    if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 's' && state.doc) { e.preventDefault(); save(false); }
  });
  let rz;
  window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => state.doc && state.doc.fit && state.doc.fit(), 150); });
  window.addEventListener('beforeunload', (e) => { if (state.doc && state.doc.dirty) { e.preventDefault(); e.returnValue = ''; } });

  const drop = $('drop');
  ['dragenter', 'dragover'].forEach((ev) => document.addEventListener(ev, (e) => { if ($('home').classList.contains('hidden')) return; e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach((ev) => document.addEventListener(ev, () => drop.classList.remove('over')));
  document.addEventListener('drop', async (e) => {
    if ($('home').classList.contains('hidden')) return;
    e.preventDefault();
    const f = e.dataTransfer.files[0];
    if (f) openBuffer(await f.arrayBuffer(), f.name);
  });

  // PC: abrir archivos con doble clic cuando la app está instalada
  if ('launchQueue' in window) {
    launchQueue.setConsumer(async (params) => {
      const h = params.files && params.files[0];
      if (!h) return;
      const f = await h.getFile();
      openBuffer(await f.arrayBuffer(), f.name, h);
    });
  }

  // Excel: recordar el formato regional elegido
  const origOpen = OF.ExcelDoc.prototype.open;
  OF.ExcelDoc.prototype.open = function (...a) {
    try { this.locale = localStorage.getItem('folio-locale') || 'cl'; } catch (e) {}
    $('localeLbl').textContent = this.locale === 'cl' ? 'Números: 1.234,56' : 'Números: 1,234.56';
    return origOpen.apply(this, a);
  };

  // Archivo recibido con "Compartir → Folio"
  async function openShared() {
    if (!/[?&]compartido=1/.test(location.search) || !('caches' in window)) return;
    history.replaceState(null, '', location.pathname);
    try {
      const cache = await caches.open('folio-compartido');
      const res = await cache.match('compartido');
      if (!res) return toast('No llegó ningún archivo. Comparte un .docx, .xlsx o .pptx.');
      const name = decodeURIComponent(res.headers.get('x-nombre') || 'archivo');
      const buf = await res.arrayBuffer();
      await cache.delete('compartido');
      openBuffer(buf, name);
    } catch (e) { toast('No se pudo abrir el archivo compartido.'); }
  }
  openShared();

  // PC: pegar un archivo copiado (Ctrl+V) en la pantalla de inicio
  document.addEventListener('paste', async (e) => {
    if ($('home').classList.contains('hidden')) return;
    const f = e.clipboardData && e.clipboardData.files && e.clipboardData.files[0];
    if (!f) return;
    e.preventDefault();
    openBuffer(await f.arrayBuffer(), f.name);
  });

  if ('serviceWorker' in navigator && location.protocol !== 'file:') {
    navigator.serviceWorker.register('sw.js').catch((e) => console.warn('Sin modo offline:', e));
  }
  renderRecent();
})();
