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
    const q = ($('recentSearch').value || '').trim().toLowerCase();
    const all = await listRecent();
    const items = q ? all.filter((it) => it.name.toLowerCase().includes(q)) : all;
    list.innerHTML = '';
    $('recentEmpty').classList.toggle('hidden', items.length > 0);
    $('recentEmpty').textContent = q && all.length ? 'Ningún archivo reciente coincide con la búsqueda.' : 'Los archivos que abras quedarán aquí para volver a ellos sin conexión.';
    list.classList.toggle('hidden', !items.length);
    for (const it of items) {
      const when = new Date(it.at).toLocaleString('es-CL', { day: 'numeric', month: 'short', hour: '2-digit', minute: '2-digit' });
      const li = OF.el('li', null,
        OF.el('button', { class: 'open', onclick: () => openBuffer(it.data, it.name, null, it.id, it.draft) },
          OF.el('span', { class: 'tag ' + it.kind }, KINDS[it.kind].label),
          OF.el('span', { style: { minWidth: 0 } }, OF.el('div', { class: 'nm' }, it.name), OF.el('div', { class: 'dt' }, it.draft ? OF.el('span', { class: 'draft' }, 'Cambios sin guardar') : null, (it.draft ? ' · ' : '') + (it.saved ? 'Guardado ' : 'Abierto ') + when))),
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
    if (['doc', 'ppt'].includes(ext)) throw new Error(`Los archivos .${ext} usan el formato antiguo de Office (de 1997–2003). Ábrelo en Office, Google Drive o LibreOffice y guárdalo como .${ext}x para editarlo aquí.`);
    throw new Error('Folio abre archivos .docx, .xlsx y .pptx.');
  }

  // hojas de cálculo antiguas u otras (.xls, .ods, .csv): se convierten a .xlsx al abrir
  const CONVERT = ['xls', 'xlt', 'ods', 'csv', 'tsv'];
  // SheetJS corre en un worker: si un archivo manipulado intentara algo, queda encerrado ahí y se corta a los 30 s
  function convertSheet(buffer, name) {
    const ext = name.split('.').pop().toLowerCase();
    let data = new Uint8Array(buffer);
    const opts = { type: 'array', cellFormula: true, cellNF: true, cellDates: false };
    if (ext === 'csv' || ext === 'tsv') {
      opts.type = 'string';
      data = new TextDecoder(/\ufffd/.test(new TextDecoder('utf-8').decode(data)) ? 'windows-1252' : 'utf-8').decode(data);
      const first = data.split('\n')[0];
      if (ext === 'csv' && /;/.test(first) && !/,/.test(first)) opts.FS = ';';
      if (ext === 'tsv') opts.FS = '\t';
    }
    return new Promise((ok, bad) => {
      const w = new Worker('js/convert-worker.js');
      const t = setTimeout(() => { w.terminate(); bad(new Error('el archivo tardó demasiado en convertirse')); }, 30000);
      w.onmessage = (e) => { clearTimeout(t); w.terminate(); e.data.ok ? ok(e.data.buf) : bad(new Error(e.data.msg)); };
      w.onerror = (e) => { clearTimeout(t); w.terminate(); bad(new Error(e.message || 'error del convertidor')); };
      w.postMessage({ data, opts });
    });
  }
  async function openBuffer(buffer, name, handle, recentId, draft) {
    const ext0 = name.split('.').pop().toLowerCase();
    if (CONVERT.includes(ext0) && !draft) {
      try { buffer = await convertSheet(buffer, name); } catch (e) { console.error(e); return toast('No se pudo convertir el archivo: ' + (e.message || e), 6000); }
      name = name.replace(/\.[^.]+$/, '') + '.xlsx';
      handle = null;
      setTimeout(() => toast(`Convertido a ${name}: se conservan valores, fórmulas y formatos de número; colores y bordes del archivo antiguo pueden perderse. Guardar crea el .xlsx.`, 7000), 600);
    }
    if (!(await confirmLeave())) return;
    let kind;
    try { kind = kindOf(name); } catch (e) { return toast(e.message, 6000); }
    if (kind === 'excel' && !window.HyperFormula) await waitFor(() => window.HyperFormula, 3000);
    showEditor(kind, name);
    const view = $('view');
    view.className = '';
    view.innerHTML = '<div class="loading">Abriendo…</div>';
    const doc = new (KINDS[kind].cls())();
    resetHistory();
    try {
      await doc.open((draft ? draft.bytes : buffer).slice(0), view);
      if (draft) {
        if (draft.html && doc.searchRoots) { const roots = doc.searchRoots(); if (roots.length === draft.html.length) roots.forEach((r, i) => { const kids = OF.parseKids(draft.html[i]); if (kids) r.replaceChildren(...kids); }); doc.afterRestore && doc.afterRestore(); }
        doc.dirty = true;
        if (doc.slides) doc.slides.forEach((x, i) => doc.updateThumb && doc.updateThumb(i, true));
      }
    } catch (e) {
      console.error(e);
      showHome();
      return toast('No se pudo abrir el archivo: ' + (e.message || e), 6000);
    }
    doc.onchange = () => { updateState(); scheduleAutosave(); };
    doc.onselect = refreshUI;
    Object.assign(state, { doc, name, kind, handle: handle || null, recentId: recentId || cryptoId() });
    updateState();
    if (draft) toast('Se recuperaron tus cambios sin guardar.', 4000);
    else putRecent({ id: state.recentId, name, kind, data: buffer, at: Date.now(), saved: false });
    refreshUI();
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

  const TEMPLATES = [
    { file: 'factura.xlsx', name: 'Factura', kind: 'excel', sub: 'IVA 19% y totales automáticos' },
    { file: 'cotizacion.docx', name: 'Cotización', kind: 'word', sub: 'Para enviar a clientes' },
    { file: 'presupuesto.xlsx', name: 'Presupuesto mensual', kind: 'excel', sub: 'Planificado vs. real' },
    { file: 'inventario.xlsx', name: 'Control de inventario', kind: 'excel', sub: 'Avisa cuándo reponer' },
    { file: 'informe.docx', name: 'Informe', kind: 'word', sub: 'Con títulos, tabla y listas' },
    { file: 'presentacion.pptx', name: 'Presentación de proyecto', kind: 'ppt', sub: '5 diapositivas con gráfico' },
  ];
  function renderTemplates() {
    const box = $('templates');
    for (const t of TEMPLATES) {
      const b = OF.el('button', { class: 'tpl', type: 'button' },
        OF.el('span', { class: 'tag ' + t.kind }, KINDS[t.kind].label),
        OF.el('span', { class: 'tpl-t' }, OF.el('b', null, t.name), OF.el('small', null, t.sub)));
      b.addEventListener('click', async () => {
        try {
          const buf = await (await fetch('assets/plantillas/' + t.file)).arrayBuffer();
          const stamp = new Date().toLocaleDateString('es-CL').replace(/\//g, '-');
          openBuffer(buf, `${t.name} ${stamp}.${KINDS[t.kind].ext}`);
        } catch (e) { toast('No se pudo abrir la plantilla.'); }
      });
      box.append(b);
    }
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
    clearTimeout(autoT);
    await putRecent({ id: state.recentId, name: state.name, kind: state.kind, data: await blob.arrayBuffer(), at: Date.now(), saved: true, draft: null });
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
    bullets: 'M9 6h11M9 12h11M9 18h11M4.5 6h.01M4.5 12h.01M4.5 18h.01', numbers: 'M10 6h10M10 12h10M10 18h10M4 5l1.5-1V9M3.5 15.5a1.4 1.4 0 0 1 2.6.6c0 1-2.6 2-2.6 2.9h2.8',
    indentL: 'M3 6h18M11 10h10M11 14h10M3 18h18M7 9l-3 3 3 3', indentR: 'M3 6h18M11 10h10M11 14h10M3 18h18M4 9l3 3-3 3', spacing: 'M11 6h10M11 12h10M11 18h10M5 4v16M3 6l2-2 2 2M3 18l2 2 2-2',
    pagebreak: 'M5 3v5h14V3M5 21v-5h14v5M3 12h2M8 12h3M14 12h3M19 12h2', header: 'M5 3h14v18H5zM8 7h8', footer: 'M5 3h14v18H5zM8 17h8', pagenum: 'M5 3h14v18H5zM11 15l2-1.5V19', footnote: 'M4 5h16M4 9h16M4 13h10M4 18h6M15 16l1.5-1V20', cond: 'M3 4h18v16H3zM3 10h18M3 15h18M10 4v16', rotate: 'M20 11a8 8 0 1 0-2.3 5.7M20 4v7h-7', textbox: 'M4 5h16v14H4zM9 9h6M12 9v7', shapes: 'M4 13h8v8H4zM17 3l4 7h-8zM17 21a4 4 0 1 0 0-8 4 4 0 0 0 0 8', chart: 'M4 20V10M10 20V4M16 20v-7M2 20h20',
    sortAZ: 'M4 6h7M4 12h5M4 18h3M17 4v16M14 17l3 3 3-3', sortZA: 'M4 6h3M4 12h5M4 18h7M17 4v16M14 17l3 3 3-3', filter: 'M3 5h18l-7 8v6l-4 2v-8z', fill: 'M12 4v12M7 11l5 5 5-5M5 20h14',
    border: 'M4 4h16v16H4zM4 12h16M12 4v16', merge: 'M3 5h18v14H3zM8 12h8M11 9l-3 3 3 3M13 9l3 3-3 3', freeze: 'M3 4h18v16H3zM3 9h18M9 4v16', width: 'M3 12h18M7 8l-4 4 4 4M17 8l4 4-4 4', height: 'M12 3v18M8 7l4-4 4 4M8 17l4 4 4-4',
    sum: 'M18 5H6l6 7-6 7h12', bg: 'M3 5h18v14H3zM3 15l5-4 4 3 3-2 6 4', hide: 'M3 3l18 18M10.6 10.6a2 2 0 0 0 2.8 2.8M9.9 5.1A9 9 0 0 1 21 12a12 12 0 0 1-2.4 3.4M6.6 6.6A12 12 0 0 0 3 12a9 9 0 0 0 11.3 6.3', sheet: 'M4 4h16v16H4zM4 9h16M9 9v11', dup: 'M8 8h12v12H8zM4 16V4h12',
    up: 'M12 19V5M6 11l6-6 6 6', down: 'M12 5v14M6 13l6 6 6-6', trash: 'M4 7h16M10 11v6M14 11v6M6 7l1 13h10l1-13M9 7V4h6v3', locale: 'M4 7h16M4 12h10M4 17h7',
  };
  const COLORS = ['#000000', '#595959', '#A6A6A6', '#FFFFFF', '#C00000', '#FF0000', '#ED7D31', '#FFC000', '#FFFF00', '#70AD47', '#00B050', '#00B0F0', '#4472C4', '#002060', '#7030A0', '#1F4E79'];
  const FILLS = ['#FFF2CC', '#E2EFDA', '#DDEBF7', '#FCE4D6', '#EDEDED', '#FFD966', '#A9D08E', '#9BC2E6', '#F4B084', '#BFBFBF', '#FFC000', '#70AD47', '#4472C4', '#ED7D31', '#C00000', '#1F4E79'];
  const BGS = ['#FFFFFF', '#F2F2F2', '#D9D9D9', '#262626', '#000000', '#DDEBF7', '#1F4E79', '#E2EFDA', '#375623', '#FFF2CC', '#FCE4D6', '#C00000', '#7030A0', '#002060', '#F8F4EC', '#0F2A3F'];
  const SIZES = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 24, 28, 32, 36, 40, 44, 48, 54, 60, 72];
  const HIGHLIGHTS = [['yellow', '#FFFF00'], ['green', '#00FF00'], ['cyan', '#00FFFF'], ['magenta', '#FF00FF'], ['red', '#FF0000'], ['lightGray', '#C0C0C0']];
  const FONTS = ['Calibri', 'Calibri Light', 'Aptos', 'Arial', 'Times New Roman', 'Cambria', 'Georgia', 'Garamond', 'Verdana', 'Tahoma', 'Segoe UI', 'Century Gothic', 'Trebuchet MS', 'Courier New'];

  function icon(d) {
    const s = document.createElementNS('http://www.w3.org/2000/svg', 'svg');
    s.setAttribute('viewBox', '0 0 24 24'); s.setAttribute('fill', 'none'); s.setAttribute('stroke', 'currentColor'); s.setAttribute('stroke-width', '1.8'); s.setAttribute('stroke-linecap', 'round'); s.setAttribute('stroke-linejoin', 'round');
    const p = document.createElementNS('http://www.w3.org/2000/svg', 'path'); p.setAttribute('d', d); s.append(p);
    return s;
  }
  async function run(fn) {
    if (!state.doc) return;
    OF.checkpoint();
    try { await fn(state.doc); } catch (e) { console.error(e); toast('No se pudo completar: ' + (e.message || e), 5000); }
    refreshUI();
  }
  const closePop = () => OF.closePopup();
  const openPop = (btn, build) => OF.popup(btn, (p) => build(p));
  function swatches(list, onPick, noneLabel) {
    const wrap = OF.el('div', { class: 'sw-grid' });
    for (const c of list) wrap.append(OF.el('button', { class: 'sw', title: c, 'aria-label': c, style: { background: c }, onclick: () => { closePop(); onPick(c); } }));
    return [noneLabel ? OF.el('button', { class: 'pop-item', onclick: () => { closePop(); onPick(null); } }, noneLabel) : null, wrap].filter(Boolean);
  }
  function menuItems(items) {
    return items.map((it) => (it === '-' ? OF.el('hr') : OF.el('button', { class: 'pop-item' + (it[2] ? ' danger' : ''), onclick: () => { closePop(); run(it[1]); } }, it[0])));
  }
  let rbId = 0;
  function rb(label, title, onclick, opts = {}) {
    const b = OF.el('button', { class: 'rb' + (opts.cls ? ' ' + opts.cls : ''), id: 'rb' + rbId++, title, 'aria-label': title, type: 'button' });
    if (opts.fs) b.dataset.fs = opts.fs;
    if (typeof label === 'string') b.append(OF.el('span', { class: 'rb-t' }, label)); else b.append(label);
    if (opts.text) b.append(OF.el('span', { class: 'rb-l' }, opts.text));
    if (opts.caret) b.append(OF.el('span', { class: 'rb-c', 'aria-hidden': 'true' }, '▾'));
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', (e) => { e.stopPropagation(); onclick(b); });
    return b;
  }
  const sep = () => OF.el('span', { class: 'rb-sep' });
  const menuBtn = (iconD, title, text, items) => rb(icon(iconD), title, (b) => openPop(b, (p) => p.append(...(typeof items === 'function' ? items() : menuItems(items)))), { text, caret: true });

  // selector de tamaño de tabla
  function tableGrid(b, onPick) {
    openPop(b, (p) => {
      const lbl = OF.el('div', { class: 'pop-h' }, 'Insertar tabla');
      const g = OF.el('div', { class: 'tgrid' });
      const R = 8, Cn = 8;
      const cells = [];
      for (let r = 0; r < R; r++) for (let c = 0; c < Cn; c++) {
        const x = OF.el('button', { class: 'tg', type: 'button', 'aria-label': `${r + 1} por ${c + 1}` });
        x.addEventListener('pointerenter', () => { cells.forEach((y) => y.el.classList.toggle('on', y.r <= r && y.c <= c)); lbl.textContent = `Tabla de ${r + 1} × ${c + 1}`; });
        x.addEventListener('click', () => { closePop(); run(() => onPick(r + 1, c + 1)); });
        cells.push({ el: x, r, c });
        g.append(x);
      }
      p.append(lbl, g);
    });
  }
  function fontControls(kind) {
    const fontBtn = rb(OF.el('span', { class: 'rb-font', 'data-lbl': 'font' }, 'Calibri'), 'Fuente', (b) => openPop(b, (p) => {
      p.classList.add('fonts');
      const doc = state.doc;
      const list = Array.from(new Set([...(doc.fonts || []), ...FONTS]));
      for (const f of list) p.append(OF.el('button', { class: 'pop-item', style: { fontFamily: `"${f}", sans-serif` }, onclick: () => { closePop(); run((d) => (d.setFont ? d.setFont(f) : d.applyStyle({ font: { name: f } }))); } }, f));
    }), { caret: true, cls: 'wide' });
    const sizeBtn = rb(OF.el('span', { class: 'rb-size', 'data-lbl': 'size' }, '11'), 'Tamaño de letra', (b) => openPop(b, (p) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Tamaño'));
      const g = OF.el('div', { class: 'size-grid' });
      for (const s of SIZES) g.append(OF.el('button', { class: 'pop-item', onclick: () => { closePop(); run((d) => d.setSize(s)); } }, String(s).replace('.', ',')));
      p.append(g);
    }), { caret: true });
    const out = [fontBtn, sizeBtn, sep(),
      rb('B', 'Negrita (Ctrl+N)', () => run((d) => d.toggle('b')), { cls: 'b', fs: 'b' }),
      rb('I', 'Cursiva (Ctrl+K)', () => run((d) => d.toggle('i')), { cls: 'i', fs: 'i' }),
      rb('U', 'Subrayado (Ctrl+U)', () => run((d) => d.toggle('u')), { cls: 'u', fs: 'u' }),
      rb(OF.el('span', { class: 'rb-a' }, 'A'), 'Color de texto', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Color de texto'), ...swatches(COLORS, (c) => run((d) => d.setColor(c)), 'Automático'))), { caret: true })];
    if (kind === 'word') out.push(rb(OF.el('span', { class: 'rb-hl' }, 'ab'), 'Resaltar', (b) => openPop(b, (p) => {
      p.append(OF.el('div', { class: 'pop-h' }, 'Resaltado'), OF.el('button', { class: 'pop-item', onclick: () => { closePop(); run((d) => d.setHighlight(null)); } }, 'Sin resaltado'));
      const g = OF.el('div', { class: 'sw-grid' });
      for (const [n, c] of HIGHLIGHTS) g.append(OF.el('button', { class: 'sw', title: n, 'aria-label': n, style: { background: c }, onclick: () => { closePop(); run((d) => d.setHighlight(n, c)); } }));
      p.append(g);
    }), { caret: true }));
    if (kind === 'excel') out.push(rb(OF.el('span', { class: 'rb-fill' }), 'Color de relleno', (b) => openPop(b, (p) => p.append(OF.el('div', { class: 'pop-h' }, 'Relleno de celda'), ...swatches(FILLS, (c) => run((d) => d.setFill(c)), 'Sin relleno'))), { caret: true }));
    return out;
  }
  function paragraphControls(kind) {
    const out = [];
    for (const [a, t] of [['left', 'Alinear a la izquierda'], ['center', 'Centrar'], ['right', 'Alinear a la derecha']]) out.push(rb(icon(ICON[a]), t, () => run((d) => d.align(a)), { fs: 'align:' + a }));
    if (kind !== 'excel') {
      out.push(rb(icon(ICON.justify), 'Justificar', () => run((d) => d.align('justify')), { fs: 'align:justify' }));
      out.push(sep());
      out.push(rb(icon(ICON.bullets), 'Viñetas', () => run((d) => d.toggleList('bullet')), { fs: 'list' }));
      out.push(rb(icon(ICON.numbers), 'Numeración', () => run((d) => d.toggleList('number'))));
      out.push(rb(icon(ICON.indentL), 'Disminuir sangría', () => run((d) => d.indent(-1))));
      out.push(rb(icon(ICON.indentR), 'Aumentar sangría', () => run((d) => d.indent(1))));
      out.push(menuBtn(ICON.spacing, 'Interlineado', null, () => {
        const items = [1, 1.15, 1.5, 2, 2.5, 3].map((m) => [String(m).replace('.', ','), (d) => d.lineSpacing(m)]);
        if (kind === 'word') items.push('-', ['Agregar espacio entre párrafos', (d) => d.paraSpacing(true)], ['Quitar espacio entre párrafos', (d) => d.paraSpacing(false)]);
        return [OF.el('div', { class: 'pop-h' }, 'Interlineado'), ...menuItems(items)];
      }));
    }
    return out;
  }

  const RIBBONS = {
    word: () => ({
      Inicio: [...fontControls('word'), sep(), ...paragraphControls('word'), sep(),
        rb(OF.el('span', { class: 'rb-style', 'data-lbl': 'style' }, 'Normal'), 'Estilos', (b) => openPop(b, (p) => {
          p.classList.add('styles');
          for (const st of state.doc.styleMenu()) p.append(OF.el('button', { class: 'pop-item sty-' + st.key.replace(/\s/g, ''), onclick: () => { closePop(); run((d) => d.setStyle(st.key)); } }, st.label));
        }), { caret: true, cls: 'wide' })],
      Insertar: [
        rb(icon(ICON.image), 'Insertar imagen', async () => { const f = await OF.pickImage(); if (f) run((d) => d.insertImage(f)); }, { text: 'Imagen' }),
        rb(icon(ICON.table), 'Insertar tabla', (b) => tableGrid(b, (r, c) => state.doc.insertTable(r, c)), { text: 'Tabla', caret: true }),
        rb(icon(ICON.pagebreak), 'Salto de página', () => run((d) => d.pageBreak()), { text: 'Salto de página' }),
        sep(),
        menuBtn(ICON.header, 'Encabezado', 'Encabezado', [['Editar encabezado', (d) => d.editHF('header')], ['Primera página diferente (activar/quitar)', (d) => d.toggleFirstPage()], ['Pares e impares diferentes (activar/quitar)', (d) => d.toggleEvenOdd()], '-', ['Quitar encabezado', (d) => d.removeHF('header'), true]]),
        menuBtn(ICON.footer, 'Pie de página', 'Pie de página', [['Editar pie de página', (d) => d.editHF('footer')], ['Quitar pie de página', (d) => d.removeHF('footer'), true]]),
        rb(icon(ICON.footnote), 'Insertar nota al pie en el cursor', () => run((d) => d.insertFootnote()), { text: 'Nota al pie' }),
        menuBtn(ICON.pagenum, 'Número de página', 'Nº de página', [['Al pie de página: «Página 1 de N»', (d) => d.pageNumber('footer')], ['En el encabezado: «Página 1 de N»', (d) => d.pageNumber('header')]]),
        sep(),
        menuBtn(ICON.rows, 'Filas de la tabla', 'Filas', [['Insertar fila arriba', (d) => d.tableRow('above')], ['Insertar fila abajo', (d) => d.tableRow('below')], ['Eliminar fila', (d) => d.tableRow('del'), true]]),
        menuBtn(ICON.cols, 'Columnas de la tabla', 'Columnas', [['Insertar columna a la izquierda', (d) => d.tableCol('left')], ['Insertar columna a la derecha', (d) => d.tableCol('right')], ['Eliminar columna', (d) => d.tableCol('del'), true]]),
        rb(icon(ICON.trash), 'Eliminar la tabla donde está el cursor', () => run((d) => d.deleteTable()), { text: 'Eliminar tabla' }),
      ],
    }),
    excel: () => ({
      Inicio: [...fontControls('excel'), menuBtn(ICON.border, 'Bordes', null, [['Todos los bordes', (d) => d.setBorders('all')], ['Borde exterior', (d) => d.setBorders('outer')], ['Borde exterior grueso', (d) => d.setBorders('thick')], ['Borde inferior', (d) => d.setBorders('bottom')], ['Doble borde inferior', (d) => d.setBorders('doubleBottom')], '-', ['Sin bordes', (d) => d.setBorders('none')]]), sep(),
        ...paragraphControls('excel'),
        menuBtn(ICON.cond, 'Formato condicional', null, [['Mayor que…', (d) => d.cfDialog('gt')], ['Menor que…', (d) => d.cfDialog('lt')], ['Entre…', (d) => d.cfDialog('between')], ['Igual a…', (d) => d.cfDialog('eq')], ['Texto que contiene…', (d) => d.cfDialog('text')], ['Valores duplicados', (d) => d.cfQuick('dup')], '-', ['10 superiores', (d) => d.cfQuick('top')], ['10 inferiores', (d) => d.cfQuick('bottom')], ['Por encima del promedio', (d) => d.cfQuick('above')], ['Por debajo del promedio', (d) => d.cfQuick('below')], '-', ['Barras de datos azules', (d) => d.cfQuick('bar', '638EC6')], ['Barras de datos verdes', (d) => d.cfQuick('bar', '63BE7B')], ['Barras de datos rojas', (d) => d.cfQuick('bar', 'F8696B')], ['Escala verde – amarillo – rojo', (d) => d.cfQuick('scale', ['63BE7B', 'FFEB84', 'F8696B'])], ['Escala rojo – amarillo – verde', (d) => d.cfQuick('scale', ['F8696B', 'FFEB84', '63BE7B'])], ['Escala blanco – azul', (d) => d.cfQuick('scale', ['FFFFFF', '5A8AC6'])], ['Iconos: flechas', (d) => d.cfQuick('icons', '3Arrows')], ['Iconos: semáforo', (d) => d.cfQuick('icons', '3TrafficLights1')], '-', ['Administrar reglas…', (d) => d.cfManager()], ['Borrar reglas de las celdas seleccionadas', (d) => d.clearCF('sel'), true], ['Borrar reglas de toda la hoja', (d) => d.clearCF('sheet'), true]]),
        rb(icon(ICON.wrap), 'Ajustar texto en la celda', () => run((d) => d.wrapText()), { fs: 'wrap' }),
        menuBtn(ICON.merge, 'Combinar celdas', null, [['Combinar y centrar', (d) => d.merge('center')], ['Combinar celdas', (d) => d.merge('merge')], ['Separar celdas', (d) => d.merge('split')]]), sep(),
        menuBtn(ICON.num, 'Formato de número', null, [['General', (d) => d.numFormat('General')], ['Número  1.234,56', (d) => d.numFormat('#,##0.00')], ['Número entero  1.234', (d) => d.numFormat('#,##0')], ['Moneda  $1.234', (d) => d.numFormat('"$"#,##0')], ['Moneda con decimales  $1.234,50', (d) => d.numFormat('"$"#,##0.00')], ['Porcentaje  12%', (d) => d.numFormat('0%')], ['Porcentaje  12,5%', (d) => d.numFormat('0.0%')], ['Fecha  15-03-2026', (d) => d.numFormat('dd-mm-yyyy')], ['Fecha larga  15 de marzo de 2026', (d) => d.numFormat('d "de" mmmm "de" yyyy')], ['Hora  14:30', (d) => d.numFormat('hh:mm')], ['Texto', (d) => d.numFormat('@')]]),
        menuBtn(ICON.sum, 'Autosuma', null, [['Suma', (d) => d.autoSum('SUM')], ['Promedio', (d) => d.autoSum('AVERAGE')], ['Contar números', (d) => d.autoSum('COUNT')], ['Máximo', (d) => d.autoSum('MAX')], ['Mínimo', (d) => d.autoSum('MIN')]]),
        rb(icon(ICON.range), 'Seleccionar rango: toca la primera y la última celda', (b) => { const d = state.doc; d.rangeMode = !d.rangeMode; b.classList.toggle('on', d.rangeMode); toast(d.rangeMode ? 'Toca la última celda del rango' : 'Selección de una celda'); }, { text: 'Rango' })],
      Insertar: [
        rb(icon(ICON.image), 'Insertar imagen', async () => { const f = await OF.pickImage(); if (f) run((d) => d.insertImage(f)); }, { text: 'Imagen' }),
        menuBtn(ICON.chart, 'Insertar gráfico con los datos seleccionados', 'Gráfico', [['Columnas', (d) => d.insertChart('col')], ['Barras', (d) => d.insertChart('bar')], ['Líneas', (d) => d.insertChart('line')], ['Circular', (d) => d.insertChart('pie')]]),
        sep(),
        menuBtn(ICON.rows, 'Filas', 'Filas', [['Insertar fila arriba', (d) => d.rowOp('above')], ['Insertar fila abajo', (d) => d.rowOp('below')], ['Eliminar fila', (d) => d.rowOp('del'), true]]),
        menuBtn(ICON.cols, 'Columnas', 'Columnas', [['Insertar columna a la izquierda', (d) => d.colOp('left')], ['Insertar columna a la derecha', (d) => d.colOp('right')], ['Eliminar columna', (d) => d.colOp('del'), true]]),
        rb(icon(ICON.sheet), 'Agregar hoja', () => run((d) => d.sheetOp('add')), { text: 'Hoja' }),
      ],
      Datos: [
        rb(icon(ICON.sortAZ), 'Ordenar de A a Z (o de menor a mayor) por la columna seleccionada', () => run((d) => d.sortRange('asc')), { text: 'A → Z' }),
        rb(icon(ICON.sortZA), 'Ordenar de Z a A (o de mayor a menor) por la columna seleccionada', () => run((d) => d.sortRange('desc')), { text: 'Z → A' }),
        rb(icon(ICON.filter), 'Activar o quitar filtros', () => run((d) => d.toggleFilter()), { text: 'Filtro' }),
        sep(),
        rb(icon(ICON.fill), 'Rellenar hacia abajo (Ctrl+D)', () => run((d) => d.fillDir('down')), { text: 'Rellenar ↓' }),
        rb(icon(ICON.fill), 'Rellenar hacia la derecha (Ctrl+R)', () => run((d) => d.fillDir('right')), { text: 'Rellenar →', cls: 'rot' }),
      ],
      Vista: [
        menuBtn(ICON.freeze, 'Inmovilizar paneles', 'Inmovilizar', [['Inmovilizar fila superior', (d) => d.freeze('row')], ['Inmovilizar primera columna', (d) => d.freeze('col')], ['Inmovilizar hasta la celda seleccionada', (d) => d.freeze('cell')], ['Movilizar paneles', (d) => d.freeze('none')]]),
        rb(icon(ICON.width), 'Ancho de columna', () => run((d) => d.promptSize('col')), { text: 'Ancho' }),
        rb(icon(ICON.height), 'Alto de fila', () => run((d) => d.promptSize('row')), { text: 'Alto' }),
        rb(icon(ICON.width), 'Autoajustar el ancho de las columnas seleccionadas', () => run((d) => d.autofitSelection()), { text: 'Autoajustar' }),
        sep(),
        rb(icon(ICON.locale), 'Cambiar el formato de números y fórmulas', () => { $('localeBtn').click(); }, { text: 'Formato regional' }),
      ],
    }),
    ppt: () => ({
      Inicio: [rb(icon(ICON.slide), 'Nueva diapositiva', (b) => state.doc.newSlideMenu(b), { text: 'Nueva', caret: true }), sep(), ...fontControls('ppt'), sep(), ...paragraphControls('ppt')],
      Insertar: [
        rb(icon(ICON.slide), 'Nueva diapositiva', (b) => state.doc.newSlideMenu(b), { text: 'Diapositiva', caret: true }),
        rb(icon(ICON.textbox), 'Cuadro de texto', () => run((d) => d.insertTextBox()), { text: 'Cuadro de texto' }),
        menuBtn(ICON.shapes, 'Formas', 'Formas', [['Rectángulo', (d) => d.insertShape('rect')], ['Rectángulo redondeado', (d) => d.insertShape('roundRect')], ['Círculo', (d) => d.insertShape('ellipse')], ['Triángulo', (d) => d.insertShape('triangle')], ['Flecha', (d) => d.insertShape('rightArrow')], ['Estrella', (d) => d.insertShape('star5')], ['Línea', (d) => d.insertShape('line')]]),
        rb(icon(ICON.image), 'Insertar imagen', async () => { const f = await OF.pickImage(); if (f) run((d) => d.insertImage(f)); }, { text: 'Imagen' }),
        rb(icon(ICON.table), 'Insertar tabla', (b) => tableGrid(b, (r, c) => state.doc.insertTable(r, c)), { text: 'Tabla', caret: true }),
        sep(),
        menuBtn(ICON.rows, 'Filas de la tabla', 'Filas', [['Insertar fila arriba', (d) => d.tableRow('above')], ['Insertar fila abajo', (d) => d.tableRow('below')], ['Eliminar fila', (d) => d.tableRow('del'), true]]),
        menuBtn(ICON.cols, 'Columnas de la tabla', 'Columnas', [['Insertar columna a la izquierda', (d) => d.tableCol('left')], ['Insertar columna a la derecha', (d) => d.tableCol('right')], ['Eliminar columna', (d) => d.tableCol('del'), true]]),
      ],
      Diseño: [
        rb(icon(ICON.bg), 'Fondo de la diapositiva', (b) => openPop(b, (p) => {
          let all = false;
          const chk = OF.el('input', { type: 'checkbox' });
          chk.addEventListener('change', () => { all = chk.checked; });
          p.append(OF.el('div', { class: 'pop-h' }, 'Fondo'), ...swatches(BGS, (c) => run((d) => d.setBackground(c, all)), 'Fondo del diseño'), OF.el('label', { class: 'pop-chk' }, chk, ' Aplicar a todas las diapositivas'));
        }), { text: 'Fondo', caret: true }),
        sep(),
        rb(icon(ICON.dup), 'Duplicar diapositiva', () => run((d) => d.slideOp('dup')), { text: 'Duplicar' }),
        rb(icon(ICON.up), 'Mover diapositiva hacia arriba', () => run((d) => d.slideOp('up')), { text: 'Subir' }),
        rb(icon(ICON.down), 'Mover diapositiva hacia abajo', () => run((d) => d.slideOp('down')), { text: 'Bajar' }),
        rb(icon(ICON.hide), 'Ocultar o mostrar la diapositiva en la presentación', () => run((d) => d.toggleHidden()), { text: 'Ocultar' }),
        rb(icon(ICON.trash), 'Eliminar diapositiva', () => run((d) => d.slideOp('delete')), { text: 'Eliminar', cls: 'danger' }),
      ],
    }),
  };
  let ribbonTabs = {}, currentTab = 'Inicio';
  function buildRibbon(kind) {
    ribbonTabs = RIBBONS[kind]();
    const tabs = $('rtabs');
    tabs.innerHTML = '';
    for (const name of Object.keys(ribbonTabs)) {
      const t = OF.el('button', { class: 'rtab', role: 'tab', type: 'button' }, name);
      t.addEventListener('mousedown', (e) => e.preventDefault());
      t.addEventListener('click', () => showTab(name));
      tabs.append(t);
    }
    showTab('Inicio');
  }
  function showTab(name) {
    currentTab = name;
    const bar = $('ribbon');
    bar.innerHTML = '';
    bar.append(...ribbonTabs[name]);
    $('rtabs').querySelectorAll('.rtab').forEach((t) => { t.classList.toggle('on', t.textContent === name); t.setAttribute('aria-selected', String(t.textContent === name)); });
    refreshUI();
  }

  /* --------------------------- estado de la cinta --------------------------- */
  function refreshFormat() {
    const d = state.doc;
    if (!d || !d.formatState) return;
    let st;
    try { st = d.formatState(); } catch (e) { st = null; }
    const bar = $('ribbon');
    bar.querySelectorAll('[data-fs]').forEach((b) => {
      const k = b.dataset.fs;
      const on = st && (k.startsWith('align:') ? st.align === k.slice(6) : !!st[k]);
      b.classList.toggle('on', !!on);
      b.setAttribute('aria-pressed', String(!!on));
    });
    if (!st) return;
    const f = bar.querySelector('[data-lbl=font]'); if (f && st.font) f.textContent = st.font;
    const s = bar.querySelector('[data-lbl=size]'); if (s && st.size) s.textContent = String(st.size).replace('.', ',');
    const y = bar.querySelector('[data-lbl=style]'); if (y && st.style) y.textContent = st.style;
  }
  function refreshStatus() {
    const d = state.doc;
    if (!d) return;
    let t = '';
    try { t = d.statusText ? d.statusText() : ''; } catch (e) {}
    $('status').textContent = t;
  }
  let uiT;
  function refreshUI() { clearTimeout(uiT); uiT = setTimeout(() => { refreshFormat(); refreshStatus(); updateUndo(); }, 60); }
  document.addEventListener('selectionchange', () => { if (state.doc && state.kind !== 'excel') refreshUI(); });

  /* --------------------------- deshacer / rehacer --------------------------- */
  const hist = { undo: [], redo: [], last: 0, typing: 0, kind: null, busy: false };
  function resetHistory() { hist.undo = []; hist.redo = []; hist.kind = null; updateUndo(); }
  function updateUndo() { $('undoBtn').disabled = !hist.undo.length; $('redoBtn').disabled = !hist.redo.length; }
  OF.checkpoint = function () {
    const d = state.doc;
    if (!d || !d.snapshot || hist.busy) return;
    const now = Date.now();
    if (hist.kind === 'op' && now - hist.last < 80) return;
    try { hist.undo.push(d.snapshot()); } catch (e) { console.warn('Sin instantánea', e); return; }
    if (hist.undo.length > 60) hist.undo.shift();
    hist.redo = [];
    hist.last = now; hist.kind = 'op';
    updateUndo();
    scheduleAutosave();
  };
  OF.typingCheckpoint = function (e) {
    if (e && e.inputType && /^history/.test(e.inputType)) { e.preventDefault(); e.inputType === 'historyUndo' ? undo() : redo(); return; }
    if (e && e.inputType && !/^(insert|delete|format)/.test(e.inputType)) return;
    const now = Date.now();
    if (hist.kind !== 'typing' || now - hist.typing > 1500) { OF.checkpoint(); hist.kind = 'typing'; }
    hist.typing = now;
    scheduleAutosave();
  };
  async function stepHistory(from, to, label) {
    const d = state.doc;
    if (!d || hist.busy || !from.length) return;
    hist.busy = true;
    try {
      to.push(d.snapshot());
      const s = from.pop();
      await d.restore(s);
      d.dirty = true; updateState();
      hist.kind = null;
    } catch (e) { console.error(e); toast('No se pudo ' + label + '.'); }
    hist.busy = false;
    updateUndo(); refreshUI(); scheduleAutosave();
  }
  const undo = () => stepHistory(hist.undo, hist.redo, 'deshacer');
  const redo = () => stepHistory(hist.redo, hist.undo, 'rehacer');

  /* ------------------------- guardado automático ------------------------- */
  let autoT, autoBusy = false;
  function scheduleAutosave() { clearTimeout(autoT); autoT = setTimeout(autosave, 2500); }
  async function autosave() {
    const d = state.doc;
    if (!d || !d.dirty || !d.snapshot || autoBusy) return;
    autoBusy = true;
    try {
      const s = d.snapshot();
      const z = new JSZip();
      for (const k in s.files) z.files[k] = s.files[k];
      const bytes = await z.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE', compressionOptions: { level: 1 } });
      const rec = await tx('readonly', (st) => st.get(state.recentId));
      if (rec && state.doc === d) await putRecent({ ...rec, draft: { bytes, html: s.html ? s.html.map(OF.serializeKids) : null, at: Date.now() }, at: Date.now() });
    } catch (e) { console.warn('Autoguardado falló', e); }
    autoBusy = false;
  }

  /* ----------------------------- buscar ----------------------------- */
  function openFind(withReplace) {
    const fb = $('findbar');
    fb.classList.remove('hidden');
    fb.classList.toggle('rep', !!withReplace || fb.classList.contains('rep'));
    const sel = getSelection().toString();
    if (sel && sel.length < 80 && !sel.includes('\n')) $('findQ').value = sel;
    $('findQ').focus(); $('findQ').select();
    if (state.doc) state.doc._fresh = true;
    if ($('findQ').value) doFind(1);
  }
  function closeFind() { $('findbar').classList.add('hidden'); $('findCount').textContent = ''; }
  function showCount(r) { $('findCount').textContent = r.n ? `${r.i + 1} de ${r.n}` : $('findQ').value ? 'Sin resultados' : ''; }
  function doFind(dir) { const d = state.doc; if (!d || !d.find) return; showCount(d.find($('findQ').value, dir)); $('findQ').focus(); }
  $('findBtn').addEventListener('click', () => ($('findbar').classList.contains('hidden') ? openFind(true) : closeFind()));
  $('findClose').addEventListener('click', closeFind);
  $('findNext').addEventListener('click', () => doFind(1));
  $('findPrev').addEventListener('click', () => doFind(-1));
  let findT;
  $('findQ').addEventListener('input', () => { clearTimeout(findT); findT = setTimeout(() => { if (state.doc) state.doc._fresh = true; doFind(1); }, 200); });
  $('findQ').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); doFind(e.shiftKey ? -1 : 1); } if (e.key === 'Escape') closeFind(); });
  $('findR').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); $('repOne').click(); } if (e.key === 'Escape') closeFind(); });
  $('repOne').addEventListener('click', () => { const d = state.doc; if (!d) return; showCount(d.replaceOne($('findQ').value, $('findR').value)); updateUndo(); });
  $('repAll').addEventListener('click', () => { const d = state.doc; if (!d || !$('findQ').value) return; const n = d.replaceAll($('findQ').value, $('findR').value); toast(n ? `${n} reemplazo${n === 1 ? '' : 's'}` : 'Sin resultados'); showCount({ n: 0 }); updateUndo(); });

  /* ----------------------------- imprimir ----------------------------- */
  function printDoc() {
    const d = state.doc;
    if (!d) return;
    closeMenu();
    const css = d.printSetup ? d.printSetup() : '';
    let st = document.getElementById('printPage');
    if (!st) { st = OF.el('style', { id: 'printPage' }); document.head.append(st); }
    st.textContent = css;
    document.body.classList.add('printing', 'print-' + state.kind);
    const done = () => { document.body.classList.remove('printing', 'print-word', 'print-excel', 'print-ppt'); st.textContent = ''; d.fit && d.fit(); window.removeEventListener('afterprint', done); };
    window.addEventListener('afterprint', done);
    setTimeout(() => { window.print(); done(); }, 50);
  }
  $('printBtn').addEventListener('click', printDoc);

  /* ------------------------------- vistas ------------------------------- */
  function showEditor(kind, name) {
    $('home').classList.add('hidden');
    $('editor').classList.remove('hidden');
    $('editor').dataset.kind = kind;
    $('fname').textContent = name;
    $('presentBtn').classList.toggle('hidden', kind !== 'ppt');
    $('localeBtn').classList.toggle('hidden', kind !== 'excel');
    const hints = {
      word: 'Escribe directamente sobre la página. Selecciona texto para darle formato. Ctrl+Z deshace y Ctrl+F busca.',
      excel: 'Toca una celda para seleccionarla y otra vez para editar. Arrastra el cuadrito verde para rellenar. Las fórmulas se escriben en español con punto y coma: =SUMA(A1:A5; B1).',
      ppt: 'Toca un cuadro de texto para escribir. Usa «⠿ Mover» para moverlo y el círculo de la esquina para cambiar su tamaño. La columna izquierda muestra las diapositivas.',
    };
    $('menuHint').textContent = hints[kind];
    closeFind();
    $('status').textContent = '';
    buildRibbon(kind);
    history.pushState({ editor: true }, '');
  }
  async function showHome(fromPop) {
    const d = state.doc;
    if (d && d.dirty) {
      clearTimeout(autoT);
      await autosave();
      toast('Tus cambios quedaron como borrador en Recientes. Para tener el archivo, toca Guardar.', 4500);
    }
    void fromPop;
    state.doc = null;
    resetHistory();
    closeFind();
    OF.closePopup(); OF.closeImageBar();
    $('editor').classList.add('hidden');
    $('home').classList.remove('hidden');
    $('view').innerHTML = '';
    $('view').className = '';
    closeMenu();
    renderRecent();
  }
  async function confirmLeave() {
    if (!state.doc || !state.doc.dirty) return true;
    clearTimeout(autoT);
    await autosave();
    return true;
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
  const showZoom = () => { const d = state.doc; if (d) $('statusZoom').textContent = Math.round((d.userZoom || 1) * 100) + '%'; };
  $('zoomIn').addEventListener('click', () => { if (state.doc) { state.doc.zoom(1.15); showZoom(); } });
  $('zoomOut').addEventListener('click', () => { if (state.doc) { state.doc.zoom(1 / 1.15); showZoom(); } });
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
    toast(d.locale === 'cl' ? 'Formato chileno: 1.234,56 y fórmulas en español (=SUMA(A1;B1))' : 'Formato inglés: 1,234.56 y fórmulas en inglés (=SUM(A1,B1))', 4500);
    try { localStorage.setItem('folio-locale', d.locale); } catch (e) {}
    d.showSheet(d.cur);
  });
  document.addEventListener('keydown', (e) => {
    if (!state.doc || $('editor').classList.contains('hidden')) return;
    const k = e.key.toLowerCase(), mod = e.ctrlKey || e.metaKey;
    if (!mod) return;
    const inField = e.target.matches && e.target.matches('input, textarea');
    if (k === 's' || k === 'g') { e.preventDefault(); save(false); }
    else if (k === 'p') { e.preventDefault(); printDoc(); }
    else if (k === 'f' || k === 'b') { e.preventDefault(); openFind(false); }
    else if (k === 'h' || k === 'l') { e.preventDefault(); openFind(true); }
    else if (!inField && k === 'z' && !e.shiftKey) { e.preventDefault(); undo(); }
    else if (!inField && (k === 'y' || (k === 'z' && e.shiftKey))) { e.preventDefault(); redo(); }
  }, true);
  document.addEventListener('beforeinput', (e) => { if (state.doc && /^history/.test(e.inputType) && !e.target.matches('input, textarea')) { e.preventDefault(); e.inputType === 'historyUndo' ? undo() : redo(); } }, true);
  $('undoBtn').addEventListener('click', undo);
  $('redoBtn').addEventListener('click', redo);
  $('recentSearch').addEventListener('input', () => renderRecent());
  let rz;
  window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => state.doc && state.doc.fit && state.doc.fit(), 150); });
  window.addEventListener('beforeunload', (e) => { if (state.doc && state.doc.dirty) { autosave(); e.preventDefault(); e.returnValue = ''; } });
  document.addEventListener('visibilitychange', () => { if (document.visibilityState === 'hidden' && state.doc && state.doc.dirty) autosave(); });

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
  renderTemplates();
  renderRecent();
})();
