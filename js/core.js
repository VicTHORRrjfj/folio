/* Núcleo compartido: utilidades XML/ZIP y sincronización de texto editado con el XML original.
   Principio: nunca regenerar el documento. Solo se cambia el texto dentro de los nodos
   existentes; estilos, imágenes y todo lo demás quedan intactos. */
(function () {
  const OF = (window.OF = window.OF || {});

  OF.NS = {
    w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
    p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
    r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    s: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
    mc: 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    xml: 'http://www.w3.org/XML/1998/namespace',
  };

  // Office no usa entradas de carpeta dentro del ZIP: no crearlas
  if (window.JSZip && JSZip.defaults) JSZip.defaults.createFolders = false;

  OF.parseXml = function (text) {
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length) throw new Error('XML inválido');
    return doc;
  };

  OF.serializeXml = function (doc) {
    let s = new XMLSerializer().serializeToString(doc);
    if (!s.startsWith('<?xml')) s = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s;
    return s;
  };

  // Hijos directos con un nombre local (y namespace opcional)
  OF.kids = function (el, local, ns) {
    const out = [];
    if (!el) return out;
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === local && (!ns || c.namespaceURI === ns)) out.push(c);
    }
    return out;
  };
  OF.kid = function (el, local, ns) {
    if (!el) return null;
    for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
      if (c.localName === local && (!ns || c.namespaceURI === ns)) return c;
    }
    return null;
  };
  // Descendiente por ruta de nombres locales: path(el,'spPr','xfrm','off')
  OF.path = function (el, ...names) {
    let cur = el;
    for (const n of names) { cur = OF.kid(cur, n); if (!cur) return null; }
    return cur;
  };
  OF.all = function (el, local, ns) {
    if (!el) return [];
    return Array.from(ns ? el.getElementsByTagNameNS(ns, local) : el.getElementsByTagName('*')).filter(
      (e) => ns || e.localName === local
    );
  };
  OF.attr = function (el, name, def) {
    if (!el) return def;
    const v = el.getAttribute(name);
    return v === null ? def : v;
  };

  // Rutas relativas dentro del paquete OOXML
  OF.resolvePath = function (base, target) {
    if (target.startsWith('/')) return target.slice(1);
    const parts = base.split('/');
    parts.pop();
    for (const seg of target.split('/')) {
      if (seg === '..') parts.pop();
      else if (seg !== '.') parts.push(seg);
    }
    return parts.join('/');
  };
  OF.relsPathFor = function (partPath) {
    const i = partPath.lastIndexOf('/');
    return partPath.slice(0, i + 1) + '_rels/' + partPath.slice(i + 1) + '.rels';
  };
  OF.readRels = async function (zip, partPath) {
    const f = zip.file(OF.relsPathFor(partPath));
    const map = {};
    if (!f) return map;
    const doc = OF.parseXml(await f.async('string'));
    for (const r of OF.all(doc, 'Relationship')) {
      map[r.getAttribute('Id')] = {
        type: r.getAttribute('Type') || '',
        target: r.getAttribute('TargetMode') === 'External' ? r.getAttribute('Target') : OF.resolvePath(partPath, r.getAttribute('Target')),
        external: r.getAttribute('TargetMode') === 'External',
      };
    }
    return map;
  };

  OF.mimeFor = function (path) {
    const ext = path.split('.').pop().toLowerCase();
    return (
      { png: 'image/png', jpg: 'image/jpeg', jpeg: 'image/jpeg', gif: 'image/gif', bmp: 'image/bmp', svg: 'image/svg+xml', webp: 'image/webp', tif: 'image/tiff', tiff: 'image/tiff', emf: 'image/emf', wmf: 'image/wmf' }[ext] ||
      'application/octet-stream'
    );
  };

  OF.setText = function (tEl, text) {
    tEl.textContent = text;
    if (/^\s|\s$/.test(text) || /\s\s/.test(text)) tEl.setAttributeNS(OF.NS.xml, 'xml:space', 'preserve');
  };

  /* ---------------------------------------------------------------------------
     TextSync: relaciona párrafos/fragmentos del XML con elementos del DOM editable
     y, al guardar, aplica los cambios sobre el XML.

     cfg = { ns, P:'p', R:'r', T:'t', PPR:'pPr', wrappers:Set(['hyperlink',...]) }
     El renderizador debe marcar:
       elemento de párrafo DOM  -> data-p="<id>"
       elemento de fragmento    -> data-r="<id>"
       contenido no editable    -> data-x (se ignora, se conserva desde el XML)
  --------------------------------------------------------------------------- */
  class TextSync {
    constructor(cfg) {
      this.cfg = cfg;
      this.paras = []; // id -> elemento XML w:p / a:p
      this.runs = []; // id -> elemento XML w:r / a:r (con un solo w:t)
      this.runPara = []; // run id -> para id
      this.paraRuns = []; // para id -> [run ids]
      this.mapped = new Set(); // ids de párrafos que tienen un DOM asociado
    }
    addPara(el) {
      const id = this.paras.length;
      this.paras.push(el);
      this.paraRuns.push([]);
      return id;
    }
    addRun(pid, el) {
      const id = this.runs.length;
      this.runs.push(el);
      this.runPara.push(pid);
      this.paraRuns[pid].push(id);
      return id;
    }
    // Fragmentos con texto de un párrafo (directos o dentro de envoltorios como hipervínculos)
    textRunsOf(pEl) {
      const { ns, R, T, wrappers } = this.cfg;
      const out = [];
      const visit = (el) => {
        for (let c = el.firstElementChild; c; c = c.nextElementSibling) {
          if (c.namespaceURI !== ns) continue;
          if (c.localName === R) {
            if (OF.kid(c, T, ns)) out.push(c);
          } else if (wrappers.has(c.localName)) visit(c);
        }
      };
      visit(pEl);
      return out;
    }
    textOfRun(rEl) {
      const t = OF.kid(rEl, this.cfg.T, this.cfg.ns);
      return t ? t.textContent : '';
    }

    // Lee el DOM editado y lo vuelca al XML. Devuelve true si hubo cambios.
    apply(domParas) {
      const cfg = this.cfg;
      const groups = new Map();
      let lastPid = null;
      for (const dp of domParas) {
        let pid = dp.hasAttribute('data-p') ? +dp.getAttribute('data-p') : lastPid;
        if (pid == null || !this.mapped.has(pid)) continue;
        lastPid = pid;
        if (!groups.has(pid)) groups.set(pid, []);
        const items = collectItems(dp);
        items.pfmt = dp.dataset.pfmt ? JSON.parse(dp.dataset.pfmt) : null;
        groups.get(pid).push(items);
      }
      const result = new Map(); // pid -> párrafos XML resultantes (para operaciones posteriores)
      // ocurrencias de cada fragmento en todo el DOM
      const occ = new Map();
      for (const list of groups.values()) for (const items of list) for (const it of items) if (it.r != null) occ.set(it.r, (occ.get(it.r) || 0) + 1);

      let changed = false;
      const rebuild = [];
      for (const pid of this.mapped) {
        const list = groups.get(pid);
        const orig = this.paraRuns[pid];
        const simple =
          list && list.length === 1 && list[0].every((it) => it.r != null) &&
          list[0].length === orig.length && list[0].every((it, i) => it.r === orig[i] && occ.get(it.r) === 1);
        if (simple) {
          for (const it of list[0]) {
            const t = OF.kid(this.runs[it.r], cfg.T, cfg.ns);
            if (t.textContent !== it.text) { OF.setText(t, it.text); changed = true; }
            if (it.fmt && cfg.applyFmt) { cfg.applyFmt(this.runs[it.r], it.fmt); changed = true; }
          }
          if (list[0].pfmt && cfg.applyPFmt) { cfg.applyPFmt(this.paras[pid], list[0].pfmt); changed = true; }
          result.set(pid, [this.paras[pid]]);
        } else rebuild.push(pid);
      }
      this.result = result;
      if (!rebuild.length) return changed;

      // anclas: contenido no textual (imágenes, tabulaciones, marcadores) ubicado tras un fragmento
      const anchorsAfter = new Map(), lead = new Map();
      for (const pid of rebuild) {
        const p = this.paras[pid];
        const runSet = new Set(this.paraRuns[pid].map((r) => this.runs[r]));
        let cur = null;
        const L = [];
        for (let c = p.firstElementChild; c; c = c.nextElementSibling) {
          if (c.localName === cfg.PPR || c.localName === 'endParaRPr') continue;
          if (runSet.has(c)) { cur = this.runs.indexOf(c); continue; }
          // envoltorio que contiene fragmentos de texto -> se reconstruye vía sus fragmentos
          let containsText = false;
          for (const r of runSet) if (c.contains(r)) { containsText = true; break; }
          if (containsText) {
            const inner = this.paraRuns[pid].filter((r) => c.contains(this.runs[r]));
            cur = inner[inner.length - 1];
            continue;
          }
          if (cur == null) L.push(c);
          else { if (!anchorsAfter.has(cur)) anchorsAfter.set(cur, []); anchorsAfter.get(cur).push(c); }
        }
        lead.set(pid, L);
      }

      const seen = new Map();
      const leadDone = new Set();
      const ensureLead = (pid, target) => {
        if (leadDone.has(pid) || !lead.has(pid)) return;
        leadDone.add(pid);
        for (const n of lead.get(pid)) target.appendChild(n);
      };
      const rebuildSet = new Set(rebuild);
      const newParasByPid = new Map();

      for (const pid of rebuild) {
        const src = this.paras[pid];
        const list = groups.get(pid) || [];
        const made = [];
        list.forEach((items, gi) => {
          const np = src.cloneNode(false);
          // Word: los párrafos nuevos no deben repetir los identificadores internos del original
          if (gi > 0) for (const at of Array.from(np.attributes)) if (/^(paraId|textId)$/.test(at.localName)) np.removeAttributeNode(at);
          const ppr = OF.kid(src, cfg.PPR, cfg.ns);
          if (ppr) np.appendChild(ppr.cloneNode(true));
          if (gi === 0) ensureLead(pid, np);
          let lastRun = null, wrapSrc = null, wrapClone = null;
          for (const it of items) {
            let rid = it.r;
            let tmplRun;
            if (rid != null) tmplRun = this.runs[rid];
            else tmplRun = lastRun != null ? this.runs[lastRun] : this.paraRuns[pid].length ? this.runs[this.paraRuns[pid][0]] : null;
            let nr;
            if (tmplRun) {
              nr = tmplRun.cloneNode(true);
              OF.setText(OF.kid(nr, cfg.T, cfg.ns), it.text);
            } else {
              nr = src.ownerDocument.createElementNS(cfg.ns, cfg.prefix + ':' + cfg.R);
              const markProps = cfg.markProps && cfg.markProps(src);
              if (markProps) {
                const rp = src.ownerDocument.createElementNS(cfg.ns, cfg.prefix + ':rPr');
                for (const at of Array.from(markProps.attributes)) rp.setAttribute(at.name, at.value);
                for (const ch of Array.from(markProps.children)) rp.appendChild(ch.cloneNode(true));
                nr.appendChild(rp);
              }
              const t = src.ownerDocument.createElementNS(cfg.ns, cfg.prefix + ':' + cfg.T);
              OF.setText(t, it.text);
              nr.appendChild(t);
            }
            if (it.fmt && cfg.applyFmt) cfg.applyFmt(nr, it.fmt);
            if (rid != null) {
              const opid = this.runPara[rid];
              if (opid !== pid) ensureLead(opid, np);
            }
            // conservar envoltorio (p.ej. hipervínculo) si el fragmento original estaba dentro de uno
            const parent = tmplRun && tmplRun.parentNode;
            const origPara = tmplRun ? this.paras[this.runPara[this.runs.indexOf(tmplRun)]] : src;
            if (parent && parent !== origPara && rid != null) {
              if (wrapSrc !== parent) { wrapSrc = parent; wrapClone = parent.cloneNode(false); np.appendChild(wrapClone); }
              wrapClone.appendChild(nr);
            } else {
              wrapSrc = null; wrapClone = null;
              np.appendChild(nr);
            }
            if (rid != null) {
              lastRun = rid;
              seen.set(rid, (seen.get(rid) || 0) + 1);
              if (seen.get(rid) === occ.get(rid) && anchorsAfter.has(rid)) {
                for (const n of anchorsAfter.get(rid)) np.appendChild(n);
                anchorsAfter.delete(rid);
              }
            }
          }
          // PowerPoint: endParaRPr debe ir al final del párrafo
          const epr = OF.kid(src, 'endParaRPr', cfg.ns);
          if (epr) np.appendChild(epr.cloneNode(true));
          if (items.pfmt && cfg.applyPFmt) cfg.applyPFmt(np, items.pfmt);
          made.push(np);
        });
        newParasByPid.set(pid, made);
        result.set(pid, made);
      }
      // anclas que quedaron sin ubicar (su texto fue borrado): al final del párrafo de origen
      for (const [rid, nodes] of anchorsAfter) {
        const made = newParasByPid.get(this.runPara[rid]);
        if (made && made.length) {
          const last = made[made.length - 1];
          const epr = OF.kid(last, 'endParaRPr', cfg.ns);
          for (const n of nodes) last.insertBefore(n, epr);
        }
      }
      for (const pid of rebuild) {
        if (!leadDone.has(pid)) {
          const made = newParasByPid.get(pid);
          if (made && made.length) { const first = made[0]; const ppr = OF.kid(first, cfg.PPR, cfg.ns); for (const n of lead.get(pid) || []) first.insertBefore(n, ppr ? ppr.nextSibling : first.firstChild); }
        }
      }
      // reemplazar en el XML
      for (const pid of rebuild) {
        const src = this.paras[pid];
        const parent = src.parentNode;
        const made = newParasByPid.get(pid);
        for (const np of made) parent.insertBefore(np, src);
        parent.removeChild(src);
        if (!made.length && cfg.mustHaveParagraph && cfg.mustHaveParagraph(parent) && !OF.kids(parent, cfg.P, cfg.ns).length) {
          const keep = src.cloneNode(false);
          const ppr = OF.kid(src, cfg.PPR, cfg.ns);
          if (ppr) keep.appendChild(ppr.cloneNode(true));
          parent.appendChild(keep);
        }
      }
      return true;
    }
  }
  OF.TextSync = TextSync;

  function collectItems(dp) {
    const items = [];
    const walk = (node, fmt) => {
      for (let c = node.firstChild; c; c = c.nextSibling) {
        if (c.nodeType === 3) {
          const t = clean(c.nodeValue);
          if (t) items.push({ r: null, text: t, fmt });
        } else if (c.nodeType === 1) {
          if (c.hasAttribute('data-x')) continue;
          let f = fmt;
          if (c.hasAttribute('data-fmt')) f = { ...(fmt || {}), ...JSON.parse(c.getAttribute('data-fmt')) };
          if (c.hasAttribute('data-r')) {
            items.push({ r: +c.getAttribute('data-r'), text: clean(c.textContent), fmt: f });
            continue;
          }
          if (c.hasAttribute('data-p') && c !== dp) continue; // párrafo anidado (no debería ocurrir)
          walk(c, f);
        }
      }
    };
    walk(dp, null);
    return items;
  }
  function clean(s) {
    return s.replace(/​/g, '').replace(/ /g, ' ');
  }
  OF.cleanText = clean;

  /* ------------------------- formato sobre la selección ------------------------- */
  // Recuerda la última selección dentro de un área editable (los botones de la barra la usan)
  OF.lastRange = null;
  document.addEventListener('selectionchange', () => {
    const sel = getSelection();
    if (!sel.rangeCount) return;
    const r = sel.getRangeAt(0);
    const host = (r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentElement);
    if (host && host.closest && host.closest('[contenteditable=true]')) OF.lastRange = r.cloneRange();
  });
  OF.currentRange = function () {
    const sel = getSelection();
    if (sel.rangeCount) {
      const r = sel.getRangeAt(0);
      const h = r.commonAncestorContainer.nodeType === 1 ? r.commonAncestorContainer : r.commonAncestorContainer.parentElement;
      if (h && h.closest && h.closest('[contenteditable=true]')) return r;
    }
    return OF.lastRange && OF.lastRange.startContainer.isConnected ? OF.lastRange : null;
  };

  // Nodos de texto editables dentro del rango
  OF.textNodesIn = function (range) {
    const root = range.commonAncestorContainer.nodeType === 1 ? range.commonAncestorContainer : range.commonAncestorContainer.parentElement;
    const out = [];
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    let n;
    if (root.nodeType === 3) return [root];
    while ((n = w.nextNode())) {
      if (!n.nodeValue || !range.intersectsNode(n)) continue;
      if (n.parentElement.closest('[data-x]') || !n.parentElement.closest('[contenteditable=true]')) continue;
      if (!n.parentElement.closest('[data-p]')) continue;
      out.push(n);
    }
    if (range.startContainer.nodeType === 3 && range.startContainer === range.endContainer && !out.includes(range.startContainer)) out.push(range.startContainer);
    return out;
  };

  // Aísla el texto seleccionado en sus propios <span> (copias del original, mismo data-r)
  // y les marca el formato pedido. El XML se actualiza al guardar.
  OF.formatSelection = function (range, fmt, css) {
    if (!range || range.collapsed) return false;
    const nodes = OF.textNodesIn(range);
    if (!nodes.length) return false;
    const done = [];
    for (let node of nodes) {
      let start = node === range.startContainer ? range.startOffset : 0;
      let end = node === range.endContainer ? range.endOffset : node.nodeValue.length;
      if (start >= end) continue;
      if (end < node.nodeValue.length) node.splitText(end);
      if (start > 0) node = node.splitText(start);
      let span = node.parentElement;
      if (!(span.tagName === 'SPAN' && (span.hasAttribute('data-r') || span.hasAttribute('data-fmt')))) {
        const ns = document.createElement('span');
        node.before(ns);
        ns.append(node);
        span = ns;
      } else {
        const before = [], after = [];
        let seen = false;
        for (const c of Array.from(span.childNodes)) { if (c === node) { seen = true; continue; } (seen ? after : before).push(c); }
        if (before.length) { const b = span.cloneNode(false); before.forEach((c) => b.append(c)); span.before(b); }
        if (after.length) { const a = span.cloneNode(false); after.forEach((c) => a.append(c)); span.after(a); }
      }
      const prev = span.dataset.fmt ? JSON.parse(span.dataset.fmt) : {};
      span.dataset.fmt = JSON.stringify({ ...prev, ...fmt });
      Object.assign(span.style, css);
      done.push(node);
    }
    if (!done.length) return false;
    const r = document.createRange();
    r.setStart(done[0], 0);
    r.setEnd(done[done.length - 1], done[done.length - 1].nodeValue.length);
    const sel = getSelection();
    sel.removeAllRanges();
    sel.addRange(r);
    OF.lastRange = r.cloneRange();
    return true;
  };
  // ¿Toda la selección ya tiene esta propiedad? (para alternar negrita, etc.)
  OF.selectionHas = function (range, test) {
    const nodes = OF.textNodesIn(range).filter((n) => n.nodeValue.trim());
    return nodes.length > 0 && nodes.every((n) => test(getComputedStyle(n.parentElement)));
  };
  // Párrafos (data-p) tocados por la selección
  OF.parasIn = function (range) {
    const set = new Set();
    const startP = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement).closest('[data-p]');
    if (startP) set.add(startP);
    for (const n of OF.textNodesIn(range)) { const p = n.parentElement.closest('[data-p]'); if (p) set.add(p); }
    return Array.from(set);
  };

  /* ------------------------- XML en el orden de Office ------------------------- */
  // Devuelve el hijo "name" de parent, creándolo en la posición que exige el esquema
  OF.ordered = function (parent, ns, prefix, name, order) {
    let el = OF.kid(parent, name, ns);
    if (el) return el;
    el = parent.ownerDocument.createElementNS(ns, prefix ? prefix + ':' + name : name);
    const idx = order.indexOf(name);
    let before = null;
    for (let c = parent.firstElementChild; c; c = c.nextElementSibling) {
      const ci = order.indexOf(c.localName);
      if (ci > idx || (ci === -1 && idx !== -1 && false)) { before = c; break; }
    }
    parent.insertBefore(el, before);
    return el;
  };
  OF.removeKids = function (parent, names, ns) {
    for (const c of Array.from(parent ? parent.children : [])) if (names.includes(c.localName) && (!ns || c.namespaceURI === ns)) parent.removeChild(c);
  };

  /* ------------------------- paquete: relaciones y tipos ------------------------- */
  OF.relTarget = function (from, to) {
    const a = from.split('/'); a.pop();
    const b = to.split('/');
    while (a.length && b.length && a[0] === b[0]) { a.shift(); b.shift(); }
    return '../'.repeat(a.length) + b.join('/');
  };
  OF.addRel = async function (zip, partPath, type, targetPath, external) {
    const relPath = OF.relsPathFor(partPath);
    const f = zip.file(relPath);
    const doc = f ? OF.parseXml(await f.async('string')) : OF.parseXml(`<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="${OF.NS.rel}"/>`);
    const ids = OF.all(doc, 'Relationship').map((r) => +(r.getAttribute('Id') || '').replace(/\D/g, '') || 0);
    const id = 'rId' + (Math.max(0, ...ids) + 1);
    const r = doc.createElementNS(OF.NS.rel, 'Relationship');
    r.setAttribute('Id', id);
    r.setAttribute('Type', type);
    r.setAttribute('Target', external ? targetPath : OF.relTarget(partPath, targetPath));
    if (external) r.setAttribute('TargetMode', 'External');
    doc.documentElement.appendChild(r);
    zip.file(relPath, OF.serializeXml(doc));
    return id;
  };
  OF.contentTypes = async function (zip, fn) {
    const doc = OF.parseXml(await zip.file('[Content_Types].xml').async('string'));
    fn(doc);
    zip.file('[Content_Types].xml', OF.serializeXml(doc));
  };
  OF.ensureDefaultCT = function (doc, ext, mime) {
    if (OF.all(doc, 'Default').some((d) => (d.getAttribute('Extension') || '').toLowerCase() === ext)) return;
    const d = doc.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Default');
    d.setAttribute('Extension', ext);
    d.setAttribute('ContentType', mime);
    doc.documentElement.insertBefore(d, doc.documentElement.firstChild);
  };
  OF.addOverrideCT = function (doc, part, ct) {
    const o = doc.createElementNS('http://schemas.openxmlformats.org/package/2006/content-types', 'Override');
    o.setAttribute('PartName', '/' + part);
    o.setAttribute('ContentType', ct);
    doc.documentElement.appendChild(o);
  };
  OF.uniquePart = function (zip, dir, base, ext) {
    let i = 1;
    while (zip.file(`${dir}/${base}${i}.${ext}`)) i++;
    return `${dir}/${base}${i}.${ext}`;
  };
  // Inserta un fragmento XML (con sus espacios de nombres) en un documento
  OF.fragment = function (doc, xml) {
    const tmp = OF.parseXml(xml);
    return doc.importNode(tmp.documentElement, true);
  };

  /* ------------------------------- imágenes ------------------------------- */
  // Reduce fotos grandes (cámara) para que el documento no pese demasiado
  OF.prepareImage = async function (file) {
    const url = URL.createObjectURL(file);
    try {
      const img = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = () => rej(new Error('No se pudo leer la imagen')); i.src = url; });
      let w = img.naturalWidth, h = img.naturalHeight;
      let ext = /png/i.test(file.type) ? 'png' : /gif/i.test(file.type) ? 'gif' : 'jpeg';
      let bytes = new Uint8Array(await file.arrayBuffer());
      const max = 1800;
      if (w > max || h > max || bytes.length > 1.5e6 || !/png|jpe?g|gif/i.test(file.type)) {
        const k = Math.min(1, max / Math.max(w, h));
        const c = document.createElement('canvas');
        c.width = Math.round(w * k); c.height = Math.round(h * k);
        const g = c.getContext('2d');
        if (ext !== 'png') { g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height); }
        g.drawImage(img, 0, 0, c.width, c.height);
        const keepPng = ext === 'png' && bytes.length < 3e6;
        const blob = await new Promise((r) => c.toBlob(r, keepPng ? 'image/png' : 'image/jpeg', 0.86));
        bytes = new Uint8Array(await blob.arrayBuffer());
        ext = keepPng ? 'png' : 'jpeg';
        w = c.width; h = c.height;
      }
      return { bytes, ext, w, h, mime: 'image/' + ext, name: file.name || 'imagen.' + ext };
    } finally { URL.revokeObjectURL(url); }
  };
  OF.pickImage = function () {
    return new Promise((res) => {
      const inp = OF.el('input', { type: 'file', accept: 'image/*' });
      inp.addEventListener('change', () => res(inp.files[0] || null));
      inp.click();
    });
  };

  /* ------------------------- barra flotante de imagen ------------------------- */
  OF.imageBar = function (target, actions) {
    OF.closeImageBar();
    const bar = OF.el('div', { class: 'imgbar', role: 'toolbar', 'aria-label': 'Imagen' });
    for (const [label, fn, danger] of actions) {
      const b = OF.el('button', { class: danger ? 'danger' : '', type: 'button' }, label);
      b.addEventListener('mousedown', (e) => e.preventDefault());
      b.addEventListener('click', (e) => { e.stopPropagation(); fn(); });
      bar.append(b);
    }
    document.body.append(bar);
    const place = () => {
      if (!target.isConnected) return OF.closeImageBar();
      const r = target.getBoundingClientRect();
      const top = r.bottom + 8 + bar.offsetHeight > innerHeight ? Math.max(8, r.top - bar.offsetHeight - 8) : r.bottom + 8;
      bar.style.left = Math.max(8, Math.min(innerWidth - bar.offsetWidth - 8, r.left + r.width / 2 - bar.offsetWidth / 2)) + 'px';
      bar.style.top = top + 'px';
    };
    place();
    OF._imgBar = { bar, place };
    return bar;
  };
  OF.closeImageBar = function () { if (OF._imgBar) { OF._imgBar.bar.remove(); OF._imgBar = null; } };
  document.addEventListener('scroll', () => OF._imgBar && OF._imgBar.place(), true);

  /* --------------------------- utilidades de UI --------------------------- */
  OF.el = function (tag, attrs, ...children) {
    const e = document.createElement(tag);
    if (attrs) for (const k in attrs) {
      if (k === 'style' && typeof attrs[k] === 'object') Object.assign(e.style, attrs[k]);
      else if (k.startsWith('on')) e.addEventListener(k.slice(2), attrs[k]);
      else if (attrs[k] !== false && attrs[k] != null) e.setAttribute(k, attrs[k] === true ? '' : attrs[k]);
    }
    for (const c of children.flat()) if (c != null) e.append(c);
    return e;
  };

  // Pegar siempre como texto plano (evita meter formato ajeno al documento)
  /* --------------------------- pegar con formato --------------------------- */
  const toHex = (c) => {
    const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)(?:,\s*([\d.]+))?/.exec(c || '');
    if (!m || (m[4] != null && +m[4] === 0)) return null;
    return '#' + [m[1], m[2], m[3]].map((x) => (+x).toString(16).padStart(2, '0')).join('').toUpperCase();
  };
  OF.toHex = toHex;
  // Carga HTML del portapapeles en un marco aislado (sin scripts) para leer los estilos calculados
  OF.withHtml = function (html, fn) {
    const ifr = document.createElement('iframe');
    ifr.setAttribute('sandbox', 'allow-same-origin');
    ifr.style.cssText = 'position:fixed;left:-10000px;top:0;width:900px;height:700px;visibility:hidden;border:0';
    document.body.append(ifr);
    try {
      const d = ifr.contentDocument;
      d.open();
      // valores centinela: si el contenido copiado no define fuente/tamaño/color, no se inventan
      d.write('<style>html{font-family:__ninguna__;font-size:13.37px;color:rgb(1,2,3)}</style>' + html.replace(/<script[\s\S]*?<\/script>/gi, '').replace(/<img\b[^>]*>/gi, '').replace(/<link\b[^>]*>/gi, ''));
      d.close();
      return fn(d, ifr.contentWindow);
    } finally { ifr.remove(); }
  };
  const fmtOf = (el, win, body) => {
    const cs = win.getComputedStyle(el);
    let u = false;
    for (let a = el; a && a !== body; a = a.parentElement) if (/underline/.test(win.getComputedStyle(a).textDecorationLine)) { u = true; break; }
    const f = { b: +cs.fontWeight >= 600 || cs.fontWeight === 'bold', i: cs.fontStyle === 'italic', u };
    const col = toHex(cs.color); if (col && col !== '#010203') f.color = col;
    const px = parseFloat(cs.fontSize); if (px && Math.abs(px - 13.37) > 0.01) f.sz = Math.round(px * 0.75 * 2) / 2;
    const fam = (cs.fontFamily || '').split(',')[0].replace(/["']/g, '').trim();
    if (fam && fam !== '__ninguna__' && !/^(serif|sans-serif|monospace|system-ui|cursive|fantasy|-webkit)/i.test(fam)) f.font = fam;
    return f;
  };
  // HTML -> [{align, runs:[{text, fmt}]}]
  OF.htmlToParas = function (html) {
    return OF.withHtml(html, (d, win) => {
      const paras = [];
      let cur = null;
      const BLOCK = /^(P|DIV|H[1-6]|LI|TR|TABLE|TBODY|THEAD|UL|OL|BLOCKQUOTE|PRE|SECTION|ARTICLE|CENTER|DL|DT|DD)$/;
      const alignOf = (el) => { const ta = el && el !== d.body ? win.getComputedStyle(el).textAlign : ''; return /center/.test(ta) ? 'center' : /right|end/.test(ta) ? 'right' : /justify/.test(ta) ? 'justify' : /left|start/.test(ta) ? 'left' : null; };
      const open = (el) => { cur = { runs: [], align: alignOf(el) }; paras.push(cur); };
      const walk = (node, block) => {
        for (let c = node.firstChild; c; c = c.nextSibling) {
          if (c.nodeType === 3) {
            let t = c.nodeValue.replace(/[\r\n]+/g, ' ').replace(/[ \t]+/g, ' ').replace(/ /g, ' ');
            if (!t.trim() && (!cur || !cur.runs.length)) continue;
            if (!cur) open(block);
            if (!cur.runs.length) t = t.replace(/^ +/, '');
            if (t) cur.runs.push({ text: t, fmt: fmtOf(c.parentElement, win, d.body) });
          } else if (c.nodeType === 1) {
            const tag = c.tagName.toUpperCase();
            if (/^(STYLE|SCRIPT|HEAD|TITLE|META|XML|O:P)$/.test(tag) && tag !== 'O:P') continue;
            if (win.getComputedStyle(c).display === 'none') continue;
            if (tag === 'BR') { open(block); continue; }
            if (tag === 'TD' || tag === 'TH') { if (cur && cur.runs.length) cur.runs.push({ text: '\t', fmt: null }); else if (!cur) open(c); walk(c, c); continue; }
            if (BLOCK.test(tag)) {
              cur = null;
              if (tag === 'LI') { open(c); cur.runs.push({ text: '• ', fmt: null }); }
              walk(c, c);
              cur = null;
              continue;
            }
            walk(c, block);
          }
        }
      };
      walk(d.body, d.body);
      for (const p of paras) { const last = p.runs[p.runs.length - 1]; if (last && last.text) last.text = last.text.replace(/ +$/, ''); p.runs = p.runs.filter((r) => r.text); }
      while (paras.length && !paras[paras.length - 1].runs.length) paras.pop();
      while (paras.length && !paras[0].runs.length) paras.shift();
      return paras;
    });
  };
  OF.fmtCss = function (f) {
    if (!f) return {};
    const css = { fontWeight: f.b ? '700' : '400', fontStyle: f.i ? 'italic' : 'normal', textDecoration: f.u ? 'underline' : 'none' };
    if (f.color) css.color = f.color;
    if (f.sz) css.fontSize = f.sz + 'pt';
    if (f.font) css.fontFamily = `"${f.font}", Calibri, Carlito, Arial, sans-serif`;
    return css;
  };
  // Separa el <span> donde está el cursor para poder insertar fragmentos entre medio
  function caretSlot(range) {
    let node = range.startContainer, off = range.startOffset;
    if (node.nodeType === 3) {
      const span = node.parentElement.closest('span[data-r], span[data-fmt]');
      if (span && span.closest('[data-p]')) {
        const after = node.splitText(off);
        const right = span.cloneNode(false);
        let move = false;
        for (const ch of Array.from(span.childNodes)) { if (ch === after) move = true; if (move) right.append(ch); }
        span.after(right);
        const r = document.createRange();
        if (!right.textContent) { right.remove(); r.setStartAfter(span); }
        else if (!span.textContent) { span.remove(); r.setStartBefore(right); }
        else r.setStartAfter(span);
        r.collapse(true);
        return r;
      }
    }
    return range;
  }
  OF.insertParas = function (paras) {
    const sel = getSelection();
    let range = OF.currentRange();
    if (!range) return false;
    if (!range.collapsed) { range.deleteContents(); }
    paras.forEach((p, i) => {
      if (i > 0) { document.execCommand('insertParagraph'); range = sel.getRangeAt(0); }
      range = caretSlot(range);
      const frag = document.createDocumentFragment();
      let last = null;
      for (const run of p.runs) {
        const s = document.createElement('span');
        if (run.fmt) { s.dataset.fmt = JSON.stringify(run.fmt); Object.assign(s.style, OF.fmtCss(run.fmt)); }
        s.textContent = run.text;
        frag.append(s);
        last = s;
      }
      if (last) range.insertNode(frag);
      const host = (range.startContainer.nodeType === 1 ? range.startContainer : range.startContainer.parentElement).closest('[data-p]');
      if (host && p.align) {
        const jc = { left: 'left', center: 'center', right: 'right', justify: 'both' }[p.align];
        host.dataset.pfmt = JSON.stringify({ ...(host.dataset.pfmt ? JSON.parse(host.dataset.pfmt) : {}), jc });
        host.style.textAlign = p.align;
      }
      const nr = document.createRange();
      if (last) nr.setStartAfter(last); else nr.setStart(range.startContainer, range.startOffset);
      nr.collapse(true);
      sel.removeAllRanges();
      sel.addRange(nr);
    });
    OF.lastRange = sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
    return true;
  };

  // Áreas de texto editables (Word y PowerPoint): pegado con formato y atajos
  OF.plainPaste = function (root, doc) {
    root.addEventListener('paste', (e) => {
      e.preventDefault();
      const cd = e.clipboardData || window.clipboardData;
      const html = cd.getData('text/html');
      const text = cd.getData('text/plain');
      let paras = null;
      try { if (html) paras = OF.htmlToParas(html); } catch (err) { console.warn('Pegado con formato falló, se usa texto', err); }
      if (!paras || !paras.length) paras = text.replace(/\r\n?/g, '\n').split('\n').map((t) => ({ runs: t ? [{ text: t, fmt: null }] : [], align: null }));
      while (paras.length > 1 && !paras[paras.length - 1].runs.length) paras.pop();
      OF.insertParas(paras);
      root.dispatchEvent(new Event('input', { bubbles: true }));
    });
    root.addEventListener('drop', (e) => e.preventDefault());
    root.addEventListener('keydown', (e) => {
      // atajos de formato: se aplican al archivo (no como etiquetas del navegador)
      if ((e.ctrlKey || e.metaKey) && ['b', 'i', 'u', 'n', 'k', 's'].includes(e.key.toLowerCase()) && !(e.key.toLowerCase() === 's')) {
        e.preventDefault();
        const k = { b: 'b', n: 'b', i: 'i', k: 'i', u: 'u' }[e.key.toLowerCase()];
        if (doc && doc.toggle && k) doc.toggle(k);
      }
    });
  };
})();
