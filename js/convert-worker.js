/* Convierte .xls/.ods/.csv a .xlsx en un hilo aparte: el convertidor no toca la página. */
self.onmessage = (e) => {
  try {
    importScripts('../vendor/xlsx.full.min.js');
    const { data, opts } = e.data;
    const wb = XLSX.read(data, opts);
    const out = XLSX.write(wb, { type: 'array', bookType: 'xlsx', compression: true });
    const buf = out.buffer || out;
    self.postMessage({ ok: true, buf }, [buf]);
  } catch (err) {
    self.postMessage({ ok: false, msg: String((err && err.message) || err) });
  }
};
