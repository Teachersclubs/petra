/* ===== المرفقات: صور، PDF، صور الماسح الضوئي ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const store = { letter: [], internal: [], thanks: [], circular: [], assign: [], claim: [] };
  let uid = 0, getType = () => 'letter', onChange = () => {};

  const readDataUrl = f => new Promise((res, rej) => { const fr = new FileReader(); fr.onload = () => res(fr.result); fr.onerror = rej; fr.readAsDataURL(f); });

  /** يحوّل أي صورة إلى JPEG بخلفية بيضاء وبحد أقصى للأبعاد */
  function toJpeg(url, maxSide) {
    return new Promise((res, rej) => {
      const im = new Image();
      im.onload = () => {
        const w0 = im.naturalWidth || 800, h0 = im.naturalHeight || 800;
        const k = Math.min(1, maxSide / Math.max(w0, h0));
        const c = document.createElement('canvas');
        c.width = Math.max(1, Math.round(w0 * k)); c.height = Math.max(1, Math.round(h0 * k));
        const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
        x.drawImage(im, 0, 0, c.width, c.height);
        res({ url: c.toDataURL('image/jpeg', 0.88), w: c.width, h: c.height });
      };
      im.onerror = () => rej(new Error('img'));
      im.src = url;
    });
  }

  async function addImages(type, files) {
    for (const f of files) {
      if (!/^image\//.test(f.type)) { g.toast('«' + f.name + '» ليس صورة'); continue; }
      try {
        const page = await toJpeg(await readDataUrl(f), 1800);
        store[type].push({ id: ++uid, name: f.name.replace(/\.[^.]+$/, ''), pages: [page] });
      } catch (e) { g.toast('تعذّر قراءة «' + f.name + '»'); }
    }
  }

  async function addPdf(type, file) {
    if (!g.pdfjsLib) { g.toast('مكتبة PDF غير محمّلة'); return; }
    g.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    const pdf = await g.pdfjsLib.getDocument({ data: await file.arrayBuffer() }).promise;
    const n = Math.min(pdf.numPages, 30);
    if (pdf.numPages > n) g.toast('أُضيفت أول ' + n + ' صفحة فقط من الملف');
    const pages = [];
    for (let i = 1; i <= n; i++) {
      const pg = await pdf.getPage(i);
      const vp0 = pg.getViewport({ scale: 1 });
      const vp = pg.getViewport({ scale: Math.min(2, 1700 / Math.max(vp0.width, vp0.height)) });
      const c = document.createElement('canvas'); c.width = Math.round(vp.width); c.height = Math.round(vp.height);
      const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
      await pg.render({ canvasContext: x, viewport: vp }).promise;
      pages.push({ url: c.toDataURL('image/jpeg', 0.86), w: c.width, h: c.height });
    }
    store[type].push({ id: ++uid, name: file.name.replace(/\.pdf$/i, ''), pages });
  }

  /** كل صفحات المرفقات بشكل مسطّح مع عنوان */
  function flat(type) {
    const out = [];
    store[type].forEach((a, i) => a.pages.forEach((p, j) => out.push(Object.assign({}, p, {
      cap: 'مرفق ' + (i + 1) + (a.name ? ': ' + a.name : '') + (a.pages.length > 1 ? ' (صفحة ' + (j + 1) + ' من ' + a.pages.length + ')' : '')
    }))));
    return out;
  }

  function renderList() {
    const type = getType(), ul = document.getElementById('attList'), list = store[type];
    document.getElementById('attCount').textContent = list.length ? '(' + list.length + ')' : '';
    ul.replaceChildren();
    list.forEach((a, i) => {
      const li = document.createElement('li');
      li.innerHTML = '<img alt=""><span class="att-name"></span><span class="att-ctl">' +
        '<button type="button" data-a="up" title="تقديم">↑</button><button type="button" data-a="down" title="تأخير">↓</button>' +
        '<button type="button" data-a="del" title="حذف">✕</button></span>';
      li.querySelector('img').src = a.pages[0].url;
      li.querySelector('.att-name').textContent = (i + 1) + '. ' + a.name + (a.pages.length > 1 ? ' — ' + a.pages.length + ' صفحة' : '');
      li.querySelectorAll('button').forEach(b => b.addEventListener('click', () => {
        const act = b.dataset.a;
        if (act === 'del') list.splice(i, 1);
        else if (act === 'up' && i > 0) [list[i - 1], list[i]] = [list[i], list[i - 1]];
        else if (act === 'down' && i < list.length - 1) [list[i + 1], list[i]] = [list[i], list[i + 1]];
        renderList(); onChange();
      }));
      ul.appendChild(li);
    });
  }

  async function handle(inp, fn) {
    const files = Array.from(inp.files); inp.value = '';
    if (!files.length) return;
    g.toast('جارٍ إضافة المرفقات…');
    try { await fn(getType(), files); } catch (e) { console.error(e); g.toast('تعذّرت قراءة الملف'); return; }
    renderList(); onChange(); g.toast('تمت إضافة المرفقات');
  }

  function init(gt, oc) {
    getType = gt; onChange = oc;
    document.getElementById('attImg').addEventListener('change', e => handle(e.target, addImages));
    document.getElementById('attScan').addEventListener('change', e => handle(e.target, addImages));
    document.getElementById('attPdf').addEventListener('change', e => handle(e.target, async (t, files) => {
      for (const f of files) {
        try { await addPdf(t, f); } catch (err) { console.error(err); g.toast('تعذّر فتح «' + f.name + '» (ملف PDF غير صالح أو محمي)'); }
      }
    }));
    renderList();
  }

  function setAll(o) { ['letter', 'internal', 'thanks', 'circular', 'assign', 'claim'].forEach(k => { store[k] = Array.isArray(o && o[k]) ? o[k] : []; uid = Math.max(uid, ...store[k].map(a => a.id || 0), 0); }); }
  g.App = Object.assign(App, { attach: { init, flat, renderList, store, setAll } });
})(window);
