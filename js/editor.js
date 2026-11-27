/* ===== تبويب «تحميل كتاب للتعديل»: تفكيك PDF وتحريره (صفحات + نصوص + صور + أشكال) ===== */
(function (g) {
  'use strict';
  const App = g.App, doc = document;
  const $ = (s, r) => (r || doc).querySelector(s), $$ = (s, r) => Array.from((r || doc).querySelectorAll(s));
  const MM = 72 / 25.4; // نقطة لكل مم
  const E = { kind: null, name: '', srcs: [], pages: [], cur: 0, psel: new Set(), z: 1, fit: true, tool: 'select', sel: null, imgs: {}, undo: [], redo: [], word: null };
  let uid = 0; const nid = () => 'i' + (++uid);
  const S = () => App.getState();
  const toast = m => g.toast(m);
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const readBuf = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(f); });
  const readUrl = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  const blobUrl = (b) => URL.createObjectURL(b);
  const dl = (blob, name) => App.saveFile(blob, name);
  const baseName = () => (E.name || 'كتاب').replace(/\.(pdf|docx?|)$/i, '');
  const loadImg = src => new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = src; });

  /* ================= التراجع ================= */
  const snap = () => JSON.stringify(E.pages.map(p => Object.assign({}, p, { items: p.items })));
  function pushUndo() { E.undo.push(snap()); if (E.undo.length > 60) E.undo.shift(); E.redo = []; }
  function restore(s) { E.pages = JSON.parse(s); E.cur = Math.min(E.cur, E.pages.length - 1); E.sel = null; E.psel.clear(); renderAll(); }
  function undo() { if (!E.undo.length) return; E.redo.push(snap()); restore(E.undo.pop()); }
  function redo() { if (!E.redo.length) return; E.undo.push(snap()); restore(E.redo.pop()); }

  /* ================= تحميل الملفات ================= */
  const isPdf = f => /pdf$/i.test(f.type) || /\.pdf$/i.test(f.name);
  const isDocx = f => /\.docx$/i.test(f.name) || /officedocument\.wordprocessingml/.test(f.type);
  async function openFile(f) {
    if (!f) return;
    if (/\.doc$/i.test(f.name)) { toast('صيغة .doc القديمة غير مدعومة. احفظ الملف بصيغة .docx من Word ثم ارفعه'); return; }
    if (isPdf(f)) return App.run('جارٍ فتح ملف PDF…', () => loadPdf(f));
    if (isDocx(f)) return App.run('جارٍ فتح ملف Word…', () => App.editorWord.load(f, E));
    if (/^image\//.test(f.type)) { toast('الصور تُضاف من «إضافة صورة» داخل الملف المفتوح'); return; }
    toast('الملفات المدعومة: PDF أو Word (.docx)');
  }

  async function loadPdf(f, append) {
    if (!g.pdfjsLib) throw new Error('مكتبة PDF غير محمّلة');
    g.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js';
    const bytes = new Uint8Array(await readBuf(f));
    const pdf = await g.pdfjsLib.getDocument({ data: bytes.slice() }).promise;
    const si = E.srcs.length; E.srcs.push({ bytes, pdf, name: f.name });
    const np = [];
    for (let i = 0; i < pdf.numPages; i++) {
      const p = await pdf.getPage(i + 1), v = p.view, r0 = p.rotate || 0;
      np.push({ s: si, p: i, uw: v[2] - v[0], uh: v[3] - v[1], r0, rot: 0, items: [] });
    }
    if (append === undefined) { // ملف جديد
      E.kind = 'pdf'; E.name = f.name; E.srcs = [E.srcs[si]]; E.pages = np.map(x => Object.assign(x, { s: 0 }));
      E.cur = 0; E.psel.clear(); E.sel = null; E.undo = []; E.redo = []; E.word = null; E.fit = true;
    } else { pushUndo(); const at = append === 'end' ? E.pages.length : E.cur + 1; E.pages.splice(at, 0, ...np); E.cur = at; }
    showStage(); renderAll(); syncForm();
    toast(append === undefined ? 'تم فتح الملف: ' + np.length + ' صفحة' : 'تم دمج ' + np.length + ' صفحة');
  }

  /* ================= أبعاد وعرض الصفحة ================= */
  const rotOf = pg => ((pg.r0 + pg.rot) % 360 + 360) % 360;
  const dims = pg => (rotOf(pg) % 180 ? [pg.uh, pg.uw] : [pg.uw, pg.uh]);
  const bgCache = new Map();
  async function renderPageCanvas(pg, scale) {
    const key = pg.s + ':' + pg.p + ':' + rotOf(pg) + ':' + scale.toFixed(2);
    if (bgCache.has(key)) return bgCache.get(key);
    const c = doc.createElement('canvas');
    const [w, h] = dims(pg); c.width = Math.max(1, Math.round(w * scale)); c.height = Math.max(1, Math.round(h * scale));
    const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height);
    if (pg.s >= 0) {
      const p = await E.srcs[pg.s].pdf.getPage(pg.p + 1);
      const vp = p.getViewport({ scale, rotation: rotOf(pg) });
      await p.render({ canvasContext: x, viewport: vp }).promise;
    }
    bgCache.set(key, c); if (bgCache.size > 40) bgCache.delete(bgCache.keys().next().value);
    return c;
  }

  /* ================= عناصر الصفحة (نصوص/صور/أشكال) ================= */
  const FONTS = () => (App.FONTS || []).map(f => f[0]);
  function newItem(t, o) { return Object.assign({ id: nid(), t, x: 40, y: 40, w: 160, h: 40, rot: 0, op: 100 }, o); }
  const textDef = { txt: 'اكتب النص هنا', fs: 16, ff: 'Tajawal', col: '#000000', b: false, i: false, u: false, al: 'right', bg: '' };
  function itemCss(it, z) {
    let s = 'left:' + it.x * z + 'px;top:' + it.y * z + 'px;width:' + it.w * z + 'px;';
    if (it.t !== 'text') s += 'height:' + it.h * z + 'px;';
    if (it.rot) s += 'transform:rotate(' + it.rot + 'deg);';
    s += 'opacity:' + (it.op / 100) + ';';
    return s;
  }
  function itemEl(it, z) {
    const el = doc.createElement('div');
    el.className = 'ed-it ed-' + it.t; el.dataset.id = it.id;
    let css = itemCss(it, z);
    if (it.t === 'text') {
      css += 'font-size:' + it.fs * z + 'px;color:' + it.col + ';font-family:\'' + it.ff + '\',Tahoma,Arial,sans-serif;font-weight:' + (it.b ? 700 : 400) + ';font-style:' + (it.i ? 'italic' : 'normal') + ';text-decoration:' + (it.u ? 'underline' : 'none') + ';text-align:' + it.al + ';' + (it.bg ? 'background:' + it.bg + ';' : '') + (it.bd ? 'border:' + it.bd * z + 'px solid ' + it.col + ';padding:' + 4 * z + 'px;box-sizing:border-box;' : '');
      el.textContent = it.txt;
    } else if (it.t === 'img') {
      el.innerHTML = '<img alt="" draggable="false" src="' + E.imgs[it.img] + '">';
    } else if (it.t === 'rect' || it.t === 'white' || it.t === 'hl') {
      const fill = it.t === 'white' ? (it.fill || '#ffffff') : it.t === 'hl' ? (it.fill || '#ffe600') : it.fill;
      css += (fill ? 'background:' + fill + ';' : '') + (it.sw ? 'border:' + it.sw * z + 'px solid ' + (it.stroke || '#000') + ';' : '');
      if (it.t === 'hl') css += 'mix-blend-mode:multiply;';
    } else if (it.t === 'ell') {
      css += (it.fill ? 'background:' + it.fill + ';' : '') + 'border-radius:50%;' + (it.sw ? 'border:' + it.sw * z + 'px solid ' + (it.stroke || '#000') + ';' : '');
    } else if (it.t === 'line' || it.t === 'pen') {
      const pts = it.t === 'line' ? [[it.x1, it.y1], [it.x2, it.y2]] : it.pts;
      const xs = pts.map(p => p[0]), ys = pts.map(p => p[1]);
      const x0 = Math.min.apply(null, xs), y0 = Math.min.apply(null, ys), pad = (it.sw || 2) + 2;
      const w = Math.max.apply(null, xs) - x0, h = Math.max.apply(null, ys) - y0;
      css = 'left:' + (x0 - pad) * z + 'px;top:' + (y0 - pad) * z + 'px;width:' + (w + 2 * pad) * z + 'px;height:' + (h + 2 * pad) * z + 'px;opacity:' + it.op / 100 + ';';
      el.innerHTML = '<svg width="100%" height="100%" viewBox="' + (x0 - pad) + ' ' + (y0 - pad) + ' ' + (w + 2 * pad) + ' ' + (h + 2 * pad) + '" preserveAspectRatio="none" style="overflow:visible"><polyline fill="none" stroke="' + (it.stroke || '#d00000') + '" stroke-width="' + (it.sw || 2) + '" stroke-linecap="round" stroke-linejoin="round" points="' + pts.map(p => p.join(',')).join(' ') + '"/></svg>';
    }
    el.setAttribute('style', css);
    return el;
  }
  function boxOf(it) {
    if (it.t === 'line') return { x: Math.min(it.x1, it.x2), y: Math.min(it.y1, it.y2), w: Math.abs(it.x2 - it.x1), h: Math.abs(it.y2 - it.y1) };
    if (it.t === 'pen') { const xs = it.pts.map(p => p[0]), ys = it.pts.map(p => p[1]); const x = Math.min.apply(null, xs), y = Math.min.apply(null, ys); return { x, y, w: Math.max.apply(null, xs) - x, h: Math.max.apply(null, ys) - y }; }
    return { x: it.x, y: it.y, w: it.w, h: it.h };
  }
  const curPage = () => E.pages[E.cur];
  const findItem = id => { for (const p of E.pages) { const it = p.items.find(i => i.id === id); if (it) return { it, pg: p }; } return null; };
  const selItem = () => { const f = E.sel && findItem(E.sel); return f && f.it; };

  /* ================= عرض المسرح ================= */
  const stage = () => $('#editStage');
  function showStage() {
    doc.body.classList.add('ed-has-file');
    const st = stage(); if (!st) return;
    if (E.kind === 'pdf') {
      st.innerHTML =
        '<div class="ed-top" id="edTop">' +
        '<div class="ed-tools" id="edToolBtns">' +
        [['select', 'تحديد/تحريك', '<path d="M5 3l14 8-6 2-3 6z"/>'], ['edit', 'تعديل نص موجود', '<path d="M4 7h10M4 12h7M4 17h5M15 20l5-5-2-2-5 5v2z"/>'], ['text', 'نص', '<path d="M5 6V4h14v2M12 4v16M9 20h6"/>'], ['img', 'صورة', '<rect x="3" y="4" width="18" height="16" rx="2"/><circle cx="9" cy="10" r="2"/><path d="M21 16l-5-5-8 9"/>'],
          ['rect', 'مستطيل', '<rect x="4" y="6" width="16" height="12" rx="1"/>'], ['ell', 'دائرة', '<ellipse cx="12" cy="12" rx="8" ry="6"/>'], ['line', 'خط', '<path d="M5 19L19 5"/>'], ['pen', 'قلم حر', '<path d="M4 20l4-1 11-11-3-3L5 16z"/>'],
          ['hl', 'تظليل', '<path d="M4 20h16M7 16l8-8 3 3-8 8H7z"/>'], ['white', 'محو/تغطية', '<path d="M4 15l8-9 8 8-6 6H7z"/>']]
          .map(t => '<button type="button" class="ed-tb' + (t[0] === E.tool ? ' on' : '') + '" data-tool="' + t[0] + '" title="' + t[1] + '"><svg viewBox="0 0 24 24" width="18" height="18" fill="none" stroke="currentColor" stroke-width="1.9" stroke-linecap="round" stroke-linejoin="round">' + t[2] + '</svg><span>' + t[1] + '</span></button>').join('') +
        '</div><div class="ed-ctl">' +
        '<button type="button" class="ed-tb" id="edUndo" title="تراجع (Ctrl+Z)">↶</button><button type="button" class="ed-tb" id="edRedo" title="إعادة (Ctrl+Y)">↷</button>' +
        '<button type="button" class="ed-tb" id="edPrev" title="الصفحة السابقة">‹</button><output id="edNav">1 / 1</output><button type="button" class="ed-tb" id="edNext" title="الصفحة التالية">›</button><button type="button" class="ed-tb" id="edZo">−</button><output id="edZv">100%</output><button type="button" class="ed-tb" id="edZi">+</button><button type="button" class="ed-tb" id="edZf">ملاءمة</button>' +
        '</div></div>' +
        '<div class="ed-body"><div class="ed-thumbs" id="edThumbs"></div><div class="ed-view" id="edView"><div class="ed-page" id="edPage"><canvas id="edBg"></canvas><div class="ed-layer" id="edLayer"></div></div></div></div>';
      bindStage();
    } else if (E.kind === 'docx') App.editorWord.showStage(st);
    st.hidden = false;
    if (doc.body.classList.contains('mode-edit')) { $('#sheetHolder').hidden = true; }
  }
  const BLANK_OPTS = '<option value="a4p">A4 عمودي</option><option value="a4l">A4 أفقي</option><option value="a3p">A3 عمودي</option><option value="a3l">A3 أفقي</option><option value="letter">Letter</option>';
  const BSIZE = { a4p: [595.28, 841.89], a4l: [841.89, 595.28], a3p: [841.89, 1190.55], a3l: [1190.55, 841.89], letter: [612, 792] };
  function newBlank(kind, size) {
    const [w, h] = BSIZE[size] || BSIZE.a4p;
    E.name = kind === 'pdf' ? 'مستند جديد.pdf' : 'مستند جديد.docx'; E.srcs = []; E.undo = []; E.redo = []; E.sel = null; E.psel.clear(); E.cur = 0; E.fit = true;
    if (kind === 'pdf') { E.kind = 'pdf'; E.word = null; E.pages = [{ s: -1, p: -1, uw: w, uh: h, r0: 0, rot: 0, items: [] }]; showStage(); renderAll(); syncForm(); }
    else { E.kind = 'docx'; E.pages = []; E.word = { html: '<p><br></p>' }; show(true); }
    toast('تم إنشاء مستند فارغ');
  }
  function showEmpty() {
    const st = stage(); if (!st) return;
    st.innerHTML = '<div class="ed-empty" id="edDrop"><div class="ed-empty-box"><svg viewBox="0 0 24 24" width="54" height="54" fill="none" stroke="currentColor" stroke-width="1.5" stroke-linecap="round" stroke-linejoin="round"><path d="M14 3H7a2 2 0 0 0-2 2v14a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V8z"/><path d="M14 3v5h5M12 17v-6M9 13l3-3 3 3"/></svg><h3>اسحب ملف الكتاب هنا</h3><p>ملف PDF أو Word (.docx) — أو اضغط للاختيار من جهازك</p><button type="button" class="btn primary" id="edPickBig">اختيار ملف</button><div class="ed-or">أو ابدأ بصفحة فارغة</div><div class="ed-blank" data-blank><select class="ed-bsize">' + BLANK_OPTS + '</select><button type="button" class="btn" data-newblank="pdf">+ صفحة فارغة (PDF)</button><button type="button" class="btn" data-newblank="docx">+ مستند Word فارغ</button></div></div></div>';
    $('#edPickBig').addEventListener('click', () => $('#edFile').click());
    $$('[data-newblank]', st).forEach(b => b.addEventListener('click', () => newBlank(b.dataset.newblank, $('.ed-bsize', b.closest('[data-blank]')).value)));
    const di = App.editorWord && App.editorWord.draftInfo(); if (di) { const b = doc.createElement('button'); b.type = 'button'; b.className = 'btn'; b.style.marginTop = '10px'; b.textContent = 'استعادة آخر مسودة Word: ' + di.name; b.addEventListener('click', () => App.editorWord.restoreDraft(E)); $('.ed-empty-box', st).appendChild(b); }
    st.hidden = false; doc.body.classList.remove('ed-has-file');
  }

  async function renderAll() {
    if (E.kind !== 'pdf') return;
    renderThumbs(); await renderPage(); syncForm();
  }
  let rpTok = 0;
  async function renderPage() {
    const pg = curPage(); if (!pg) return;
    const tok = ++rpTok, view = $('#edView'); if (!view) return;
    const [w, h] = dims(pg);
    if (E.fit) E.z = Math.max(0.25, Math.min(3, (view.clientWidth - 36) / w));
    const z = E.z, dpr = Math.min(2, g.devicePixelRatio || 1);
    const pgEl = $('#edPage'); pgEl.style.width = w * z + 'px'; pgEl.style.height = h * z + 'px';
    $('#edZv').textContent = Math.round(z * 100 / 1.333) + '%';
    const c = await renderPageCanvas(pg, z * dpr); if (tok !== rpTok) return;
    const bg = $('#edBg'); bg.width = c.width; bg.height = c.height; bg.style.width = w * z + 'px'; bg.style.height = h * z + 'px';
    bg.getContext('2d').drawImage(c, 0, 0);
    drawLayer();
    $('#edNav') && ($('#edNav').textContent = (E.cur + 1) + ' / ' + E.pages.length);
  }
  function drawLayer() {
    const pg = curPage(), L = $('#edLayer'); if (!pg || !L) return;
    L.replaceChildren();
    pg.items.forEach(it => {
      const el = itemEl(it, E.z); L.appendChild(el);
      if (it.t === 'text') { const hh = el.offsetHeight / E.z; if (hh) it.h = hh; }
    });
    drawSel();
  }
  function drawSel() {
    $$('.ed-hd').forEach(n => n.remove()); $$('.ed-it.sel').forEach(n => n.classList.remove('sel'));
    const it = selItem(), L = $('#edLayer'); if (!it || !L) return;
    const el = L.querySelector('[data-id="' + it.id + '"]'); if (!el) return;
    el.classList.add('sel');
    const b = boxOf(it), z = E.z;
    if (it.t === 'line') {
      [['x1', 'y1'], ['x2', 'y2']].forEach(k => { const h = doc.createElement('i'); h.className = 'ed-hd ed-pt'; h.dataset.pt = k[0]; h.style.left = it[k[0]] * z + 'px'; h.style.top = it[k[1]] * z + 'px'; L.appendChild(h); });
    } else if (it.t !== 'pen') {
      const h = it.t === 'text' ? el.offsetHeight / z : it.h;
      const frame = doc.createElement('div'); frame.className = 'ed-hd ed-frame';
      frame.setAttribute('style', 'left:' + it.x * z + 'px;top:' + it.y * z + 'px;width:' + it.w * z + 'px;height:' + h * z + 'px;' + (it.rot ? 'transform:rotate(' + it.rot + 'deg);' : ''));
      ['nw', 'ne', 'sw', 'se'].concat(it.t === 'text' ? ['e', 'w'] : ['n', 's', 'e', 'w']).forEach(d => { const k = doc.createElement('i'); k.className = 'ed-hd ed-rs ed-rs-' + d; k.dataset.rs = d; frame.appendChild(k); });
      const ro = doc.createElement('i'); ro.className = 'ed-hd ed-ro'; ro.dataset.ro = '1'; ro.title = 'تدوير'; frame.appendChild(ro);
      L.appendChild(frame);
    }
  }
  function select(id) { E.sel = id || null; drawSel(); syncProps(); }

  /* ================= المصغّرات ================= */
  const thumbCache = new Map();
  async function thumbOf(pg) {
    const key = pg.s + ':' + pg.p + ':' + rotOf(pg);
    if (thumbCache.has(key)) return thumbCache.get(key);
    const [w] = dims(pg), c = await renderPageCanvas(pg, 150 / w);
    const u = c.toDataURL('image/jpeg', 0.7); thumbCache.set(key, u); return u;
  }
  function renderThumbs() {
    const box = $('#edThumbs'); if (!box) return;
    box.replaceChildren();
    E.pages.forEach((pg, i) => {
      const d = doc.createElement('div'); d.className = 'ed-th' + (i === E.cur ? ' cur' : '') + (E.psel.has(i) ? ' chk' : ''); d.draggable = true; d.dataset.i = i;
      const [w, h] = dims(pg);
      d.innerHTML = '<div class="ed-th-img" style="aspect-ratio:' + w + '/' + h + '"><img alt=""></div><label class="ed-th-n"><input type="checkbox"' + (E.psel.has(i) ? ' checked' : '') + '> ' + (i + 1) + (pg.items.length ? ' ✎' : '') + '</label>';
      thumbOf(pg).then(u => { const im = $('img', d); im && (im.src = u); });
      d.addEventListener('click', e => { if (e.target.tagName === 'INPUT') { e.target.checked ? E.psel.add(i) : E.psel.delete(i); d.classList.toggle('chk', e.target.checked); syncForm(); return; } goPage(i); });
      d.addEventListener('dragstart', e => { e.dataTransfer.setData('text/x-ed-page', String(i)); e.dataTransfer.effectAllowed = 'move'; });
      d.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('text/x-ed-page')) { e.preventDefault(); d.classList.add('over'); } });
      d.addEventListener('dragleave', () => d.classList.remove('over'));
      d.addEventListener('drop', e => { d.classList.remove('over'); const f = e.dataTransfer.getData('text/x-ed-page'); if (f === '') return; e.preventDefault(); e.stopPropagation(); movePage(+f, i); });
      box.appendChild(d);
    });
  }
  function goPage(i) { E.cur = Math.max(0, Math.min(E.pages.length - 1, i)); E.sel = null; E.fit = E.fit; $$('.ed-th').forEach((t, k) => t.classList.toggle('cur', k === E.cur)); renderPage(); const t = $('.ed-th.cur'); t && t.scrollIntoView({ block: 'nearest' }); syncProps(); }

  /* ================= عمليات الصفحات ================= */
  function movePage(from, to) { if (from === to) return; pushUndo(); const [p] = E.pages.splice(from, 1); E.pages.splice(to, 0, p); E.cur = to; E.psel.clear(); renderAll(); }
  const targets = () => (E.psel.size ? Array.from(E.psel).sort((a, b) => a - b) : [E.cur]);
  const pageOps = {
    up() { if (E.cur > 0) movePage(E.cur, E.cur - 1); },
    down() { if (E.cur < E.pages.length - 1) movePage(E.cur, E.cur + 1); },
    first() { movePage(E.cur, 0); }, last() { movePage(E.cur, E.pages.length - 1); },
    rotL() { pushUndo(); targets().forEach(i => { E.pages[i].rot = (E.pages[i].rot + 270) % 360; }); relayoutItems(); renderAll(); },
    rotR() { pushUndo(); targets().forEach(i => { E.pages[i].rot = (E.pages[i].rot + 90) % 360; }); relayoutItems(); renderAll(); },
    del() {
      const t = targets(); if (t.length >= E.pages.length) { toast('لا يمكن حذف كل الصفحات'); return; }
      if (!confirm('حذف ' + t.length + ' صفحة؟')) return; pushUndo();
      E.pages = E.pages.filter((_, i) => !t.includes(i)); E.cur = Math.min(E.cur, E.pages.length - 1); E.psel.clear(); E.sel = null; renderAll();
    },
    dup() { pushUndo(); const t = targets(); let off = 0; t.forEach(i => { const c = JSON.parse(JSON.stringify(E.pages[i + off])); c.items.forEach(x => { x.id = nid(); }); E.pages.splice(i + off + 1, 0, c); off++; }); E.psel.clear(); E.cur = t[0] + 1; renderAll(); },
    blank() { pushUndo(); const ref = curPage(); E.pages.splice(E.cur + 1, 0, { s: -1, p: -1, uw: ref ? ref.uw : 595.28, uh: ref ? ref.uh : 841.89, r0: 0, rot: ref ? ref.rot : 0, items: [] }); E.cur++; renderAll(); },
    rev() { pushUndo(); E.pages.reverse(); E.cur = E.pages.length - 1 - E.cur; E.psel.clear(); renderAll(); },
    all() { E.psel = new Set(E.pages.map((_, i) => i)); renderThumbs(); syncForm(); },
    none() { E.psel.clear(); renderThumbs(); syncForm(); },
    odd() { E.psel = new Set(E.pages.map((_, i) => i).filter(i => i % 2 === 0)); renderThumbs(); syncForm(); },
    even() { E.psel = new Set(E.pages.map((_, i) => i).filter(i => i % 2 === 1)); renderThumbs(); syncForm(); }
  };
  function relayoutItems() { E.pages.forEach(p => { /* العناصر تبقى بإحداثيات الصفحة المعروضة؛ نُبقيها داخل الحدود */ const [w, h] = dims(p); p.items.forEach(it => { if (it.x != null && it.x > w) it.x = Math.max(0, w - it.w); if (it.y != null && it.y > h) it.y = Math.max(0, h - it.h); }); }); }

  function parseRange(str, n) {
    const out = [];
    String(str || '').split(/[,،\s]+/).filter(Boolean).forEach(part => {
      const m = /^(\d+)(?:\s*[-–]\s*(\d+))?$/.exec(part.replace(/[٠-٩]/g, d => '٠١٢٣٤٥٦٧٨٩'.indexOf(d)));
      if (!m) return; let a = +m[1], b = m[2] ? +m[2] : a; if (a > b) [a, b] = [b, a];
      for (let i = a; i <= b; i++) if (i >= 1 && i <= n && !out.includes(i - 1)) out.push(i - 1);
    });
    return out;
  }

  /* ================= تفاعل المسرح: رسم/تحريك/تحجيم ================= */
  function bindStage() {
    const st = stage();
    $$('[data-tool]', st).forEach(b => b.addEventListener('click', () => setTool(b.dataset.tool)));
    $('#edPrev').addEventListener('click', () => goPage(E.cur - 1)); $('#edNext').addEventListener('click', () => goPage(E.cur + 1));
    $('#edUndo').addEventListener('click', undo); $('#edRedo').addEventListener('click', redo);
    $('#edZi').addEventListener('click', () => { E.fit = false; E.z = Math.min(4, E.z * 1.2); renderPage(); });
    $('#edZo').addEventListener('click', () => { E.fit = false; E.z = Math.max(0.2, E.z / 1.2); renderPage(); });
    $('#edZf').addEventListener('click', () => { E.fit = true; renderPage(); });
    const L = $('#edLayer');
    L.addEventListener('pointerdown', onDown);
    L.addEventListener('dblclick', onDbl);
    L.addEventListener('pointermove', hoverEdit);
    const view = $('#edView');
    view.addEventListener('wheel', e => { if (e.ctrlKey) { e.preventDefault(); E.fit = false; E.z = Math.max(0.2, Math.min(4, E.z * (e.deltaY < 0 ? 1.1 : 0.9))); renderPage(); } }, { passive: false });
    view.addEventListener('dragover', e => { if (e.dataTransfer.types.includes('Files')) e.preventDefault(); });
    view.addEventListener('drop', e => { if (e.dataTransfer.files && e.dataTransfer.files.length) { e.preventDefault(); handleDropFiles(e.dataTransfer.files, e); } });
  }
  function setTool(t) {
    E.tool = t; $$('[data-tool]', stage()).forEach(b => b.classList.toggle('on', b.dataset.tool === t));
    const L = $('#edLayer'); if (L) L.dataset.tool = t;
    if (t === 'img') $('#edImgFile').click();
  }
  const pt = e => { const r = $('#edLayer').getBoundingClientRect(); return [(e.clientX - r.left) / E.z, (e.clientY - r.top) / E.z]; };
  let drag = null;
  function onDown(e) {
    if (e.button) return;
    const L = $('#edLayer'), [px, py] = pt(e), pg = curPage();
    const hd = e.target.closest('.ed-hd');
    if (E.tool === 'edit' && !hd) { e.preventDefault(); editLineAt(px, py); return; }
    if (E.tool === 'select' || (hd && E.sel)) {
      const it = selItem();
      if (hd && it) {
        pushUndo(); L.setPointerCapture(e.pointerId);
        if (hd.dataset.rs) drag = { k: 'rs', d: hd.dataset.rs, it, x0: px, y0: py, o: Object.assign({}, it) };
        else if (hd.dataset.ro) { const b = boxOf(it); drag = { k: 'ro', it, cx: b.x + b.w / 2, cy: b.y + (it.t === 'text' ? it.h : b.h) / 2 }; }
        else if (hd.dataset.pt) drag = { k: 'pt', it, p: hd.dataset.pt };
        e.preventDefault(); return;
      }
      const el = e.target.closest('.ed-it');
      if (el) {
        const f = findItem(el.dataset.id); if (!f) return;
        if (el.isContentEditable) return;
        select(f.it.id); pushUndo(); L.setPointerCapture(e.pointerId);
        drag = { k: 'mv', it: f.it, x0: px, y0: py, o: JSON.parse(JSON.stringify(f.it)), moved: false };
        e.preventDefault();
      } else select(null);
      return;
    }
    // أدوات الإنشاء
    e.preventDefault(); pushUndo(); L.setPointerCapture(e.pointerId);
    if (E.tool === 'text') {
      const it = newItem('text', Object.assign({ x: px, y: py, w: 200 }, textDef)); pg.items.push(it); setTool('select'); drawLayer(); select(it.id);
      const el = $('[data-id="' + it.id + '"]', L); startInline(el, it, true); drag = null; return;
    }
    if (E.tool === 'pen') { const it = newItem('pen', { pts: [[px, py]], sw: 2, stroke: '#d00000' }); pg.items.push(it); drag = { k: 'pen', it }; drawLayer(); return; }
    if (E.tool === 'line') { const it = newItem('line', { x1: px, y1: py, x2: px, y2: py, sw: 2, stroke: '#d00000' }); pg.items.push(it); drag = { k: 'line', it }; drawLayer(); return; }
    const base = { rect: { fill: '', stroke: '#d00000', sw: 2 }, ell: { fill: '', stroke: '#d00000', sw: 2 }, white: { fill: '#ffffff', sw: 0 }, hl: { fill: '#ffe600', sw: 0 } }[E.tool];
    if (base) { const it = newItem(E.tool, Object.assign({ x: px, y: py, w: 1, h: 1 }, base)); pg.items.push(it); drag = { k: 'new', it, x0: px, y0: py }; drawLayer(); }
  }
  function onMove(e) {
    if (!drag) return;
    const [px, py] = pt(e), it = drag.it;
    if (drag.k === 'mv') {
      const dx = px - drag.x0, dy = py - drag.y0; drag.moved = drag.moved || Math.abs(dx) + Math.abs(dy) > 1;
      if (it.t === 'line') { it.x1 = drag.o.x1 + dx; it.x2 = drag.o.x2 + dx; it.y1 = drag.o.y1 + dy; it.y2 = drag.o.y2 + dy; }
      else if (it.t === 'pen') it.pts = drag.o.pts.map(p => [p[0] + dx, p[1] + dy]);
      else { it.x = drag.o.x + dx; it.y = drag.o.y + dy; }
    } else if (drag.k === 'new') {
      it.x = Math.min(px, drag.x0); it.y = Math.min(py, drag.y0); it.w = Math.abs(px - drag.x0); it.h = Math.abs(py - drag.y0);
    } else if (drag.k === 'pen') it.pts.push([px, py]);
    else if (drag.k === 'line') { it.x2 = px; it.y2 = py; if (e.shiftKey) { if (Math.abs(px - it.x1) > Math.abs(py - it.y1)) it.y2 = it.y1; else it.x2 = it.x1; } }
    else if (drag.k === 'pt') { it[drag.p] = px; it[drag.p.replace('x', 'y')] = py; }
    else if (drag.k === 'ro') { let a = Math.atan2(py - drag.cy, px - drag.cx) * 180 / Math.PI + 90; if (e.shiftKey) a = Math.round(a / 15) * 15; it.rot = Math.round(a * 10) / 10; }
    else if (drag.k === 'rs') {
      const o = drag.o, d = drag.d, a = -(o.rot || 0) * Math.PI / 180, cx = o.x + o.w / 2, cy = o.y + (o.t === 'text' ? o.h : o.h) / 2;
      let lx = (px - cx) * Math.cos(a) - (py - cy) * Math.sin(a), ly = (px - cx) * Math.sin(a) + (py - cy) * Math.cos(a);
      let x0 = o.x, y0 = o.y, x1 = o.x + o.w, y1 = o.y + o.h; const ch = o.t === 'text' ? o.h : o.h;
      const L2 = cx + lx - 0, T = cy + ly; // نقطة في فضاء غير مُدار
      const gx = cx + lx, gy = cy + ly;
      if (d.includes('w')) x0 = Math.min(gx, x1 - 8); if (d.includes('e')) x1 = Math.max(gx, x0 + 8);
      if (it.t !== 'text') { if (d.includes('n')) y0 = Math.min(gy, y1 - 8); if (d.includes('s')) y1 = Math.max(gy, y0 + 8); }
      if (o.t === 'img' && d.length === 2) { // الحفاظ على النسبة
        const ar = o.w / o.h; const nw = x1 - x0; y1 = y0 + nw / ar;
      }
      it.x = x0; it.y = y0; it.w = x1 - x0; if (it.t !== 'text') it.h = y1 - y0;
      void L2; void T; void ch;
    }
    liveUpdate(it);
  }
  function liveUpdate(it) {
    const L = $('#edLayer'); const old = $('[data-id="' + it.id + '"]', L); if (!old) return;
    const el = itemEl(it, E.z); old.replaceWith(el); if (it.t === 'text') it.h = el.offsetHeight / E.z;
    drawSel(); syncProps(true);
  }
  function onUp() {
    if (!drag) return;
    const d = drag; drag = null;
    if (d.k === 'mv' && !d.moved) { E.undo.pop(); }
    if (d.k === 'new' && (d.it.w < 3 || d.it.h < 3)) { const pg = curPage(); pg.items = pg.items.filter(x => x !== d.it); drawLayer(); return; }
    if (d.k === 'new' && d.it.t === 'white') { d.it.fill = sampleBg(d.it.x - 3, d.it.y + d.it.h / 2, d.it.x + d.it.w + 3, d.it.y + d.it.h / 2); drawLayer(); }
    if (d.k === 'new' || d.k === 'pen' || d.k === 'line') { setToolKeep(); select(d.it.id); }
    renderThumbsLight();
  }
  const setToolKeep = () => { /* تبقى الأداة نشطة للرسم المتكرر؛ يعود للتحديد بعد النص فقط */ };
  function renderThumbsLight() { const c = $('.ed-th.cur .ed-th-n'); if (c) c.lastChild.textContent = ' ' + (E.cur + 1) + (curPage().items.length ? ' ✎' : ''); }
  doc.addEventListener('pointermove', onMove);
  doc.addEventListener('pointerup', onUp);
  doc.addEventListener('pointercancel', onUp);

  function startInline(el, it, selectAll) {
    if (!el) return; el.contentEditable = 'true'; el.classList.add('editing'); el.focus();
    if (selectAll) { const r = doc.createRange(); r.selectNodeContents(el); const s = g.getSelection(); s.removeAllRanges(); s.addRange(r); }
    const done = () => { el.removeEventListener('blur', done); el.contentEditable = 'false'; el.classList.remove('editing'); const t = el.innerText.replace(/\n$/, ''); if (t !== it.txt) { pushUndo(); it.txt = t || ' '; } drawLayer(); syncProps(); };
    el.addEventListener('blur', done);
    el.addEventListener('keydown', ev => { if (ev.key === 'Escape') { ev.preventDefault(); el.blur(); } ev.stopPropagation(); });
  }
  function onDbl(e) {
    const el = e.target.closest('.ed-it.ed-text'); if (!el) return;
    const f = findItem(el.dataset.id); if (f) { select(f.it.id); startInline(el, f.it, false); }
  }

  /* ================= لوحة خصائص العنصر ================= */
  function syncProps(light) {
    const box = $('#edProps'); if (!box) return;
    const it = selItem();
    if (!it) { box.innerHTML = '<p class="hint">اضغط على أي عنصر مضاف (نص، صورة، شكل) لتعديل خصائصه، أو اختر أداة من الشريط أعلى المعاينة ثم اضغط على الصفحة.</p>'; box.dataset.cur = ''; return; }
    if (light && box.dataset.cur === it.id) { updateNums(it); return; }
    box.dataset.cur = it.id;
    const row = (l, inner) => '<label class="field"><span>' + l + '</span>' + inner + '</label>';
    const num = (k, l, min, max, step) => row(l, '<input type="number" data-p="' + k + '" min="' + min + '" max="' + max + '" step="' + (step || 1) + '">');
    const col = (k, l) => row(l, '<input type="color" data-p="' + k + '">');
    let h = '';
    if (it.t === 'text') {
      h += row('النص', '<textarea rows="3" data-p="txt"></textarea>');
      h += '<div class="row2">' + row('نوع الخط', '<select data-p="ff">' + FONTS().map(f => '<option>' + esc(f) + '</option>').join('') + '</select>') + num('fs', 'الحجم (نقطة)', 4, 400, 0.5) + '</div>';
      h += '<div class="row2">' + col('col', 'لون الخط') + row('خلفية النص', '<input type="color" data-p="bgc"><label class="check"><input type="checkbox" data-p="bgon"> <span>تفعيل</span></label>') + '</div>';
      h += '<div class="ed-fmt"><button type="button" data-fmt="b"><b>B</b></button><button type="button" data-fmt="i"><i>I</i></button><button type="button" data-fmt="u"><u>U</u></button><button type="button" data-al="right">يمين</button><button type="button" data-al="center">وسط</button><button type="button" data-al="left">يسار</button></div>';
    } else if (it.t === 'img') {
      h += '<p class="hint">اسحب زوايا الصورة لتغيير الحجم بنسبتها، واسحب المقبض العلوي للتدوير.</p>';
    } else if (it.t === 'line' || it.t === 'pen') {
      h += '<div class="row2">' + col('stroke', 'اللون') + num('sw', 'السماكة', 0.5, 40, 0.5) + '</div>';
    } else {
      h += '<div class="row2">' + row('التعبئة', '<input type="color" data-p="fill"><label class="check"><input type="checkbox" data-p="fillon"> <span>تفعيل</span></label>') + '</div>';
      if (it.t === 'rect' || it.t === 'ell') h += '<div class="row2">' + col('stroke', 'لون الإطار') + num('sw', 'سماكة الإطار', 0, 40, 0.5) + '</div>';
    }
    h += '<div class="row2">' + num('op', 'الشفافية %', 5, 100, 5) + (it.t === 'line' || it.t === 'pen' ? '' : num('rot', 'التدوير °', -360, 360, 1)) + '</div>';
    if (it.t !== 'line' && it.t !== 'pen') h += '<div class="row2">' + num('mx', 'س (مم)', -500, 800, 0.5) + num('my', 'ص (مم)', -500, 1200, 0.5) + '</div><div class="row2">' + num('mw', 'العرض (مم)', 1, 800, 0.5) + (it.t === 'text' ? '' : num('mh', 'الارتفاع (مم)', 1, 1200, 0.5)) + '</div>';
    h += '<div class="ed-btnrow"><button type="button" class="btn sm" data-act="front">للأمام</button><button type="button" class="btn sm" data-act="back">للخلف</button><button type="button" class="btn sm" data-act="cx">توسيط أفقي</button><button type="button" class="btn sm" data-act="cy">توسيط رأسي</button>' +
      '<button type="button" class="btn sm" data-act="dup">تكرار</button><button type="button" class="btn sm" data-act="allpages">نسخ لكل الصفحات</button><button type="button" class="btn sm danger" data-act="del">حذف</button></div>';
    box.innerHTML = h; updateNums(it);
  }
  function updateNums(it) {
    const box = $('#edProps'); if (!box) return;
    const set = (k, v) => { const n = $('[data-p="' + k + '"]', box); if (n && doc.activeElement !== n) { if (n.type === 'checkbox') n.checked = !!v; else n.value = v; } };
    ['txt', 'ff', 'fs', 'col', 'op', 'rot', 'stroke', 'sw'].forEach(k => set(k, it[k] == null ? '' : it[k]));
    set('bgc', it.bg || '#ffff00'); set('bgon', !!it.bg); set('fill', it.fill || (it.t === 'white' ? '#ffffff' : '#ffe600')); set('fillon', !!it.fill);
    if (it.t !== 'line' && it.t !== 'pen') { const b = boxOf(it); set('mx', +(b.x / MM).toFixed(1)); set('my', +(b.y / MM).toFixed(1)); set('mw', +(b.w / MM).toFixed(1)); set('mh', +(b.h / MM).toFixed(1)); }
  }
  function bindProps() {
    const box = $('#edProps');
    box.addEventListener('input', e => {
      const n = e.target.closest('[data-p]'); const it = selItem(); if (!n || !it) return;
      const k = n.dataset.p, v = n.type === 'checkbox' ? n.checked : n.type === 'number' ? parseFloat(n.value) : n.value;
      if (n.type === 'number' && isNaN(v)) return;
      if (!box._u) { pushUndo(); box._u = setTimeout(() => { box._u = null; }, 600); }
      if (k === 'mx') it.x = v * MM; else if (k === 'my') it.y = v * MM; else if (k === 'mw') it.w = v * MM; else if (k === 'mh') it.h = v * MM;
      else if (k === 'bgc') { if (it.bg) it.bg = v; } else if (k === 'bgon') it.bg = v ? ($('[data-p="bgc"]', box).value || '#ffff00') : '';
      else if (k === 'fillon') it.fill = v ? ($('[data-p="fill"]', box).value) : ''; else if (k === 'fill') { it.fill = v; const f = $('[data-p="fillon"]', box); if (f) f.checked = true; }
      else it[k] = v;
      liveUpdate(it);
    });
    box.addEventListener('click', e => {
      const it = selItem(); if (!it) return;
      const f = e.target.closest('[data-fmt]'), a = e.target.closest('[data-al]'), c = e.target.closest('[data-act]');
      if (f) { pushUndo(); it[f.dataset.fmt] = !it[f.dataset.fmt]; liveUpdate(it); }
      if (a) { pushUndo(); it.al = a.dataset.al; liveUpdate(it); }
      if (c) itemAct(c.dataset.act, it);
    });
  }
  function itemAct(act, it) {
    const pg = curPage(), [w, h] = dims(pg);
    pushUndo();
    if (act === 'del') { pg.items = pg.items.filter(x => x !== it); E.sel = null; drawLayer(); syncProps(); renderThumbsLight(); return; }
    if (act === 'front') { pg.items = pg.items.filter(x => x !== it).concat(it); }
    else if (act === 'back') { pg.items = [it].concat(pg.items.filter(x => x !== it)); }
    else if (act === 'dup') { const c = JSON.parse(JSON.stringify(it)); c.id = nid(); shift(c, 12, 12); pg.items.push(c); E.sel = c.id; }
    else if (act === 'cx') { const b = boxOf(it); shift(it, (w - b.w) / 2 - b.x, 0); }
    else if (act === 'cy') { const b = boxOf(it); shift(it, 0, (h - b.h) / 2 - b.y); }
    else if (act === 'allpages') {
      E.pages.forEach(p => { if (p === pg) return; const [pw, ph] = dims(p); const c = JSON.parse(JSON.stringify(it)); c.id = nid(); const b = boxOf(it); shift(c, (pw - w) / 2, 0); void ph; void b; p.items.push(c); });
      toast('نُسخ العنصر إلى كل الصفحات');
    }
    drawLayer(); syncProps(); renderThumbs();
  }
  function shift(it, dx, dy) {
    if (it.t === 'line') { it.x1 += dx; it.x2 += dx; it.y1 += dy; it.y2 += dy; }
    else if (it.t === 'pen') it.pts = it.pts.map(p => [p[0] + dx, p[1] + dy]);
    else { it.x += dx; it.y += dy; }
  }

  /* ================= إضافة صور وعناصر جاهزة ================= */
  async function addImageItem(url, opt) {
    const pg = curPage(); if (!pg) return;
    const im = await loadImg(url), [pw, ph] = dims(pg);
    const id = 'g' + (++uid); E.imgs[id] = url;
    let w = Math.min(pw * 0.4, im.naturalWidth * 0.75), h = w * im.naturalHeight / im.naturalWidth;
    if (opt && opt.w) { w = opt.w; h = w * im.naturalHeight / im.naturalWidth; }
    pushUndo();
    const it = newItem('img', { img: id, x: (pw - w) / 2, y: (ph - h) / 2, w, h });
    pg.items.push(it); setTool('select'); drawLayer(); select(it.id); renderThumbsLight();
  }
  async function handleDropFiles(files, ev) {
    for (const f of Array.from(files)) {
      if (/^image\//.test(f.type)) await addImageItem(await readUrl(f));
      else if (isPdf(f)) await loadPdf(f, ev && ev.shiftKey ? 'end' : 'end');
    }
  }
  async function imageAsPages(url, name) {
    const im = await loadImg(url); const land = im.naturalWidth > im.naturalHeight;
    let uw = land ? 841.89 : 595.28, uh = land ? 595.28 : 841.89;
    const k = Math.min(uw / im.naturalWidth, uh / im.naturalHeight); const w = im.naturalWidth * k, h = im.naturalHeight * k;
    const id = 'g' + (++uid); E.imgs[id] = url;
    return { s: -1, p: -1, uw, uh, r0: 0, rot: 0, items: [newItem('img', { img: id, x: (uw - w) / 2, y: (uh - h) / 2, w, h, name })] };
  }
  async function pullAttachments() {
    const st = App.attach && App.attach.store && App.attach.store.edit; if (!st || !st.length) return;
    const list = st.splice(0, st.length); App.attach.renderList && App.attach.renderList();
    if (E.kind === null) { // لا ملف: ابنِ ملفاً جديداً من الصور
      E.kind = 'pdf'; E.name = 'مسح ضوئي.pdf'; E.srcs = []; E.pages = []; E.cur = 0; E.undo = []; showStage();
    }
    if (E.kind !== 'pdf') { toast('أضف الصور داخل ملف PDF مفتوح'); return; }
    pushUndo();
    const asPage = ($('input[name="edAttMode"]:checked') || {}).value !== 'item';
    for (const a of list) for (const p of a.pages) {
      if (asPage || !curPage()) { const at = E.pages.length ? E.cur + 1 : 0; E.pages.splice(at, 0, await imageAsPages(p.url, a.name)); E.cur = at; }
      else await addImageItem(p.url);
    }
    renderAll(); toast('أُضيفت الصور');
  }

  /* ================= النصوص ↔ الإضافات السريعة ================= */
  function addTextAll(opt) {
    pushUndo();
    E.pages.forEach((pg, i) => {
      const [w, h] = dims(pg);
      const txt = opt.fmt.replace('{n}', i + 1).replace('{N}', E.pages.length);
      const it = newItem('text', Object.assign({}, textDef, { txt, fs: opt.fs, col: opt.col, ff: opt.ff, w: opt.w || 200, al: 'center', op: opt.op || 100, rot: opt.rot || 0 }));
      it.x = (w - it.w) / 2; it.y = opt.top != null ? opt.top : h - 30 - opt.fs; if (opt.pos === 'top') it.y = 24; pg.items.push(it);
    });
    renderAll();
  }
  function addWatermark(txt, fs, col, op) {
    pushUndo();
    E.pages.forEach(pg => {
      const [w, h] = dims(pg); const it = newItem('text', Object.assign({}, textDef, { txt, fs, col, b: true, al: 'center', w: w * 0.9, op, rot: -35 }));
      it.x = (w - it.w) / 2; it.y = h / 2 - fs; pg.items.push(it);
    });
    renderAll();
  }

  /* ================= التصدير: بناء PDF بـ pdf-lib ================= */
  async function rasterItem(it) {
    // يعيد {url, x, y, w, h} بالنقاط (المربع المحيط بعد التدوير) مع صورة PNG شفافة
    const K = 3; let b = boxOf(it), pad = 0;
    if (it.t === 'line' || it.t === 'pen') pad = (it.sw || 2) + 2;
    let bw = b.w + 2 * pad, bh = b.h + 2 * pad, bx = b.x - pad, by = b.y - pad;
    const c0 = doc.createElement('canvas');
    if (it.t === 'text') {
      const host = doc.createElement('div'); host.style.cssText = 'position:fixed;left:-9999px;top:0;';
      const el = itemEl(Object.assign({}, it, { x: 0, y: 0, rot: 0, op: 100 }), 1); el.style.left = '0'; el.style.top = '0'; host.appendChild(el); doc.body.appendChild(host);
      bh = el.offsetHeight; bw = it.w;
      const cv = await g.html2canvas(el, { backgroundColor: null, scale: K, useCORS: true, logging: false });
      host.remove(); c0.width = cv.width; c0.height = cv.height; c0.getContext('2d').drawImage(cv, 0, 0);
    } else {
      c0.width = Math.max(1, Math.round(bw * K)); c0.height = Math.max(1, Math.round(bh * K)); const x = c0.getContext('2d'); x.scale(K, K);
      if (it.t === 'img') { const im = await loadImg(E.imgs[it.img]); x.drawImage(im, 0, 0, bw, bh); }
      else if (it.t === 'rect' || it.t === 'white' || it.t === 'hl') {
        const fill = it.t === 'white' ? (it.fill || '#fff') : it.t === 'hl' ? (it.fill || '#ffe600') : it.fill;
        if (fill) { x.fillStyle = fill; x.fillRect(0, 0, bw, bh); }
        if (it.sw) { x.strokeStyle = it.stroke || '#000'; x.lineWidth = it.sw; x.strokeRect(it.sw / 2, it.sw / 2, bw - it.sw, bh - it.sw); }
      } else if (it.t === 'ell') {
        x.beginPath(); x.ellipse(bw / 2, bh / 2, Math.max(1, bw / 2 - (it.sw || 0) / 2), Math.max(1, bh / 2 - (it.sw || 0) / 2), 0, 0, Math.PI * 2);
        if (it.fill) { x.fillStyle = it.fill; x.fill(); } if (it.sw) { x.strokeStyle = it.stroke || '#000'; x.lineWidth = it.sw; x.stroke(); }
      } else {
        const pts = it.t === 'line' ? [[it.x1, it.y1], [it.x2, it.y2]] : it.pts;
        x.strokeStyle = it.stroke || '#d00000'; x.lineWidth = it.sw || 2; x.lineCap = 'round'; x.lineJoin = 'round'; x.beginPath();
        pts.forEach((p, i) => { const px = p[0] - bx, py = p[1] - by; i ? x.lineTo(px, py) : x.moveTo(px, py); }); x.stroke();
      }
    }
    // تدوير + شفافية
    const rot = (it.rot || 0) * Math.PI / 180, cw = Math.abs(bw * Math.cos(rot)) + Math.abs(bh * Math.sin(rot)), ch = Math.abs(bw * Math.sin(rot)) + Math.abs(bh * Math.cos(rot));
    const c1 = doc.createElement('canvas'); c1.width = Math.max(1, Math.round(cw * K)); c1.height = Math.max(1, Math.round(ch * K));
    const x1 = c1.getContext('2d'); x1.translate(c1.width / 2, c1.height / 2); x1.rotate(rot); x1.drawImage(c0, -c0.width / 2, -c0.height / 2);
    const mx = bx + bw / 2, my = by + bh / 2;
    return { url: c1.toDataURL('image/png'), x: mx - cw / 2, y: my - ch / 2, w: cw, h: ch, op: it.op / 100, hl: it.t === 'hl' };
  }
  const bytesOf = url => { const b = atob(url.split(',')[1]), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };

  async function buildPdf(idx, setMsg) {
    const PL = g.PDFLib; if (!PL) throw new Error('مكتبة تحرير PDF غير محمّلة');
    const out = await PL.PDFDocument.create(), srcDocs = new Map(), list = idx || E.pages.map((_, i) => i);
    let n = 0;
    for (const i of list) {
      const pg = E.pages[i]; n++; setMsg && setMsg('جارٍ تجهيز الصفحة ' + n + ' من ' + list.length + '…');
      let page, ox = 0, oy = 0, W = pg.uw, H = pg.uh;
      if (pg.s >= 0) {
        let sd = srcDocs.get(pg.s); if (!sd) { sd = await PL.PDFDocument.load(E.srcs[pg.s].bytes, { ignoreEncryption: true }); srcDocs.set(pg.s, sd); }
        const [cp] = await out.copyPages(sd, [pg.p]); page = out.addPage(cp);
        const cb = page.getCropBox(); ox = cb.x; oy = cb.y; W = cb.width; H = cb.height;
      } else page = out.addPage([pg.uw, pg.uh]);
      const R = rotOf(pg); page.setRotation(PL.degrees(R));
      for (const it of pg.items) {
        const r = await rasterItem(it);
        const png = await out.embedPng(bytesOf(r.url));
        const dx = r.x, dy = r.y + r.h; // الزاوية السفلية اليسرى المعروضة
        let ux, uy;
        if (R === 0) { ux = dx; uy = H - dy; } else if (R === 90) { ux = dy; uy = dx; } else if (R === 180) { ux = W - dx; uy = dy; } else { ux = W - dy; uy = H - dx; }
        page.drawImage(png, { x: ox + ux, y: oy + uy, width: r.w, height: r.h, rotate: PL.degrees(R), opacity: r.op, blendMode: r.hl ? PL.BlendMode.Multiply : undefined });
      }
    }
    const mt = ($('#edMetaT') || {}).value, ma = ($('#edMetaA') || {}).value; out.setTitle(mt || baseName()); if (ma) out.setAuthor(ma); out.setProducer('محرر الكتب الرسمية');
    return new Blob([await out.save()], { type: 'application/pdf' });
  }

  /* ================= تفكيك وتحويل ================= */
  /* ---------- استخراج النص بدقة (عربي) ---------- */
  const AR = /[؀-ۿݐ-ݿﭐ-﷿ﹰ-﻿]/;
  const fixAr = s => s.replace(/ا([أإآ])ل/g, 'ال$1').replace(/ـ/g, '').replace(/[‎‏‪-‮]/g, '');
  const lineCache = new Map();
  async function fillLines(res, p, rot) {
      const vp = p.getViewport({ scale: 1, rotation: rot }), tc = await p.getTextContent();
      const boldOf = fn => { try { const f = p.commonObjs.get(fn); return /bold|black|heavy|semibold|demi/i.test((f && f.name) || ''); } catch (e) { return false; } };
      const bc = {}; res.items = tc.items.filter(t => t.str && t.str.trim() && !(t.transform[1] && Math.abs(t.transform[1]) > Math.abs(t.transform[0]) * 0.8)).map(t => {
        const [a, b2, , , e, f] = t.transform, size = Math.hypot(a, b2) || t.height || 12;
        const p1 = vp.convertToViewportPoint(e, f), p2 = vp.convertToViewportPoint(e + t.width, f + size);
        const str = fixAr(t.str), bold = bc[t.fontName] != null ? bc[t.fontName] : (bc[t.fontName] = boldOf(t.fontName));
        return { str, x0: Math.min(p1[0], p2[0]), x1: Math.max(p1[0], p2[0]), y0: Math.min(p1[1], p2[1]), y1: Math.max(p1[1], p2[1]), size, bold, ar: AR.test(str) };
      });
      const its = res.items.slice().sort((a, b) => (a.y0 + a.y1) - (b.y0 + b.y1)), groups = [];
      its.forEach(it => { const cy = (it.y0 + it.y1) / 2, g0 = groups[groups.length - 1]; if (g0 && Math.abs(cy - g0.cy) < Math.max(it.size, g0.size) * 0.45) { g0.items.push(it); g0.cy = (g0.cy * (g0.items.length - 1) + cy) / g0.items.length; g0.size = Math.max(g0.size, it.size); } else groups.push({ cy, size: it.size, items: [it] }); });
      groups.forEach(g0 => {
        let asc = g0.items.slice().sort((a, b) => a.x0 - b.x0); const rtl = asc.some(i => i.ar) && asc.filter(i => i.ar).length * 2 >= asc.filter(i => /[A-Za-z]/.test(i.str)).length;
        let seq = asc;
        if (rtl) { asc = asc.filter(i => !(!i.ar && /^[A-Za-z]$/.test(i.str) && (i.x1 - i.x0) < i.size * 0.3)); }
        if (rtl) { seq = []; let blk = []; for (let i = asc.length - 1; i >= 0; i--) { const it = asc[i]; if (!it.ar) blk.unshift(it); else { seq.push.apply(seq, blk); blk = []; seq.push(it); } } seq.push.apply(seq, blk); }
        let text = '';
        seq.forEach((it, i) => { if (i) { const pv = seq[i - 1], gap = Math.max(pv.x0, it.x0) - Math.min(pv.x1, it.x1); if (gap > it.size * (it.str === '\u0627' && pv.ar ? 0.6 : 0.12)) text += ' '; } text += it.str.replace(/\s+/g, ' ').trim(); });
        const bs = seq.filter(i => i.bold).length;
        res.lines.push({ text: text.trim(), seq, rtl, x0: Math.min.apply(null, asc.map(i => i.x0)), x1: Math.max.apply(null, asc.map(i => i.x1)), y0: Math.min.apply(null, asc.map(i => i.y0)), y1: Math.max.apply(null, asc.map(i => i.y1)), size: g0.size, bold: bs * 2 > seq.length });
      });
  }
  async function pageLines(pg) {
    const key = pg.s + ':' + pg.p + ':' + rotOf(pg);
    if (lineCache.has(key)) return lineCache.get(key);
    const res = { lines: [], items: [], W: dims(pg)[0], H: dims(pg)[1] };
    if (pg.s >= 0) await fillLines(res, await E.srcs[pg.s].pdf.getPage(pg.p + 1), rotOf(pg));
    lineCache.set(key, res); if (lineCache.size > 60) lineCache.delete(lineCache.keys().next().value);
    return res;
  }
  async function textOfPage(pg) { return (await pageLines(pg)).lines.map(l => l.text); }
  async function openPdfData(f) { g.pdfjsLib.GlobalWorkerOptions.workerSrc = 'vendor/pdf.worker.min.js'; return g.pdfjsLib.getDocument({ data: new Uint8Array(await readBuf(f)) }).promise; }
  async function pdfToImages(f, scale, setMsg) {
    const pdf = await openPdfData(f), out = [];
    for (let i = 1; i <= pdf.numPages; i++) { setMsg && setMsg('تحويل الصفحة ' + i + ' من ' + pdf.numPages + '…'); const p = await pdf.getPage(i), vp = p.getViewport({ scale: scale || 1.6 }), c = doc.createElement('canvas'); c.width = vp.width; c.height = vp.height; const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); await p.render({ canvasContext: x, viewport: vp }).promise; out.push(c.toDataURL('image/jpeg', 0.9)); }
    return out;
  }
  async function pdfTextHTML(f, setMsg) {
    const pdf = await openPdfData(f); let h = '';
    for (let i = 1; i <= pdf.numPages; i++) {
      setMsg && setMsg('استخراج نص الصفحة ' + i + ' من ' + pdf.numPages + '…');
      const p = await pdf.getPage(i), v = p.getViewport({ scale: 1 }), res = { lines: [], items: [], W: v.width, H: v.height };
      await fillLines(res, p, 0);
      const paras = paragraphsOf(res);
      if (!paras.length) h += '<p dir="rtl"><i>[صفحة ' + i + ': لا يوجد نص قابل للاستخراج — صفحة ممسوحة ضوئياً]</i></p>';
      paras.forEach(pa => { const txt = pa.lines.map(l => l.text).join(' '), b = pa.lines[0].bold; h += '<p dir="' + (pa.rtl ? 'rtl' : 'ltr') + '" style="text-align:' + (pa.center ? 'center' : pa.rtl ? 'right' : 'left') + '">' + (b ? '<b>' : '') + esc(txt) + (b ? '</b>' : '') + '</p>'; });
      if (i < pdf.numPages) h += '<hr data-pb="1" style="border:0;border-top:2px dashed #999"><p><br></p>';
    }
    return h;
  }
  function paragraphsOf(L) {
    const W = L.W, lines = L.lines.slice().sort((a, b) => a.y0 - b.y0), maxW = Math.max.apply(null, lines.map(l => l.x1 - l.x0).concat(1));
    const paras = []; let cur = null;
    const center = l => Math.abs((l.x0 + l.x1) / 2 - W / 2) < W * 0.03 && (l.x1 - l.x0) < W * 0.7 && l.x0 > W * 0.12;
    lines.forEach(l => {
      const c = center(l);
      if (cur) {
        const pv = cur.last, gap = l.y0 - pv.y1, same = Math.abs(l.size - pv.size) < 1.2 && l.bold === pv.bold && c === cur.center && l.rtl === pv.rtl;
        const edge = l.rtl ? Math.abs(l.x1 - pv.x1) < W * 0.04 : Math.abs(l.x0 - pv.x0) < W * 0.04;
        if (same && !c && edge && gap < l.size * 0.55 && (pv.x1 - pv.x0) > maxW * 0.8) { cur.lines.push(l); cur.last = l; return; }
      }
      cur = { lines: [l], last: l, center: c, before: paras.length ? Math.max(0, Math.round((l.y0 - (paras[paras.length - 1].last.y1)) * 20)) : Math.round(l.y0 * 10), rtl: l.rtl }; paras.push(cur);
    });
    return paras;
  }
  async function toWordDoc(setMsg, mode) {
    const D = g.docx, secs = [];
    const jpgOf = async (pg, px) => { const b = await pageToBlob(pg, 'jpg', px, 0.85); return new Uint8Array(await b.arrayBuffer()); };
    if (mode === 'visual') {
      for (let i = 0; i < E.pages.length; i++) {
        setMsg && setMsg('تحويل الصفحة ' + (i + 1) + ' من ' + E.pages.length + '…');
        const pg = E.pages[i], [w, h] = dims(pg), data = await jpgOf(pg, 1700);
        secs.push({ properties: { page: { size: { width: Math.round(w * 20), height: Math.round(h * 20) }, margin: { top: 0, bottom: 0, left: 0, right: 0 } } }, children: [new D.Paragraph({ spacing: { before: 0, after: 0, line: 240 }, children: [new D.ImageRun({ type: 'jpg', data, transformation: { width: Math.round(w * 4 / 3), height: Math.round(h * 4 / 3) - 2 } })] })] });
      }
      return D.Packer.toBlob(new D.Document({ sections: secs }));
    }
    const kids = [];
    for (let i = 0; i < E.pages.length; i++) {
      setMsg && setMsg('استخراج نص الصفحة ' + (i + 1) + ' من ' + E.pages.length + '…');
      const pg = E.pages[i], L = await pageLines(pg), paras = paragraphsOf(L);
      if (!paras.length) {
        const data = await jpgOf(pg, 1500), [w, h] = dims(pg), pw = 600;
        kids.push(new D.Paragraph({ alignment: D.AlignmentType.CENTER, children: [new D.ImageRun({ type: 'jpg', data, transformation: { width: pw, height: Math.round(pw * h / w) } })] }));
      }
      paras.forEach((pa, pi) => {
        const runs = []; const rtl = pa.rtl;
        pa.lines.forEach((ln, li) => {
          if (li) runs.push(new D.TextRun({ text: ' ' }));
          const seq = ln.seq; let cur = null;
          const flush = () => { if (cur) { runs.push(new D.TextRun({ text: cur.t, bold: cur.b, size: cur.sz, font: 'Arial', rightToLeft: cur.ar || undefined })); cur = null; } };
          const parts = ln.text.split(' '), sz = Math.max(16, Math.round(ln.size * 2));
          // نبني الأجزاء من النص المرتّب منطقياً بنمط الخط السائد للسطر
          cur = { t: ln.text, b: ln.bold, sz, ar: AR.test(ln.text) }; void parts; flush();
        });
        kids.push(new D.Paragraph({ bidirectional: rtl, alignment: pa.center ? D.AlignmentType.CENTER : D.AlignmentType.START, spacing: { before: Math.min(pi === 0 ? 0 : pa.before, 720), after: 40, line: 276 }, children: runs }));
      });
      if (i < E.pages.length - 1) kids.push(new D.Paragraph({ children: [new D.PageBreak()] }));
    }
    const pg0 = E.pages[0], [w0, h0] = dims(pg0);
    return D.Packer.toBlob(new D.Document({ styles: { default: { document: { run: { font: 'Arial', size: 26 } } } }, sections: [{ properties: { page: { size: { width: Math.round(w0 * 20), height: Math.round(h0 * 20) }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children: kids }] }));
  }
  async function pagesToImages(idx, type, setMsg) {
    const out = [];
    for (const i of idx) { setMsg && setMsg('تحويل الصفحة ' + (i + 1) + ' إلى صورة…'); out.push({ i, blob: await pageToBlob(E.pages[i], type) }); }
    return out;
  }
  async function pageToBlob(pg, type, px, q) {
    const [w] = dims(pg), bg = await renderPageCanvas(pg, (px || 2200) / w), c = doc.createElement('canvas'); c.width = bg.width; c.height = bg.height;
    const x = c.getContext('2d'); x.drawImage(bg, 0, 0); const k = c.width / w;
    for (const it of pg.items) { const r = await rasterItem(it); const im = await loadImg(r.url); x.globalAlpha = r.op; x.drawImage(im, r.x * k, r.y * k, r.w * k, r.h * k); x.globalAlpha = 1; }
    return new Promise(res => c.toBlob(res, type === 'jpg' ? 'image/jpeg' : 'image/png', q || 0.92));
  }
  async function zipBlobs(files) { const Z = new g.JSZip(); files.forEach(f => Z.file(f.name, f.blob)); return Z.generateAsync({ type: 'blob' }); }

  const tools = {
    async extractSel() { const t = targets(); await App.run('جارٍ استخراج الصفحات…', async m => { await dl(await buildPdf(t, m), baseName() + '-استخراج.pdf'); }); },
    async extractRange() { const v = $('#edRange').value, idx = parseRange(v, E.pages.length); if (!idx.length) { toast('اكتب مدى صحيحاً مثل: 1-3,5'); return; } await App.run('جارٍ الاستخراج…', async m => { await dl(await buildPdf(idx, m), baseName() + '-صفحات.pdf'); }); },
    async splitAll() { await App.run('جارٍ تقسيم الملف…', async m => { const f = []; for (let i = 0; i < E.pages.length; i++) { m('الصفحة ' + (i + 1) + ' من ' + E.pages.length); f.push({ name: baseName() + '-صفحة-' + (i + 1) + '.pdf', blob: await buildPdf([i]) }); } await dl(await zipBlobs(f), baseName() + '-صفحات-منفصلة.zip'); }); },
    async splitEvery() { const n = Math.max(1, parseInt($('#edEvery').value, 10) || 1); await App.run('جارٍ تقسيم الملف…', async m => { const f = []; for (let i = 0, k = 1; i < E.pages.length; i += n, k++) { m('الجزء ' + k); const idx = []; for (let j = i; j < Math.min(E.pages.length, i + n); j++) idx.push(j); f.push({ name: baseName() + '-جزء-' + k + '.pdf', blob: await buildPdf(idx) }); } await dl(await zipBlobs(f), baseName() + '-أجزاء.zip'); }); },
    async toWord() { await App.run('جارٍ تحويل PDF إلى Word…', async m => { await dl(await toWordDoc(m), baseName() + '.docx'); toast('تم التحويل — النص فقط دون التنسيق الأصلي'); }); },
    async toWordImg() { await App.run('جارٍ تحويل PDF إلى Word…', async m => { await dl(await toWordDoc(m, 'visual'), baseName() + '-مطابق.docx'); toast('تم — كل صفحة صورة مطابقة للأصل'); }); },
    async imgPage() { await App.run('جارٍ التحويل…', async () => { await dl(await pageToBlob(curPage(), 'png'), baseName() + '-صفحة-' + (E.cur + 1) + '.png'); }); },
    async imgAll() { await App.run('جارٍ تحويل الصفحات إلى صور…', async m => { const r = await pagesToImages(targets().length > 1 ? targets() : E.pages.map((_, i) => i), 'png', m); await dl(await zipBlobs(r.map(x => ({ name: baseName() + '-صفحة-' + (x.i + 1) + '.png', blob: x.blob }))), baseName() + '-صور.zip'); }); },
    async text() {
      const lines = await textOfPage(curPage()); const ta = $('#edText'); ta.value = lines.length ? lines.join('\n') : '(لا يوجد نص قابل للتحديد في هذه الصفحة — الصفحة صورة ممسوحة)'; ta.hidden = false;
    },
    async textAll() { await App.run('جارٍ استخراج النص…', async () => { const out = []; for (let i = 0; i < E.pages.length; i++) out.push('— صفحة ' + (i + 1) + ' —\n' + (await textOfPage(E.pages[i])).join('\n')); const ta = $('#edText'); ta.value = out.join('\n\n'); ta.hidden = false; }); },
    copyText() { const ta = $('#edText'); ta.select(); (navigator.clipboard ? navigator.clipboard.writeText(ta.value) : Promise.resolve(doc.execCommand('copy'))).then(() => toast('تم نسخ النص')); },
    pageNums() { addTextAll({ pos: $('#edPnPos').value, fmt: $('#edPnFmt').value || '{n}', fs: +$('#edPnSize').value || 11, col: $('#edPnCol').value, ff: $('#edPnFont').value, w: 160 }); },
    wm() { const t = $('#edWmTxt').value.trim(); if (!t) { toast('اكتب نص العلامة المائية'); return; } addWatermark(t, +$('#edWmSize').value || 70, $('#edWmCol').value, +$('#edWmOp').value || 15); },
    logo(which) { const s = S(); const u = s && s.logos[which] && s.logos[which].data; if (!u) { toast('لا يوجد شعار محفوظ — أضفه من تبويب «الشعارات»'); return; } addImageItem(u, { w: 80 }); },
    person(who, what) { const s = S(), p = s && s.people[who]; const u = p && p[what === 'sign' ? 'signImg' : 'stampImg']; if (!u) { toast('لم يُضف ' + (what === 'sign' ? 'توقيع' : 'ختم') + ' في الإعدادات'); return; } addImageItem(u, { w: 90 }); },
    async merge(where) { $('#edMergeFile')._where = where; $('#edMergeFile').click(); },
    sig() { sigPad(); }, stamp(txt, col) { addStamp(txt, col || '#c00000'); }, search() { doSearch(); }, hl() { hlHits(); },
    async flat() { const px = +$('#edFlatQ').value || 1500; await App.run('جارٍ إنشاء PDF مسطّح…', async m => { await dl(await buildFlatPdf(px, m), baseName() + '-مسطّح.pdf'); toast('تم — الصفحات صور، النص المخفي تحت التغطية أُزيل'); }); },
    imgFile() { $('#edImgFile').click(); },
    imgPageFile() { $('#edImgPageFile').click(); }
  };


  /* ================= أدوات احترافية ================= */
  function px1(x, y) {
    const bg = $('#edBg'), k = bg.width / dims(curPage())[0];
    const d = bg.getContext('2d').getImageData(Math.min(bg.width - 1, Math.max(0, Math.round(x * k))), Math.min(bg.height - 1, Math.max(0, Math.round(y * k))), 1, 1).data; return [d[0], d[1], d[2]];
  }
  function sampleBg(x, y, x2, y2) {
    const bg = $('#edBg'); if (!bg) return '#ffffff';
    try {
      const pts = [[x, y], [x2 == null ? x : x2, y2 == null ? y : y2], [x, y - 4], [x2 == null ? x : x2, (y2 == null ? y : y2) - 4], [x, y + 4]].map(q => px1(q[0], q[1]));
      let best = pts[0], bd = 1e9; pts.forEach(a => { const s = pts.reduce((m, c) => m + Math.abs(a[0] - c[0]) + Math.abs(a[1] - c[1]) + Math.abs(a[2] - c[2]), 0); if (s < bd) { bd = s; best = a; } });
      return '#' + best.map(v => v.toString(16).padStart(2, '0')).join('');
    } catch (e) { return '#ffffff'; }
  }
  const lineHit = (L, px, py) => L.lines.find(l => px >= l.x0 - 2 && px <= l.x1 + 2 && py >= l.y0 - 2 && py <= l.y1 + l.size * 0.3);
  async function editLineAt(px, py) {
    const pg = curPage(), L = await pageLines(pg), ln = lineHit(L, px, py);
    if (!ln) { toast('اضغط على سطر نصي موجود في الصفحة (الصفحات الممسوحة ضوئياً لا تحوي نصاً قابلاً للتحديد)'); return; }
    pushUndo();
    const w = (ln.x1 - ln.x0) * 1.12 + 8;
    const cover = newItem('white', { x: ln.x0 - 2, y: ln.y0 - 1, w: ln.x1 - ln.x0 + 4, h: ln.y1 - ln.y0 + ln.size * 0.32, fill: sampleBg(ln.x0 - 5, (ln.y0 + ln.y1) / 2, ln.x1 + 5, (ln.y0 + ln.y1) / 2), sw: 0 });
    const it = newItem('text', Object.assign({}, textDef, { txt: ln.text, fs: Math.round(ln.size * 10) / 10, b: ln.bold, w, al: ln.rtl ? 'right' : 'left', ff: 'Tajawal', y: ln.y0 - ln.size * 0.16 }));
    it.x = ln.rtl ? ln.x1 + 3 - w : ln.x0 - 1;
    pg.items.push(cover, it); setTool('select'); drawLayer(); select(it.id);
    startInline($('#edLayer [data-id="' + it.id + '"]'), it, false);
  }
  let hoverTok = 0;
  async function hoverEdit(e) {
    if (E.tool !== 'edit' || drag) return;
    const tok = ++hoverTok, L = await pageLines(curPage()); if (tok !== hoverTok) return;
    const [px, py] = pt(e), ln = lineHit(L, px, py), lay = $('#edLayer'); let h = $('#edHover');
    if (!ln) { h && h.remove(); return; }
    if (!h) { h = doc.createElement('div'); h.id = 'edHover'; h.className = 'ed-hover'; lay.appendChild(h); }
    h.style.cssText = 'left:' + (ln.x0 - 2) * E.z + 'px;top:' + (ln.y0 - 1) * E.z + 'px;width:' + (ln.x1 - ln.x0 + 4) * E.z + 'px;height:' + (ln.y1 - ln.y0 + ln.size * 0.3) * E.z + 'px';
  }
  function sigPad() {
    let d = $('#edSigDlg');
    if (!d) {
      d = doc.createElement('dialog'); d.id = 'edSigDlg';
      d.innerHTML = '<form method="dialog" class="ed-sig"><h4>ارسم توقيعك بالماوس أو الإصبع</h4><canvas id="edSigC" width="560" height="220"></canvas><div class="ed-btnrow"><label>اللون <input type="color" id="edSigCol" value="#0a2a6b"></label><label>السماكة <input type="range" id="edSigW" min="1" max="8" value="3"></label><button type="button" class="btn sm" id="edSigClr">مسح</button><button type="button" class="btn sm primary" id="edSigOk">إضافة للصفحة</button><button class="btn sm">إغلاق</button></div></form>';
      doc.body.appendChild(d);
      const c = $('#edSigC'), x = c.getContext('2d'); let on = false, bb = null;
      d._clear = () => { x.clearRect(0, 0, c.width, c.height); bb = null; };
      const P = e => { const r = c.getBoundingClientRect(); return [(e.clientX - r.left) * c.width / r.width, (e.clientY - r.top) * c.height / r.height]; };
      c.addEventListener('pointerdown', e => { on = true; c.setPointerCapture(e.pointerId); const [a, b] = P(e); x.strokeStyle = $('#edSigCol').value; x.lineWidth = +$('#edSigW').value; x.lineCap = x.lineJoin = 'round'; x.beginPath(); x.moveTo(a, b); x.lineTo(a + 0.1, b + 0.1); x.stroke(); });
      c.addEventListener('pointermove', e => { if (!on) return; const [a, b] = P(e); x.lineTo(a, b); x.stroke(); bb = bb ? [Math.min(bb[0], a), Math.min(bb[1], b), Math.max(bb[2], a), Math.max(bb[3], b)] : [a, b, a, b]; });
      c.addEventListener('pointerup', () => { on = false; });
      $('#edSigClr').addEventListener('click', d._clear);
      $('#edSigOk').addEventListener('click', () => {
        if (!bb) { toast('ارسم التوقيع أولاً'); return; }
        const pad = 8, sx = Math.max(0, bb[0] - pad), sy = Math.max(0, bb[1] - pad), sw = Math.min(c.width - sx, bb[2] - bb[0] + 2 * pad), sh = Math.min(c.height - sy, bb[3] - bb[1] + 2 * pad);
        const o = doc.createElement('canvas'); o.width = sw; o.height = sh; o.getContext('2d').drawImage(c, sx, sy, sw, sh, 0, 0, sw, sh);
        d.close(); addImageItem(o.toDataURL('image/png'), { w: Math.max(40, sw * 0.4) });
      });
    }
    d._clear(); d.showModal();
  }
  function addStamp(txt, col) {
    const pg = curPage(); if (!pg) return; const [w, h] = dims(pg), fs = 30, ww = Math.max(90, txt.length * fs * 0.62 + 24);
    pushUndo(); const it = newItem('text', Object.assign({}, textDef, { txt, fs, col, b: true, bd: 2.5, al: 'center', w: ww, rot: -10, x: (w - ww) / 2, y: h * 0.3 }));
    pg.items.push(it); drawLayer(); select(it.id);
  }
  async function buildFlatPdf(px, setMsg) {
    const PL = g.PDFLib, out = await PL.PDFDocument.create();
    for (let i = 0; i < E.pages.length; i++) {
      setMsg && setMsg('تسطيح الصفحة ' + (i + 1) + ' من ' + E.pages.length + '…');
      const pg = E.pages[i], [w, h] = dims(pg), jb = await pageToBlob(pg, 'jpg', px, px > 1800 ? 0.9 : px > 1200 ? 0.8 : 0.65), im = await out.embedJpg(new Uint8Array(await jb.arrayBuffer()));
      const pp = out.addPage([w, h]); pp.drawImage(im, { x: 0, y: 0, width: w, height: h });
    }
    out.setTitle(($('#edMetaT') || {}).value || baseName()); out.setProducer('محرر الكتب الرسمية');
    return new Blob([await out.save()], { type: 'application/pdf' });
  }
  let hits = [];
  async function doSearch() {
    const q = fixAr($('#edSearch').value.trim()).toLowerCase(), box = $('#edResults'); hits = [];
    if (!q) { toast('اكتب كلمة للبحث'); return; }
    await App.run('جارٍ البحث في الصفحات…', async () => {
      for (let i = 0; i < E.pages.length; i++) { const L = await pageLines(E.pages[i]); L.lines.forEach(ln => { if (ln.text.toLowerCase().includes(q)) { const its = ln.seq.filter(s => s.str.toLowerCase().includes(q)); hits.push({ i, text: ln.text, boxes: (its.length ? its : [ln]).map(s => ({ x0: s.x0, x1: s.x1, y0: s.y0, y1: s.y1, size: s.size || ln.size })) }); } }); }
    });
    box.innerHTML = hits.length ? '<div class="hint">' + hits.length + ' نتيجة — اضغط للانتقال</div>' + hits.slice(0, 80).map((h, k) => '<button type="button" class="ed-hit" data-hit="' + k + '">ص ' + (h.i + 1) + ': ' + esc(h.text.slice(0, 60)) + '</button>').join('') : '<div class="hint">لا توجد نتائج (الصفحات الممسوحة لا تحوي نصاً قابلاً للبحث)</div>';
  }
  function hlHits() {
    if (!hits.length) { toast('ابحث أولاً'); return; } pushUndo();
    hits.forEach(h => h.boxes.forEach(b => E.pages[h.i].items.push(newItem('hl', { x: b.x0 - 1, y: b.y0 - 1, w: b.x1 - b.x0 + 2, h: b.y1 - b.y0 + (b.size || 10) * 0.3, fill: '#ffe600', sw: 0 }))));
    renderAll(); toast('تم تظليل ' + hits.length + ' نتيجة');
  }

  /* ================= النموذج الجانبي ================= */
  function buildForm() {
    const host = $('#edForm'); if (!host) return;
    const F = FONTS().map(f => '<option>' + esc(f) + '</option>').join('');
    host.innerHTML =
      '<input type="file" id="edFile" accept=".pdf,application/pdf,.docx,application/vnd.openxmlformats-officedocument.wordprocessingml.document" hidden>' +
      '<input type="file" id="edImgFile" accept="image/*" multiple hidden><input type="file" id="edImgPageFile" accept="image/*" multiple hidden><input type="file" id="edMergeFile" accept="application/pdf,.pdf" hidden>' +
      '<fieldset class="group"><legend>١) الملف</legend>' +
      '<div class="ed-drop" id="edSideDrop"><b>اسحب ملف PDF أو Word هنا</b><span>أو</span><button type="button" class="btn primary" id="edPick">اختيار ملف من الجهاز</button></div>' +
      '<div class="ed-blank" data-blank><select class="ed-bsize">' + BLANK_OPTS + '</select><button type="button" class="btn sm" data-newblank="pdf">+ صفحة فارغة (PDF)</button><button type="button" class="btn sm" data-newblank="docx">+ مستند Word فارغ</button></div>' +
      '<p class="hint" id="edInfo">لم يُفتح ملف بعد. يُدعم: PDF و Word (.docx).</p>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" id="edScanBtn">الماسح الضوئي</button><button type="button" class="btn sm" id="edCamBtn">تصوير بالكاميرا</button><button type="button" class="btn sm danger" id="edClose" hidden>إغلاق الملف</button></div>' +
      '<div class="ed-att" data-edit-only="pdf"><span>الصور الملتقطة/الممسوحة تُضاف:</span><label class="check"><input type="radio" name="edAttMode" value="page" checked> <span>صفحة جديدة</span></label><label class="check"><input type="radio" name="edAttMode" value="item"> <span>صورة على الصفحة الحالية</span></label></div>' +
      '</fieldset>' +
      '<fieldset class="group" data-edit-only="pdf"><legend>٢) الصفحات <small id="edSelInfo"></small></legend>' +
      '<p class="hint">اضغط صفحة في الشريط الجانبي للتعديل عليها، وعلّم المربعات لاختيار عدة صفحات (تنطبق عليها العمليات). اسحب الصفحات لإعادة ترتيبها.</p>' +
      '<div class="ed-grid">' + [['up', '↑ تقديم'], ['down', '↓ تأخير'], ['first', 'إلى الأول'], ['last', 'إلى الأخير'], ['rotL', '⟲ تدوير'], ['rotR', '⟳ تدوير'], ['dup', 'تكرار'], ['blank', 'صفحة فارغة'], ['rev', 'عكس الترتيب'], ['del', 'حذف']].map(b => '<button type="button" class="btn sm' + (b[0] === 'del' ? ' danger' : '') + '" data-page="' + b[0] + '">' + b[1] + '</button>').join('') + '</div>' +
      '<div class="ed-grid">' + [['all', 'تحديد الكل'], ['none', 'إلغاء التحديد'], ['odd', 'الفردية'], ['even', 'الزوجية']].map(b => '<button type="button" class="btn sm ghost" data-page="' + b[0] + '">' + b[1] + '</button>').join('') + '</div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="merge" data-arg="end">دمج PDF في النهاية</button><button type="button" class="btn sm" data-tl="merge" data-arg="cur">دمج بعد الحالية</button><button type="button" class="btn sm" data-tl="imgPageFile">صورة كصفحة</button></div>' +
      '</fieldset>' +
      '<fieldset class="group" data-edit-only="pdf"><legend>٣) خصائص العنصر المحدد</legend><div id="edProps"></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="imgFile">+ إضافة صورة</button></div></fieldset>' +
      '<fieldset class="group" data-edit-only="pdf"><legend>٤) إضافات سريعة من الموقع</legend>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="logo" data-arg="club">شعار النادي</button><button type="button" class="btn sm" data-tl="logo" data-arg="ministry">شعار الوزارة</button>' +
      '<button type="button" class="btn sm" data-tl="person" data-arg="head|sign">توقيع الرئيس</button><button type="button" class="btn sm" data-tl="person" data-arg="head|stamp">ختم الرئيس</button><button type="button" class="btn sm" data-tl="person" data-arg="manager|sign">توقيع المدير</button><button type="button" class="btn sm" data-tl="person" data-arg="manager|stamp">ختم المدير</button></div>' +
      '<details class="t2-el"><summary>ترقيم الصفحات / رأس وتذييل</summary><div><label class="field"><span>الموضع</span><select id="edPnPos"><option value="bottom">أسفل الصفحة (تذييل)</option><option value="top">أعلى الصفحة (رأس)</option></select></label><div class="row2"><label class="field"><span>الصيغة</span><input id="edPnFmt" type="text" value="صفحة {n} من {N}"></label><label class="field"><span>الحجم</span><input id="edPnSize" type="number" value="11" min="6" max="40"></label></div><div class="row2"><label class="field"><span>اللون</span><input id="edPnCol" type="color" value="#333333"></label><label class="field"><span>الخط</span><select id="edPnFont">' + F + '</select></label></div><button type="button" class="btn sm" data-tl="pageNums">إضافة الترقيم لكل الصفحات</button></div></details>' +
      '<details class="t2-el"><summary>علامة مائية نصية</summary><div><label class="field"><span>النص</span><input id="edWmTxt" type="text" value="نسخة للمراجعة"></label><div class="row2"><label class="field"><span>الحجم</span><input id="edWmSize" type="number" value="70"></label><label class="field"><span>الشفافية %</span><input id="edWmOp" type="number" value="15" min="5" max="100"></label></div><label class="field"><span>اللون</span><input id="edWmCol" type="color" value="#888888"></label><button type="button" class="btn sm" data-tl="wm">إضافة لكل الصفحات</button></div></details>' +
      '</fieldset>' +
      '<fieldset class="group" data-edit-only="pdf"><legend>٥) التفكيك والتحويل</legend>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="extractSel">استخراج المحدد كملف</button><button type="button" class="btn sm" data-tl="splitAll">كل صفحة ملف (ZIP)</button></div>' +
      '<div class="row2"><label class="field"><span>استخراج مدى</span><input id="edRange" type="text" placeholder="1-3,5,8-10"></label><label class="field"><span>تقسيم كل N صفحة</span><input id="edEvery" type="number" value="2" min="1"></label></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="extractRange">استخراج المدى</button><button type="button" class="btn sm" data-tl="splitEvery">تقسيم لأجزاء (ZIP)</button></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm primary" data-tl="toWord" title="نص عربي قابل للتعديل مع الفقرات والمحاذاة">تحويل إلى Word (نص قابل للتعديل)</button><button type="button" class="btn sm" data-tl="toWordImg" title="كل صفحة صورة مطابقة للأصل">Word مطابق للأصل (صور)</button><button type="button" class="btn sm" data-tl="imgPage">الصفحة كصورة PNG</button><button type="button" class="btn sm" data-tl="imgAll">كل الصفحات صور (ZIP)</button></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="text">نص الصفحة</button><button type="button" class="btn sm" data-tl="textAll">نص كل الصفحات</button><button type="button" class="btn sm" data-tl="copyText">نسخ النص</button></div>' +
      '<textarea id="edText" rows="6" hidden readonly></textarea></fieldset>' +
      '<fieldset class="group" data-edit-only="pdf"><legend>٦) أدوات احترافية</legend>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-tl="sig">✍ رسم توقيع</button></div>' +
      '<div class="hint">أختام نصية جاهزة:</div><div class="ed-btnrow">' + [['معتمد', '#0a7a2f'], ['نسخة طبق الأصل', '#0a3a8a'], ['سري', '#c00000'], ['مسودة', '#c06000'], ['تم الدفع', '#0a7a2f'], ['صورة', '#555555']].map(s => '<button type="button" class="btn sm" data-tl="stamp" data-arg="' + s[0] + '|' + s[1] + '">' + s[0] + '</button>').join('') + '</div>' +
      '<div class="row2"><label class="field"><span>بحث في نص الكتاب</span><input id="edSearch" type="text" placeholder="كلمة أو عبارة"></label><div class="ed-btnrow" style="align-items:flex-end"><button type="button" class="btn sm" data-tl="search">بحث</button><button type="button" class="btn sm" data-tl="hl">تظليل النتائج</button></div></div><div id="edResults"></div>' +
      '<div class="row2"><label class="field"><span>عنوان الملف (خصائص)</span><input id="edMetaT" type="text"></label><label class="field"><span>المؤلف</span><input id="edMetaA" type="text"></label></div>' +
      '<div class="row2"><label class="field"><span>جودة التسطيح/الضغط</span><select id="edFlatQ"><option value="2200">عالية</option><option value="1500" selected>متوسطة</option><option value="1000">منخفضة (أصغر حجماً)</option></select></label><div class="ed-btnrow" style="align-items:flex-end"><button type="button" class="btn sm" data-tl="flat">PDF مسطّح/مضغوط</button></div></div>' +
      '<p class="hint">«تعديل نص موجود»: اختر الأداة من الشريط ثم اضغط سطراً في الصفحة. «المسطّح» يحوّل الصفحات إلى صور فيُزيل أي نص مغطّى بمستطيل أسود أو أبيض (مفيد للحجب).</p></fieldset>' +
      '<div data-edit-only="docx" id="edWordForm"></div>' +
      '<p class="hint ed-foot">التنزيل والطباعة والإرسال من الأزرار العلوية تعمل على هذا الملف. التعديلات تبقى في الصفحة المفتوحة فقط (لا تُحفظ بعد إغلاقها)، فنزّل الملف قبل الخروج.</p>';
    bindForm();
    if (App.editorWord) App.editorWord.buildForm($('#edWordForm'), E);
  }
  function bindForm() {
    const host = $('#edForm');
    $('#edFile').addEventListener('change', e => { const f = e.target.files[0]; e.target.value = ''; openFile(f); });
    $('#edPick').addEventListener('click', () => $('#edFile').click());
    $$('[data-newblank]', host).forEach(b => b.addEventListener('click', () => { if (E.kind && !confirm('سيُغلق الملف الحالي وتضيع التعديلات غير المنزّلة. متابعة؟')) return; newBlank(b.dataset.newblank, $('.ed-bsize', b.closest('[data-blank]')).value); }));
    $('#edScanBtn').addEventListener('click', () => $('#btnScan').click());
    $('#edCamBtn').addEventListener('click', () => $('#btnCam').click());
    $('#edClose').addEventListener('click', () => { if (!confirm('إغلاق الملف؟ ستفقد التعديلات غير المنزّلة.')) return; closeFile(); });
    const drop = $('#edSideDrop');
    ['dragover', 'dragenter'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
    ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, () => drop.classList.remove('over')));
    drop.addEventListener('drop', e => { e.preventDefault(); openFile(e.dataTransfer.files[0]); });
    host.addEventListener('click', e => {
      const p = e.target.closest('[data-page]'); if (p && pageOps[p.dataset.page]) { pageOps[p.dataset.page](); return; }
      const t = e.target.closest('[data-tl]'); if (t && tools[t.dataset.tl]) { const a = (t.dataset.arg || '').split('|'); tools[t.dataset.tl](a[0], a[1]); }
      const h = e.target.closest('[data-hit]'); if (h && hits[+h.dataset.hit]) goPage(hits[+h.dataset.hit].i);
    });
    $('#edImgFile').addEventListener('change', async e => { const fs = Array.from(e.target.files); e.target.value = ''; for (const f of fs) await addImageItem(await readUrl(f)); setTool('select'); });
    $('#edImgPageFile').addEventListener('change', async e => { const fs = Array.from(e.target.files); e.target.value = ''; if (E.kind !== 'pdf') return; pushUndo(); for (const f of fs) { E.pages.splice(E.cur + 1, 0, await imageAsPages(await readUrl(f), f.name)); E.cur++; } renderAll(); });
    $('#edMergeFile').addEventListener('change', async e => { const f = e.target.files[0], w = e.target._where || 'end'; e.target.value = ''; if (f) await App.run('جارٍ دمج الملف…', () => loadPdf(f, w)); });
    bindProps();
  }
  function closeFile() {
    E.kind = null; E.name = ''; E.srcs = []; E.pages = []; E.undo = []; E.redo = []; E.sel = null; E.psel.clear(); E.word = null; bgCache.clear(); thumbCache.clear();
    showEmpty(); syncForm();
  }
  function syncForm() {
    const kind = E.kind;
    $$('[data-edit-only]').forEach(n => { n.hidden = !(kind && (n.dataset.editOnly === kind)); });
    const info = $('#edInfo'); if (info) info.textContent = kind ? ('الملف المفتوح: ' + E.name + (kind === 'pdf' ? ' — ' + E.pages.length + ' صفحة' : ' — مستند Word')) : 'لم يُفتح ملف بعد. يُدعم: PDF و Word (.docx).';
    const c = $('#edClose'); if (c) c.hidden = !kind;
    const si = $('#edSelInfo'); if (si) si.textContent = E.psel.size ? '(' + E.psel.size + ' محددة)' : '';
    const nav = $('#edNav'); if (nav) nav.textContent = (E.cur + 1) + ' / ' + E.pages.length;
    if (kind === 'pdf') syncProps();
  }

  /* ================= دخول/خروج التبويب والتصدير ================= */
  function show(on) {
    doc.body.classList.toggle('mode-edit', on);
    const st = stage(); if (!st) return;
    if (!on) { if (E.kind === 'docx' && App.editorWord) App.editorWord.snapshot(); st.hidden = true; $('#sheetHolder').hidden = false; return; }
    $('#sheetHolder').hidden = true;
    if (!E.kind) showEmpty(); else { showStage(); renderAll(); }
    syncForm();
  }
  async function getFile(fmt) {
    if (!E.kind) throw new Error('افتح ملف PDF أو Word أولاً');
    if (E.kind === 'pdf') {
      if (fmt === 'pdf') return { blob: await buildPdf(null, App.setBusyMsg), name: baseName() + '-معدّل.pdf' };
      return { blob: await toWordDoc(App.setBusyMsg), name: baseName() + '.docx' };
    }
    return App.editorWord.getFile(fmt, E);
  }
  async function printIt() {
    if (!E.kind) { toast('افتح ملفاً أولاً'); return; }
    await App.run('جارٍ تجهيز الطباعة…', async () => {
      if (E.kind === 'pdf') {
        const b = await buildPdf(null, App.setBusyMsg), url = blobUrl(b);
        const f = doc.createElement('iframe'); f.style.cssText = 'position:fixed;right:0;bottom:0;width:1px;height:1px;border:0;opacity:0'; f.src = url; doc.body.appendChild(f);
        f.onload = () => { setTimeout(() => { try { f.contentWindow.focus(); f.contentWindow.print(); } catch (e) { g.open(url); } }, 300); setTimeout(() => { f.remove(); URL.revokeObjectURL(url); }, 120000); };
      } else App.editorWord.print(E);
    });
  }

  function init() {
    buildForm();
    doc.addEventListener('keydown', e => {
      if (!doc.body.classList.contains('mode-edit') || E.kind !== 'pdf') return;
      const tag = (e.target.tagName || '').toLowerCase(); if (tag === 'input' || tag === 'textarea' || tag === 'select' || e.target.isContentEditable) return;
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'z') { e.preventDefault(); e.shiftKey ? redo() : undo(); }
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'y') { e.preventDefault(); redo(); }
      else if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'd') { const it = selItem(); if (it) { e.preventDefault(); itemAct('dup', it); } }
      else if (e.key === 'Delete' || e.key === 'Backspace') { const it = selItem(); if (it) { e.preventDefault(); itemAct('del', it); } }
      else if (e.key.startsWith('Arrow')) { const it = selItem(); if (!it) return; e.preventDefault(); const st = e.shiftKey ? 10 : 1; pushUndo(); shift(it, e.key === 'ArrowLeft' ? -st : e.key === 'ArrowRight' ? st : 0, e.key === 'ArrowUp' ? -st : e.key === 'ArrowDown' ? st : 0); liveUpdate(it); }
      else if (e.key === 'PageDown') goPage(E.cur + 1); else if (e.key === 'PageUp') goPage(E.cur - 1);
    });
    g.addEventListener('resize', () => { if (E.kind === 'pdf' && E.fit && doc.body.classList.contains('mode-edit')) renderPage(); });
  }

  g.App = Object.assign(App, { editor: { init, show, getFile, print: printIt, pull: pullAttachments, E, openFile, close: closeFile, pdfToImages, pdfTextHTML } });
})(window);
