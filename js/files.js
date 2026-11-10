/* ===== تجهيز الملفات وتنزيلها: مؤشر تحميل، منع التكرار، ذاكرة مؤقتة، تنزيل موثوق ===== */
(function (g) {
  'use strict';
  const App = g.App;
  let working = false, msgEl = null, ovEl = null;

  const frame = () => new Promise(r => requestAnimationFrame(() => setTimeout(r, 0)));
  function overlay() {
    if (ovEl) return ovEl;
    ovEl = document.createElement('div');
    ovEl.id = 'busy'; ovEl.hidden = true; ovEl.setAttribute('role', 'status'); ovEl.setAttribute('aria-live', 'polite');
    ovEl.innerHTML = '<div class="busy-card"><span class="spin" aria-hidden="true"></span><span id="busyMsg"></span></div>';
    document.body.appendChild(ovEl); msgEl = ovEl.querySelector('#busyMsg');
    return ovEl;
  }
  function setMsg(m) { overlay(); msgEl.textContent = m; }

  /** تشغيل عملية طويلة: مؤشر تحميل + منع الضغط المتكرر + عدم تجميد الواجهة */
  async function run(label, fn) {
    if (working) { g.toast('هناك عملية قيد التنفيذ، انتظر حتى تنتهي'); return false; }
    working = true; document.body.classList.add('is-busy');
    const o = overlay(); setMsg(label); o.hidden = false;
    try { await frame(); const r = await fn(setMsg); return r === undefined ? true : r; }
    catch (e) { console.error(e); g.toast('تعذّر تنفيذ العملية: ' + ((e && e.message) || e)); return false; }
    finally { working = false; document.body.classList.remove('is-busy'); o.hidden = true; }
  }

  /* ---------- ذاكرة مؤقتة للملفات: لا يُعاد البناء ما دامت البيانات لم تتغير ---------- */
  function hash(str) { let h = 5381; for (let i = 0; i < str.length; i++) h = ((h << 5) + h + str.charCodeAt(i)) | 0; return (h >>> 0).toString(36) + ':' + str.length; }
  function fingerprint(S, type, fmt) {
    const c = Object.assign({}, S); delete c.send; delete c.theme; delete c.activeDoc;
    const att = App.attach.flat(type).map(p => [p.cap, p.w, p.h, p.url.length, p.url.slice(-60)]);
    return hash(JSON.stringify([type, fmt, c, att]));
  }
  const cache = new Map();
  async function getFile(S, type, fmt) {
    const key = fingerprint(S, type, fmt), name = App.fileName(S, type) + '.' + fmt;
    if (cache.has(key)) { const b = cache.get(key); cache.delete(key); cache.set(key, b); return { blob: b, name, cached: true }; }
    setMsg(fmt === 'pdf' ? 'جارٍ تجهيز ملف PDF…' : 'جارٍ تجهيز ملف Word…');
    await frame();
    const blob = fmt === 'docx' ? await App.buildDocx(S, type) : await App.buildPdf(S, type);
    cache.set(key, blob); while (cache.size > 4) cache.delete(cache.keys().next().value);
    return { blob, name, cached: false };
  }

  /* ---------- حفظ الملف ---------- */
  let lastUrl = '';
  function readyBar(blob, name) {
    let bar = document.getElementById('dlBar');
    if (!bar) { bar = document.createElement('div'); bar.id = 'dlBar'; document.body.appendChild(bar); }
    if (lastUrl) URL.revokeObjectURL(lastUrl);
    lastUrl = URL.createObjectURL(blob);
    bar.innerHTML = '';
    const t = document.createElement('span'); t.textContent = 'الملف جاهز: '; bar.appendChild(t);
    const a = document.createElement('a'); a.href = lastUrl; a.download = name; a.textContent = 'إن لم يبدأ التنزيل اضغط هنا'; a.rel = 'noopener'; bar.appendChild(a);
    const x = document.createElement('button'); x.type = 'button'; x.textContent = '×'; x.setAttribute('aria-label', 'إغلاق'); x.onclick = () => { bar.hidden = true; }; bar.appendChild(x);
    bar.hidden = false; clearTimeout(readyBar.t); readyBar.t = setTimeout(() => { bar.hidden = true; }, 25000);
  }

  async function viaCapability(blob, name) {
    // داخل صفحة منشورة على claude.ai يُسمح بالتنزيل عبر قدرة downloads فقط
    try {
      if (!(g.claude && typeof g.claude.use === 'function')) return null;
      const dl = await Promise.race([g.claude.use('downloads'), new Promise(r => setTimeout(() => r(null), 1500))]);
      if (!dl) return null;
      await dl.save({ filename: name, data: blob });
      return 'saved';
    } catch (e) { return e && e.code === 'declined' ? 'declined' : null; }
  }

  /** ينزّل الملف باسمه (عربي وفراغات مسموحة). يعيد true إذا بدأ التنزيل */
  async function saveFile(blob, name) {
    name = String(name || 'file').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').trim();
    const cap = await viaCapability(blob, name);
    if (cap === 'saved') return true;
    if (cap === 'declined') return false;
    const typed = blob.type ? blob : new Blob([blob], { type: /\.pdf$/i.test(name) ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document' });
    const a = document.createElement('a');
    a.href = URL.createObjectURL(typed); a.download = name; a.rel = 'noopener'; a.style.display = 'none';
    document.body.appendChild(a); a.click();
    setTimeout(() => { a.remove(); URL.revokeObjectURL(a.href); }, 60000);
    readyBar(typed, name);
    return true;
  }

  g.App = Object.assign(App, { run, getFile, saveFile, setBusyMsg: setMsg, download: (b, n) => saveFile(b, n) });
})(window);
