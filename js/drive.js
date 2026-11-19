/* ===== Google Drive: تنزيل الملف ثم فتح صفحة Drive لرفعه (بدون إعدادات أو حسابات) ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const DRIVE_URL = 'https://drive.google.com/drive/my-drive';

  async function go(S, btn) {
    if (!(await App.approve('الرفع إلى Drive'))) return;
    // تُفتح النافذة فورًا من نقرة المستخدم حتى لا يحجبها المتصفح
    const win = window.open('about:blank', '_blank');
    let ok = false;
    try {
      ok = await App.run('جارٍ تجهيز الملف لرفعه إلى Drive…', async () => {
        const f = await App.getFile(S, S.activeDoc, S.send.format === 'docx' ? 'docx' : 'pdf');
        if (!(await App.saveFile(f.blob, f.name))) throw new Error('أُلغي التنزيل');
        g.toast('نُزّل الملف — في صفحة Drive اختر «جديد ← رفع ملف» أو اسحب الملف إليها');
      });
    } finally {
      if (ok) { if (win) win.location.href = DRIVE_URL; else window.open(DRIVE_URL, '_blank'); }
      else if (win) win.close();
    }
  }

  /* تحرير مباشر في Google Docs: يرفع ملف Word إلى Drive الخاص بالمستخدم مع التحويل إلى مستند Google ثم يفتحه.
     يحتاج «معرّف عميل OAuth» مجانيًا من Google Cloud (يُحفظ في هذا المتصفح فقط). */
  const DOCS_NEW = 'https://docs.google.com/document/create';
  const LS_ID = 'gdocsClientId';
  const getId = () => { try { return localStorage.getItem(LS_ID) || ''; } catch (e) { return ''; } };
  const setId = v => { try { v ? localStorage.setItem(LS_ID, v) : localStorage.removeItem(LS_ID); } catch (e) { /* لا شيء */ } };
  let tokenClient = null, tokenId = '', token = null, tokenExp = 0;

  function loadGis() {
    if (window.google && google.accounts && google.accounts.oauth2) return Promise.resolve();
    return new Promise((res, rej) => {
      const sc = document.createElement('script');
      sc.src = 'https://accounts.google.com/gsi/client'; sc.async = true;
      sc.onload = () => res(); sc.onerror = () => rej(new Error('تعذّر تحميل خدمة تسجيل الدخول من Google (تحقق من الاتصال بالإنترنت)'));
      document.head.appendChild(sc);
    });
  }
  // يجب استدعاؤها مباشرةً بعد نقرة المستخدم حتى لا تُحجب نافذة الإذن
  function getToken(clientId) {
    if (token && Date.now() < tokenExp && tokenId === clientId) return Promise.resolve(token);
    return loadGis().then(() => new Promise((res, rej) => {
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: clientId, scope: 'https://www.googleapis.com/auth/drive.file',
        callback: r => {
          if (r.error) return rej(new Error(r.error === 'access_denied' ? 'لم يتم منح الإذن' : 'تعذّر تسجيل الدخول: ' + r.error));
          token = r.access_token; tokenId = clientId; tokenExp = Date.now() + (r.expires_in - 60) * 1000; res(token);
        },
        error_callback: e => rej(new Error(e && e.type === 'popup_closed' ? 'أُغلقت نافذة تسجيل الدخول' : 'تعذّر فتح نافذة تسجيل الدخول (اسمح بالنوافذ المنبثقة)'))
      });
      tokenClient.requestAccessToken({ prompt: token ? '' : undefined });
    }));
  }
  async function uploadAsDoc(blob, name, tk) {
    const B = 'tce' + Math.random().toString(36).slice(2);
    const meta = JSON.stringify({ name: name.replace(/\.docx$/i, ''), mimeType: 'application/vnd.google-apps.document' });
    const body = new Blob(['--' + B + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' + meta + '\r\n--' + B +
      '\r\nContent-Type: application/vnd.openxmlformats-officedocument.wordprocessingml.document\r\n\r\n', blob, '\r\n--' + B + '--']);
    const r = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
      method: 'POST', headers: { Authorization: 'Bearer ' + tk, 'Content-Type': 'multipart/related; boundary=' + B }, body
    });
    if (!r.ok) { if (r.status === 401) token = null; throw new Error('فشل الرفع إلى Drive (' + r.status + ')'); }
    return (await r.json()).id;
  }

  async function editInDocs(S) {
    const id = getId();
    if (!id) { openDlg(); return; }
    if (!(await App.approve('التحرير في Google Docs'))) return;
    const win = window.open('about:blank', '_blank');
    let url = '';
    try {
      const tk = await getToken(id);               // أولًا: الإذن (من نقرة المستخدم)
      url = await App.run('جارٍ فتح الكتاب في Google Docs…', async () => {
        const f = await App.getFile(S, S.activeDoc, 'docx');
        const fid = await uploadAsDoc(f.blob, f.name, tk);
        return 'https://docs.google.com/document/d/' + fid + '/edit';
      });
    } catch (e) {
      g.toast((e && e.message) || 'تعذّر الفتح في Google Docs', 7000);
    } finally {
      if (url) { if (win) win.location.href = url; else window.open(url, '_blank'); }
      else if (win) win.close();
    }
  }

  /* ---- نافذة إعداد معرّف العميل ---- */
  function openDlg() {
    const d = document.getElementById('gdocsDlg'); document.getElementById('gdId').value = getId();
    document.getElementById('gdOrigin').textContent = location.origin === 'null' ? '(افتح الموقع من رابط https وليس من ملف)' : location.origin;
    d.showModal();
  }
  function setupDlg(getS) {
    const d = document.getElementById('gdocsDlg');
    document.getElementById('gdSave').addEventListener('click', () => {
      const v = document.getElementById('gdId').value.trim();
      if (v && !/\.apps\.googleusercontent\.com$/.test(v)) { g.toast('المعرّف يجب أن ينتهي بـ apps.googleusercontent.com'); return; }
      setId(v); d.close(); g.toast(v ? 'تم الحفظ — اضغط أيقونة Google Docs مجددًا' : 'تم حذف المعرّف');
    });
    document.getElementById('gdCancel').addEventListener('click', () => d.close());
    document.getElementById('gdManual').addEventListener('click', async () => {
      d.close();
      const S = getS(); if (!(await App.approve('التحرير في Google Docs'))) return;
      const win = window.open('about:blank', '_blank'); let ok = false;
      try {
        ok = await App.run('جارٍ تجهيز ملف Word…', async () => {
          const f = await App.getFile(S, S.activeDoc, 'docx');
          if (!(await App.saveFile(f.blob, f.name))) throw new Error('أُلغي التنزيل');
          g.toast('نُزّل الملف — في المحرر: ملف ← فتح (Ctrl+O) ← تحميل ← اختر الملف', 9000);
        });
      } finally { if (ok) { if (win) win.location.href = DOCS_NEW; else window.open(DOCS_NEW, '_blank'); } else if (win) win.close(); }
    });
    const b = document.getElementById('btnGdocsSetup'); if (b) b.addEventListener('click', openDlg);
  }

  function init(getS) {
    document.getElementById('btnDrive').addEventListener('click', e => go(getS(), e.currentTarget));
    document.getElementById('btnGdocs').addEventListener('click', () => editInDocs(getS()));
    setupDlg(getS);
  }
  g.App = Object.assign(App, { drive: { init } });
})(window);
