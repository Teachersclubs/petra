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

  /* تحرير في محرر مستندات Google: ينزّل ملف Word ثم يفتح Google Docs ليُرفع ويُحرَّر */
  const DOCS_URL = 'https://docs.google.com/document/create'; // يفتح محرر مستند جديد مباشرةً
  async function editInDocs(S) {
    if (!(await App.approve('التحرير في Google Docs'))) return;
    const win = window.open('about:blank', '_blank');
    let ok = false;
    try {
      ok = await App.run('جارٍ تجهيز ملف Word للتحرير في Google Docs…', async () => {
        const f = await App.getFile(S, S.activeDoc, 'docx');
        if (!(await App.saveFile(f.blob, f.name))) throw new Error('أُلغي التنزيل');
        g.toast('نُزّل الملف — في محرر Google Docs: ملف ← فتح (Ctrl+O) ← تحميل ← اختر الملف، فيُفتح للتحرير', 9000);
      });
    } finally {
      if (ok) { if (win) win.location.href = DOCS_URL; else window.open(DOCS_URL, '_blank'); }
      else if (win) win.close();
    }
  }

  function init(getS) {
    document.getElementById('btnDrive').addEventListener('click', e => go(getS(), e.currentTarget));
    document.getElementById('btnGdocs').addEventListener('click', () => editInDocs(getS()));
  }
  g.App = Object.assign(App, { drive: { init } });
})(window);
