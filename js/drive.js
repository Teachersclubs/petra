/* ===== Google Drive: تنزيل الملف ثم فتح صفحة Drive لرفعه (بدون إعدادات أو حسابات) ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const DRIVE_URL = 'https://drive.google.com/drive/my-drive';

  async function go(S, btn) {
    // تُفتح النافذة فورًا من نقرة المستخدم حتى لا يحجبها المتصفح
    const win = window.open('about:blank', '_blank');
    btn.disabled = true;
    try {
      g.toast('جارٍ تجهيز الملف…');
      const type = S.activeDoc, fmt = S.send.format === 'docx' ? 'docx' : 'pdf';
      const blob = fmt === 'docx' ? await App.buildDocx(S, type) : await App.buildPdf(S, type);
      App.download(blob, App.fileName(S, type) + '.' + fmt);
      g.toast('نُزّل الملف — في صفحة Drive اختر «جديد ← رفع ملف» أو اسحب الملف إليها');
      if (win) win.location.href = DRIVE_URL; else window.open(DRIVE_URL, '_blank');
    } catch (e) {
      console.error(e); if (win) win.close();
      g.toast('تعذّر تجهيز الملف: ' + (e.message || e));
    } finally { btn.disabled = false; }
  }

  function init(getS) { document.getElementById('btnDrive').addEventListener('click', e => go(getS(), e.currentTarget)); }
  g.App = Object.assign(App, { drive: { init } });
})(window);
