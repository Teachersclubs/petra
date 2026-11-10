/* ===== الإرسال: بريد إلكتروني معتمد وواتساب ===== */
(function (g) {
  'use strict';
  const App = g.App;

  async function makeFile(S) {
    const type = S.activeDoc, fmt = S.send.format === 'docx' ? 'docx' : 'pdf';
    const blob = fmt === 'docx' ? await App.buildDocx(S, type) : await App.buildPdf(S, type);
    const mime = fmt === 'pdf' ? 'application/pdf' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
    return new File([blob], App.fileName(S, type) + '.' + fmt, { type: mime });
  }

  function subjectText(S) {
    const t = S.activeDoc, d = S[t];
    return (t === 'circular' ? 'تعميم رقم ' : 'كتاب رقم ') + (d.number || '') + (d.subject ? ' – ' + d.subject : '');
  }

  function clubLine(S) { return 'نادي ' + (S.clubName || ''); }

  async function guard(btn, fn) {
    btn.disabled = true;
    try { await fn(); } catch (e) { console.error(e); if (e && e.name === 'AbortError') return; g.toast('تعذّر الإرسال: ' + (e.message || e)); }
    finally { btn.disabled = false; }
  }

  async function sendMail(S, btn) {
    const email = (S.send.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { g.toast('أدخل البريد الإلكتروني المعتمد بشكل صحيح'); return; }
    await guard(btn, async () => {
      g.toast('جارٍ تجهيز الملف…');
      const file = await makeFile(S);
      App.download(file, file.name);
      const body = 'السلام عليكم ورحمة الله وبركاته،\n\nنرفق لكم ' + subjectText(S) + '.\n(الملف المرفق: ' + file.name + ' — يرجى إرفاقه من التنزيلات)\n\nوتفضلوا بقبول فائق الاحترام والتقدير.\n' + clubLine(S);
      g.toast('نُزّل الملف — أرفقه بالرسالة التي ستُفتح الآن');
      setTimeout(() => { location.href = 'mailto:' + email + '?subject=' + encodeURIComponent(subjectText(S)) + '&body=' + encodeURIComponent(body); }, 700);
    });
  }

  async function sendWhatsApp(S, btn) {
    const probe = new File([''], 'a.pdf', { type: 'application/pdf' });
    const canShare = !!(navigator.canShare && navigator.share && navigator.canShare({ files: [probe] }));
    // نافذة تُفتح فورًا حتى لا يحجبها المتصفح بعد إنشاء الملف
    const win = canShare ? null : window.open('about:blank', '_blank');
    await guard(btn, async () => {
      g.toast('جارٍ تجهيز الملف…');
      const file = await makeFile(S);
      const text = subjectText(S) + '\n' + clubLine(S);
      if (canShare && navigator.canShare({ files: [file] })) {
        await navigator.share({ files: [file], title: subjectText(S), text });
        return;
      }
      App.download(file, file.name);
      const num = (S.send.whatsapp || '').replace(/\D/g, '').replace(/^00/, '');
      const url = 'https://wa.me/' + (num ? num : '') + '?text=' + encodeURIComponent(text + '\n(الملف: ' + file.name + ' — أرفقه من التنزيلات)');
      g.toast('نُزّل الملف — أرفقه في محادثة واتساب');
      if (win) win.location.href = url; else window.open(url, '_blank');
    });
    if (win && win.location.href === 'about:blank') { try { win.close(); } catch (e) { /* تجاهل */ } }
  }

  function init(getS) {
    document.getElementById('btnMail').addEventListener('click', e => sendMail(getS(), e.currentTarget));
    document.getElementById('btnWa').addEventListener('click', e => sendWhatsApp(getS(), e.currentTarget));
  }

  g.App = Object.assign(App, { initSend: init, makeFile });
})(window);
