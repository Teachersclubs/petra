/* ===== الإرسال: بريد إلكتروني وواتساب (كمبيوتر وهاتف) ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const UA = navigator.userAgent || '';
  const isMobile = /Android|iPhone|iPad|iPod/i.test(UA) || (navigator.maxTouchPoints > 1 && /Macintosh/.test(UA));

  const fmtOf = S => (S.send.format === 'docx' ? 'docx' : 'pdf');
  function subjectText(S) {
    const t = S.activeDoc, d = S[t];
    if ((S.send.subject || '').trim()) return S.send.subject.trim();
    const kind = { circular: 'تعميم رقم ', assign: 'قرار تكليف ', claim: 'مطالبة مالية ' }[t] || 'كتاب رقم ';
    return kind + (d.number || '') + (d.subject ? ' – ' + d.subject : '');
  }
  const clubLine = S => 'نادي ' + (S.clubName || '');
  function messageText(S, fileName) {
    const m = (S.send.message || '').trim();
    return (m || 'السلام عليكم ورحمة الله وبركاته،\n\nنرفق لكم ' + subjectText(S) + '.\n\nوتفضلوا بقبول فائق الاحترام والتقدير.') + '\n' + clubLine(S) + (fileName ? '\n(الملف المرفق: ' + fileName + ')' : '');
  }

  /** مشاركة الملف عبر قائمة النظام (يرفقه تلقائيًا). يعيد true عند النجاح */
  async function tryShare(file, title, text) {
    try {
      if (!(navigator.share && navigator.canShare && navigator.canShare({ files: [file] }))) return false;
      await navigator.share({ files: [file], title, text });
      return true;
    } catch (e) {
      if (e && e.name === 'AbortError') return 'cancel';
      return false;
    }
  }
  const canShareFiles = () => { try { return !!(navigator.share && navigator.canShare && navigator.canShare({ files: [new File(['x'], 'a.pdf', { type: 'application/pdf' })] })); } catch (e) { return false; } };

  /* ---------- نافذة الخطوات ---------- */
  function dialog(title, html, buttons) {
    const dlg = document.getElementById('sendDlg');
    document.getElementById('sdTitle').textContent = title;
    document.getElementById('sdBody').innerHTML = html;
    const bt = document.getElementById('sdBtns'); bt.innerHTML = '';
    buttons.concat([{ label: 'إغلاق', value: 'close' }]).forEach(b => {
      const el = document.createElement('button'); el.type = b.value === 'close' ? 'submit' : 'button'; el.className = 'btn' + (b.primary ? ' primary' : '');
      el.textContent = b.label; if (b.onClick) el.addEventListener('click', b.onClick); bt.appendChild(el);
    });
    if (!dlg.open) dlg.showModal();
  }
  const li = a => '<ol class="steps">' + a.map(x => '<li>' + x + '</li>').join('') + '</ol>';
  const E = App.esc;

  async function sendMail(S) {
    const email = (S.send.email || '').trim();
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) { g.toast('أدخل البريد الإلكتروني المعتمد بشكل صحيح (مثال: name@example.com)'); document.querySelector('[data-bind="send.email"]').focus(); return; }
    if (!(await App.approve('الإرسال'))) return;
    const app = S.send.mailApp || 'gmail', useShare = isMobile || S.send.useShare;
    // تُفتح النافذة فورًا من نقرة المستخدم حتى لا يحجبها المتصفح بعد البناء
    const win = (!useShare && app !== 'default') ? window.open('about:blank', '_blank') : null;
    let state = 'fail';
    await App.run('جارٍ تجهيز الملف للإرسال…', async () => {
      const f = await App.getFile(S, S.activeDoc, fmtOf(S));
      const file = new File([f.blob], f.name, { type: f.blob.type });
      const subj = subjectText(S), body = messageText(S, f.name);
      if (useShare) {
        const r = await tryShare(file, subj, body);
        if (r === true) { state = 'ok'; g.toast('تم تجهيز الرسالة مع الملف المرفق'); return; }
        if (r === 'cancel') { state = 'cancel'; g.toast('أُلغيت المشاركة'); return; }
      }
      const saved = await App.saveFile(f.blob, f.name);
      if (!saved) { state = 'cancel'; g.toast('أُلغي تنزيل الملف، لم تُفتح الرسالة'); return; }
      let url;
      if (app === 'gmail') url = 'https://mail.google.com/mail/?view=cm&fs=1&to=' + encodeURIComponent(email) + '&su=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(body);
      else if (app === 'outlook') url = 'https://outlook.office.com/mail/deeplink/compose?to=' + encodeURIComponent(email) + '&subject=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(body);
      else url = 'mailto:' + email + '?subject=' + encodeURIComponent(subj) + '&body=' + encodeURIComponent(body);
      const openIt = () => { if (url.indexOf('mailto:') === 0) location.href = url; else window.open(url, '_blank'); };
      if (win) win.location.href = url; else if (url.indexOf('mailto:') === 0) setTimeout(openIt, 400);
      state = win || url.indexOf('mailto:') === 0 ? 'ok' : 'popup';
      dialog('تم تجهيز رسالة البريد', '<p>' + E('الملف «' + f.name + '» نُزّل إلى جهازك وفُتحت رسالة جاهزة إلى ') + '<bdi dir="ltr">' + E(email) + '</bdi>.</p>' +
        li(['في نافذة الرسالة اضغط على أيقونة <b>المشبك (إرفاق)</b>.', 'اختر الملف <b>' + E(f.name) + '</b> من مجلد التنزيلات.', 'راجع الرسالة ثم اضغط <b>إرسال</b>.']) +
        (state === 'popup' ? '<p class="warn-fit">حجب المتصفح فتح النافذة؛ اضغط الزر أدناه.</p>' : ''),
        [{ label: 'فتح رسالة البريد', primary: true, onClick: openIt }, { label: 'تنزيل الملف مرة أخرى', onClick: () => App.saveFile(f.blob, f.name) }]);
    });
    if (win && state !== 'ok') { try { if (win.location.href === 'about:blank') win.close(); } catch (e) { /* تجاهل */ } }
    if (state === 'fail') g.toast('تعذّر تجهيز الإرسال بالبريد، حاول مرة أخرى');
  }

  async function sendWhatsApp(S) {
    if (!(await App.approve('الإرسال'))) return;
    const useShare = isMobile || S.send.useShare;
    const win = !useShare ? window.open('about:blank', '_blank') : null;
    let state = 'fail';
    await App.run('جارٍ تجهيز الملف لإرساله عبر واتساب…', async () => {
      const f = await App.getFile(S, S.activeDoc, fmtOf(S));
      const file = new File([f.blob], f.name, { type: f.blob.type });
      const text = subjectText(S) + '\n' + clubLine(S);
      if (useShare) {
        const r = await tryShare(file, subjectText(S), text);
        if (r === true) { state = 'ok'; g.toast('تم فتح المشاركة مع الملف'); return; }
        if (r === 'cancel') { state = 'cancel'; g.toast('أُلغيت المشاركة'); return; }
      }
      const saved = await App.saveFile(f.blob, f.name);
      if (!saved) { state = 'cancel'; g.toast('أُلغي تنزيل الملف'); return; }
      const num = (S.send.whatsapp || '').replace(/\D/g, '').replace(/^00/, '');
      const msg = text + '\n(الملف: ' + f.name + ')';
      const url = isMobile
        ? 'https://wa.me/' + num + '?text=' + encodeURIComponent(msg)
        : 'https://web.whatsapp.com/send?' + (num ? 'phone=' + num + '&' : '') + 'text=' + encodeURIComponent(msg);
      const openIt = () => window.open(url, '_blank');
      if (win) win.location.href = url;
      state = win ? 'ok' : 'popup';
      dialog('تم تجهيز رسالة واتساب', '<p>' + E('الملف «' + f.name + '» نُزّل إلى جهازك وفُتح واتساب ويب بنص الرسالة (يتضمن موضوع الكتاب).') + '</p>' +
        li(['اختر المحادثة المطلوبة إن لم تُحدَّد (أدخل رقم واتساب في هذه الصفحة لتُفتح مباشرة).', 'اضغط على أيقونة <b>المشبك (+)</b> ثم <b>مستند</b>.', 'اختر الملف <b>' + E(f.name) + '</b> من مجلد التنزيلات (أو اسحبه من شريط التنزيلات إلى المحادثة).', 'اضغط <b>إرسال</b>.']) +
        (state === 'popup' ? '<p class="warn-fit">حجب المتصفح فتح واتساب؛ اضغط الزر أدناه.</p>' : ''),
        [{ label: 'فتح واتساب ويب', primary: true, onClick: openIt }, { label: 'تنزيل الملف مرة أخرى', onClick: () => App.saveFile(f.blob, f.name) }]);
    });
    if (win && state !== 'ok') { try { if (win.location.href === 'about:blank') win.close(); } catch (e) { /* تجاهل */ } }
    if (state === 'fail') g.toast('تعذّر تجهيز الإرسال عبر واتساب، حاول مرة أخرى');
  }

  function init(getS) {
    document.getElementById('btnMail').addEventListener('click', () => sendMail(getS()));
    document.getElementById('btnWa').addEventListener('click', () => sendWhatsApp(getS()));
    if (!canShareFiles()) { const cb = document.querySelector('[data-bind="send.useShare"]'); if (cb) { cb.disabled = true; cb.closest('label').title = 'غير مدعومة في هذا المتصفح'; } }
  }

  g.App = Object.assign(App, { initSend: init });
})(window);
