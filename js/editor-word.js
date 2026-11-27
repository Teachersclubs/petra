/* ===== محرر Word المتقدم (تبويب «تحميل كتاب للتعديل»): قوائم وأدوات علوية، صفحات A4 حقيقية، مسطرة وهوامش، ذيل الموقع، إدراج صور وملفات ===== */
(function (g) {
  'use strict';
  const App = g.App, doc = document;
  const $ = (s, r) => (r || doc).querySelector(s), $$ = (s, r) => Array.from((r || doc).querySelectorAll(s));
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const readBuf = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(f); });
  const readUrl = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  const readTxt = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsText(f, 'utf-8'); });
  const toast = m => g.toast(m);
  const ed = () => $('#wdPaper');
  const FONTS = () => (App.FONTS || []).map(f => f[0]);
  const CM = 96 / 2.54, SIZES = { a4: [210, 297], a3: [297, 420], letter: [215.9, 279.4] };
  let E0 = null, savedRange = null, selImg = null, IMGMAX = 560, TBLW = 9500;
  const pageDefaults = () => ({ size: 'a4', orient: 'p', mt: 2, mb: 2, ml: 2, mr: 2, guides: true, ruler: true, zoom: 100, siteFooter: false, hdr: '', ftr: '', pn: false });
  const PG = () => E0.word.page || (E0.word.page = pageDefaults());
  function dimsPx() { const p = PG(), s = SIZES[p.size] || SIZES.a4, W = s[0] / 25.4 * 96, H = s[1] / 25.4 * 96; return p.orient === 'l' ? [H, W] : [W, H]; }

  /* ---------- تحميل / مسودة ---------- */
  async function load(f, E) {
    E0 = E;
    if (!g.mammoth) throw new Error('مكتبة Word غير محمّلة');
    const r = await g.mammoth.convertToHtml({ arrayBuffer: await readBuf(f) }, { styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Heading 1'] => h2:fresh"] });
    E.kind = 'docx'; E.name = f.name; E.pages = []; E.srcs = []; E.undo = []; E.redo = []; E.sel = null; E.psel.clear();
    E.word = { html: r.value || '<p><br></p>', page: pageDefaults() };
    App.editor.show(true);
    toast('تم فتح مستند Word' + (r.messages.length ? ' (بعض التنسيقات المعقدة قد تُبسَّط)' : ''));
  }
  const DK = 'wd-draft-v1';
  const draftInfo = () => { try { const d = JSON.parse(localStorage.getItem(DK) || 'null'); return d && d.html ? { name: d.name || 'مسودة', ts: d.ts } : null; } catch (e) { return null; } };
  function saveDraft(quiet) { try { if (!E0 || E0.kind !== 'docx' || !ed()) return; localStorage.setItem(DK, JSON.stringify({ name: E0.name, html: snapshot(), page: PG(), ts: Date.now() })); if (!quiet) toast('تم حفظ المسودة في هذا المتصفح'); } catch (e) { if (!quiet) toast('تعذّر حفظ المسودة (الحجم كبير)'); } }
  function restoreDraft(E) { try { const d = JSON.parse(localStorage.getItem(DK)); E0 = E; E.kind = 'docx'; E.name = d.name; E.pages = []; E.srcs = []; E.word = { html: d.html, page: Object.assign(pageDefaults(), d.page || {}) }; App.editor.show(true); toast('تمت استعادة المسودة'); } catch (e) { toast('لا توجد مسودة'); } }
  function blank(E) { E0 = E; E.word = { html: '<p><br></p>', page: pageDefaults() }; }
  setInterval(() => { if (E0 && E0.kind === 'docx' && E0.word && E0.word.dirty && ed()) { E0.word.dirty = false; saveDraft(true); } }, 4000);

  /* ---------- لقطة نظيفة للمحتوى (بدون فواصل الترقيم) ---------- */
  function snapshot() {
    const p = ed(); if (!p || !E0 || !E0.word) return E0 && E0.word ? E0.word.html : '';
    const c = p.cloneNode(true); $$('[data-pg]', c).forEach(n => { n.style.marginTop = n.dataset.pgb || ''; if (!n.getAttribute('style')) n.removeAttribute('style'); n.removeAttribute('data-pg'); n.removeAttribute('data-pgb'); });
    $$('.sel-img', c).forEach(n => n.classList.remove('sel-img'));
    return (E0.word.html = c.innerHTML);
  }

  /* ---------- الواجهة ---------- */
  const IC = { u: '↶', r: '↷', pr: '🖨', b: '<b>B</b>', i: '<i>I</i>', un: '<u>U</u>', s: '<s>S</s>' };
  const MENUS = [
    ['ملف', [['جديد (مستند فارغ)', 'new'], ['فتح ملف…', 'open'], ['-'], ['تنزيل Word (.docx)', 'dlWord', 'Ctrl+S'], ['تنزيل PDF', 'dlPdf'], ['طباعة', 'print', 'Ctrl+P'], ['تنزيل نص (.txt)', 'dlTxt'], ['تنزيل صفحة ويب (.html)', 'dlHtml'], ['-'], ['حفظ مسودة في المتصفح', 'saveDraft'], ['إغلاق الملف', 'close']]],
    ['تعديل', [['تراجع', 'undo', 'Ctrl+Z'], ['إعادة', 'redo', 'Ctrl+Y'], ['-'], ['قص', 'cut'], ['نسخ', 'copy'], ['لصق', 'paste'], ['لصق كنص عادي', 'pastePlain'], ['تحديد الكل', 'selAll'], ['-'], ['بحث واستبدال…', 'find', 'Ctrl+F'], ['مسح التنسيق', 'clearFmt']]],
    ['عرض', [['حدود صفحات A4 (تسطير)', 'tGuides'], ['المسطرة والهوامش', 'tRuler'], ['ذيل الصفحة (من إعدادات الموقع)', 'tFooter'], ['-'], ['تكبير', 'zoomIn'], ['تصغير', 'zoomOut'], ['100%', 'zoom100'], ['ملاءمة العرض', 'zoomFit']]],
    ['إدراج', [['صورة من الجهاز…', 'insImg'], ['صورة من رابط…', 'insImgUrl'], ['شعار النادي', 'insLogoC'], ['شعار الوزارة', 'insLogoM'], ['توقيع الرئيس', 'insSignH'], ['ختم الرئيس', 'insStampH'], ['توقيع المدير', 'insSignM'], ['ختم المدير', 'insStampM'], ['-'], ['صفحات PDF كصور…', 'insPdfImg'], ['نص ملف PDF…', 'insPdfTxt'], ['محتوى ملف Word…', 'insDocx'], ['ملف نصي…', 'insTxt'], ['-'], ['جدول 2×2', 'tbl2'], ['جدول 3×3', 'tbl3'], ['جدول 4×4', 'tbl4'], ['جدول 5×5', 'tbl5'], ['جدول مخصص…', 'tblN'], ['-'], ['خط فاصل', 'hr'], ['فاصل صفحة', 'pb'], ['رابط…', 'link'], ['تاريخ اليوم', 'date'], ['مربع ملاحظة', 'callout'], ['قائمة مهام ☐', 'check'], ['-'], ['رموز…', 'symbols']]],
    ['تنسيق', [['غامق', 'bold', 'Ctrl+B'], ['مائل', 'italic', 'Ctrl+I'], ['تسطير', 'underline', 'Ctrl+U'], ['يتوسطه خط', 'strike'], ['منخفض', 'sub'], ['مرتفع', 'sup'], ['-'], ['عنوان 1', 'h1'], ['عنوان 2', 'h2'], ['عنوان 3', 'h3'], ['نص عادي', 'p'], ['اقتباس', 'quote'], ['-'], ['محاذاة يمين', 'jr'], ['توسيط', 'jc'], ['محاذاة يسار', 'jl'], ['ضبط', 'jf'], ['-'], ['اتجاه عربي RTL', 'rtl'], ['اتجاه إنجليزي LTR', 'ltr'], ['-'], ['مسافة أول السطر', 'fi'], ['زيادة المسافة البادئة', 'indent'], ['إنقاص المسافة البادئة', 'outdent']]],
    ['جدول', [['إضافة صف أسفل', 'rowB'], ['إضافة صف أعلى', 'rowA'], ['إضافة عمود يسار', 'colL'], ['إضافة عمود يمين', 'colR'], ['-'], ['حذف الصف', 'delRow'], ['حذف العمود', 'delCol'], ['حذف الجدول', 'delTbl'], ['-'], ['تظليل الخلية', 'shade'], ['تظليل الصف', 'shadeRow'], ['صف العناوين (غامق)', 'hdrRow'], ['-'], ['إظهار/إخفاء الحدود', 'borders'], ['توزيع الأعمدة بالتساوي', 'eqCols']]],
    ['أدوات', [['إحصاءات المستند', 'stats'], ['تحويل الأرقام إلى هندية ٠١٢', 'toHindi'], ['تحويل الأرقام إلى عربية 012', 'toWest'], ['حذف الفقرات الفارغة', 'cleanEmpty'], ['حذف المسافات المزدوجة', 'cleanSp'], ['توحيد الخط لكل المستند', 'allFont'], ['توحيد حجم الخط لكل المستند', 'allSize'], ['-'], ['تفعيل/إيقاف المدقق الإملائي', 'spell'], ['استعادة المسودة المحفوظة', 'restore']]]
  ];
  function showStage(st) {
    const P = PG();
    st.innerHTML =
      '<div class="wd-menus" id="wdMenus">' + MENUS.map((m, i) => '<div class="wd-mn"><button type="button" class="wd-mb" data-m="' + i + '">' + m[0] + '</button><div class="wd-dd" hidden>' + m[1].map(it => it[0] === '-' ? '<hr>' : '<button type="button" data-a="' + it[1] + '"><span>' + it[0] + '</span>' + (it[2] ? '<kbd>' + it[2] + '</kbd>' : '') + '</button>').join('') + '</div></div>').join('') +
      '<span class="wd-name" id="wdName"></span></div>' +
      '<div class="wd-bar" id="wdBar">' +
      '<button type="button" class="ed-tb wd-b" data-cmd="undo" title="تراجع">↶</button><button type="button" class="ed-tb wd-b" data-cmd="redo" title="إعادة">↷</button><button type="button" class="ed-tb" data-a="print" title="طباعة">🖨</button><span class="wd-sep"></span>' +
      '<select id="wdBlock" title="النمط"><option value="p">نص عادي</option><option value="h1">عنوان 1</option><option value="h2">عنوان 2</option><option value="h3">عنوان 3</option><option value="blockquote">اقتباس</option></select>' +
      '<select id="wdFont" title="الخط">' + FONTS().map(f => '<option>' + esc(f) + '</option>').join('') + '</select>' +
      '<button type="button" class="ed-tb" id="wdSzM" title="تصغير الخط">−</button><select id="wdSize" title="الحجم">' + [8, 9, 10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 36, 48, 72].map(n => '<option' + (n === 14 ? ' selected' : '') + '>' + n + '</option>').join('') + '</select><button type="button" class="ed-tb" id="wdSzP" title="تكبير الخط">+</button><span class="wd-sep"></span>' +
      [['bold', '<b>B</b>', 'غامق'], ['italic', '<i>I</i>', 'مائل'], ['underline', '<u>U</u>', 'تسطير'], ['strikeThrough', '<s>S</s>', 'يتوسطه خط'], ['subscript', 'x₂', 'منخفض'], ['superscript', 'x²', 'مرتفع']].map(b => '<button type="button" class="ed-tb wd-b" data-cmd="' + b[0] + '" title="' + b[2] + '">' + b[1] + '</button>').join('') +
      '<label class="ed-tb" title="لون الخط">A<input type="color" id="wdCol" value="#000000"></label><label class="ed-tb" title="تظليل النص">🖍<input type="color" id="wdHl" value="#ffff00"></label><span class="wd-sep"></span>' +
      [['justifyRight', '⇥', 'محاذاة يمين'], ['justifyCenter', '≡', 'توسيط'], ['justifyLeft', '⇤', 'محاذاة يسار'], ['justifyFull', '☰', 'ضبط'], ['insertUnorderedList', '•', 'قائمة نقطية'], ['insertOrderedList', '1.', 'قائمة رقمية'], ['outdent', '←', 'إنقاص المسافة'], ['indent', '→', 'زيادة المسافة']].map(b => '<button type="button" class="ed-tb wd-b" data-cmd="' + b[0] + '" title="' + b[2] + '">' + b[1] + '</button>').join('') +
      '<button type="button" class="ed-tb" data-a="rtl" title="اتجاه عربي">RTL</button><button type="button" class="ed-tb" data-a="ltr" title="اتجاه إنجليزي">LTR</button>' +
      '<select id="wdLs" title="تباعد الأسطر"><option value="">تباعد الأسطر</option><option>1</option><option>1.15</option><option>1.5</option><option>2</option><option>2.5</option></select>' +
      '<select id="wdSp" title="فراغ بعد الفقرة"><option value="">فراغ بعد</option><option value="0">0</option><option value="6">6 نقاط</option><option value="12">12 نقطة</option><option value="18">18 نقطة</option></select><span class="wd-sep"></span>' +
      '<button type="button" class="ed-tb" data-a="insImg" title="إدراج صورة">🖼</button><button type="button" class="ed-tb" data-a="tbl3" title="إدراج جدول 3×3">▦</button><button type="button" class="ed-tb" data-a="link" title="رابط">🔗</button><button type="button" class="ed-tb" data-a="pb" title="فاصل صفحة">⏎</button><button type="button" class="ed-tb wd-b" data-cmd="removeFormat" title="مسح التنسيق">Tx</button></div>' +
      '<div class="wd-ctx" id="wdCtx" hidden></div>' +
      '<div class="wd-find" id="wdFind" hidden><input id="wdFq" type="text" placeholder="بحث"><input id="wdFr" type="text" placeholder="استبدال بـ"><label class="check"><input id="wdFc" type="checkbox"> <span>مطابقة الحالة</span></label><button type="button" class="btn sm" data-a="fPrev">السابق</button><button type="button" class="btn sm" data-a="fNext">التالي</button><button type="button" class="btn sm" data-a="fRep">استبدال</button><button type="button" class="btn sm" data-a="fAll">استبدال الكل</button><output id="wdFn"></output><button type="button" class="btn sm" data-a="fClose">✕</button></div>' +
      '<div class="wd-view" id="wdView"><div class="wd-zoomwrap" id="wdZoom"><div class="wd-ruler" id="wdRuler"></div><div class="wd-page" id="wdPage"><div class="wd-paper" id="wdPaper" contenteditable="true" dir="rtl" spellcheck="false"></div><div class="wd-over" id="wdOver"></div></div></div></div>' +
      '<div class="wd-status" id="wdStatus"><span id="wdPos"></span><span id="wdCnt"></span><span class="wd-sp"></span>' +
      '<button type="button" class="ed-tb" data-a="insImg">🖼 صورة</button><button type="button" class="ed-tb" data-a="insPdfImg">📄 PDF كصور</button><button type="button" class="ed-tb" data-a="insDocx">📝 ملف Word</button><button type="button" class="ed-tb" data-a="insTxt">📃 ملف نصي</button><button type="button" class="ed-tb" data-a="tbl3">▦ جدول</button><button type="button" class="ed-tb" data-a="pb">⏎ فاصل صفحة</button><button type="button" class="ed-tb" data-a="tFooter" id="wdFtBtn">▁ ذيل الصفحة</button><span class="wd-sp"></span>' +
      '<button type="button" class="ed-tb" data-a="zoomOut">−</button><input type="range" id="wdZr" min="40" max="200" step="5" value="' + P.zoom + '"><button type="button" class="ed-tb" data-a="zoomIn">+</button><output id="wdZv"></output></div>' +
      '<input type="file" id="wdImgIn" accept="image/*" multiple hidden><input type="file" id="wdPdfImgIn" accept=".pdf,application/pdf" hidden><input type="file" id="wdPdfTxtIn" accept=".pdf,application/pdf" hidden><input type="file" id="wdDocxIn" accept=".docx" hidden><input type="file" id="wdTxtIn" accept=".txt,text/plain" hidden><input type="file" id="wdOpenIn" accept=".pdf,.docx" hidden><input type="file" id="wdRepImg" accept="image/*" hidden>';
    $('#wdName').textContent = E0.name || '';
    const p = ed(); p.innerHTML = E0.word.html; if (!p.textContent.trim() && !p.querySelector('img,table')) p.innerHTML = '<p><br></p>';
    bind(); applyPage(true);
  }
  const saveSel = () => { const s = g.getSelection(); if (s.rangeCount && ed() && ed().contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange(); };
  const restoreSel = () => { const p = ed(); p.focus({ preventScroll: true }); if (savedRange) { const s = g.getSelection(); s.removeAllRanges(); s.addRange(savedRange); } };
  function cmd(c, v) { restoreSel(); doc.execCommand('styleWithCSS', false, true); doc.execCommand(c, false, v == null ? null : v); saveSel(); changed(); }
  function changed() { E0.word.dirty = true; schedule(); }
  let ptimer = 0; const schedule = () => { clearTimeout(ptimer); ptimer = setTimeout(() => { paginate(); stats(); }, 220); };

  /* ---------- الصفحة والتقسيم A4 ---------- */
  function footerInfo() {
    const S = App.getState(), P = PG();
    if (!P.siteFooter || !App.footHasContent(S)) return null;
    const html = App.footerHTML(S, { footSize: S.footer.size, footGap: null }); if (!html) return null;
    const [W] = dimsPx(), h = doc.createElement('div'); h.style.cssText = 'position:fixed;left:-9999px;top:0;width:' + (W - (P.ml + P.mr) * CM) + 'px;font-family:' + S.footer.font + ',Arial'; h.innerHTML = html; doc.body.appendChild(h);
    const hh = h.offsetHeight; h.remove(); return { html, h: hh };
  }
  const mbEff = fi => Math.max(PG().mb * CM, fi ? fi.h + 0.8 * CM : 0);
  function applyPage(first) {
    const P = PG(), [W, H] = dimsPx(), pg = $('#wdPage'), p = ed(), fi = footerInfo();
    pg.style.width = W + 'px'; p.style.cssText = 'padding:' + (P.mt * CM) + 'px ' + (P.mr * CM) + 'px ' + mbEff(fi) + 'px ' + (P.ml * CM) + 'px;min-height:' + H + 'px';
    IMGMAX = Math.max(100, Math.round(W - (P.ml + P.mr) * CM)); TBLW = Math.round((W - (P.ml + P.mr) * CM) * 15);
    $('#wdZoom').style.zoom = P.zoom / 100; $('#wdZv').textContent = P.zoom + '%'; $('#wdZr').value = P.zoom;
    $('#wdRuler').hidden = !P.ruler; buildRuler();
    const b = $('#wdFtBtn'); b && b.classList.toggle('on', P.siteFooter);
    syncSide(); paginate();
  }
  function paginate() {
    const P = PG(), [W, H] = dimsPx(), p = ed(), over = $('#wdOver'); if (!p || !over) return;
    const fi = footerInfo(), mt = P.mt * CM, mb = mbEff(fi), ch = H - mt - mb;
    p.style.paddingBottom = mb + 'px';
    Array.from(p.children).forEach(c => { if (c.dataset.pg) { c.style.marginTop = c.dataset.pgb || ''; delete c.dataset.pg; delete c.dataset.pgb; } });
    let guard = 0;
    Array.from(p.children).forEach(c => {
      if (!c.offsetHeight) return; const y = c.offsetTop, h = c.offsetHeight, k = Math.max(0, Math.floor(y / H)), top = k * H + mt, bot = (k + 1) * H - mb; let push = 0;
      if (y < top && k > 0) push = top - y; else if (y > bot) push = (k + 1) * H + mt - y; else if (y + h > bot && h <= ch && y > top + 2) push = (k + 1) * H + mt - y;
      if (push > 0.5 && guard++ < 3000) { const cs = parseFloat(getComputedStyle(c).marginTop) || 0; c.dataset.pgb = c.style.marginTop || ''; c.dataset.pg = '1'; c.style.marginTop = (cs + push) + 'px'; }
    });
    const last = p.lastElementChild, content = last ? last.offsetTop + last.offsetHeight : mt, pages = Math.max(1, Math.ceil((content + mb) / H));
    p.style.minHeight = pages * H + 'px'; $('#wdPage').dataset.pages = pages; E0.word.pages = pages;
    let o = '';
    for (let k = 0; k < pages; k++) {
      if (k && P.guides) o += '<div class="wd-gap" style="top:' + (k * H - 7) + 'px"></div>';
      if (P.guides) o += '<div class="wd-pn" style="top:' + (k * H + 4) + 'px">' + (k + 1) + '</div>';
      if (fi) o += '<div class="wd-ft" style="top:' + ((k + 1) * H - fi.h - 0.6 * CM) + 'px;left:' + P.ml * CM + 'px;right:' + P.mr * CM + 'px" dir="rtl">' + fi.html + '</div>';
      if (P.hdr) o += '<div class="wd-hd" style="top:' + (k * H + 0.7 * CM) + 'px;left:' + P.ml * CM + 'px;right:' + P.mr * CM + 'px">' + esc(P.hdr) + '</div>';
      if (P.ftr || P.pn) o += '<div class="wd-tx" style="top:' + ((k + 1) * H - 0.75 * CM - (fi ? fi.h + 0.5 * CM : 0)) + 'px;left:' + P.ml * CM + 'px;right:' + P.mr * CM + 'px">' + (P.ftr ? esc(P.ftr) : '') + (P.ftr && P.pn ? ' &nbsp;|&nbsp; ' : '') + (P.pn ? 'صفحة ' + (k + 1) + ' من ' + pages : '') + '</div>';
    }
    over.innerHTML = o; over.style.height = pages * H + 'px'; $('#wdPage').style.height = pages * H + 'px'; $('#wdPage').classList.toggle('guides', !!P.guides);
    updPos();
  }
  function buildRuler() {
    const P = PG(), [W] = dimsPx(), r = $('#wdRuler'); if (!r) return; r.style.width = W + 'px';
    let t = ''; for (let i = 0; i <= Math.floor(W / CM); i++) t += '<span style="left:' + (i * CM) + 'px">' + (i % 2 === 0 ? i : '') + '</span>';
    r.innerHTML = t + '<i class="wd-rm l" style="left:' + P.ml * CM + 'px" title="الهامش الأيسر"></i><i class="wd-rm r" style="right:' + P.mr * CM + 'px" title="الهامش الأيمن"></i>';
  }
  function updPos() {
    const P = PG(), [, H] = dimsPx(), s = g.getSelection(); let k = 1;
    if (s.rangeCount && ed().contains(s.anchorNode)) { const r = s.getRangeAt(0).getBoundingClientRect(), pr = ed().getBoundingClientRect(), z = P.zoom / 100; if (r && (r.height || r.top)) k = Math.max(1, Math.floor((r.top - pr.top) / z / H) + 1); }
    const el = $('#wdPos'); if (el) el.textContent = 'صفحة ' + Math.min(k, E0.word.pages || 1) + ' من ' + (E0.word.pages || 1) + '  •  ' + (P.size === 'a4' ? 'A4' : P.size.toUpperCase()) + (P.orient === 'l' ? ' أفقي' : ' عمودي');
  }
  function stats() { const t = (ed().innerText || '').trim(), el = $('#wdCnt'); if (el) el.textContent = (t ? t.split(/\s+/).length : 0) + ' كلمة  •  ' + t.length + ' حرف'; updPos(); return { words: t ? t.split(/\s+/).length : 0, chars: t.length }; }
  function setPage(o) { Object.assign(PG(), o); applyPage(); }
  function setZoom(z) { setPage({ zoom: Math.max(40, Math.min(200, Math.round(z))) }); }

  /* ---------- الربط والأوامر ---------- */
  const blocksSel = () => { restoreSel(); const s = g.getSelection(); if (!s.rangeCount) return []; const r = s.getRangeAt(0), all = $$('p,h1,h2,h3,h4,h5,h6,li,blockquote,div', ed()).filter(n => r.intersectsNode(n) && !n.querySelector('p,h1,h2,h3,h4,h5,h6,li,div,table')); if (all.length) return all; let n = s.anchorNode; while (n && n.parentNode !== ed()) n = n.parentNode; return n && n.nodeType === 1 ? [n] : []; };
  function setDir(d) { blocksSel().forEach(n => { n.setAttribute('dir', d); n.style.textAlign = d === 'rtl' ? 'right' : 'left'; }); changed(); }
  const html = h => { restoreSel(); doc.execCommand('insertHTML', false, h); saveSel(); changed(); };
  function insertImg(url, w) {
    return new Promise(res => { const im = new Image(); im.onload = () => { const ww = Math.min(w || 400, im.naturalWidth || 400, IMGMAX); html('<img src="' + url + '" style="width:' + ww + 'px;max-width:100%">'); res(); }; im.onerror = () => { toast('تعذّر قراءة الصورة'); res(); }; im.src = url; });
  }
  function insertTable(r, c) {
    let h = '<table style="border-collapse:collapse;width:100%"><tbody>';
    for (let i = 0; i < r; i++) { h += '<tr>'; for (let j = 0; j < c; j++) h += '<td style="border:1px solid #444;padding:4px 6px"><br></td>'; h += '</tr>'; }
    html(h + '</tbody></table><p><br></p>');
  }
  const curCell = () => { restoreSel(); let n = g.getSelection().anchorNode; while (n && n !== ed()) { if (n.nodeType === 1 && /^(TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; } return null; };
  function tableOp(op) {
    const td = curCell(); if (!td) { toast('ضع المؤشر داخل خلية جدول أولاً'); return; }
    const tr = td.parentNode, tb = tr.parentNode, tbl = td.closest('table'), ci = Array.from(tr.children).indexOf(td), col = ($('#wdShadeC') || {}).value || '#d9e2f3';
    const mk = () => { const c = doc.createElement('td'); c.style.cssText = 'border:1px solid #444;padding:4px 6px'; c.innerHTML = '<br>'; return c; };
    if (op === 'rowB' || op === 'rowA') { const n = doc.createElement('tr'); Array.from(tr.children).forEach(() => n.appendChild(mk())); op === 'rowB' ? tr.after(n) : tr.before(n); }
    else if (op === 'colL' || op === 'colR') $$('tr', tbl).forEach(r => { const c = mk(), ref = r.children[ci]; ref ? (op === 'colR' ? ref.before(c) : ref.after(c)) : r.appendChild(c); });
    else if (op === 'delRow') { tr.remove(); if (!$$('tr', tbl).length) tbl.remove(); }
    else if (op === 'delCol') { $$('tr', tbl).forEach(r => r.children[ci] && r.children[ci].remove()); if (!tr.children.length) tbl.remove(); }
    else if (op === 'delTbl') tbl.remove();
    else if (op === 'shade') td.style.backgroundColor = col;
    else if (op === 'shadeRow') Array.from(tr.children).forEach(c => { c.style.backgroundColor = col; });
    else if (op === 'hdrRow') Array.from(tr.children).forEach(c => { c.style.fontWeight = 700; c.style.backgroundColor = c.style.backgroundColor || '#eeeeee'; });
    else if (op === 'borders') { const on = $$('td,th', tbl).some(c => /none|^$/.test(c.style.borderStyle || '') && !/solid/.test(c.style.border)); $$('td,th', tbl).forEach(c => { c.style.border = on ? '1px solid #444' : '1px solid transparent'; }); }
    else if (op === 'eqCols') { const n = tr.children.length; $$('tr', tbl).forEach(r => Array.from(r.children).forEach(c => { c.style.width = (100 / n).toFixed(2) + '%'; })); }
    changed(); ctxUpdate();
  }
  const dateTxt = () => { const iso = App.todayISO ? App.todayISO() : new Date().toISOString().slice(0, 10), [y, m, d] = iso.split('-'); return d + '/' + m + '/' + y + 'م' + (App.formatHijri ? ' الموافق ' + App.formatHijri(iso) : ''); };
  const pick = (id, fn) => { const i = $('#' + id); i.onchange = async () => { const fs = Array.from(i.files); i.value = ''; if (fs.length) await fn(fs); }; };
  const runBusy = (m, fn) => App.run(m, async () => { await fn((x) => App.setMsg && App.setMsg(x)); });
  function find(next, back) {
    const q = $('#wdFq').value, cs = $('#wdFc').checked; clearHl(); if (!q) { $('#wdFn').textContent = ''; return; }
    const w = doc.createTreeWalker(ed(), NodeFilter.SHOW_TEXT), rs = []; const Q = cs ? q : q.toLowerCase();
    while (w.nextNode()) { const t = w.currentNode, v = cs ? t.nodeValue : t.nodeValue.toLowerCase(); for (let i = v.indexOf(Q); i >= 0; i = v.indexOf(Q, i + Q.length)) { const r = doc.createRange(); r.setStart(t, i); r.setEnd(t, i + q.length); rs.push(r); } }
    find.rs = rs; $('#wdFn').textContent = rs.length + ' نتيجة'; if (!rs.length) return;
    if (g.CSS && CSS.highlights && g.Highlight) CSS.highlights.set('wdfind', new Highlight(...rs));
    find.i = next ? ((find.i == null ? -1 : find.i) + (back ? -1 : 1) + rs.length) % rs.length : 0;
    const r = rs[find.i], s = g.getSelection(); s.removeAllRanges(); s.addRange(r); const el = r.startContainer.parentElement; el && el.scrollIntoView({ block: 'center' });
  }
  const clearHl = () => { try { g.CSS && CSS.highlights && CSS.highlights.delete('wdfind'); } catch (e) { } };
  function replaceOne() { const rs = find.rs || [], s = g.getSelection(); if (find.i != null && rs[find.i] && s.rangeCount && !s.isCollapsed) { doc.execCommand('insertText', false, $('#wdFr').value); changed(); } find(false); }
  function replaceAll() {
    const q = $('#wdFq').value, rep = $('#wdFr').value, cs = $('#wdFc').checked; if (!q) return; let n = 0; const w = doc.createTreeWalker(ed(), NodeFilter.SHOW_TEXT), ns = []; while (w.nextNode()) ns.push(w.currentNode);
    const re = new RegExp(q.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), cs ? 'g' : 'gi'); ns.forEach(t => { t.nodeValue = t.nodeValue.replace(re, () => { n++; return rep; }); });
    $('#wdFn').textContent = 'تم استبدال ' + n; changed();
  }
  function mapText(fn) { const w = doc.createTreeWalker(ed(), NodeFilter.SHOW_TEXT), ns = []; while (w.nextNode()) ns.push(w.currentNode); ns.forEach(t => { t.nodeValue = fn(t.nodeValue); }); changed(); }
  const dlText = () => App.saveFile(new Blob([ed().innerText], { type: 'text/plain;charset=utf-8' }), nameOf() + '.txt');
  const dlHtml = () => App.saveFile(new Blob(['<!doctype html><html dir="rtl" lang="ar"><meta charset="utf-8"><body>' + snapshot() + '</body></html>'], { type: 'text/html;charset=utf-8' }), nameOf() + '.html');
  const dlWord = () => App.run('جارٍ تجهيز ملف Word…', async () => { App.saveFile(await toDocx(), nameOf() + '-معدّل.docx'); });
  const dlPdf = () => App.run('جارٍ تجهيز ملف PDF…', async () => { App.saveFile(await toPdf(), nameOf() + '-معدّل.pdf'); });
  function setBlock(t) { restoreSel(); doc.execCommand('formatBlock', false, t); changed(); }
  const ACT = {
    new() { if (confirm('سيُستبدل المحتوى الحالي بمستند فارغ. متابعة؟')) { ed().innerHTML = '<p><br></p>'; changed(); } },
    open() { $('#wdOpenIn').click(); }, dlWord, dlPdf, print() { print(E0); }, dlTxt: dlText, dlHtml, saveDraft() { saveDraft(); }, close() { App.editor.close(); },
    undo() { cmd('undo'); }, redo() { cmd('redo'); }, cut() { restoreSel(); doc.execCommand('cut'); changed(); },
    copy() { restoreSel(); doc.execCommand('copy'); },
    async paste() { try { const t = await navigator.clipboard.readText(); restoreSel(); doc.execCommand('insertText', false, t); changed(); } catch (e) { toast('اضغط Ctrl+V للصق (المتصفح يمنع اللصق من القائمة)'); } },
    async pastePlain() { ACT.paste(); }, selAll() { ed().focus(); doc.execCommand('selectAll'); saveSel(); },
    find() { const f = $('#wdFind'); f.hidden = false; $('#wdFq').focus(); const s = g.getSelection().toString(); if (s) $('#wdFq').value = s; }, fNext() { find(true); }, fPrev() { find(true, true); }, fRep: replaceOne, fAll: replaceAll, fClose() { $('#wdFind').hidden = true; clearHl(); },
    clearFmt() { cmd('removeFormat'); },
    tGuides() { setPage({ guides: !PG().guides }); }, tRuler() { setPage({ ruler: !PG().ruler }); },
    tFooter() { const S = App.getState(); if (!PG().siteFooter && !App.footHasContent(S)) { toast('لا يوجد ذيل مُعدّ في الموقع — أضفه من تبويب الإعدادات/الذيل'); return; } setPage({ siteFooter: !PG().siteFooter }); },
    zoomIn() { setZoom(PG().zoom + 10); }, zoomOut() { setZoom(PG().zoom - 10); }, zoom100() { setZoom(100); }, zoomFit() { const v = $('#wdView'), [W] = dimsPx(); setZoom((v.clientWidth - 40) / W * 100); },
    insImg() { saveSel(); $('#wdImgIn').click(); },
    async insImgUrl() { const u = prompt('رابط الصورة (https://…)'); if (!u) return; try { const b = await (await fetch(u)).blob(); await insertImg(await readUrl(b)); } catch (e) { toast('تعذّر تحميل الصورة من الرابط (قد يمنعه الموقع المصدر) — نزّلها ثم أدرجها من الجهاز'); } },
    insLogoC() { logo('club'); }, insLogoM() { logo('ministry'); }, insSignH() { person('head', 'signImg'); }, insStampH() { person('head', 'stampImg'); }, insSignM() { person('manager', 'signImg'); }, insStampM() { person('manager', 'stampImg'); },
    insPdfImg() { saveSel(); $('#wdPdfImgIn').click(); }, insPdfTxt() { saveSel(); $('#wdPdfTxtIn').click(); }, insDocx() { saveSel(); $('#wdDocxIn').click(); }, insTxt() { saveSel(); $('#wdTxtIn').click(); },
    tbl2() { insertTable(2, 2); }, tbl3() { insertTable(3, 3); }, tbl4() { insertTable(4, 4); }, tbl5() { insertTable(5, 5); },
    tblN() { const r = parseInt(prompt('عدد الصفوف', '3')), c = parseInt(prompt('عدد الأعمدة', '3')); if (r > 0 && c > 0) insertTable(Math.min(r, 60), Math.min(c, 12)); },
    hr() { html('<hr style="border:0;border-top:1px solid #444"><p><br></p>'); }, pb() { html('<hr data-pb="1" style="border:0;border-top:2px dashed #999;page-break-after:always"><p><br></p>'); },
    link() { saveSel(); const u = prompt('عنوان الرابط (https://…)'); if (u) cmd('createLink', u); }, date() { restoreSel(); doc.execCommand('insertText', false, dateTxt()); changed(); },
    callout() { html('<table style="border-collapse:collapse;width:100%"><tbody><tr><td style="border:1px solid #c9a227;background:#fff8e1;padding:8px 10px">ملاحظة: <br></td></tr></tbody></table><p><br></p>'); },
    check() { html('<p>☐ <br></p>'); },
    symbols() { const s = prompt('اكتب أو الصق الرموز المطلوبة، أو اختر من:\n✓ ✗ © ® ™ § ° ± × ÷ « » ٪ ؛ ؟ ← → • — … ☐ ☑ ★ ☎ ✉'); if (s) { restoreSel(); doc.execCommand('insertText', false, s); changed(); } },
    bold() { cmd('bold'); }, italic() { cmd('italic'); }, underline() { cmd('underline'); }, strike() { cmd('strikeThrough'); }, sub() { cmd('subscript'); }, sup() { cmd('superscript'); },
    h1() { setBlock('h1'); }, h2() { setBlock('h2'); }, h3() { setBlock('h3'); }, p() { setBlock('p'); }, quote() { setBlock('blockquote'); },
    jr() { cmd('justifyRight'); }, jc() { cmd('justifyCenter'); }, jl() { cmd('justifyLeft'); }, jf() { cmd('justifyFull'); },
    rtl() { setDir('rtl'); }, ltr() { setDir('ltr'); }, fi() { blocksSel().forEach(n => { n.style.textIndent = n.style.textIndent ? '' : '1.5cm'; }); changed(); },
    indent() { cmd('indent'); }, outdent() { cmd('outdent'); },
    rowB() { tableOp('rowB'); }, rowA() { tableOp('rowA'); }, colL() { tableOp('colL'); }, colR() { tableOp('colR'); }, delRow() { tableOp('delRow'); }, delCol() { tableOp('delCol'); }, delTbl() { tableOp('delTbl'); },
    shade() { tableOp('shade'); }, shadeRow() { tableOp('shadeRow'); }, hdrRow() { tableOp('hdrRow'); }, borders() { tableOp('borders'); }, eqCols() { tableOp('eqCols'); },
    stats() { const s = stats(); toast('الكلمات: ' + s.words + ' — الأحرف: ' + s.chars + ' — الصفحات: ' + (E0.word.pages || 1)); },
    toHindi() { mapText(t => t.replace(/[0-9]/g, d => '٠١٢٣٤٥٦٧٨٩'[d])); }, toWest() { mapText(t => t.replace(/[٠-٩]/g, d => String(d.charCodeAt(0) - 1632))); },
    cleanEmpty() { $$('p,div', ed()).forEach(n => { if (!n.textContent.trim() && !n.querySelector('img,table,hr,br:not(:only-child)') && n.parentNode === ed() && n !== ed().lastElementChild) n.remove(); }); changed(); },
    cleanSp() { mapText(t => t.replace(/[  ]{2,}/g, ' ')); },
    allFont() { const f = $('#wdFont').value; $$('*', ed()).forEach(n => { if (n.style && n.style.fontFamily) n.style.fontFamily = ''; }); ed().style.fontFamily = f; $$('p,h1,h2,h3,li,td,th,blockquote', ed()).forEach(n => { n.style.fontFamily = f; }); changed(); },
    allSize() { const s = $('#wdSize').value; $$('*', ed()).forEach(n => { if (n.style && n.style.fontSize) n.style.fontSize = ''; }); $$('p,li,td,th,blockquote', ed()).forEach(n => { n.style.fontSize = s + 'pt'; }); changed(); },
    spell() { const p = ed(); p.spellcheck = !p.spellcheck; p.blur(); p.focus(); toast('المدقق الإملائي: ' + (p.spellcheck ? 'مفعّل' : 'متوقف')); },
    restore() { if (draftInfo()) restoreDraft(E0); else toast('لا توجد مسودة محفوظة'); }
  };
  const logo = k => { const u = App.getState().logos[k] && App.getState().logos[k].data; u ? insertImg(u, 90) : toast('لا يوجد شعار محفوظ'); };
  const person = (w, f) => { const p = App.getState().people[w], u = p && p[f]; u ? insertImg(u, 110) : toast('لم يُضف في الإعدادات'); };
  function fontSizeStep(d) { const s = $('#wdSize'), i = Math.max(0, Math.min(s.options.length - 1, s.selectedIndex + d)); s.selectedIndex = i; applySize(s.value); }
  function applySize(v) { restoreSel(); doc.execCommand('styleWithCSS', false, false); doc.execCommand('fontSize', false, 7); $$('font[size="7"]', ed()).forEach(f => { f.removeAttribute('size'); f.style.fontSize = v + 'pt'; }); saveSel(); changed(); }
  function ctxUpdate() {
    const c = $('#wdCtx'); if (!c) return; let h = '';
    if (selImg && ed().contains(selImg)) {
      h = '<b>صورة:</b>' + [25, 50, 75, 100].map(n => '<button type="button" class="ed-tb" data-iw="' + n + '">' + n + '%</button>').join('') + '<label class="ed-tb">عرض px <input type="number" id="wdIw" min="20" max="2000" value="' + Math.round(selImg.getBoundingClientRect().width / (PG().zoom / 100)) + '" style="width:70px"></label><button type="button" class="ed-tb" data-ia="r">يمين</button><button type="button" class="ed-tb" data-ia="c">وسط</button><button type="button" class="ed-tb" data-ia="l">يسار</button><button type="button" class="ed-tb" data-ia="b">إطار</button><button type="button" class="ed-tb" data-ia="rep">استبدال</button><button type="button" class="ed-tb" data-ia="del">حذف</button>';
    } else if (curCellQuiet()) {
      h = '<b>جدول:</b><button type="button" class="ed-tb" data-tb="rowB">+صف</button><button type="button" class="ed-tb" data-tb="colR">+عمود</button><button type="button" class="ed-tb" data-tb="delRow">−صف</button><button type="button" class="ed-tb" data-tb="delCol">−عمود</button><label class="ed-tb">لون <input type="color" id="wdShadeC" value="#d9e2f3"></label><button type="button" class="ed-tb" data-tb="shade">تظليل الخلية</button><button type="button" class="ed-tb" data-tb="shadeRow">تظليل الصف</button><button type="button" class="ed-tb" data-tb="hdrRow">صف عناوين</button><button type="button" class="ed-tb" data-tb="borders">الحدود</button><button type="button" class="ed-tb" data-tb="eqCols">تساوي الأعمدة</button><button type="button" class="ed-tb" data-tb="delTbl">حذف الجدول</button>';
    }
    c.hidden = !h; c.innerHTML = h;
  }
  const curCellQuiet = () => { const s = g.getSelection(); let n = s.anchorNode; while (n && n !== ed()) { if (n.nodeType === 1 && /^(TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; } return null; };
  function bind() {
    const root = $('#edStage') || doc, p = ed();
    $$('.wd-mb').forEach(b => b.addEventListener('mousedown', e => { e.preventDefault(); saveSel(); const dd = b.nextElementSibling, open = !dd.hidden; $$('.wd-dd').forEach(d => d.hidden = true); dd.hidden = open; }));
    doc.addEventListener('mousedown', e => { if (!e.target.closest('.wd-mn')) $$('.wd-dd').forEach(d => d.hidden = true); });
    $$('[data-a]').forEach(b => { b.addEventListener('mousedown', e => { if (!/INPUT|SELECT/.test(e.target.tagName)) e.preventDefault(); saveSel(); }); b.addEventListener('click', () => { $$('.wd-dd').forEach(d => d.hidden = true); const f = ACT[b.dataset.a]; f && f(); }); });
    $$('.wd-b').forEach(b => { b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => cmd(b.dataset.cmd)); });
    $('#wdFont').addEventListener('change', e => cmd('fontName', e.target.value));
    $('#wdSize').addEventListener('change', e => applySize(e.target.value)); $('#wdSzM').onclick = () => fontSizeStep(-1); $('#wdSzP').onclick = () => fontSizeStep(1);
    $('#wdSzM').onmousedown = $('#wdSzP').onmousedown = e => e.preventDefault();
    $('#wdBlock').addEventListener('change', e => setBlock(e.target.value));
    $('#wdCol').addEventListener('input', e => cmd('foreColor', e.target.value)); $('#wdHl').addEventListener('input', e => cmd('hiliteColor', e.target.value));
    $('#wdLs').addEventListener('change', e => { if (e.target.value) { blocksSel().forEach(n => { n.style.lineHeight = e.target.value; }); changed(); } e.target.value = ''; });
    $('#wdSp').addEventListener('change', e => { if (e.target.value !== '') { blocksSel().forEach(n => { n.style.marginBottom = e.target.value + 'pt'; }); changed(); } e.target.value = ''; });
    $('#wdZr').addEventListener('input', e => setZoom(+e.target.value));
    $('#wdFq').addEventListener('input', () => find(false)); $('#wdFq').addEventListener('keydown', e => { if (e.key === 'Enter') { e.preventDefault(); find(true, e.shiftKey); } if (e.key === 'Escape') ACT.fClose(); });
    pick('wdImgIn', async fs => { for (const f of fs) { restoreSel(); await insertImg(await readUrl(f)); } });
    pick('wdRepImg', async fs => { if (selImg) { selImg.src = await readUrl(fs[0]); changed(); } });
    pick('wdPdfImgIn', fs => App.run('جارٍ تحويل صفحات PDF إلى صور…', async () => { const im = await App.editor.pdfToImages(fs[0], 1.6, m => App.setMsg && App.setMsg(m)); restoreSel(); for (const u of im) { await insertImg(u, IMGMAX); } toast('أُدرجت ' + im.length + ' صفحة'); }));
    pick('wdPdfTxtIn', fs => App.run('جارٍ استخراج نص PDF…', async () => { const h = await App.editor.pdfTextHTML(fs[0], m => App.setMsg && App.setMsg(m)); html(h); toast('أُدرج نص الملف'); }));
    pick('wdDocxIn', fs => App.run('جارٍ قراءة ملف Word…', async () => { const r = await g.mammoth.convertToHtml({ arrayBuffer: await readBuf(fs[0]) }); html(r.value || '<p><br></p>'); toast('أُدرج محتوى الملف'); }));
    pick('wdTxtIn', async fs => { const t = await readTxt(fs[0]); html(t.split(/\r?\n/).map(l => '<p dir="auto">' + (esc(l) || '<br>') + '</p>').join('')); });
    pick('wdOpenIn', fs => App.editor.openFile(fs[0]));
    doc.addEventListener('selectionchange', () => { if (ed()) { saveSel(); updPos(); } });
    $('#wdCtx').addEventListener('mousedown', e => { if (!/INPUT/.test(e.target.tagName)) e.preventDefault(); });
    $('#wdCtx').addEventListener('click', e => {
      const t = e.target.closest('button'); if (!t) return;
      if (t.dataset.tb) return tableOp(t.dataset.tb);
      if (t.dataset.iw && selImg) { selImg.style.width = Math.round(IMGMAX * t.dataset.iw / 100) + 'px'; selImg.style.maxWidth = '100%'; changed(); ctxUpdate(); }
      const a = t.dataset.ia; if (!a || !selImg) return;
      if (a === 'del') { selImg.remove(); selImg = null; changed(); ctxUpdate(); } else if (a === 'rep') $('#wdRepImg').click();
      else if (a === 'b') { selImg.style.border = selImg.style.border ? '' : '1px solid #444'; changed(); }
      else { const pr = selImg.parentNode; if (pr && pr !== ed()) pr.style.textAlign = { r: 'right', c: 'center', l: 'left' }[a]; else { const w = doc.createElement('p'); selImg.before(w); w.appendChild(selImg); w.style.textAlign = { r: 'right', c: 'center', l: 'left' }[a]; } changed(); }
    });
    $('#wdCtx').addEventListener('input', e => { if (e.target.id === 'wdIw' && selImg) { selImg.style.width = (+e.target.value || 100) + 'px'; selImg.style.maxWidth = '100%'; changed(); } });
    p.addEventListener('click', e => { $$('.sel-img', p).forEach(n => n.classList.remove('sel-img')); selImg = e.target.tagName === 'IMG' ? e.target : null; if (selImg) selImg.classList.add('sel-img'); ctxUpdate(); });
    p.addEventListener('keyup', () => ctxUpdate());
    p.addEventListener('input', changed);
    p.addEventListener('paste', e => { const items = e.clipboardData && e.clipboardData.items; if (!items) return; for (const it of items) if (/^image\//.test(it.type)) { e.preventDefault(); readUrl(it.getAsFile()).then(u => insertImg(u)); return; } });
    p.addEventListener('keydown', e => {
      if (!(e.ctrlKey || e.metaKey)) return; const k = e.key.toLowerCase();
      if (k === 's') { e.preventDefault(); dlWord(); } else if (k === 'p') { e.preventDefault(); print(E0); } else if (k === 'f') { e.preventDefault(); ACT.find(); }
    });
    p.addEventListener('dragover', e => { if (e.dataTransfer && Array.from(e.dataTransfer.types || []).includes('Files')) e.preventDefault(); });
    p.addEventListener('drop', async e => { const fs = Array.from((e.dataTransfer && e.dataTransfer.files) || []).filter(f => /^image\//.test(f.type)); if (fs.length) { e.preventDefault(); for (const f of fs) await insertImg(await readUrl(f)); } });
    // سحب الهوامش على المسطرة
    $('#wdRuler').addEventListener('pointerdown', e => {
      const h = e.target.closest('.wd-rm'); if (!h) return; e.preventDefault(); const left = h.classList.contains('l'), z = PG().zoom / 100, r = $('#wdRuler').getBoundingClientRect(), [W] = dimsPx();
      const mv = ev => { const x = (ev.clientX - r.left) / z, cm = Math.round((left ? x : W - x) / CM * 10) / 10; PG()[left ? 'ml' : 'mr'] = Math.max(0.5, Math.min(8, cm)); applyPage(); };
      const up = () => { doc.removeEventListener('pointermove', mv); doc.removeEventListener('pointerup', up); changed(); };
      doc.addEventListener('pointermove', mv); doc.addEventListener('pointerup', up);
    });
    g.addEventListener('resize', () => { if (ed()) paginate(); });
  }

  /* ---------- النموذج الجانبي ---------- */
  function buildForm(host, E) {
    E0 = E;
    host.innerHTML =
      '<fieldset class="group"><legend>٢) تعديل مستند Word</legend><p class="hint">قوائم وأدوات التحرير في أعلى المعاينة (مثل Google Docs). كل التعديلات تظهر مباشرة على صفحات A4.</p></fieldset>' +
      '<fieldset class="group"><legend>٣) إعدادات الصفحة</legend>' +
      '<div class="row2"><label class="field"><span>حجم الورق</span><select id="wsSize"><option value="a4">A4</option><option value="a3">A3</option><option value="letter">Letter</option></select></label><label class="field"><span>الاتجاه</span><select id="wsOri"><option value="p">عمودي</option><option value="l">أفقي</option></select></label></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-mp="2,2,2,2">هوامش عادية 2سم</button><button type="button" class="btn sm" data-mp="1.27,1.27,1.27,1.27">ضيقة</button><button type="button" class="btn sm" data-mp="2.54,2.54,3.17,3.17">واسعة</button></div>' +
      '<div class="row2"><label class="field"><span>أعلى (سم)</span><input id="wsMt" type="number" step="0.1" min="0.3" max="8"></label><label class="field"><span>أسفل (سم)</span><input id="wsMb" type="number" step="0.1" min="0.3" max="8"></label></div>' +
      '<div class="row2"><label class="field"><span>يمين (سم)</span><input id="wsMr" type="number" step="0.1" min="0.3" max="8"></label><label class="field"><span>يسار (سم)</span><input id="wsMl" type="number" step="0.1" min="0.3" max="8"></label></div>' +
      '<label class="check"><input id="wsGuides" type="checkbox"> <span>تسطير الصفحة (حدود A4 وأرقامها)</span></label>' +
      '<label class="check"><input id="wsRuler" type="checkbox"> <span>إظهار المسطرة وسحب الهوامش</span></label></fieldset>' +
      '<fieldset class="group"><legend>٤) الذيل والرأس</legend>' +
      '<label class="check"><input id="wsSiteFt" type="checkbox"> <span>إضافة ذيل الصفحة من إعدادات الموقع (في كل صفحة)</span></label>' +
      '<label class="field"><span>نص رأس الصفحة</span><input id="wsHdr" type="text" placeholder="اختياري"></label>' +
      '<label class="field"><span>نص تذييل إضافي</span><input id="wsFtr" type="text" placeholder="اختياري"></label>' +
      '<label class="check"><input id="wsPn" type="checkbox"> <span>ترقيم الصفحات (صفحة X من N)</span></label></fieldset>' +
      '<fieldset class="group"><legend>٥) تحويل وتنزيل</legend><div class="ed-btnrow"><button type="button" class="btn sm primary" id="wsPdf">تحويل إلى PDF</button><button type="button" class="btn sm" id="wsTxt">نص (.txt)</button><button type="button" class="btn sm" id="wsHtml">صفحة ويب</button><button type="button" class="btn sm" id="wsFind">بحث واستبدال</button></div><p class="hint">اختصارات: Ctrl+S تنزيل Word • Ctrl+P طباعة • Ctrl+F بحث.</p></fieldset>';
    const on = (id, ev, fn) => $('#' + id, host).addEventListener(ev, fn), num = (id, k) => on(id, 'input', e => { if (E0.word && e.target.value !== '') { setPage({ [k]: Math.max(0.3, Math.min(8, parseFloat(e.target.value) || 2)) }); changed(); } });
    on('wsSize', 'change', e => { E0.word && setPage({ size: e.target.value }); changed(); }); on('wsOri', 'change', e => { E0.word && setPage({ orient: e.target.value }); changed(); });
    num('wsMt', 'mt'); num('wsMb', 'mb'); num('wsMr', 'mr'); num('wsMl', 'ml');
    $$('[data-mp]', host).forEach(b => b.addEventListener('click', () => { if (!E0.word) return; const [t, bt, r, l] = b.dataset.mp.split(',').map(Number); setPage({ mt: t, mb: bt, mr: r, ml: l }); changed(); }));
    on('wsGuides', 'change', e => E0.word && setPage({ guides: e.target.checked })); on('wsRuler', 'change', e => E0.word && setPage({ ruler: e.target.checked }));
    on('wsSiteFt', 'change', e => { if (!E0.word) return; if (e.target.checked && !App.footHasContent(App.getState())) { toast('لا يوجد ذيل مُعدّ في الموقع'); e.target.checked = false; return; } setPage({ siteFooter: e.target.checked }); changed(); });
    on('wsHdr', 'input', e => { E0.word && setPage({ hdr: e.target.value }); changed(); }); on('wsFtr', 'input', e => { E0.word && setPage({ ftr: e.target.value }); changed(); }); on('wsPn', 'change', e => { E0.word && setPage({ pn: e.target.checked }); changed(); });
    on('wsPdf', 'click', () => E0.word && ed() && dlPdf()); on('wsTxt', 'click', () => ed() && dlText()); on('wsHtml', 'click', () => ed() && dlHtml()); on('wsFind', 'click', () => ed() && ACT.find());
  }
  function syncSide() {
    const P = PG(), v = (id, x) => { const e = $('#' + id); if (e && doc.activeElement !== e) e.value = x; }, c = (id, x) => { const e = $('#' + id); if (e) e.checked = !!x; };
    v('wsSize', P.size); v('wsOri', P.orient); v('wsMt', P.mt); v('wsMb', P.mb); v('wsMr', P.mr); v('wsMl', P.ml); v('wsHdr', P.hdr); v('wsFtr', P.ftr);
    c('wsGuides', P.guides); c('wsRuler', P.ruler); c('wsSiteFt', P.siteFooter); c('wsPn', P.pn);
  }
  const nameOf = () => (E0.name || 'مستند').replace(/\.(docx?|pdf)$/i, '');

  /* ---------- HTML → docx ---------- */
  const bytes = url => { const b = atob(url.split(',')[1]), u = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) u[i] = b.charCodeAt(i); return u; };
  const hex = c => { const m = /rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(c || ''); return m ? [m[1], m[2], m[3]].map(x => (+x).toString(16).padStart(2, '0')).join('').toUpperCase() : /^#([0-9a-f]{6})$/i.test(c || '') ? c.slice(1).toUpperCase() : null; };
  const isAr = s => /[؀-ۿ]/.test(s);
  function runsOf(node, fmt, D, out) {
    node.childNodes.forEach(ch => {
      if (ch.nodeType === 3) {
        const t = ch.nodeValue.replace(/\s+/g, ' '); if (!t) return;
        out.push(new D.TextRun({ text: t, bold: fmt.b, italics: fmt.i, underline: fmt.u ? {} : undefined, strike: fmt.s, color: fmt.c, font: fmt.f, size: fmt.sz, highlight: fmt.h ? 'yellow' : undefined, superScript: fmt.sup, subScript: fmt.sub, rightToLeft: isAr(t) || undefined, shading: fmt.bg ? { type: D.ShadingType.CLEAR, fill: fmt.bg, color: 'auto' } : undefined }));
      } else if (ch.nodeType === 1) {
        const tg = ch.tagName, st = ch.style, f = Object.assign({}, fmt);
        if (tg === 'BR') { out.push(new D.TextRun({ break: 1 })); return; }
        if (tg === 'IMG') { try { const src = ch.getAttribute('src') || ''; if (!/^data:image/.test(src)) return; const w = Math.round(ch.getBoundingClientRect().width || ch.naturalWidth || 200), h = Math.round(ch.getBoundingClientRect().height || ch.naturalHeight || 100); out.push(new D.ImageRun({ type: /jpe?g/.test(src.slice(0, 20)) ? 'jpg' : /gif/.test(src.slice(0, 20)) ? 'gif' : 'png', data: bytes(src), transformation: { width: Math.min(IMGMAX, w), height: Math.round(Math.min(IMGMAX, w) * h / Math.max(1, w)) } })); } catch (e) { /* تخطّي */ } return; }
        if (tg === 'B' || tg === 'STRONG' || +st.fontWeight >= 600 || st.fontWeight === 'bold') f.b = true;
        if (tg === 'I' || tg === 'EM' || st.fontStyle === 'italic') f.i = true;
        if (tg === 'U' || /underline/.test(st.textDecoration || st.textDecorationLine)) f.u = true;
        if (tg === 'S' || tg === 'STRIKE' || /line-through/.test(st.textDecoration || st.textDecorationLine)) f.s = true;
        if (tg === 'SUP') f.sup = true; if (tg === 'SUB') f.sub = true;
        if (st.color) f.c = hex(st.color) || f.c;
        if (st.backgroundColor) { const h = hex(st.backgroundColor); if (h === 'FFFF00') f.h = true; else if (h) f.bg = h; }
        if (st.fontFamily) f.f = st.fontFamily.split(',')[0].replace(/["']/g, '').trim();
        if (st.fontSize && /pt$/.test(st.fontSize)) f.sz = Math.round(parseFloat(st.fontSize) * 2); else if (st.fontSize && /px$/.test(st.fontSize)) f.sz = Math.round(parseFloat(st.fontSize) * 1.5);
        if (tg === 'FONT') { if (ch.getAttribute('face')) f.f = ch.getAttribute('face'); if (ch.getAttribute('color')) f.c = hex(ch.getAttribute('color')) || f.c; }
        if (tg === 'A') { const kids = []; runsOf(ch, Object.assign(f, { c: '0563C1', u: true }), D, kids); out.push(new D.ExternalHyperlink({ link: ch.getAttribute('href') || '#', children: kids })); return; }
        runsOf(ch, f, D, out);
      }
    });
  }
  const HEAD = { H1: [D => D.HeadingLevel.HEADING_1, 40], H2: [D => D.HeadingLevel.HEADING_2, 32], H3: [D => D.HeadingLevel.HEADING_3, 28], H4: [D => D.HeadingLevel.HEADING_4, 26], H5: [D => D.HeadingLevel.HEADING_5, 24], H6: [D => D.HeadingLevel.HEADING_6, 22] };
  function blockProps(el, D, base) {
    const lh = parseFloat(el.style.lineHeight), mb = el.style.marginBottom, ti = el.style.textIndent;
    const dir = el.getAttribute('dir') || (isAr(el.textContent) ? 'rtl' : 'ltr'), ta = (el.style.textAlign || '').toLowerCase();
    const rtl = dir === 'rtl'; let al;
    // في فقرة bidi: START = يمين للعربي، END = يسار
    if (ta === 'center') al = D.AlignmentType.CENTER; else if (ta === 'justify') al = D.AlignmentType.JUSTIFIED;
    else if (ta === 'left') al = rtl ? D.AlignmentType.END : D.AlignmentType.START; else if (ta === 'right') al = rtl ? D.AlignmentType.START : D.AlignmentType.END;
    const ml = parseFloat(el.style.marginRight || el.style.paddingRight || 0), mrl = parseFloat(el.style.marginLeft || el.style.paddingLeft || 0);
    const ind = rtl ? ml : mrl;
    return Object.assign({ bidirectional: rtl, alignment: al, spacing: { after: mb ? Math.round(parseFloat(mb) * (/px$/.test(mb) ? 15 : 20)) : 100, line: lh && lh < 10 ? Math.round(lh * 240) : 300 }, indent: (ind > 0 || ti) ? { start: ind > 0 ? Math.round(ind * 15) : undefined, firstLine: ti ? Math.round(parseFloat(ti) * (/cm$/.test(ti) ? 567 : /px$/.test(ti) ? 15 : 20)) : undefined } : undefined }, base || {});
  }
  function listItems(list, D, kids, level, ctr) {
    const ord = list.tagName === 'OL', inst = ++ctr.n;
    Array.from(list.children).forEach(li => {
      if (li.tagName !== 'LI') return;
      const runs = []; const copy = li.cloneNode(true); $$('ul,ol', copy).forEach(n => n.remove());
      runsOf(copy, { sz: 28 }, D, runs);
      kids.push(new D.Paragraph(Object.assign(blockProps(li.parentNode, D, { children: runs, spacing: { after: 60 } }), ord ? { numbering: { reference: 'edOl', level: Math.min(level, 3), instance: inst } } : { bullet: { level: Math.min(level, 3) } })));
      Array.from(li.children).filter(c => /^(UL|OL)$/.test(c.tagName)).forEach(c => listItems(c, D, kids, level + 1, ctr));
    });
  }
  function tableOf(t, D) {
    const rows = Array.from(t.rows), ncol = Math.max.apply(null, rows.map(r => r.cells.length).concat(1)), rtl = (t.closest('[dir]') || ed()).getAttribute('dir') !== 'ltr';
    const W = TBLW, cw = Math.floor(W / ncol);
    return new D.Table({
      width: { size: W, type: D.WidthType.DXA }, columnWidths: Array(ncol).fill(cw), visuallyRightToLeft: rtl,
      rows: rows.map(r => new D.TableRow({ children: Array.from(r.cells).map(c => {
        const kids = []; blocksOf(c, D, kids, { n: 0 }); if (!kids.length) kids.push(new D.Paragraph({ children: [] }));
        const bg = hex(c.style.backgroundColor);
        return new D.TableCell({ width: { size: cw * (c.colSpan || 1), type: D.WidthType.DXA }, columnSpan: c.colSpan > 1 ? c.colSpan : undefined, children: kids, shading: bg ? { type: D.ShadingType.CLEAR, fill: bg, color: 'auto' } : undefined });
      }) }))
    });
  }
  function blocksOf(root, D, kids, ctr) {
    let loose = [];
    const flush = () => { if (!loose.length) return; const w = doc.createElement('div'); loose.forEach(n => w.appendChild(n.cloneNode(true))); const runs = []; runsOf(w, { sz: 28 }, D, runs); if (runs.length && w.textContent.trim() || w.querySelector('img')) kids.push(new D.Paragraph(blockProps(root, D, { children: runs }))); loose = []; };
    root.childNodes.forEach(n => {
      if (n.nodeType === 3) { loose.push(n); return; }
      if (n.nodeType !== 1) return;
      const tg = n.tagName;
      if (/^(P|H[1-6]|BLOCKQUOTE)$/.test(tg)) {
        flush(); const runs = [], h = HEAD[tg], bq = tg === 'BLOCKQUOTE';
        runsOf(n, h ? { b: true, sz: h[1] } : { sz: 28, i: bq || undefined }, D, runs);
        kids.push(new D.Paragraph(blockProps(n, D, { children: runs, heading: h ? h[0](D) : undefined, indent: bq ? { start: 720 } : undefined })));
      } else if (tg === 'UL' || tg === 'OL') { flush(); listItems(n, D, kids, 0, ctr); }
      else if (tg === 'TABLE') { flush(); kids.push(tableOf(n, D)); kids.push(new D.Paragraph({ children: [] })); }
      else if (tg === 'HR') { flush(); kids.push(n.dataset.pb ? new D.Paragraph({ children: [new D.PageBreak()] }) : new D.Paragraph({ border: { bottom: { style: D.BorderStyle.SINGLE, size: 6, color: '888888', space: 1 } }, children: [] })); }
      else if (tg === 'DIV' || tg === 'SECTION' || tg === 'ARTICLE') { flush(); if (n.querySelector('p,h1,h2,h3,h4,h5,h6,ul,ol,table,div')) blocksOf(n, D, kids, ctr); else { const runs = []; runsOf(n, { sz: 28 }, D, runs); if (n.textContent.trim() || n.querySelector('img')) kids.push(new D.Paragraph(blockProps(n, D, { children: runs }))); } }
      else loose.push(n);
    });
    flush();
  }

  /* ---------- التصدير ---------- */
  function cleanClone() { const c = ed().cloneNode(true); $$('[data-pg]', c).forEach(n => { n.style.marginTop = n.dataset.pgb || ''; n.removeAttribute('data-pg'); n.removeAttribute('data-pgb'); }); $$('.sel-img', c).forEach(n => n.classList.remove('sel-img')); c.removeAttribute('contenteditable'); c.removeAttribute('id'); return c; }
  async function toDocx() {
    const D = g.docx, kids = [], P = PG(), [W, H] = dimsPx(), sz = SIZES[P.size] || SIZES.a4, tw = mm => Math.round(mm / 25.4 * 1440);
    const hdr = (P.hdr || '').trim(), ftr = (P.ftr || '').trim(), pn = !!P.pn, S = App.getState();
    const host = doc.createElement('div'), clone = cleanClone();
    clone.style.cssText = 'width:' + IMGMAX + 'px;padding:0;min-height:0;font:14pt Arial;direction:rtl'; host.style.cssText = 'position:fixed;left:-99999px;top:0;'; host.appendChild(clone); doc.body.appendChild(host);
    try { blocksOf(clone, D, kids, { n: 0 }); } finally { host.remove(); }
    if (!kids.length) kids.push(new D.Paragraph({ children: [] }));
    const lvl = i => ({ level: i, format: D.LevelFormat.DECIMAL, text: '%' + (i + 1) + '.', alignment: D.AlignmentType.START });
    const cm = x => Math.round(x * 567), run = (t, extra) => new D.TextRun(Object.assign({ text: t, font: 'Arial', size: 20, color: '555555', rightToLeft: isAr(t) || undefined }, extra || {}));
    let footer = null, header = null;
    const pnRun = () => new D.TextRun({ children: ['صفحة ', D.PageNumber.CURRENT, ' من ', D.PageNumber.TOTAL_PAGES], font: 'Arial', size: 20, color: '555555', rightToLeft: true });
    if (P.siteFooter && App.buildFooter) { footer = await App.buildFooter(S); }
    if (!footer && (ftr || pn)) footer = new D.Footer({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, bidirectional: true, children: [].concat(ftr ? [run(ftr + (pn ? '   |   ' : ''))] : [], pn ? [pnRun()] : []) })] });
    if (hdr || (P.siteFooter && pn)) header = new D.Header({ children: [new D.Paragraph({ alignment: D.AlignmentType.CENTER, bidirectional: isAr(hdr) || true, children: [].concat(hdr ? [run(hdr + (P.siteFooter && pn ? '   |   ' : ''))] : [], P.siteFooter && pn ? [pnRun()] : []) })] });
    const fh = footer && P.siteFooter ? cm(Math.max(P.mb, 2)) : cm(P.mb);
    const d = new D.Document({ numbering: { config: [{ reference: 'edOl', levels: [0, 1, 2, 3].map(lvl) }] }, styles: { default: { document: { run: { font: 'Arial', size: 28 } } } }, sections: [{ properties: { page: { size: { width: tw(sz[0]), height: tw(sz[1]), orientation: P.orient === 'l' ? D.PageOrientation.LANDSCAPE : D.PageOrientation.PORTRAIT }, margin: { top: cm(P.mt), bottom: fh, left: cm(P.ml), right: cm(P.mr), header: 400, footer: 300 } } }, headers: header ? { default: header } : undefined, footers: footer ? { default: footer } : undefined, children: kids }] });
    return D.Packer.toBlob(d);
  }
  async function toPdf() {
    const P = PG(), [W, H] = dimsPx(), fi = footerInfo(), host = doc.createElement('div'), clone = cleanClone();
    host.style.cssText = 'position:fixed;left:-99999px;top:0;width:' + W + 'px;background:#fff;';
    const pages = E0.word.pages || 1; clone.style.cssText = ed().style.cssText + ';width:' + W + 'px;box-sizing:border-box;background:#fff;color:#000;direction:rtl;position:relative;min-height:' + pages * H + 'px';
    host.appendChild(clone); doc.body.appendChild(host);
    const wrap = doc.createElement('div'); wrap.style.cssText = 'position:relative;width:' + W + 'px;height:' + pages * H + 'px;background:#fff'; host.insertBefore(wrap, clone); wrap.appendChild(clone);
    const ov = $('#wdOver').cloneNode(true); ov.style.cssText = 'position:absolute;left:0;right:0;top:0;height:' + pages * H + 'px;pointer-events:none'; $$('.wd-gap,.wd-pn', ov).forEach(n => n.remove()); wrap.appendChild(ov);
    try {
      const cv = await g.html2canvas(wrap, { scale: 2, backgroundColor: '#fff', useCORS: true, logging: false, width: W, height: pages * H });
      const J = g.jspdf.jsPDF, pdf = new J({ unit: 'px', format: [W, H], orientation: W > H ? 'l' : 'p', compress: true, hotfixes: ['px_scaling'] }), k = cv.width / W;
      for (let i = 0; i < pages; i++) {
        const c = doc.createElement('canvas'); c.width = cv.width; c.height = Math.min(Math.round(H * k), cv.height - Math.round(i * H * k)); if (c.height <= 0) break;
        const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(cv, 0, Math.round(i * H * k), cv.width, c.height, 0, 0, c.width, c.height);
        if (i) pdf.addPage([W, H], W > H ? 'l' : 'p'); pdf.addImage(c.toDataURL('image/jpeg', 0.93), 'JPEG', 0, 0, W, c.height / k);
      }
      return pdf.output('blob');
    } finally { host.remove(); }
  }
  async function getFile(fmt, E) {
    E0 = E;
    if (fmt === 'docx') return { blob: await toDocx(), name: nameOf() + '-معدّل.docx' };
    return { blob: await toPdf(), name: nameOf() + '-معدّل.pdf' };
  }
  function print(E) {
    E0 = E; const P = PG(), s = SIZES[P.size] || SIZES.a4, fi = footerInfo(), w = g.open('', '_blank'); if (!w) { toast('اسمح بالنوافذ المنبثقة للطباعة'); return; }
    const [, H] = dimsPx(), body = ed().cloneNode(true); $$('[data-pg]', body).forEach(n => { n.style.marginTop = n.dataset.pgb || ''; n.removeAttribute('data-pg'); });
    const mb = fi ? Math.max(P.mb * CM, fi.h + 0.8 * CM) / CM : P.mb;
    const foot = fi ? '<div style="position:fixed;bottom:0.4cm;left:' + P.ml + 'cm;right:' + P.mr + 'cm" dir="rtl">' + fi.html + '</div>' : '';
    const css = $$('link[rel=stylesheet],style').map(n => n.outerHTML).join('');
    w.document.write('<!doctype html><html dir="rtl" lang="ar"><meta charset="utf-8"><title>' + esc(nameOf()) + '</title>' + css + '<style>@page{size:' + s[0] + 'mm ' + s[1] + 'mm' + (P.orient === 'l' ? ' landscape' : '') + ';margin:' + P.mt + 'cm ' + P.mr + 'cm ' + mb + 'cm ' + P.ml + 'cm}body{margin:0;background:#fff;font:14pt Arial,Tahoma,sans-serif;direction:rtl}img{max-width:100%}table{border-collapse:collapse}h1{font-size:20pt}h2{font-size:16pt}h3{font-size:14pt}hr[data-pb]{page-break-after:always;border:0}</style><body>' + foot + body.innerHTML + '</body></html>');
    w.document.close(); setTimeout(() => { w.focus(); w.print(); }, 600);
  }
  g.App = Object.assign(App, { editorWord: { load, showStage, buildForm, getFile, print, snapshot, draftInfo, restoreDraft, blank } });
})(window);
