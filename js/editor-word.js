/* ===== وضع Word في تبويب «تحميل كتاب للتعديل»: محرر مستند + تصدير Word/PDF ===== */
(function (g) {
  'use strict';
  const App = g.App, doc = document;
  const $ = (s, r) => (r || doc).querySelector(s), $$ = (s, r) => Array.from((r || doc).querySelectorAll(s));
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const readBuf = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsArrayBuffer(f); });
  const readUrl = f => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(r.result); r.onerror = rej; r.readAsDataURL(f); });
  const ed = () => $('#wdPaper');
  const FONTS = () => (App.FONTS || []).map(f => f[0]);
  let E0 = null, savedRange = null;

  async function load(f, E) {
    E0 = E;
    if (!g.mammoth) throw new Error('مكتبة Word غير محمّلة');
    const r = await g.mammoth.convertToHtml({ arrayBuffer: await readBuf(f) }, { styleMap: ["p[style-name='Title'] => h1:fresh", "p[style-name='Heading 1'] => h2:fresh"] });
    E.kind = 'docx'; E.name = f.name; E.pages = []; E.srcs = []; E.undo = []; E.redo = []; E.sel = null; E.psel.clear();
    E.word = { html: r.value || '<p><br></p>', warn: r.messages.length };
    App.editor.show(true);
    g.toast('تم فتح مستند Word' + (r.messages.length ? ' (بعض التنسيقات المعقدة قد تُبسَّط)' : ''));
  }

  function showStage(st) {
    st.innerHTML =
      '<div class="wd-bar" id="wdBar">' +
      '<select id="wdFont" title="الخط">' + FONTS().map(f => '<option>' + esc(f) + '</option>').join('') + '</select>' +
      '<select id="wdSize" title="الحجم">' + [10, 11, 12, 13, 14, 16, 18, 20, 24, 28, 32, 40, 48].map(n => '<option' + (n === 14 ? ' selected' : '') + '>' + n + '</option>').join('') + '</select>' +
      '<select id="wdBlock" title="النمط"><option value="p">نص عادي</option><option value="h1">عنوان 1</option><option value="h2">عنوان 2</option><option value="h3">عنوان 3</option><option value="blockquote">اقتباس</option></select>' +
      [['bold', '<b>B</b>', 'غامق'], ['italic', '<i>I</i>', 'مائل'], ['underline', '<u>U</u>', 'تسطير'], ['strikeThrough', '<s>S</s>', 'يتوسطه خط'], ['justifyRight', '⇥', 'محاذاة يمين'], ['justifyCenter', '≡', 'توسيط'], ['justifyLeft', '⇤', 'محاذاة يسار'], ['justifyFull', '☰', 'ضبط'], ['insertUnorderedList', '•', 'قائمة نقطية'], ['insertOrderedList', '1.', 'قائمة رقمية'], ['indent', '→', 'زيادة المسافة'], ['outdent', '←', 'إنقاص المسافة'], ['subscript', 'x₂', 'منخفض'], ['superscript', 'x²', 'مرتفع'], ['removeFormat', 'Tx', 'مسح التنسيق'], ['undo', '↶', 'تراجع'], ['redo', '↷', 'إعادة']]
        .map(b => '<button type="button" class="ed-tb wd-b" data-cmd="' + b[0] + '" title="' + b[2] + '">' + b[1] + '</button>').join('') +
      '<label class="ed-tb" title="لون الخط">A<input type="color" id="wdCol" value="#000000"></label><label class="ed-tb" title="تظليل">▮<input type="color" id="wdHl" value="#ffff00"></label>' +
      '<button type="button" class="ed-tb" id="wdRtl" title="اتجاه الفقرة: عربي (يمين)">RTL</button><button type="button" class="ed-tb" id="wdLtr" title="اتجاه الفقرة: إنجليزي (يسار)">LTR</button></div>' +
      '<div class="wd-view"><div class="wd-paper" id="wdPaper" contenteditable="true" dir="rtl" spellcheck="false"></div></div>';
    const p = ed(); p.innerHTML = E0.word.html; if (!p.textContent.trim() && !p.querySelector('img,table')) p.innerHTML = '<p><br></p>';
    bind();
  }
  const saveSel = () => { const s = g.getSelection(); if (s.rangeCount && ed() && ed().contains(s.anchorNode)) savedRange = s.getRangeAt(0).cloneRange(); };
  const restoreSel = () => { ed().focus(); if (savedRange) { const s = g.getSelection(); s.removeAllRanges(); s.addRange(savedRange); } };
  function cmd(c, v) { restoreSel(); doc.execCommand('styleWithCSS', false, true); doc.execCommand(c, false, v == null ? null : v); saveSel(); }
  function bind() {
    doc.addEventListener('selectionchange', saveSel);
    $$('.wd-b').forEach(b => { b.addEventListener('mousedown', e => e.preventDefault()); b.addEventListener('click', () => cmd(b.dataset.cmd)); });
    $('#wdFont').addEventListener('change', e => cmd('fontName', e.target.value));
    $('#wdSize').addEventListener('change', e => { restoreSel(); doc.execCommand('styleWithCSS', false, false); doc.execCommand('fontSize', false, 7); $$('font[size="7"]', ed()).forEach(f => { f.removeAttribute('size'); f.style.fontSize = e.target.value + 'pt'; }); saveSel(); });
    $('#wdBlock').addEventListener('change', e => cmd('formatBlock', e.target.value));
    $('#wdCol').addEventListener('input', e => cmd('foreColor', e.target.value));
    $('#wdHl').addEventListener('input', e => cmd('hiliteColor', e.target.value));
    $('#wdRtl').addEventListener('click', () => setDir('rtl')); $('#wdLtr').addEventListener('click', () => setDir('ltr'));
    $('#wdRtl').addEventListener('mousedown', e => e.preventDefault()); $('#wdLtr').addEventListener('mousedown', e => e.preventDefault());
    ed().addEventListener('paste', e => { // لصق نص نظيف أو صور
      const items = e.clipboardData && e.clipboardData.items; if (!items) return;
      for (const it of items) if (/^image\//.test(it.type)) { e.preventDefault(); readUrl(it.getAsFile()).then(u => insertImg(u)); return; }
    });
    ed().addEventListener('input', () => { E0.word.dirty = true; });
  }
  function setDir(d) {
    restoreSel(); const s = g.getSelection(); let n = s.anchorNode; while (n && n !== ed() && !(n.nodeType === 1 && /^(P|H[1-6]|LI|DIV|BLOCKQUOTE|TD|TH)$/.test(n.tagName))) n = n.parentNode;
    if (n && n !== ed()) { n.setAttribute('dir', d); n.style.textAlign = d === 'rtl' ? 'right' : 'left'; }
  }
  function insertImg(url, w) {
    restoreSel(); const im = new Image(); im.onload = () => {
      const ww = Math.min(w || 400, im.naturalWidth);
      doc.execCommand('insertHTML', false, '<img src="' + url + '" style="width:' + ww + 'px;max-width:100%">'); saveSel();
    }; im.src = url;
  }
  function insertTable(r, c) {
    let h = '<table style="border-collapse:collapse;width:100%"><tbody>';
    for (let i = 0; i < r; i++) { h += '<tr>'; for (let j = 0; j < c; j++) h += '<td style="border:1px solid #444;padding:4px 6px"><br></td>'; h += '</tr>'; }
    restoreSel(); doc.execCommand('insertHTML', false, h + '</tbody></table><p><br></p>');
  }
  const curCell = () => { let n = g.getSelection().anchorNode; while (n && n !== ed()) { if (n.nodeType === 1 && /^(TD|TH)$/.test(n.tagName)) return n; n = n.parentNode; } return null; };
  function tableOp(op) {
    restoreSel(); const td = curCell(); if (!td) { g.toast('ضع المؤشر داخل خلية جدول أولاً'); return; }
    const tr = td.parentNode, tb = tr.parentNode, ci = Array.from(tr.children).indexOf(td);
    const mk = () => { const c = doc.createElement('td'); c.style.cssText = 'border:1px solid #444;padding:4px 6px'; c.innerHTML = '<br>'; return c; };
    if (op === 'addRow') { const n = doc.createElement('tr'); Array.from(tr.children).forEach(() => n.appendChild(mk())); tr.after(n); }
    else if (op === 'addCol') Array.from(tb.children).forEach(r => { const c = mk(); const ref = r.children[ci]; ref ? ref.after(c) : r.appendChild(c); });
    else if (op === 'delRow') { tr.remove(); if (!tb.children.length) tb.closest('table').remove(); }
    else if (op === 'delCol') { Array.from(tb.children).forEach(r => r.children[ci] && r.children[ci].remove()); if (!tr.children.length) tb.closest('table').remove(); }
    else if (op === 'delTable') td.closest('table').remove();
    E0.word.dirty = true;
  }
  function findReplace(find, rep, all) {
    if (!find) return 0; let n = 0; const w = doc.createTreeWalker(ed(), NodeFilter.SHOW_TEXT), nodes = [];
    while (w.nextNode()) nodes.push(w.currentNode);
    for (const t of nodes) { if (t.nodeValue.includes(find)) { const parts = t.nodeValue.split(find); n += parts.length - 1; t.nodeValue = parts.join(rep); if (!all) { /* الأول فقط في كل عقدة */ } } }
    return n;
  }
  function stats() { const t = ed().innerText.trim(); return { words: t ? t.split(/\s+/).length : 0, chars: t.length }; }

  function buildForm(host, E) {
    E0 = E;
    host.innerHTML =
      '<fieldset class="group"><legend>٢) تعديل مستند Word</legend>' +
      '<p class="hint">عدّل النص مباشرة في الصفحة. الأدوات في الشريط أعلى المعاينة. ملاحظة: يُستورد النص والجداول والصور والتنسيق الأساسي، وقد تُبسَّط التنسيقات المعقدة (الترويسات/التذييلات والمربعات).</p>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" id="wdAddImg">+ إضافة صورة</button><button type="button" class="btn sm" id="wdLogoC">شعار النادي</button><button type="button" class="btn sm" id="wdLogoM">شعار الوزارة</button><button type="button" class="btn sm" id="wdSignH">توقيع الرئيس</button><button type="button" class="btn sm" id="wdStampH">ختم الرئيس</button><button type="button" class="btn sm" id="wdSignM">توقيع المدير</button><button type="button" class="btn sm" id="wdStampM">ختم المدير</button></div>' +
      '<input type="file" id="wdImgFile" accept="image/*" multiple hidden></fieldset>' +
      '<fieldset class="group"><legend>٣) جدول وفقرات</legend>' +
      '<div class="row2"><label class="field"><span>صفوف</span><input id="wdRows" type="number" value="3" min="1" max="50"></label><label class="field"><span>أعمدة</span><input id="wdCols" type="number" value="3" min="1" max="12"></label></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" data-wt="ins">إدراج جدول</button><button type="button" class="btn sm" data-wt="addRow">+ صف</button><button type="button" class="btn sm" data-wt="addCol">+ عمود</button><button type="button" class="btn sm danger" data-wt="delRow">حذف صف</button><button type="button" class="btn sm danger" data-wt="delCol">حذف عمود</button><button type="button" class="btn sm danger" data-wt="delTable">حذف الجدول</button></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" id="wdHr">خط فاصل</button><button type="button" class="btn sm" id="wdPb">فاصل صفحة</button><button type="button" class="btn sm" id="wdLink">رابط</button></div></fieldset>' +
      '<fieldset class="group"><legend>٤) بحث واستبدال</legend>' +
      '<div class="row2"><label class="field"><span>بحث</span><input id="wdFind" type="text"></label><label class="field"><span>استبدال بـ</span><input id="wdRep" type="text"></label></div>' +
      '<div class="ed-btnrow"><button type="button" class="btn sm" id="wdRepAll">استبدال الكل</button><span class="hint" id="wdStat"></span></div></fieldset>' +
      '<fieldset class="group"><legend>٥) تحويل</legend><div class="ed-btnrow"><button type="button" class="btn sm primary" id="wdToPdf">تحويل إلى PDF</button><button type="button" class="btn sm" id="wdToTxt">تنزيل كنص (.txt)</button><button type="button" class="btn sm" id="wdToHtml">تنزيل كصفحة (.html)</button></div></fieldset>';
    const S = () => App.getState();
    $('#wdAddImg').addEventListener('click', () => { saveSel(); $('#wdImgFile').click(); });
    $('#wdImgFile').addEventListener('change', async e => { const fs = Array.from(e.target.files); e.target.value = ''; for (const f of fs) { restoreSel(); insertImg(await readUrl(f)); await new Promise(r => setTimeout(r, 80)); } });
    const lg = k => () => { const u = S().logos[k] && S().logos[k].data; if (!u) g.toast('لا يوجد شعار محفوظ'); else insertImg(u, 90); };
    const pe = (w, f) => () => { const p = S().people[w], u = p && p[f]; if (!u) g.toast('لم يُضف في الإعدادات'); else insertImg(u, 110); };
    $('#wdLogoC').addEventListener('click', lg('club')); $('#wdLogoM').addEventListener('click', lg('ministry'));
    $('#wdSignH').addEventListener('click', pe('head', 'signImg')); $('#wdStampH').addEventListener('click', pe('head', 'stampImg'));
    $('#wdSignM').addEventListener('click', pe('manager', 'signImg')); $('#wdStampM').addEventListener('click', pe('manager', 'stampImg'));
    host.addEventListener('mousedown', e => { if (e.target.closest('button')) saveSel(); });
    $$('[data-wt]', host).forEach(b => b.addEventListener('click', () => { if (b.dataset.wt === 'ins') insertTable(+$('#wdRows').value || 3, +$('#wdCols').value || 3); else tableOp(b.dataset.wt); }));
    $('#wdHr').addEventListener('click', () => { restoreSel(); doc.execCommand('insertHorizontalRule'); });
    $('#wdPb').addEventListener('click', () => { restoreSel(); doc.execCommand('insertHTML', false, '<hr data-pb="1" style="border:0;border-top:2px dashed #999"><p><br></p>'); });
    $('#wdLink').addEventListener('click', () => { const u = prompt('عنوان الرابط (https://…)'); if (u) cmd('createLink', u); });
    $('#wdRepAll').addEventListener('click', () => { const n = findReplace($('#wdFind').value, $('#wdRep').value, true); $('#wdStat').textContent = n ? 'تم استبدال ' + n + ' موضع' : 'لا توجد نتائج'; });
    host.addEventListener('mouseenter', () => { const s = stats(); const el = $('#wdStat'); if (el && ed() && !el.textContent) el.textContent = s.words + ' كلمة'; });
    $('#wdToPdf').addEventListener('click', () => App.run('جارٍ تحويل Word إلى PDF…', async () => { App.saveFile(await toPdf(), nameOf() + '.pdf'); }));
    $('#wdToTxt').addEventListener('click', () => App.saveFile(new Blob([ed().innerText], { type: 'text/plain;charset=utf-8' }), nameOf() + '.txt'));
    $('#wdToHtml').addEventListener('click', () => App.saveFile(new Blob(['<!doctype html><html dir="rtl" lang="ar"><meta charset="utf-8"><body>' + ed().innerHTML + '</body></html>'], { type: 'text/html;charset=utf-8' }), nameOf() + '.html'));
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
        if (tg === 'IMG') { try { const src = ch.getAttribute('src') || ''; if (!/^data:image/.test(src)) return; const w = Math.round(ch.getBoundingClientRect().width || ch.naturalWidth || 200), h = Math.round(ch.getBoundingClientRect().height || ch.naturalHeight || 100); out.push(new D.ImageRun({ type: /jpe?g/.test(src.slice(0, 20)) ? 'jpg' : /gif/.test(src.slice(0, 20)) ? 'gif' : 'png', data: bytes(src), transformation: { width: Math.min(560, w), height: Math.round(Math.min(560, w) * h / Math.max(1, w)) } })); } catch (e) { /* تخطّي */ } return; }
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
    const dir = el.getAttribute('dir') || (isAr(el.textContent) ? 'rtl' : 'ltr'), ta = (el.style.textAlign || '').toLowerCase();
    const rtl = dir === 'rtl'; let al;
    // في فقرة bidi: START = يمين للعربي، END = يسار
    if (ta === 'center') al = D.AlignmentType.CENTER; else if (ta === 'justify') al = D.AlignmentType.JUSTIFIED;
    else if (ta === 'left') al = rtl ? D.AlignmentType.END : D.AlignmentType.START; else if (ta === 'right') al = rtl ? D.AlignmentType.START : D.AlignmentType.END;
    const ml = parseFloat(el.style.marginRight || el.style.paddingRight || 0), mrl = parseFloat(el.style.marginLeft || el.style.paddingLeft || 0);
    const ind = rtl ? ml : mrl;
    return Object.assign({ bidirectional: rtl, alignment: al, spacing: { after: 100, line: 300 }, indent: ind > 0 ? { start: Math.round(ind * 15) } : undefined }, base || {});
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
    const W = 9500, cw = Math.floor(W / ncol);
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
  async function toDocx() {
    const D = g.docx, kids = [];
    blocksOf(ed(), D, kids, { n: 0 });
    if (!kids.length) kids.push(new D.Paragraph({ children: [] }));
    const lvl = i => ({ level: i, format: D.LevelFormat.DECIMAL, text: '%' + (i + 1) + '.', alignment: D.AlignmentType.START });
    const d = new D.Document({ numbering: { config: [{ reference: 'edOl', levels: [0, 1, 2, 3].map(lvl) }] }, styles: { default: { document: { run: { font: 'Arial', size: 28 } } } }, sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, bottom: 1134, left: 1134, right: 1134 } } }, children: kids }] });
    return D.Packer.toBlob(d);
  }
  async function toPdf() {
    const p = ed(), W = 794, host = doc.createElement('div');
    host.style.cssText = 'position:fixed;left:-99999px;top:0;width:' + W + 'px;background:#fff;';
    const clone = p.cloneNode(true); clone.removeAttribute('contenteditable'); clone.style.cssText = 'width:' + W + 'px;padding:60px 60px;box-sizing:border-box;background:#fff;color:#000;direction:rtl;font-size:18px;line-height:1.7;';
    host.appendChild(clone); doc.body.appendChild(host);
    try {
      const cv = await g.html2canvas(clone, { scale: 2, backgroundColor: '#fff', useCORS: true, logging: false });
      const J = g.jspdf.jsPDF, pdf = new J({ unit: 'mm', format: 'a4', compress: true }), pw = 210, ph = 297, k = pw / cv.width, pagePx = Math.floor(ph / k);
      for (let y = 0, i = 0; y < cv.height; y += pagePx, i++) {
        const c = doc.createElement('canvas'); c.width = cv.width; c.height = Math.min(pagePx, cv.height - y);
        const x = c.getContext('2d'); x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); x.drawImage(cv, 0, y, cv.width, c.height, 0, 0, c.width, c.height);
        if (i) pdf.addPage(); pdf.addImage(c.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, pw, c.height * k);
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
    E0 = E; const w = g.open('', '_blank'); if (!w) { g.toast('اسمح بالنوافذ المنبثقة للطباعة'); return; }
    w.document.write('<!doctype html><html dir="rtl" lang="ar"><meta charset="utf-8"><title>' + esc(nameOf()) + '</title><style>body{font:16px/1.7 Tahoma,Arial,sans-serif;margin:20mm;direction:rtl}img{max-width:100%}table{border-collapse:collapse;width:100%}td,th{border:1px solid #444;padding:4px 6px}hr[data-pb]{page-break-after:always;border:0}</style><body>' + ed().innerHTML + '</body></html>');
    w.document.close(); setTimeout(() => { w.focus(); w.print(); }, 400);
  }
  g.App = Object.assign(App, { editorWord: { load, showStage, buildForm, getFile, print } });
})(window);
