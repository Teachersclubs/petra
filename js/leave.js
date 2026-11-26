/* ===== تبويب «إجازة سنوية»: نموذج رسمي مطابق للأصل (معاينة + طباعة + Word) ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const TEAL = '044F5F', W_TOTAL = 10389, GRID = [2316, 1877, 1362, 42, 228, 1980, 2584]; // أعمدة الشبكة (من اليمين) كما في الأصل
  const NB = ' ';

  /* ---------- حساب الأيام والتواريخ ---------- */
  const parse = iso => { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? Date.UTC(+m[1], +m[2] - 1, +m[3]) : null; };
  // كل 24 ساعة = يوم: الفرق بالساعات بين (تاريخ+وقت البداية) و(تاريخ+وقت النهاية)؛ 23:59 تُعدّ نهاية اليوم (24:00)
  const mins = t => { const m = /^(\d{1,2}):(\d{2})$/.exec(t || ''); if (!m) return null; const v = +m[1] * 60 + +m[2]; return v === 1439 ? 1440 : v; };
  function calcDays(a, b, ta, tb) { const x = parse(a), y = parse(b), m1 = mins(ta == null ? '00:00' : ta), m2 = mins(tb == null ? '23:59' : tb); if (x == null || y == null || m1 == null || m2 == null) return ''; const ms = (y + m2 * 60000) - (x + m1 * 60000); return ms < 0 ? '' : String(Math.floor(ms / 86400000)); }
  function fmt(iso) { const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || ''); return m ? (+m[3]) + '/' + (+m[2]) + '/' + m[1] : ''; }
  const blankDate = () => NB.repeat(4) + '/' + NB.repeat(4) + '/' + NB + '202';
  const dateTxt = iso => fmt(iso) || blankDate();
  const dots = n => '.'.repeat(n);

  /* ---------- نموذج الصفوف (مصدر واحد للمعاينة والطباعة وWord) ---------- */
  // خلية: {w, b:'tdldbdrd' (t/l/b/r + d=مزدوج s=مفرد)، fill، al، va، lines:[ {segs:[{r:[runs], w, al}]} ]، h(twips)}
  // run: نص | {box:true/false} 
  const L = (...r) => ({ segs: [{ r }] });
  const M = (...segs) => ({ segs });
  const sg = (w, al, ...r) => ({ r, w, al });
  function rows(S) {
    const d = S.leave, inn = d.place === 'in', out = d.place === 'out';
    const club = 'نادي ' + (S.clubName || '').replace(/^\s*(ال)?نادي\s+/, '').trim();
    const addr = (d.address || '').trim();
    const appr = (label) => L(label, dots(16), NB.repeat(4), { box: false }, NB, 'موافق', NB.repeat(3), { box: false }, NB, 'غير موافق بسبب ', dots(14));
    const dateSign = () => M(sg(6600, 'r', 'الــتــاريــــخ : ', NB.repeat(3) + '/' + NB.repeat(3) + '/' + NB.repeat(3) + '202'), sg(3789, 'l', 'التوقيع :'));
    const R = [];
    R.push({ h: 258, cells: [{ w: 10389, b: 'tdldbdrd', fill: 'FFFFFF', lines: [L(' المادة ( 50  / أ   ) من نظام ادارة الموارد البشرية  في القطاع العام رقم 33 لسنة 2024م  وتعديلاته البند رابعاً فقرة (3 ) من كتاب تفويض الصلاحيات رقم 1/70/50791  تاريخ  : 23/9/2026م ')] }] });
    R.push({ h: 432, cells: [{ w: 5555, b: 'tdldbdrs', va: 'b', lines: [L('الاسم : ', d.name)] }, { w: 4834, b: 'tdlsbdrd', va: 'b', lines: [L('المسمى الوظيفي:  ', d.job)] }] });
    R.push({ h: 432, cells: [{ w: 2316, b: 'tdldbdrs', va: 'b', lines: [L('الادارة :')] }, { w: 3239, b: 'tdlsbdrs', va: 'b', lines: [L(d.dept)] }, { w: 2250, b: 'tdlsbdrs', va: 'b', lines: [L('الـقـســم ــ ')] }, { w: 2584, b: 'tdlsbdrd', va: 'b', lines: [L(d.section), L(club)] }] });
    R.push({ h: 645, cells: [{ w: 10389, b: 'tdldbdrd', lines: [
      M(sg(6200, 'r', 'مكان تواجد الموظف أثناء الإجازة :'), sg(4000, 'l', 'الـهـاتف الـخلــوي : ', d.phone)),
      L({ box: inn }, NB, 'داخل المملكة / العنوان  ( ', inn ? addr : NB.repeat(18), ' )'),
      L({ box: out }, NB, 'خارج المملكة ...... في : ', out && addr ? addr : dots(40))] }] });
    R.push({ h: 402, cells: [{ w: 10389, b: 'tdldrd', fill: 'BFBFBF', al: 'c', lines: [L('بيانات الاجازة ')] }] });
    R.push({ h: 467, cells: [{ w: 10389, b: 'ldrd', lines: [M(
      sg(3250, 'r', 'تاريخ ابتداء الإجازة : ', dateTxt(d.start)),
      sg(3750, 'r', 'تاريخ انتهاء الإجازة : مساء ', dateTxt(d.end)),
      sg(3389, 'l', 'عدد الأيام المطلوبة :(', d.days ? NB + d.days + NB : NB.repeat(8), ')'))] }] });
    R.push({ h: 417, cells: [{ w: 5597, b: 'tdldbd', lines: [L('تاريخ تقديم الإجازة : ', NB.repeat(2), dateTxt(d.date))] }, { w: 4792, b: 'tdbdrd', lines: [L('توقيع الموظف :')] }] });
    R.push({ cells: [{ w: 10389, b: 'tdldrd', fill: 'BFBFBF', al: 'c', va: 'c', lines: [L('الوحدة التنظيمية المعنية بالموارد البشرية و التطوير المؤسسي')] }] });
    R.push({ cells: [{ w: 4193, b: 'tdld', al: 'c', va: 'c', lines: [L('عدد أيام الإجازة المستحقة بتاريخ تقديم الإجازة')] }, { w: 1632, b: 'td', al: 'c', va: 'c', lines: [L('التاريخ')] }, { w: 4564, b: 'tdrd', al: 'c', va: 'c', lines: [L('توقيع مسؤول الموارد البشرية و التطوير المؤسسي')] }] });
    R.push({ h: 368, cells: [{ w: 4193, b: 'ldbd', al: 'c', lines: [L('(' + NB.repeat(13) + ') يوم')] }, { w: 1632, b: 'bd', al: 'c', lines: [L(NB.repeat(3) + '/' + NB.repeat(4) + '/' + NB.repeat(4) + '202')] }, { w: 4564, b: 'bdrd', lines: [L(NB)] }] });
    R.push({ h: 393, cells: [{ w: 10389, b: 'ldbdrd', lines: [L('ملاحظات :')] }] });
    R.push({ cells: [{ w: 10389, b: 'tdldbdrd', fill: 'BFBFBF', al: 'c', va: 'c', lines: [L('التنسيبات وفقاً للتسلسل الاداري المعتمد في الدائرة:')] }] });
    R.push({ cells: [{ w: 10389, b: 'tdldbdrd', lines: [appr('المسمى الوظيفي:  '), L('اسم الموظف الذي سيقوم بالعمل أثناء الإجازة : ', dots(45))] }] });
    R.push({ h: 432, cells: [{ w: 10389, b: 'tdldbdrd', lines: [dateSign()] }] });
    R.push({ h: 576, cells: [{ w: 10389, b: 'ldrd', lines: [appr('المسمى الوظيفي:  ')] }] });
    R.push({ h: 467, cells: [{ w: 10389, b: 'ldbdrd', lines: [dateSign()] }] });
    R.push({ h: 576, cells: [{ w: 10389, b: 'tdldrd', lines: [appr('المسمى الوظيفي:  ')] }] });
    R.push({ h: 432, cells: [{ w: 10389, b: 'ldbdrd', lines: [dateSign()] }] });
    R.push({ h: 576, cells: [{ w: 10389, b: 'tdldrd', lines: [appr('قرار الأمين العام / أو من يفوضه : ')] }] });
    R.push({ h: 432, cells: [{ w: 10389, b: 'ldbdrd', lines: [dateSign()] }] });
    return R;
  }
  // عدد أعمدة الشبكة التي تغطيها الخلية بحسب موضعها التراكمي
  function spans(row) {
    const bounds = [0]; GRID.reduce((a, w) => (bounds.push(a + w), a + w), 0);
    let acc = 0; return row.cells.map(c => { const s = bounds.indexOf(acc), e = bounds.findIndex((b, i) => i > s && Math.abs(b - (acc + c.w)) < 90); acc += c.w; return Math.max(1, (e < 0 ? bounds.length - 1 : e) - s); });
  }

  /* ---------- HTML ---------- */
  const tw = n => (n / 56.69).toFixed(2) + 'mm';
  function runHTML(r) { return typeof r === 'string' ? esc(r) : '<i class="lv-bx' + (r.box ? ' on' : '') + '"></i>'; }
  function lineHTML(ln) {
    if (ln.segs.length === 1) return '<div class="lv-p">' + ln.segs[0].r.map(runHTML).join('') + '</div>';
    const tot = ln.segs.reduce((a, s) => a + s.w, 0);
    return '<div class="lv-ln">' + ln.segs.map(s => '<span style="flex:0 0 ' + (s.w / tot * 100).toFixed(2) + '%;text-align:' + (s.al === 'l' ? 'left' : 'right') + '">' + s.r.map(runHTML).join('') + '</span>').join('') + '</div>';
  }
  const bcss = b => { const o = {}; for (let i = 0; i < b.length; i += 2) o[b[i]] = b[i + 1]; return ['t', 'r', 'b', 'l'].map(k => 'border-' + { t: 'top', r: 'right', b: 'bottom', l: 'left' }[k] + ':' + (o[k] === 'd' ? '3px double #000' : '1px solid #000')).join(';'); };
  function html(S) {
    const R = rows(S);
    const tr = R.map(row => { const sp = spans(row); return '<tr' + (row.h ? ' style="height:' + tw(row.h) + '"' : '') + '>' + row.cells.map((c, i) =>
      '<td colspan="' + sp[i] + '" style="' + bcss(c.b) + (c.fill ? ';background:#' + c.fill : '') + (c.al === 'c' ? ';text-align:center' : '') + ';vertical-align:' + (c.va === 'b' ? 'bottom' : c.va === 'c' ? 'middle' : 'top') + '">' + c.lines.map(lineHTML).join('') + '</td>').join('') + '</tr>'; }).join('');
    return '<div class="lv"><div class="lv-frame"></div>' +
      '<div class="lv-code">رمزالنموذج: 01, 001 - 06</div>' +
      '<img class="lv-emb" alt="" src="' + g.LEAVE_EMBLEM + '">' +
      '<div class="lv-title">نموذج طلب إجازة سنوية</div>' +
      '<table class="lv-t"><colgroup>' + GRID.map(w => '<col style="width:' + tw(w) + '">').join('') + '</colgroup><tbody>' + tr + '</tbody></table></div>';
  }

  /* ---------- Word ---------- */
  const dataBytes = u => { const b = atob(u.split(',')[1]), a = new Uint8Array(b.length); for (let i = 0; i < b.length; i++) a[i] = b.charCodeAt(i); return a; };
  async function docxBlob(S) {
    const D = g.docx, f = S.font, SZ = 26, R = rows(S), none = { style: D.BorderStyle.NONE, size: 0, color: 'auto' };
    const run = (r, o) => typeof r === 'string' ? new D.TextRun(Object.assign({ text: r, bold: true, size: SZ, font: f, rightToLeft: /[؀-ۿ]/.test(r) || undefined }, o || {})) : new D.TextRun({ text: r.box ? '☒' : '☐', font: 'Segoe UI Symbol', size: SZ });
    const para = (runs, al, extra) => new D.Paragraph(Object.assign({ bidirectional: true, alignment: al === 'c' ? D.AlignmentType.CENTER : al === 'l' ? D.AlignmentType.END : D.AlignmentType.START, spacing: { before: 0, after: 0, line: 276 }, children: runs.map(r => run(r)) }, extra || {}));
    const bd = b => { const o = {}; for (let i = 0; i < b.length; i += 2) o[b[i]] = b[i + 1]; const m = k => ({ style: o[k] === 'd' ? D.BorderStyle.DOUBLE : D.BorderStyle.SINGLE, size: o[k] === 'd' ? 6 : 4, color: '000000' }); return { top: m('t'), bottom: m('b'), left: m('l'), right: m('r') }; };
    const lineEl = (ln, al) => {
      if (ln.segs.length === 1) return [para(ln.segs[0].r, al)];
      const ws = ln.segs.map(s => s.w), tot = ws.reduce((a, b) => a + b, 0), cw = ws.map(w => Math.round(w / tot * 10100));
      return [new D.Table({ width: { size: cw.reduce((a, b) => a + b, 0), type: D.WidthType.DXA }, columnWidths: cw, visuallyRightToLeft: true, borders: { top: none, bottom: none, left: none, right: none, insideHorizontal: none, insideVertical: none },
        rows: [new D.TableRow({ children: ln.segs.map((s, i) => new D.TableCell({ width: { size: cw[i], type: D.WidthType.DXA }, borders: { top: none, bottom: none, left: none, right: none }, margins: { top: 0, bottom: 0, left: 0, right: 0 }, children: [para(s.r, s.al)] })) })] }),
        new D.Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: D.LineRuleType.EXACT }, children: [new D.TextRun({ text: '', size: 2 })] })];
    };
    const trs = R.map(row => { const sp = spans(row); return new D.TableRow({ height: row.h ? { value: row.h, rule: D.HeightRule.ATLEAST } : undefined, cantSplit: true, children: row.cells.map((c, i) => {
      const kids = []; c.lines.forEach(ln => lineEl(ln, c.al).forEach(k => kids.push(k)));
      if (kids[kids.length - 1] instanceof D.Table) kids.push(new D.Paragraph({ children: [] }));
      return new D.TableCell({ width: { size: c.w, type: D.WidthType.DXA }, columnSpan: sp[i] > 1 ? sp[i] : undefined, borders: bd(c.b), verticalAlign: c.va === 'b' ? D.VerticalAlign.BOTTOM : c.va === 'c' ? D.VerticalAlign.CENTER : D.VerticalAlign.TOP,
        shading: c.fill ? { type: D.ShadingType.CLEAR, fill: c.fill, color: 'auto' } : undefined, children: kids }); }) }); });
    const tbl = new D.Table({ width: { size: W_TOTAL, type: D.WidthType.DXA }, columnWidths: GRID, alignment: D.AlignmentType.CENTER, visuallyRightToLeft: true,
      borders: { top: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' }, bottom: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' }, left: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' }, right: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' }, insideHorizontal: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' }, insideVertical: { style: D.BorderStyle.SINGLE, size: 4, color: '000000' } }, rows: trs });
    const tb = (s, c) => ({ style: D.BorderStyle.SINGLE, size: s, color: c, space: 4 });
    const code = new D.Paragraph({ bidirectional: true, alignment: D.AlignmentType.START,
      frame: { type: 'absolute', position: { x: 720, y: 720 }, width: 2800, height: 1300, anchor: { horizontal: D.FrameAnchorType.PAGE, vertical: D.FrameAnchorType.PAGE }, rule: D.HeightRule.EXACT },
      border: { top: tb(4, '808080'), right: tb(4, '808080'), left: tb(18, TEAL), bottom: tb(18, TEAL) }, spacing: { before: 420, after: 0 },
      children: [new D.TextRun({ text: 'رمزالنموذج: 01, 001 - 06', size: 22, font: f, rightToLeft: true })] });
    const emb = new D.Paragraph({ alignment: D.AlignmentType.CENTER, spacing: { before: 0, after: 0 }, children: [new D.ImageRun({ type: 'png', data: dataBytes(g.LEAVE_EMBLEM), transformation: { width: 117, height: 147 } })] });
    const title = new D.Paragraph({ bidirectional: true, alignment: D.AlignmentType.CENTER, spacing: { before: 160, after: 120 }, children: [new D.TextRun({ text: 'نموذج طلب إجازة سنوية', bold: true, size: 28, color: TEAL, font: f, rightToLeft: true })] });
    const pb = (side) => ({ style: D.BorderStyle.SINGLE, size: 18, color: TEAL, space: 24 });
    const doc = new D.Document({ creator: 'نادي المعلمين', title: 'نموذج طلب إجازة سنوية', styles: { default: { document: { run: { font: f, size: SZ } } } },
      sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 720, bottom: 600, left: 758, right: 758, header: 0, footer: 0 },
        borders: { pageBorders: { display: D.PageBorderDisplay.ALL_PAGES, offsetFrom: D.PageBorderOffsetFrom.PAGE }, pageBorderTop: pb(), pageBorderBottom: pb(), pageBorderLeft: pb(), pageBorderRight: pb() } }, bidi: true },
        children: [code, emb, title, tbl, new D.Paragraph({ spacing: { before: 0, after: 0, line: 20, lineRule: D.LineRuleType.EXACT }, children: [] })] }] });
    return D.Packer.toBlob(doc);
  }

  g.App = Object.assign(App, { leave: { html, docx: docxBlob, calcDays, fmt, rows } });
})(window);
