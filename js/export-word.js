/* ===== تصدير Word (.docx) — ملف حقيقي بصيغة OOXML مع دعم RTL ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const PX_PER_MM = 96 / 25.4;
  const TW_PER_MM = 56.6929;
  const mm = v => Math.round(v * TW_PER_MM);

  /* المكتبة تعكس left/right في الفقرات ثنائية الاتجاه (bidi)، والتحقق أُجري بالتحويل إلى PDF */
  function align(a) {
    const A = docx.AlignmentType;
    return { right: A.RIGHT, left: A.LEFT, center: A.CENTER, justify: A.BOTH }[a] || A.RIGHT;
  }

  function fontOf(name) { return { ascii: name, hAnsi: name, cs: name, eastAsia: name }; }

  function mkRun(text, o) {
    return new docx.TextRun({
      text, font: fontOf(o.font), size: Math.round(o.size * 2), sizeComplexScript: Math.round(o.size * 2),
      bold: !!o.b, boldComplexScript: !!o.b, italics: !!o.i, italicsComplexScript: !!o.i,
      underline: o.u ? { type: docx.UnderlineType.SINGLE } : undefined,
      rightToLeft: o.rtl !== false, color: o.color ? o.color.replace('#', '') : undefined,
      break: o.br ? 1 : undefined
    });
  }

  function mkPara(runs, o) {
    o = o || {};
    return new docx.Paragraph({
      children: runs, bidirectional: true, alignment: align(o.align || 'right'),
      spacing: { before: o.before || 0, after: o.after == null ? 0 : o.after, line: o.line, lineRule: o.line ? docx.LineRuleType.AUTO : undefined },
      indent: o.indent, border: o.border, keepNext: o.keepNext, keepLines: o.keepLines,
      tabStops: o.tabStops
    });
  }

  /* --- HTML نص المستند ← فقرات --- */
  function alignOf(n) {
    const a = (n.style && n.style.textAlign) || '';
    return /^(right|left|center|justify)$/.test(a) ? a : null;
  }
  function hasBlockChild(n) { return !!n.querySelector('p,div,ul,ol,table'); }

  function inlineNode(n, fmt, para) {
    if (n.nodeType === 3) {
      const t = n.nodeValue.replace(/[\r\n\t ]+/g, ' ');
      if (t) para.runs.push(Object.assign({ text: t }, fmt));
      return;
    }
    if (n.nodeType !== 1) return;
    if (n.tagName === 'BR') { para.runs.push({ br: true }); return; }
    const f = Object.assign({}, fmt);
    if (n.tagName === 'B') f.b = true;
    if (n.tagName === 'I') f.i = true;
    if (n.tagName === 'U') f.u = true;
    n.childNodes.forEach(c => inlineNode(c, f, para));
  }

  function walkBlocks(container, out, inherit) {
    let cur = null;
    container.childNodes.forEach(n => {
      if (n.nodeType === 3 && !n.nodeValue.trim() && !cur) return;
      const tag = n.nodeType === 1 ? n.tagName : '';
      if (tag === 'P' || tag === 'DIV') {
        cur = null;
        const al = alignOf(n) || inherit;
        if (hasBlockChild(n)) walkBlocks(n, out, al);
        else { const p = { runs: [], align: al }; inlineNode(n, {}, p); out.push(p); }
      } else if (tag === 'TABLE') {
        cur = null;
        const rows = [];
        n.querySelectorAll('tr').forEach(tr => {
          const cells = [];
          Array.from(tr.children).forEach(c => {
            if (!/^(TD|TH)$/.test(c.tagName)) return;
            const blocks = []; walkBlocks(c, blocks, 'right');
            if (!blocks.length) blocks.push({ runs: [], align: 'right' });
            cells.push({ blocks, head: c.tagName === 'TH' });
          });
          if (cells.length) rows.push(cells);
        });
        if (rows.length) out.push({ table: rows });
      } else if (tag === 'UL' || tag === 'OL') {
        cur = null; let i = 0;
        n.childNodes.forEach(li => {
          if (li.tagName !== 'LI') return;
          i++;
          const p = { runs: [], align: alignOf(li) || inherit, list: tag === 'OL' ? i + '.' : '•' };
          inlineNode(li, {}, p); out.push(p);
        });
      } else {
        if (!cur) { cur = { runs: [], align: inherit }; out.push(cur); }
        inlineNode(n, {}, cur);
      }
    });
  }

  function blockToPara(b, S, line, bold) {
    while (b.runs.length && b.runs[b.runs.length - 1].br) b.runs.pop();
    const base = { font: S.bodyFont || S.font, size: S.bodySize, b: !!bold };
    const runs = [];
    if (b.list) runs.push(mkRun(b.list + '\t', base));
    b.runs.forEach(r => {
      if (r.br) runs.push(new docx.TextRun({ break: 1 }));
      else runs.push(mkRun(r.text, Object.assign({}, base, r)));
    });
    return mkPara(runs, {
      align: b.align, after: 80, line,
      indent: b.list ? { start: 567, hanging: 340 } : undefined,
      tabStops: b.list ? [{ type: docx.TabStopType.LEFT, position: 567 }] : undefined
    });
  }

  function tableFromRows(rows, S, line) {
    const cols = Math.max.apply(null, rows.map(r => r.length));
    const w = Math.floor(CONTENT_W / cols);
    const B = { style: docx.BorderStyle.SINGLE, size: 4, color: '000000' };
    return new docx.Table({
      width: { size: w * cols, type: docx.WidthType.DXA }, columnWidths: Array(cols).fill(w),
      layout: docx.TableLayoutType.FIXED, visuallyRightToLeft: true,
      borders: { top: B, bottom: B, left: B, right: B, insideHorizontal: B, insideVertical: B },
      rows: rows.map(r => new docx.TableRow({
        cantSplit: true,
        children: Array.from({ length: cols }, (_, i) => {
          const c = r[i] || { blocks: [{ runs: [], align: 'right' }], head: false };
          return new docx.TableCell({
            width: { size: w, type: docx.WidthType.DXA },
            shading: c.head ? { type: docx.ShadingType.CLEAR, fill: 'F0F0F0', color: 'auto' } : undefined,
            margins: { top: 50, bottom: 50, left: 80, right: 80 },
            children: c.blocks.map(b => blockToPara(Object.assign({}, b, { align: b.align === 'justify' ? 'right' : b.align }), S, line, c.head))
          });
        })
      }))
    });
  }

  function bodyParagraphs(html, S, defAlign) {
    const t = document.createElement('template');
    t.innerHTML = App.sanitize(html);
    const blocks = [];
    walkBlocks(t.content, blocks, defAlign || 'justify');
    const line = Math.round(S.lineHeight * 240);
    const out = [];
    blocks.forEach(b => {
      if (b.table) { out.push(tableFromRows(b.table, S, line)); out.push(mkPara([], { after: 80 })); }
      else out.push(blockToPara(b, S, line));
    });
    return out;
  }

  /* --- الصور --- */
  function dataUrlToBytes(url) {
    const b64 = url.split(',')[1]; const bin = atob(b64);
    const u = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) u[i] = bin.charCodeAt(i);
    return u;
  }
  function imgSize(url) {
    return new Promise(res => { const im = new Image(); im.onload = () => res([im.naturalWidth, im.naturalHeight]); im.onerror = () => res([1, 1]); im.src = url; });
  }
  async function logoRun(L, sizeMm) {
    if (!L.show || !L.data) return null;
    return fitImage(L.data, sizeMm, sizeMm);
  }
  async function fitImage(url, wMm, hMm) {
    const [w, h] = await imgSize(url);
    const k = Math.min(wMm * PX_PER_MM / w, hMm * PX_PER_MM / h);
    return new docx.ImageRun({
      type: /^data:image\/jpe?g/.test(url) ? 'jpg' : 'png', data: dataUrlToBytes(url),
      transformation: { width: Math.max(1, Math.round(w * k)), height: Math.max(1, Math.round(h * k)) }
    });
  }

  /* --- جداول --- */
  const NONE = { style: docx.BorderStyle.NONE, size: 0, color: 'FFFFFF' };
  const NOBORDERS = { top: NONE, bottom: NONE, left: NONE, right: NONE };
  const LINE = { style: docx.BorderStyle.SINGLE, size: 6, color: '000000' };

  function mkTable(widths, cells, opt) {
    opt = opt || {};
    const total = widths.reduce((a, b) => a + b, 0);
    return new docx.Table({
      width: { size: total, type: docx.WidthType.DXA }, columnWidths: widths,
      layout: docx.TableLayoutType.FIXED, visuallyRightToLeft: true,
      borders: { top: NONE, bottom: NONE, left: NONE, right: NONE, insideHorizontal: NONE, insideVertical: NONE },
      rows: [new docx.TableRow({
        cantSplit: true,
        children: cells.map((c, i) => new docx.TableCell({
          width: { size: widths[i], type: docx.WidthType.DXA },
          verticalAlign: opt.valign || docx.VerticalAlign.CENTER,
          borders: opt.rule ? { top: LINE, bottom: LINE, left: NONE, right: NONE } : NOBORDERS,
          margins: opt.rule ? { top: mm(2), bottom: mm(2), left: 60, right: 60 } : undefined,
          children: c
        }))
      })]
    });
  }

  const CONTENT_W = mm(176); // 210 - 2×17

  async function headerBlock(S) {
    const C = App.CONST, f = S.font;
    const [r, l] = await Promise.all([logoRun(S.logos.ministry, S.logoSize), logoRun(S.logos.club, S.logoSize)]);
    const both = r || l || (S.logos.ministry.show && S.logos.ministry.data) || (S.logos.club.show && S.logos.club.data);
    const t = (text, size, b) => mkPara([mkRun(text, { font: f, size, b })], { align: 'center', line: 300 });
    const club = mkPara([mkRun(C.CLUB_PREFIX + (S.clubName ? ' ' + S.clubName : ''), { font: f, size: 22, b: true })], { align: 'center', before: 60 });
    const textCell = [t(C.MINISTRY, 11.5), t(C.DEPT, 13), t(C.CLUBS, 13, true), club];
    if (!both) return [mkTable([CONTENT_W], [textCell])];
    const lw = mm(S.logoSize + 4);
    const logoCell = run => [mkPara(run ? [run] : [], { align: 'center' })];
    return [mkTable([lw, CONTENT_W - 2 * lw, lw], [logoCell(r), textCell, logoCell(l)])];
  }

  function spacer(twips) { return mkPara([], { after: twips }); }

  async function signCell(S, key) {
    const f = S.font, size = S.bodySize, p = S.people[key];
    const imgs = [];
    const sg = p.showSign && p.signImg ? await fitImage(p.signImg, 60, S.people.signSize) : null;
    const st = p.showStamp && p.stampImg ? await fitImage(p.stampImg, 60, S.people.stampSize) : null;
    const pair = S.people.stampPos === 'left' ? [sg, st] : [st, sg];
    pair.forEach(x => { if (x) { if (imgs.length) imgs.push(mkRun('  ', { font: f, size })); imgs.push(x); } });
    return [
      mkPara([mkRun(App.TITLES[key], { font: f, size, b: true })], { align: 'center', keepNext: true }),
      ...(imgs.length ? [mkPara(imgs, { align: 'center', before: mm(1), keepNext: true })] : []),
      mkPara([mkRun(p.name || ' ', { font: f, size })], { align: 'center', before: imgs.length ? mm(1) : mm(11) })
    ];
  }

  async function signatureBlock(S, type) {
    const w = mm(66), d = S[type];
    const who = [d.showHead ? 'head' : '', d.showManager ? 'manager' : ''].filter(Boolean);
    if (!who.length) return mkPara([], {});
    if (who.length === 1) return mkTable([CONTENT_W - w, w], [[mkPara([], {})], await signCell(S, who[0])], { valign: docx.VerticalAlign.TOP });
    return mkTable([w, CONTENT_W - 2 * w, w], [await signCell(S, 'head'), [mkPara([], {})], await signCell(S, 'manager')], { valign: docx.VerticalAlign.TOP });
  }

  async function buildMain(S, type) {
    const C = App.CONST, f = S.font, size = S.bodySize;
    const out = await headerBlock(S);
    const d = S[type];
    const sp = v => mm(v);

    if (App.isLetter(type)) {
      const lab = (t, v, a) => [mkPara([mkRun(t + ' ', { font: f, size: 11, b: true }), mkRun(v || '', { font: f, size: 11 })], { align: a })];
      const greg = App.formatGregorian(d.date);
      const w = [Math.round(CONTENT_W * 0.25), Math.round(CONTENT_W * 0.32)]; w.push(CONTENT_W - w[0] - w[1]);
      out.push(spacer(sp(3)));
      out.push(mkTable(w, [
        lab('الرقم:', d.number, 'right'),
        d.showHijri ? lab('التاريخ:', d.hijri, 'center') : [mkPara([], {})],
        d.showHijri ? lab('الموافق:', greg, 'left') : lab('التاريخ:', greg, 'left')
      ], { rule: true }));
      out.push(spacer(sp(3)));
      const to = [];
      if (type === 'internal' && d.toDir) {
        to.push(mkPara([mkRun(C.DIRECTOR + ' الأكرم', { font: f, size })], { align: 'center', line: 300 }));
        if (d.dirName.trim()) to.push(mkPara([mkRun(d.dirName.trim(), { font: f, size, b: true })], { align: 'center', line: 300 }));
      }
      [d.to0, d.to1, d.to2].forEach((on, i) => { if (on) to.push(mkPara([mkRun(C.RECIPIENTS[i], { font: f, size })], { align: 'center', line: 300 })); });
      if (type === 'internal' && d.toCustom.trim()) d.toCustom.split('\n').filter(x => x.trim()).forEach(x => to.push(mkPara([mkRun(x.trim(), { font: f, size })], { align: 'center', line: 300 })));
      out.push(...to);
      if (to.length) out.push(spacer(sp(3)));
      out.push(mkPara([
        mkRun('الموضوع:', { font: f, size, b: true, u: true }),
        mkRun(d.subject ? '  ( ' + d.subject + ' )' : '  (        )', { font: f, size, b: true })
      ], { align: 'right', after: sp(3), keepNext: true }));
    } else {
      if (d.lineShow) {
        const dash = { solid: docx.BorderStyle.SINGLE, dashed: docx.BorderStyle.DASHED, dotted: docx.BorderStyle.DOTTED, double: docx.BorderStyle.DOUBLE }[d.lineStyle] || docx.BorderStyle.SINGLE;
        const ind = Math.round((100 - d.linePct) / 100 * CONTENT_W / 2);
        out.push(new docx.Paragraph({
          children: [], bidirectional: true, spacing: { before: sp(d.lineBefore), after: sp(d.lineAfter), line: 20, lineRule: docx.LineRuleType.EXACT },
          indent: ind > 0 ? { start: ind, end: ind } : undefined,
          border: { top: { style: dash, size: Math.max(2, Math.round(d.lineWidth * 8)), space: 0, color: d.lineColor.replace('#', '') } }
        }));
      } else out.push(spacer(sp(3)));
      if (d.showDate || d.showNumUnder) {
        const dl = d.showHijri ? { h: d.hijri, g: App.formatGregorian(d.date) } : { h: '', g: App.formatGregorian(d.date) };
        const dp = (t, v) => mkPara([mkRun(t + ' ', { font: f, size: 11, b: true }), mkRun(v, { font: f, size: 11 })], { align: 'right', line: 276 });
        if (d.showDate) {
          if (dl.h) { out.push(dp('التاريخ:', dl.h + 'هـ')); out.push(dp('الموافق:', dl.g + 'م')); } else out.push(dp('التاريخ:', dl.g + 'م'));
        }
        if (d.showNumUnder) out.push(dp('الرقم:', d.number || ' '));
      }
      out.push(spacer(sp(3)));
      out.push(mkPara([mkRun('تعميم هيئة عامة', { font: f, size: d.titleSize, b: true })], { align: d.titleAlign }));
      if (d.showTitleNum) out.push(mkPara([mkRun('تعميم رقم ' + (d.number || ''), { font: f, size: 14, b: true })], { align: d.titleAlign, after: sp(6) }));
      else out.push(spacer(sp(4)));
      out.push(mkPara([
        mkRun('الموضوع:', { font: f, size, b: true, u: true }),
        mkRun(d.subject ? ' ' + d.subject : '', { font: f, size, b: true })
      ], { align: 'right', after: sp(3), keepNext: true }));
    }

    out.push(...bodyParagraphs(d.body, S, type === 'circular' ? d.bodyAlign : 'justify'));
    out.push(mkPara([mkRun(C.CLOSING, { font: f, size })], { align: 'center', before: sp(7), after: sp(8), keepNext: true }));
    out.push(await signatureBlock(S, type));
    out.push(mkPara([], {}));
    return out;
  }

  function iconPng(name, color, px) {
    return new Promise(res => {
      const k = 4, svg = App.iconSVG(name, color, px * k, 'px');
      const im = new Image();
      im.onload = () => {
        const c = document.createElement('canvas'); c.width = px * k; c.height = px * k;
        c.getContext('2d').drawImage(im, 0, 0, c.width, c.height);
        res(c.toDataURL('image/png'));
      };
      im.onerror = () => res('');
      im.src = 'data:image/svg+xml;charset=utf-8,' + encodeURIComponent(svg);
    });
  }

  async function buildFooter(S) {
    const f = S.footer;
    if (!f.show || !App.footHasContent(S)) return null;
    const jc = f.align;
    const base = { font: f.font, size: f.size, b: f.bold, color: f.color };
    const paras = [];
    const border = f.line ? { top: { style: docx.BorderStyle.SINGLE, size: Math.max(2, Math.round(f.lineWidth * 8)), space: 4, color: '000000' } } : undefined;
    let first = true;
    if (f.name.trim()) {
      paras.push(mkPara([mkRun(f.name.trim(), { font: f.font, size: f.nameSize, b: f.nameBold, color: f.color })], { align: f.nameAlign, line: 276, border, after: 20 }));
      first = false;
    }
    const items = App.footItems(S);
    if (items.length) {
      const px = Math.max(8, Math.round(f.iconSize * 96 / 72));
      const mmPx = v => Math.round(v * PX_PER_MM);
      const runs = [];
      for (let i = 0; i < items.length; i++) {
        const it = items[i];
        if (i > 0) runs.push(mkRun(' '.repeat(Math.max(2, Math.round(f.gap * 1.6))), base));
        let ic = null;
        if (it.icon) {
          const url = await iconPng(it.icon, f.color, px);
          if (url) ic = new docx.ImageRun({ type: 'png', data: dataUrlToBytes(url), transformation: { width: px, height: px } });
        }
        const txt = mkRun((it.icon || !it.label ? '' : it.label + ': ') + it.value, Object.assign({}, base, { rtl: !it.ltr }));
        if (ic && f.iconPos === 'after') { runs.push(txt); runs.push(mkRun(' ', base)); runs.push(ic); }
        else if (ic) { runs.push(ic); runs.push(mkRun(' ', base)); runs.push(txt); }
        else runs.push(txt);
      }
      paras.push(mkPara(runs, { align: jc, line: 276, border: first ? border : undefined }));
      first = false;
    }
    App.footExtra(S).forEach(x => { paras.push(mkPara([mkRun(x, base)], { align: jc, line: 276, border: first ? border : undefined })); first = false; });
    return new docx.Footer({ children: paras });
  }

  async function buildDocx(S0, type) {
    const S = App.withEffective(S0, type);
    const children = await buildMain(S, type);
    const footer = await buildFooter(S);
    const attSections = [];
    for (const p of App.attach.flat(type)) {
      const img = await fitImage(p.url, 178, 250);
      attSections.push({
        properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: mm(10), bottom: mm(10), left: mm(15), right: mm(15), footer: 0, header: 0 } }, bidi: true },
        footers: { default: new docx.Footer({ children: [new docx.Paragraph({ children: [] })] }) },
        children: [
          mkPara([mkRun(p.cap, { font: S.font, size: 10 })], { align: 'center', after: 120 }),
          mkPara([img], { align: 'center' })
        ]
      });
    }
    const doc = new docx.Document({
      creator: 'أندية المعلمين والعاملين',
      title: S[type].subject || DOC_TITLES[type],
      styles: { default: { document: { run: { font: fontOf(S.font), size: Math.round(S.bodySize * 2), rightToLeft: true, language: { value: 'ar-JO', bidirectional: 'ar-JO' } } } } },
      sections: [{
        properties: {
          page: {
            size: { width: 11906, height: 16838 },
            margin: { top: mm(10), bottom: footer ? mm(20) : mm(12), left: mm(17), right: mm(17), footer: mm(7), header: 0 }
          },
          bidi: true
        },
        footers: footer ? { default: footer } : undefined,
        children
      }].concat(attSections)
    });
    return docx.Packer.toBlob(doc);
  }

  const DOC_TITLES = { letter: 'كتاب للوزارة', internal: 'كتاب داخلي', thanks: 'كتاب شكر', circular: 'تعميم هيئة عامة' };
  const LATIN = { letter: 'ministry-letter', internal: 'internal-letter', thanks: 'thanks-letter', circular: 'general-circular' };
  /** اسم الملف = موضوع الكتاب (بعد إزالة المحارف غير المسموحة في أسماء الملفات) */
  function fileName(S, type) {
    let n = (S[type].subject || '').replace(/[\\/:*?"<>|\u0000-\u001f]+/g, ' ').replace(/\s+/g, ' ').trim().replace(/[. ]+$/, '');
    if (n.length > 80) n = n.slice(0, 80).trim();
    if (n) return n;
    const num = (S[type].number || '').replace(/[^A-Za-z0-9]+/g, '-').replace(/^-|-$/g, '');
    return LATIN[type] + (num ? '-' + num : '');
  }

  g.App = Object.assign(App, { buildDocx, fileName, DOC_TITLES });
})(window);
