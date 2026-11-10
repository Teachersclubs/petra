/* ===== توليد HTML المستند (للمعاينة والطباعة وتصدير PDF) ===== */
(function (g) {
  'use strict';
  const App = g.App;

  const esc = s => String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));

  const MINISTRY = 'وزارة التربية والتعليم وتنمية الموارد البشرية';
  const DEPT = 'إدارة النشاطات التربوية';
  const CLUBS = 'أندية المعلمين والعاملين';
  const CLUB_PREFIX = 'نــــــادي';
  const CLOSING = 'واقبلوا فائق الاحترام والتقدير.';
  const DIRECTOR = 'عطوفة مدير التربية والتعليم';
  const RECIPIENTS = [
    'معالي وزير التربية والتعليم وتنمية الموارد البشرية الأكرم',
    'عطوفة مدير إدارة النشاطات التربوية الأكرم',
    'عطوفة مدير أندية المعلمين والعاملين الأكرم'
  ];

  /* --- تنقية HTML نص المستند: وسوم مسموحة فقط --- */
  const ALLOWED = { B: 'b', STRONG: 'b', I: 'i', EM: 'i', U: 'u', BR: 'br', P: 'p', DIV: 'div', UL: 'ul', OL: 'ol', LI: 'li', H1: 'p', H2: 'p', H3: 'p', TABLE: 'table', THEAD: 'thead', TBODY: 'tbody', TR: 'tr', TD: 'td', TH: 'th' };
  function sanitize(html) {
    const t = document.createElement('template');
    t.innerHTML = html || '';
    const walk = node => {
      let out = '';
      node.childNodes.forEach(n => {
        if (n.nodeType === 3) { out += esc(n.nodeValue); return; }
        if (n.nodeType !== 1) return;
        const tag = ALLOWED[n.tagName];
        const inner = walk(n);
        let bold = false, ital = false, und = false;
        if (n.tagName === 'SPAN' || n.tagName === 'FONT') {
          const st = n.style || {};
          bold = st.fontWeight === 'bold' || +st.fontWeight >= 600;
          ital = st.fontStyle === 'italic';
          und = /underline/.test(st.textDecoration || st.textDecorationLine || '');
          let r = inner;
          if (und) r = '<u>' + r + '</u>'; if (ital) r = '<i>' + r + '</i>'; if (bold) r = '<b>' + r + '</b>';
          out += r; return;
        }
        if (!tag) { out += inner; return; }
        if (tag === 'br') { out += '<br>'; return; }
        let attr = '';
        const ta = (n.style && n.style.textAlign) || n.getAttribute('align') || '';
        if (/^(right|left|center|justify)$/.test(ta) && ['p', 'div', 'li', 'td', 'th'].includes(tag)) attr = ' style="text-align:' + ta + '"';
        out += '<' + tag + attr + '>' + inner + '</' + tag + '>';
      });
      return out;
    };
    return walk(t.content);
  }

  const PAGE_H = 1122;
  const TITLES = { head: 'رئيس الهيئة الإدارية', manager: 'مدير النادي' };
  const isLetter = t => t === 'letter' || t === 'internal' || t === 'thanks';
  const DATE_DOC = t => isLetter(t) || t === 'assign' || t === 'claim';

  /* ---------- أيقونات الذيل ---------- */
  const ICONS = {
    phone: '<path d="M22 16.92v3a2 2 0 0 1-2.18 2 19.79 19.79 0 0 1-8.63-3.07 19.5 19.5 0 0 1-6-6 19.79 19.79 0 0 1-3.07-8.67A2 2 0 0 1 4.11 2h3a2 2 0 0 1 2 1.72 12.84 12.84 0 0 0 .7 2.81 2 2 0 0 1-.45 2.11L8.09 9.91a16 16 0 0 0 6 6l1.27-1.27a2 2 0 0 1 2.11-.45 12.84 12.84 0 0 0 2.81.7A2 2 0 0 1 22 16.92z"/>',
    whatsapp: '<path d="M21 11.5a8.38 8.38 0 0 1-.9 3.8 8.5 8.5 0 0 1-7.6 4.7 8.38 8.38 0 0 1-3.8-.9L3 21l1.9-5.7a8.38 8.38 0 0 1-.9-3.8 8.5 8.5 0 0 1 4.7-7.6 8.38 8.38 0 0 1 3.8-.9h.5a8.48 8.48 0 0 1 8 8v.5z"/><path d="M9 8.6c.3 2.2 2.6 4.6 5.4 5.6l1.3-1.3-1.9-1-.9.7c-.8-.4-1.5-1.1-1.9-1.9l.7-.9-1-1.9z" fill="currentColor" stroke-width="1"/>',
    mail: '<rect x="2" y="4" width="20" height="16" rx="2"/><path d="M22 6l-10 7L2 6"/>',
    facebook: '<path d="M18 2h-3a5 5 0 0 0-5 5v3H7v4h3v8h4v-8h3l1-4h-4V7a1 1 0 0 1 1-1h3z"/>',
    pin: '<path d="M21 10c0 7-9 13-9 13s-9-6-9-13a9 9 0 0 1 18 0z"/><circle cx="12" cy="10" r="3"/>',
    globe: '<circle cx="12" cy="12" r="10"/><path d="M2 12h20"/><path d="M12 2a15.3 15.3 0 0 1 4 10 15.3 15.3 0 0 1-4 10 15.3 15.3 0 0 1-4-10 15.3 15.3 0 0 1 4-10z"/>'
  };
  /** SVG نصي؛ pt للعرض داخل HTML، أو px لتحويله إلى صورة */
  function iconSVG(name, color, size, unit) {
    return '<svg xmlns="http://www.w3.org/2000/svg" class="ic" viewBox="0 0 24 24" width="' + size + unit + '" height="' + size + unit + '" fill="none" stroke="' + color + '" stroke-width="2" stroke-linecap="round" stroke-linejoin="round" style="color:' + color + '">' + ICONS[name].replace(/currentColor/g, color) + '</svg>';
  }

  function logoImg(L, size) {
    if (!L.show || !L.data) return '';
    return '<img src="' + L.data + '" alt="" style="max-width:' + size + 'mm;max-height:' + size + 'mm">';
  }

  function headerHTML(S, E) {
    const size = E.logoSize;
    const R = logoImg(S.logos.ministry, size), L = logoImg(S.logos.club, size);
    const both = (S.logos.ministry.show && S.logos.ministry.data) || (S.logos.club.show && S.logos.club.data);
    const col = (c) => both ? '<div class="h-logo" style="width:' + size + 'mm;height:' + size + 'mm">' + c + '</div>' : '';
    return '<div class="hdr">' + col(R) +
      '<div class="h-text">' +
      '<div class="h1">' + MINISTRY + '</div>' +
      '<div class="h2">' + DEPT + '</div>' +
      '<div class="h3">' + CLUBS + '</div>' +
      '<div class="h4">' + CLUB_PREFIX + (S.clubName ? ' ' + esc(S.clubName) : '') + '</div>' +
      '</div>' + col(L) + '</div>';
  }

  function ph(text) { return '<span class="ph">' + text + '</span>'; }

  /* ---------- التوقيع والختم: هندسة موحّدة (مم) للمعاينة وPDF والطباعة وWord ---------- */
  const SG = { W: 66, TITLE_H: 7, NAME_H: 7, PADB: 5, PADZ: 2, BLANK: 11, MAXW: 60, PAGE_W: 210, CL: 17, CR: 193 };

  /** من يظهر ومن أي جهة: head يمينًا ثم manager؛ والمفرد يظهر يسارًا */
  function signers(d) {
    const who = [d.showHead ? 'head' : '', d.showManager ? 'manager' : ''].filter(Boolean);
    return who;
  }
  function boxLeft(key, who, fin) {
    return (key === 'head' && (who.length === 2 || fin)) ? SG.CR - SG.W : SG.CL;
  }

  function nameWidth(S, p) {
    const n = (p.name || '').length;
    return Math.max(12, Math.min(SG.W, n * (S.bodySize || 14) * 0.3528 * 0.5 + 4));
  }

  /** هندسة صندوق موقّع واحد: مستطيلات العناصر بإحداثيات الصندوق (مم) */
  function signGeom(S, key) {
    const o = S.people, p = o[key], pos = p.pos;
    const fit = (h, ar) => { let w = h * ar; if (w > SG.MAXW) { h = h * SG.MAXW / w; w = SG.MAXW; } return { w, h }; };
    const items = [];
    if (p.showSign && p.signImg) { const f = fit(o.signSize, p.signAR || 2.2); items.push({ k: 'sign', w: f.w, h: f.h }); }
    if (p.showStamp && p.stampImg) { const f = fit(o.stampSize, p.stampAR || 1); items.push({ k: 'stamp', w: f.w, h: f.h }); }
    // ترتيب العرض الفيزيائي من اليسار إلى اليمين
    const order = o.stampPos === 'left' ? ['stamp', 'sign'] : ['sign', 'stamp'];
    items.sort((x, y) => order.indexOf(x.k) - order.indexOf(y.k));
    const gi = o.gapImg, gn = o.gapName;
    const maxH = items.reduce((m, it) => Math.max(m, it.h), 0);
    const zoneH = items.length ? maxH + 2 * SG.PADZ : SG.BLANK;
    const total = items.reduce((t, it) => t + it.w, 0) + gi * Math.max(0, items.length - 1);
    let zoneTop, nameY;
    if (o.mode === 'below') { nameY = SG.TITLE_H + gn; zoneTop = nameY + SG.NAME_H + gn; }
    else if (o.mode === 'overlap' && items.length) { zoneTop = SG.TITLE_H; nameY = SG.TITLE_H + zoneH - 5; }
    else { zoneTop = SG.TITLE_H; nameY = SG.TITLE_H + zoneH + gn; }
    const H = (o.mode === 'below' ? zoneTop + zoneH : nameY + SG.NAME_H) + SG.PADB;
    let x = (SG.W - total) / 2;
    const rects = {};
    items.forEach(it => {
      const q = pos[it.k];
      rects[it.k] = { x: x + q.x, y: zoneTop + (zoneH - it.h) / 2 + q.y, w: it.w, h: it.h };
      x += it.w + gi;
    });
    const nw = nameWidth(S, p);
    rects.name = { x: (SG.W - nw) / 2 + pos.name.x, y: nameY + pos.name.y, w: nw, h: SG.NAME_H, text: !!(p.name || '').trim() };
    rects.title = { x: (SG.W - SG.W * 0.7) / 2, y: 0, w: SG.W * 0.7, h: SG.TITLE_H };
    return { H, nameY0: nameY, rects, zoneH };
  }

  /** يحصر العناصر داخل الصفحة والصندوق حتى لا تختفي خارج الحدود (للعرض فقط) */
  function signGeomSafe(S, key, who, fin) {
    const g = signGeom(S, key), bl = boxLeft(key, who, fin);
    ['sign', 'stamp', 'name'].forEach(k => {
      const r = g.rects[k]; if (!r) return;
      r.x = Math.max(4 - bl, Math.min(SG.PAGE_W - 4 - bl - r.w, r.x));
      r.y = Math.max(k === 'name' ? SG.TITLE_H : 0, Math.min(g.H - r.h, r.y));
    });
    return g;
  }

  /** فحص التداخل والخروج عن الصفحة؛ يعيد رسائل المشاكل (فارغة = سليم) */
  function signIssues(S, d, type) {
    const who = signers(d), out = [], all = [], fin = type === 'assign' && !!d.showFinance;
    who.forEach(key => {
      const g = signGeom(S, key), bl = boxLeft(key, who, fin);
      ['sign', 'stamp', 'name', 'title'].forEach(k => {
        const r = g.rects[k]; if (!r) return;
        const rr = { key, k, x: r.x + bl, y: r.y, w: r.w, h: r.h, text: r.text };
        all.push(rr);
        if (k === 'title') return;
        if (rr.x < 4 || rr.x + rr.w > SG.PAGE_W - 4) out.push('العنصر يخرج عن حدود الصفحة');
        if (r.y < (k === 'name' ? SG.TITLE_H - 0.01 : -0.01) || r.y + r.h > g.H + 0.01) out.push('العنصر يخرج عن حدود منطقة التوقيع');
      });
    });
    if (S.people.protect) {
      const hit = (a, b) => a.x < b.x + b.w - 0.01 && b.x < a.x + a.w - 0.01 && a.y < b.y + b.h - 0.01 && b.y < a.y + a.h - 0.01;
      for (let i = 0; i < all.length; i++) for (let j = i + 1; j < all.length; j++) {
        const a = all[i], b = all[j];
        if (a.key !== b.key && false) continue;
        const imgNameOk = S.people.mode === 'overlap' && a.key === b.key && ((a.k === 'name' && b.k !== 'title' && b.k !== 'name') || (b.k === 'name' && a.k !== 'title' && a.k !== 'name'));
        if (imgNameOk) continue;
        if ((a.k === 'name' && !a.text) || (b.k === 'name' && !b.text)) { if (a.k === 'name' || b.k === 'name') { const other = a.k === 'name' ? b : a; if (other.k !== 'title') continue; } }
        if (hit(a, b)) out.push('سيتداخل هذا العنصر مع عنصر آخر');
      }
    }
    return Array.from(new Set(out));
  }

  const f1 = v => Math.round(v * 100) / 100;
  function signBox(S, key, who, fin) {
    const g = signGeomSafe(S, key, who, fin), R = g.rects, p = S.people[key];
    let h = '<div class="sign-box sg" style="width:' + SG.W + 'mm;height:' + f1(g.H) + 'mm">';
    h += '<div class="sg-title" style="height:' + SG.TITLE_H + 'mm;line-height:' + SG.TITLE_H + 'mm">' + TITLES[key] + '</div>';
    h += '<div class="sg-name" style="top:' + f1(R.name.y) + 'mm;height:' + SG.NAME_H + 'mm;line-height:' + SG.NAME_H + 'mm;left:' + f1(R.name.x + R.name.w / 2 - SG.W / 2) + 'mm;right:' + f1(-(R.name.x + R.name.w / 2 - SG.W / 2)) + 'mm">' + esc(p.name) + '</div>';
    ['sign', 'stamp'].forEach(k => {
      const r = R[k]; if (!r) return;
      h += '<img class="' + k + '-img sg-img" src="' + (k === 'sign' ? p.signImg : p.stampImg) + '" alt="" style="left:' + f1(r.x) + 'mm;top:' + f1(r.y) + 'mm;width:' + f1(r.w) + 'mm;height:' + f1(r.h) + 'mm">';
    });
    return h + '</div>';
  }

  /** بلوك «اللجنة المالية» أسفل مكان رئيس الهيئة (قرار التكليف) */
  function financeHTML(S, d) {
    if (!d.showFinance) return '';
    const lines = (d.financeText || '').split('\n').map(x => x.trim()).filter(Boolean);
    return '<div class="fin-box" style="width:' + SG.W + 'mm"><div class="fin-title">اللجنة المالية:</div>' + (lines.length ? lines.map(l => '<div class="fin-line">' + esc(l) + '</div>').join('') : '<div class="fin-line">' + ph('&nbsp;') + '</div>') + '</div>';
  }

  /** رئيس الهيئة يمينًا ثم مدير النادي؛ لكل منهما خيار إظهار */
  function signHTML(S, type) {
    const d = S[type], who = signers(d), fin = type === 'assign' && d.showFinance;
    if (!who.length && !fin) return '';
    const col = k => '<div class="sg-col">' + (who.includes(k) ? signBox(S, k, who, fin) : '') + (k === 'head' && fin ? financeHTML(S, d) : '') + '</div>';
    if (fin && !who.includes('head')) {
      return '<div class="sign sign-two">' + col('head') + (who.includes('manager') ? col('manager') : '') + '</div>';
    }
    if (who.length === 2) return '<div class="sign sign-two">' + col('head') + col('manager') + '</div>';
    return '<div class="sign ' + (fin ? 'sign-two' : 'sign-one') + '">' + who.map(k => col(k)).join('') + '</div>';
  }

  function dateLines(d) {
    const greg = App.formatGregorian(d.date);
    return d.showHijri ? { hijri: d.hijri, greg } : { hijri: '', greg };
  }

  /** تاريخ قصير: 1/10/2026 */
  function shortDate(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    return m ? (+m[3]) + '/' + (+m[2]) + '/' + m[1] : '';
  }
  function assignDateStr(d) { return (d.assignDateText || '').trim() || shortDate(d.assignDate); }
  function fillAssign(html, d) {
    const val = assignDateStr(d);
    return html.split(App.ASSIGN_TOKEN).join('<span class="adate">' + esc(val || '................') + '</span>');
  }

  let curVariant = null; // نص الجهة في النسخة المكرّرة (null = النسخة الأساسية)

  /** نسخ كتاب الشكر: الأولى الأساسية ثم نسخة لكل سطر في «الجهات الإضافية» */
  function variants(S, type) {
    const out = [null];
    if (type === 'thanks' && S.thanks.copiesOn) S.thanks.copiesText.split('\n').map(x => x.trim()).filter(Boolean).forEach(x => out.push(x));
    return out;
  }
  function toLines(S, variant) {
    if (variant) return [variant];
    return S.thanks.toText.split('\n').map(x => x.trim()).filter(Boolean);
  }

  function dateRowHTML(d) {
    const dl = dateLines(d);
    const numCell = '<div class="dr-r"><b>الرقم:</b> ' + (d.number ? esc(d.number) : ph('&nbsp;')) + '</div>';
    let dateCells;
    if (d.showHijri) {
      dateCells = '<div class="dr-c"><b>التاريخ:</b> <bdi>' + esc(dl.hijri) + '</bdi></div>' +
        '<div class="dr-l"><b>الموافق:</b> <bdi>' + esc(dl.greg) + '</bdi></div>';
    } else {
      dateCells = '<div class="dr-c"></div><div class="dr-l"><b>التاريخ:</b> <bdi>' + esc(dl.greg) + '</bdi></div>';
    }
    return '<div class="date-row">' + numCell + dateCells + '</div>';
  }

  function assignHTML(S, E) {
    const d = S.assign;
    const body = fillAssign(sanitize(d.body), d);
    return headerHTML(S, E) + dateRowHTML(d) +
      '<div class="doc-title">قرار تكليف</div>' +
      '<div class="who-line">' + esc(d.who) + ': ' + (d.assignee ? esc(d.assignee) : ph('................')) + '</div>' +
      '<div class="body">' + (body || ph('نص القرار…')) + '</div>' +
      (d.showClosing ? '<div class="closing">' + CLOSING + '</div>' : '<div style="height:' + 8 * E.sp + 'mm"></div>') +
      signHTML(S, 'assign');
  }

  function claimHTML(S, E) {
    const d = S.claim;
    const body = sanitize(d.body);
    const dv = (d.claimDateText || '').trim() || shortDate(d.claimDate);
    const line = (lbl, val) => '<div class="cl-line"><b>' + esc(lbl) + '</b> ' + (val ? '<span>' + esc(val) + '</span>' : '<span class="cl-dots">' + ph('............................') + '</span>') + '</div>';
    return headerHTML(S, E) + (d.showDateRow ? dateRowHTML(d) : '<div style="height:' + 6 * E.sp + 'mm"></div>') +
      '<div class="doc-title">' + esc(d.title) + '</div>' +
      '<div class="body cl-body">' + (body || ph('نص المطالبة…')) + '</div>' +
      (d.showClosing ? '<div class="cl-closing">' + esc(d.closing) + '</div>' : '') +
      '<div class="cl-owner">' + line(d.lblOwner, d.owner) + line(d.lblSign, '') + (d.showClaimDate ? line(d.lblDate, dv) : '') + '</div>' +
      signHTML(S, 'claim');
  }

  function recipStyle(S, type) {
    if (type !== 'thanks') return '';
    const y = +S.thanks.toY || 0, x = +S.thanks.toX || 0;
    if (!y && !x) return '';
    return ' style="' + (y > 0 ? 'padding-top:' + y + 'mm;' : y < 0 ? 'margin-top:' + y + 'mm;' : '') + (x ? 'position:relative;left:' + x + 'mm;' : '') + '"';
  }
  function letterHTML(S, type, E) {
    const d = S[type];
    const dl = dateLines(d);
    const numCell = '<div class="dr-r"><b>الرقم:</b> ' + (d.number ? esc(d.number) : ph('&nbsp;')) + '</div>';
    let dateCells;
    if (d.showHijri) {
      dateCells = '<div class="dr-c"><b>التاريخ:</b> <bdi>' + esc(dl.hijri) + '</bdi></div>' +
        '<div class="dr-l"><b>الموافق:</b> <bdi>' + esc(dl.greg) + '</bdi></div>';
    } else {
      dateCells = '<div class="dr-c"></div><div class="dr-l"><b>التاريخ:</b> <bdi>' + esc(dl.greg) + '</bdi></div>';
    }
    const list = [];
    if (type !== 'thanks') [d.to0, d.to1, d.to2].forEach((on, i) => { if (on) list.push('<div>' + RECIPIENTS[i] + '</div>'); });
    if (type === 'internal') {
      if (d.toDir) list.unshift('<div>' + DIRECTOR + ' الأكرم</div>' + (d.dirName.trim() ? '<div class="to-name">' + esc(d.dirName.trim()) + '</div>' : ''));
      if (d.toCustom.trim()) d.toCustom.split('\n').filter(x => x.trim()).forEach(x => list.push('<div>' + esc(x.trim()) + '</div>'));
    }
    if (type === 'thanks') {
      toLines(S, curVariant).forEach(x => list.push('<div>' + esc(x) + '</div>'));
      if (!list.length) list.push('<div>' + ph('الجهة الموجه إليها الكتاب…') + '</div>');
    }
    const body = sanitize(d.body);
    return headerHTML(S, E) +
      '<div class="date-row">' + numCell + dateCells + '</div>' +
      (list.length ? '<div class="recipients"' + recipStyle(S, type) + '>' + list.join('') + '</div>' : '') +
      '<div class="subject"><u>الموضوع:</u> ' + (d.subject ? '( ' + esc(d.subject) + ' )' : ph('(&nbsp;&nbsp;&nbsp;&nbsp;)')) + '</div>' +
      '<div class="body">' + (body || ph('نص الكتاب…')) + '</div>' +
      '<div class="closing">' + CLOSING + '</div>' +
      signHTML(S, type);
  }

  function circularHTML(S, E) {
    const d = S.circular;
    const body = sanitize(d.body);
    let line = '';
    if (d.lineShow) {
      line = '<div class="c-line-wrap" style="margin:' + d.lineBefore + 'mm 0 ' + d.lineAfter + 'mm"><div class="c-line" style="border-top:' + d.lineWidth + 'pt ' + d.lineStyle + ' ' + d.lineColor + ';width:' + d.linePct + '%"></div></div>';
    }
    let dateBlock = '';
    if (d.showDate || d.showNumUnder) {
      const dl = dateLines(d);
      dateBlock = '<div class="c-date">' +
        (d.showDate ? (d.showHijri && dl.hijri ? '<div><b>التاريخ:</b> <bdi>' + esc(dl.hijri) + 'هـ</bdi></div><div><b>الموافق:</b> <bdi>' + esc(dl.greg) + 'م</bdi></div>'
          : '<div><b>التاريخ:</b> <bdi>' + esc(dl.greg) + 'م</bdi></div>') : '') +
        (d.showNumUnder ? '<div><b>الرقم:</b> ' + (d.number ? esc(d.number) : ph('&nbsp;&nbsp;&nbsp;&nbsp;')) + '</div>' : '') + '</div>';
    }
    return headerHTML(S, E) + line + dateBlock +
      '<div class="c-title" style="font-size:' + d.titleSize + 'pt;text-align:' + d.titleAlign + '">تعميم هيئة عامة</div>' +
      (d.showTitleNum ? '<div class="c-num" style="text-align:' + d.titleAlign + '">تعميم رقم ' + (d.number ? esc(d.number) : '') + '</div>' : '<div style="height:4mm"></div>') +
      '<div class="subject"><u>الموضوع:</u> ' + (d.subject ? esc(d.subject) : '') + '</div>' +
      '<div class="body" style="text-align:' + d.bodyAlign + '">' + (body || ph('نص التعميم…')) + '</div>' +
      '<div class="closing">' + CLOSING + '</div>' +
      signHTML(S, 'circular');
  }

  /* ---------- الذيل ---------- */
  /** عناصر السطر الواحد في الذيل (مشتركة مع Word) */
  function footItems(S) {
    const f = S.footer, ic = f.icons, out = [];
    const add = (icon, on, label, value, ltr) => { value = (value || '').trim(); if (value) out.push({ icon: on ? icon : '', label, value, ltr }); };
    add('phone', ic.phone, 'هاتف', f.phone, true);
    add('whatsapp', ic.whatsapp, 'واتساب', f.whatsapp, true);
    add('mail', ic.email, 'البريد الإلكتروني', f.email, true);
    add('facebook', ic.facebook, 'فيسبوك', f.facebook, true);
    add('globe', ic.web, 'الموقع الإلكتروني', f.website, true);
    add('pin', ic.location, 'العنوان', f.address, false);
    add('', false, 'فاكس', f.fax, true);
    f.social.split('\n').map(x => x.trim()).filter(Boolean).forEach(x => out.push({ icon: '', label: '', value: x, ltr: true }));
    return out;
  }
  function footExtra(S) { return S.footer.extra.split('\n').map(s => s.trim()).filter(Boolean); }
  function footHasContent(S) { const f = S.footer; return !!(f.name.trim() || footItems(S).length || footExtra(S).length); }

  function footerHTML(S, E) {
    const f = S.footer;
    if (!f.show || !footHasContent(S)) return '';
    const jc = { center: 'center', right: 'flex-start', left: 'flex-end' }[f.align] || 'center';
    const size = E.footSize;
    let h = '<div class="foot' + (f.line ? ' foot-line' : '') + '" style="font-family:' + fontCss(f.font) + ';font-size:' + size + 'pt;color:' + f.color + ';font-weight:' + (f.bold ? 700 : 400) + ';' + (f.line ? 'border-top:' + f.lineWidth + 'pt solid #000;' : '') + '">';
    if (f.name.trim()) h += '<div class="foot-name" style="font-size:' + f.nameSize + 'pt;text-align:' + f.nameAlign + ';font-weight:' + (f.nameBold ? 700 : 400) + '">' + esc(f.name.trim()) + '</div>';
    const items = footItems(S);
    if (items.length) {
      h += '<div class="foot-info" style="justify-content:' + jc + ';gap:0 ' + f.gap + 'mm">';
      items.forEach(it => {
        const icon = it.icon ? iconSVG(it.icon, f.color, f.iconSize, 'pt') : '';
        const val = it.ltr ? '<bdi dir="ltr">' + esc(it.value) + '</bdi>' : esc(it.value);
        const txt = (it.icon || !it.label ? '' : esc(it.label) + ': ') + val;
        const pos = it.icon ? f.iconPos : 'before';
        h += '<span class="fi fi-' + pos + '">' + (pos === 'after' ? '<span>' + txt + '</span>' + icon : icon + '<span>' + txt + '</span>') + '</span>';
      });
      h += '</div>';
    }
    footExtra(S).forEach(x => { h += '<div class="foot-extra" style="text-align:' + f.align + '">' + esc(x) + '</div>'; });
    return h + '</div>';
  }

  function fontCss(name) { return "'" + name + "',Tahoma,Arial,sans-serif"; }

  function styleOf(S, type, E) {
    return 'font-family:' + fontCss(S.font) + ';--body-font:' + fontCss(E.bodyFont) + ';--body-size:' + E.bodySize + 'pt;--lh:' + E.lineHeight + ';--sp:' + E.sp;
  }
  function mainHTML(S, type, E) { return type === 'circular' ? circularHTML(S, E) : type === 'assign' ? assignHTML(S, E) : type === 'claim' ? claimHTML(S, E) : letterHTML(S, type, E); }

  /* ---------- القياس التلقائي: صفحة واحدة + ضبط الذيل في سطر ---------- */
  let cache = { key: '', val: null };
  function sigKey(S, type) {
    const P = S.people, mark = v => (typeof v === 'string' && v.length > 200 ? v.length : v);
    return JSON.stringify([type, S[type], S.font, S.bodySize, S.lineHeight, S.logoSize, S.onePage, S.clubName, S.footer,
      S.logos.ministry.show, mark(S.logos.ministry.data), S.logos.club.show, mark(S.logos.club.data),
      [P.signSize, P.stampSize, P.stampPos, P.mode, P.gapName, P.gapImg], ['head', 'manager'].map(k => [P[k].name, P[k].showSign, P[k].showStamp, mark(P[k].signImg), mark(P[k].stampImg), P[k].signAR, P[k].stampAR, P[k].pos]),
      document.fonts && document.fonts.status]);
  }
  function stageEl() {
    let el = document.getElementById('fitStage');
    if (!el) {
      el = document.createElement('div'); el.id = 'fitStage'; el.setAttribute('aria-hidden', 'true');
      el.style.cssText = 'position:fixed;left:-20000px;top:0;visibility:hidden;pointer-events:none';
      document.body.appendChild(el);
    }
    return el;
  }
  function effective(S, type) {
    const key = sigKey(S, type);
    if (cache.key === key) return cache.val;
    const d = S[type], circ = type === 'circular';
    const baseSize = circ && d.bodySize > 0 ? d.bodySize : S.bodySize;
    const E = { bodySize: baseSize, lineHeight: S.lineHeight, sp: 1, logoSize: S.logoSize, footSize: S.footer.size, bodyFont: circ && d.bodyFont ? d.bodyFont : S.font, fits: true, scale: 1 };
    const stage = stageEl();
    // القياس بدون فك ترميز الصور الكبيرة (تُستبدل بصورة 1×1 وتبقى أبعادها ثابتة) لتفادي البطء والتجمّد
    const PIX = 'data:image/gif;base64,R0lGODlhAQABAAAAACH5BAEKAAEALAAAAAABAAEAAAICTAEAOw==';
    const t0 = Date.now(), budget = () => Date.now() - t0 > 400;
    const measure = () => {
      const html = '<div class="sheet measure" dir="rtl" style="' + styleOf(S, type, E) + '"><div class="sheet-main">' + mainHTML(S, type, E) + '</div><div class="sheet-foot">' + footerHTML(S, E) + '</div></div>';
      stage.innerHTML = html.replace(/src="data:image\/[^"]*"/g, 'src="' + PIX + '"');
      return stage.firstChild;
    };
    try {
      if (S.footer.show) {
        for (let i = 0; i < 4 && !budget(); i++) {
          const row = measure().querySelector('.foot-info');
          if (!row || row.scrollWidth <= row.clientWidth + 1) break;
          const next = Math.max(6, Math.floor(E.footSize * (row.clientWidth / row.scrollWidth) * 4) / 4 - 0.25);
          if (next >= E.footSize) break;
          E.footSize = next;
        }
      }
      if (S.onePage) {
        let s = 1;
        for (let n = 0; n < 16; n++) {
          if (budget()) break;
          const h = measure().offsetHeight;
          if (h <= PAGE_H - 1) { E.fits = true; break; }
          s -= 0.04;
          if (s < 0.6) { E.fits = false; break; }
          E.scale = s; E.sp = s;
          E.bodySize = Math.max(9.5, Math.round(baseSize * s * 4) / 4);
          E.lineHeight = Math.max(1.15, Math.round(S.lineHeight * (0.55 + 0.45 * s) * 100) / 100);
          E.logoSize = Math.max(14, Math.round(S.logoSize * (0.6 + 0.4 * s)));
        }
      }
    } finally { stage.innerHTML = ''; }
    cache = { key, val: E };
    return E;
  }

  /** نسخة من الحالة بالقيم الفعلية (للتصدير إلى Word) */
  function withEffective(S, type) {
    const E = effective(S, type);
    return Object.assign({}, S, { bodySize: E.bodySize, lineHeight: E.lineHeight, logoSize: E.logoSize, bodyFont: E.bodyFont, footer: Object.assign({}, S.footer, { size: E.footSize }) });
  }

  /** يعيد HTML القسمين: الرأس+المحتوى، والذيل */
  function renderDoc(S, type, vi) {
    const E = effective(S, type);
    curVariant = variants(S, type)[vi || 0] || null;
    try {
      return { style: styleOf(S, type, E), main: mainHTML(S, type, E), foot: footerHTML(S, E), eff: E };
    } finally { curVariant = null; }
  }

  function attachPagesHTML(list) {
    return list.map(p => '<div class="att-page"><div class="att-cap">' + esc(p.cap) + '</div><img src="' + p.url + '" alt=""></div>');
  }

  g.App = Object.assign(App, { SG, signGeom, signGeomSafe, signIssues, signers, boxLeft, shortDate, assignDateStr, DATE_DOC, variants, toLines, attachPagesHTML, esc, sanitize, renderDoc, effective, withEffective, footItems, footExtra, footHasContent, iconSVG, fontCss, isLetter, TITLES, CONST: { MINISTRY, DEPT, CLUBS, CLUB_PREFIX, CLOSING, RECIPIENTS, DIRECTOR } });
})(window);
