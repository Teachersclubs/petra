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
        if (n.tagName === 'IMG') {
          const src = n.getAttribute('src') || '';
          if (!/^data:image\/(png|jpe?g|webp|gif);base64,[A-Za-z0-9+\/=]+$/.test(src)) return;
          const w = Math.max(8, Math.min(170, parseFloat(n.getAttribute('data-w')) || 60));
          const ar = Math.max(0.05, Math.min(20, parseFloat(n.getAttribute('data-ar')) || 0.75));
          out += '<img src="' + src + '" data-w="' + (Math.round(w * 10) / 10) + '" data-ar="' + (Math.round(ar * 1000) / 1000) + '" style="width:' + (Math.round(w * 10) / 10) + 'mm;max-width:100%;height:auto">';
          return;
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
  const DATE_DOC = t => isLetter(t) || t === 'assign' || t === 'claim' || t === 'delegation' || t === 'members';

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
    const zoneH = items.length ? maxH + 2 * SG.PADZ : 0;
    const total = items.reduce((t, it) => t + it.w, 0) + gi * Math.max(0, items.length - 1);
    // الصفة والاسم مجموعة واحدة متلاصقة تتحرك معًا؛ الختم والتوقيع حول المجموعة حسب الوضع
    let titleY, zoneTop;
    if (o.mode === 'above') { zoneTop = 0; titleY = zoneH ? zoneH + gn : 0; }
    else if (o.mode === 'overlap') { titleY = 0; zoneTop = SG.TITLE_H + 1; }
    else { titleY = 0; zoneTop = SG.TITLE_H + SG.NAME_H + gn; }
    const nameY = titleY + SG.TITLE_H;
    let x = (SG.W - total) / 2;
    const rects = {};
    items.forEach(it => {
      const q = pos[it.k];
      rects[it.k] = { x: x + q.x, y: zoneTop + (zoneH - it.h) / 2 + q.y, w: it.w, h: it.h };
      x += it.w + gi;
    });
    const nw = nameWidth(S, p), tw = SG.W * 0.7;
    rects.name = { x: (SG.W - nw) / 2 + pos.name.x, y: nameY + pos.name.y, w: nw, h: SG.NAME_H, text: !!(p.name || '').trim() };
    rects.title = { x: (SG.W - tw) / 2 + pos.name.x, y: titleY + pos.name.y, w: tw, h: SG.TITLE_H };
    return { H: heightOf(rects), nameY0: nameY, titleY0: titleY, rects, zoneH };
  }
  function heightOf(rects) {
    let m = SG.TITLE_H + SG.NAME_H;
    ['sign', 'stamp', 'name', 'title'].forEach(k => { const r = rects[k]; if (r) m = Math.max(m, r.y + r.h); });
    return m + SG.PADB;
  }

  /** يحصر العناصر داخل الصفحة حتى لا تختفي خارج الحدود (للعرض والتصدير) */
  function signGeomSafe(S, key, who, fin) {
    const g = signGeom(S, key), bl = boxLeft(key, who, fin), R = g.rects;
    const lim = (lo, hi, v, w) => Math.max(lo, Math.min(hi - w, v));
    ['sign', 'stamp'].forEach(k => {
      const r = R[k]; if (!r) return;
      r.x = lim(4 - bl, SG.PAGE_W - 4 - bl, r.x, r.w); r.y = Math.max(0, r.y);
    });
    // مجموعة الصفة + الاسم تُحصر معًا
    const t = R.title, n = R.name;
    let mn = Math.min(t.x, n.x), mx = Math.max(t.x + t.w, n.x + n.w), dx = 0, dy = 0;
    if (mn + bl < 4) dx = 4 - bl - mn; else if (mx + bl > SG.PAGE_W - 4) dx = SG.PAGE_W - 4 - bl - mx;
    if (t.y < 0) dy = -t.y;
    t.x += dx; n.x += dx; t.y += dy; n.y += dy;
    g.H = heightOf(R);
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
        if (rr.x < 4 || rr.x + rr.w > SG.PAGE_W - 4) out.push('العنصر يخرج عن حدود الصفحة');
        if (r.y < -0.01) out.push('العنصر يخرج عن حدود منطقة التوقيع');
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
    const T = R.title, N = R.name;
    let h = '<div class="sign-box sg" data-key="' + key + '" style="width:' + SG.W + 'mm;height:' + f1(g.H) + 'mm">';
    h += '<div class="sg-title sg-mv" data-key="' + key + '" data-el="name" style="top:' + f1(T.y) + 'mm;height:' + SG.TITLE_H + 'mm;line-height:' + SG.TITLE_H + 'mm;left:' + f1(T.x + T.w / 2 - SG.W / 2) + 'mm;right:' + f1(-(T.x + T.w / 2 - SG.W / 2)) + 'mm">' + TITLES[key] + '</div>';
    h += '<div class="sg-name sg-mv" data-key="' + key + '" data-el="name" style="top:' + f1(N.y) + 'mm;height:' + SG.NAME_H + 'mm;line-height:' + SG.NAME_H + 'mm;left:' + f1(N.x + N.w / 2 - SG.W / 2) + 'mm;right:' + f1(-(N.x + N.w / 2 - SG.W / 2)) + 'mm">' + (p.name ? esc(p.name) : '') + '</div>';
    ['sign', 'stamp'].forEach(k => {
      const r = R[k]; if (!r) return;
      h += '<img class="' + k + '-img sg-img sg-mv" data-key="' + key + '" data-el="' + k + '" src="' + (k === 'sign' ? p.signImg : p.stampImg) + '" alt="" style="left:' + f1(r.x) + 'mm;top:' + f1(r.y) + 'mm;width:' + f1(r.w) + 'mm;height:' + f1(r.h) + 'mm">';
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

  /** عنوان التعميم: «السادة أعضاء الهيئة العامة لنادي …» أو الجملة التي يكتبها المستخدم */
  function circTitle(S) {
    const c = (S.circular.titleCustom || '').trim();
    if (c) return c;
    const n = (S.clubName || '').replace(/^\s*(ال)?نادي\s+/, '').trim();
    return 'السادة أعضاء الهيئة العامة' + (n ? ' لنادي ' + n : ' للنادي');
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
      '<div class="c-title" style="font-size:' + d.titleSize + 'pt;text-align:' + d.titleAlign + '">' + esc(circTitle(S)) + '</div>' +
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
      h += '<div class="foot-info" style="justify-content:' + jc + ';gap:0 ' + (E.footGap == null ? f.gap : E.footGap) + 'mm">';
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
  /* ---------- الوفد المشارك ---------- */
  const DG_KINDS = { tournament: ['بطولة', 'اسم البطولة'], qualifiers: ['تصفيات', 'اسم التصفيات'], game: ['لعبة', 'اسم اللعبة'] };
  const DG_COLS = {
    admins: [['no', 'م', 8, 8], ['name', 'الاسم', 40, 60], ['ministryNo', 'الرقم الوزاري', 17, 30], ['nationalNo', 'الرقم الوطني', 25, 30], ['role', 'الصفة', 18, 40]],
    players: [['no', 'م', 8, 8], ['name', 'الاسم', 40, 60], ['ministryNo', 'الرقم الوزاري', 17, 30], ['memberNo', 'رقم العضوية', 17, 30], ['nationalNo', 'الرقم الوطني', 25, 30], ['joinDate', 'تاريخ الانتساب', 21, 30], ['duration', 'مدة الانتساب', 26, 44], ['role', 'الصفة', 18, 40]]
  };
  /** أعمدة الجدول الحالية بالترتيب (أساسية + مخصصة) */
  function dgCols(S, kind) {
    const d = S.delegation, ids = d[kind === 'admins' ? 'colsAdmins' : 'colsPlayers'], cu = d[kind === 'admins' ? 'customAdmins' : 'customPlayers'];
    return ids.map(id => { const st = DG_COLS[kind].find(c => c[0] === id), cc = cu.find(c => c.id === id); return st || (cc ? [id, cc.label, 17, 30] : null); }).filter(Boolean);
  }
  const ar = (n, one, two, few, many) => n === 1 ? one : n === 2 ? two : (n >= 3 && n <= 10) ? n + ' ' + few : n + ' ' + many;
  /** مدة الانتساب من تاريخ الانتساب حتى تاريخ الكتاب: سنوات وأشهر وأيام */
  function durationParts(fromISO, toISO) {
    const a = /^(\d{4})-(\d{2})-(\d{2})$/.exec(fromISO || ''), b = /^(\d{4})-(\d{2})-(\d{2})$/.exec(toISO || '');
    if (!a || !b) return null;
    const sy = +a[1], sm = +a[2], sd = +a[3], ey = +b[1], em = +b[2], ed = +b[3];
    if (Date.UTC(ey, em - 1, ed) < Date.UTC(sy, sm - 1, sd)) return null;
    let y = ey - sy, m = em - sm, d = ed - sd;
    if (d < 0) { m--; d += new Date(ey, em - 1, 0).getDate(); }
    if (m < 0) { y--; m += 12; }
    return { y, m, d };
  }
  function durationText(fromISO, toISO, showDays) {
    const p = durationParts(fromISO, toISO);
    if (!p) return '';
    const out = [];
    if (p.y) out.push(ar(p.y, 'سنة', 'سنتان', 'سنوات', 'سنة'));
    if (p.m) out.push(ar(p.m, 'شهر', 'شهران', 'أشهر', 'شهرًا'));
    if (p.d && (showDays !== false || !out.length)) out.push(ar(p.d, 'يوم', 'يومان', 'أيام', 'يومًا'));
    return out.length ? out.join(' و ') : 'أقل من يوم';
  }
  function dgClubPhrase(S) {
    const n = (S.clubName || '').replace(/^\s*(ال)?نادي\s+/, '').trim();
    if (!n) return 'للنادي';
    return 'لنادي ' + (/^(معلمي|العاملين)/.test(n) ? '' : 'معلمي ') + n;
  }
  function dgIntro(S) {
    const d = S.delegation, k = DG_KINDS[d.kind] || DG_KINDS.qualifiers;
    const club = dgClubPhrase(S).replace(/^ل(?=نادي)/, '').replace(/^للنادي$/, 'النادي');
    return 'نرفق لكم طيّاً كشفاً بأسماء اللاعبين وأعضاء الوفد المشارك عن ' + club +
      ' في ' + k[0] + (d.kindName.trim() ? ' ' + d.kindName.trim() : '') +
      (d.eventName.trim() ? '، ضمن فعاليات ' + d.eventName.trim() : '') + (d.venue.trim() ? '، والمقامة في ' + d.venue.trim() : '') + '، وذلك للاطلاع واتخاذ ما يلزم.';
  }
  /** قيمة خلية: يعيد النص المعروض لعمود في صف */
  function dgCell(S, kind, r, i, col) {
    const d = S.delegation;
    if (col === 'no') return String(i + 1);
    if (col === 'joinDate') return shortDate(r.joinDate);
    if (col === 'duration') return durationText(r.joinDate, d.date, d.showDays);
    if (col === 'role' && /^(لاعب|لاعبة)$/.test(r.role || '') && (r.game || '').trim()) return r.role + ' ' + r.game.trim();
    if (!/^(name|ministryNo|memberNo|nationalNo|role)$/.test(col)) return (r.x && r.x[col]) || '';
    return r[col] || '';
  }
  /** عرض الأعمدة بالمليمتر (المجموع 176): لا تُقسَّم كلمة ولا رقم؛ ويأخذ الاسم المتبقي */
  /** يختار حجم خط الجدول وعرض الأعمدة معًا: يصغّر الخط (حتى 55%) بدل أن تنكسر الكلمات حرفًا حرفًا */
  function dgFit(S, kind) {
    const cols = dgCols(S, kind), rows = S.delegation[kind], TOT = 176, PAD = 4.5, NO = 8, MINNAME = 24;
    const ni = cols.findIndex(c => c[0] === 'name'), isNo = i => cols[i][0] === 'no';
    // strict: يجب أن يظهر الاسم كاملًا في سطر واحد (كتابة أفقية)، والمدة/الصفة في سطرين كحد أقصى
    const needAt = (f, strict) => cols.map(c => {
      if (c[0] === 'no') return NO;
      const HC = 2.15 * f, CC = 2.25 * f;
      let w = Math.max.apply(null, c[1].split(' ').map(x => x.length)) * HC + PAD;
      rows.forEach((r, i) => {
        const t = dgCell(S, kind, r, i, c[0]); if (!t) return;
        if (c[0] === 'name') w = Math.max(w, (strict ? t.length : Math.min(t.length, 30)) * CC + PAD);
        else if (c[0] === 'duration') w = Math.max(w, Math.min(Math.ceil(t.length / 2), 14) * CC + PAD);
        else if (!/^(ministryNo|memberNo|nationalNo|joinDate)$/.test(c[0])) w = Math.max(w, Math.max.apply(null, t.split(' ').map(x => x.length)) * CC + PAD);
        else w = Math.max(w, t.length * CC + PAD);
      });
      return w;
    });
    const sum = n => n.reduce((a, v) => a + v, 0), must = n => sum(n) - (ni >= 0 ? n[ni] : 0) + (ni >= 0 ? Math.min(n[ni], MINNAME) : 0);
    let f = 1, need, tight = false, nw = false;
    for (f = 1; f >= 0.65; f = Math.round((f - 0.05) * 100) / 100) { need = needAt(f, true); if (sum(need) <= TOT) { nw = true; break; } }
    if (!nw) {
      for (f = 1; f >= 0.55; f = Math.round((f - 0.05) * 100) / 100) { need = needAt(f, false); if (must(need) <= TOT) break; }
      if (f < 0.55) { f = 0.55; need = needAt(f, false); tight = true; }
    }
    const flex = need.reduce((a, v, i) => a + (isNo(i) ? 0 : v), 0), noW = need.reduce((a, v, i) => a + (isNo(i) ? v : 0), 0);
    let w;
    if (noW + flex <= TOT) {
      w = need.map((v, i) => ni >= 0 ? (i === ni ? v + (TOT - noW - flex) : v) : isNo(i) ? v : v + (TOT - noW - flex) * v / flex);
    } else {
      const nameW = ni >= 0 ? Math.min(need[ni], MINNAME) : 0, others = flex - (ni >= 0 ? need[ni] : 0);
      const k = others > 0 ? (TOT - noW - nameW) / others : 1;
      w = need.map((v, i) => isNo(i) ? v : i === ni ? nameW : v * Math.max(k, 0.3));
      const t = w.reduce((a, v) => a + v, 0), g = (TOT - noW) / (t - noW);
      w = w.map((v, i) => isNo(i) ? v : v * g);
    }
    w = w.map(v => Math.round(v * 10) / 10);
    const fix = ni >= 0 ? ni : w.findIndex((v, i) => !isNo(i)), fi = fix < 0 ? 0 : fix;
    w[fi] = Math.round((w[fi] + TOT - w.reduce((a, v) => a + v, 0)) * 10) / 10;
    return { w, f, tight, nw };
  }
  const dgWidths = (S, kind) => dgFit(S, kind).w;
  function dgTableHTML(S, kind, title) {
    const cols = dgCols(S, kind), rows = S.delegation[kind], fit = dgFit(S, kind), w = fit.w;
    const th = cols.map(c => '<th>' + esc(c[1]) + '</th>').join('');
    const body = rows.map((r, i) => '<tr>' + cols.map(c => {
      const v = dgCell(S, kind, r, i, c[0]);
      return '<td' + (c[0] === 'name' ? ' class="dg-name"' : /^(joinDate|ministryNo|memberNo|nationalNo)$/.test(c[0]) ? ' class="dg-num"' : '') + '>' + (v ? (/^(joinDate|ministryNo|memberNo|nationalNo)$/.test(c[0]) ? '<bdi>' + esc(v) + '</bdi>' : esc(v)) : '') + '</td>';
    }).join('') + '</tr>').join('');
    return '<table class="dg' + (fit.tight ? ' dg-tight' : '') + (fit.nw ? ' dg-nw' : '') + '" dir="rtl" style="--dgf:' + fit.f + '"><colgroup>' + w.map(x => '<col style="width:' + x + 'mm">').join('') + '</colgroup>' +
      '<thead>' + (title == null ? '' : '<tr class="dg-t"><th colspan="' + cols.length + '">' + (title.trim() ? esc(title) : '&nbsp;') + '</th></tr>') + '<tr class="dg-h">' + th + '</tr></thead><tbody>' + body + '</tbody></table>';
  }
  function delegationHTML(S, E) {
    const d = S.delegation;
    const list = [];
    [d.to0, d.to1, d.to2].forEach((on, i) => { if (on) list.push('<div>' + RECIPIENTS[i] + '</div>'); });
    d.toCustom.split('\n').filter(x => x.trim()).forEach(x => list.push('<div>' + esc(x.trim()) + '</div>'));
    const body = sanitize(d.body);
    const intro = d.showIntro ? (d.introText.trim() ? esc(d.introText.trim()).replace(/\n/g, '<br>') : esc(dgIntro(S))) : '';
    return headerHTML(S, E) + dateRowHTML(d) +
      (list.length ? '<div class="recipients">' + list.join('') + '</div>' : '') +
      '<div class="subject"><u>الموضوع:</u> ' + (d.subject ? '( ' + esc(d.subject) + ' )' : ph('(&nbsp;&nbsp;&nbsp;&nbsp;)')) + '</div>' +
      (intro ? '<div class="body dg-intro">' + intro + '</div>' : '') +
      (body ? '<div class="body dg-body">' + body + '</div>' : '') +
      (d.showAdmins ? dgTableHTML(S, 'admins', d.titleAdmin) : '') +
      (d.showPlayers ? dgTableHTML(S, 'players', d.titlePlayers) : '') +
      (d.showClosing ? '<div class="closing">' + CLOSING + '</div>' : '<div style="height:' + 4 * E.sp + 'mm"></div>') +
      signHTML(S, 'delegation');
  }

  /* ---------- تقرير المنتسبين الجدد (يعيد استخدام جدول الوفد) ---------- */
  const MB_COLS = ['no', 'name', 'ministryNo', 'nationalNo'];
  function memView(S) {
    const m = S.members;
    return Object.assign({}, S, { delegation: { admins: m.rows, colsAdmins: MB_COLS, customAdmins: [], players: [], colsPlayers: ['no', 'name'], customPlayers: [], date: m.date, showDays: true } });
  }
  function mbIntro(S) {
    const m = S.members;
    return 'يرجى التكرم بالإيعاز إلى من يلزم لاتخاذ الإجراءات اللازمة باقتطاع مبلغ ' + (m.amount.trim() || 'دينار واحد') + ' شهرياً من رواتب المنتسبين الجدد ' + dgClubPhrase(S) + '، الواردة أسماؤهم في الكشف المرفق أدناه، وذلك اعتباراً من تاريخه، وفقاً للتعليمات والإجراءات المعتمدة.';
  }
  function membersHTML(S, E) {
    const d = S.members, list = [];
    [d.to0, d.to1, d.to2].forEach((on, i) => { if (on) list.push('<div>' + RECIPIENTS[i] + '</div>'); });
    d.toCustom.split('\n').filter(x => x.trim()).forEach(x => list.push('<div>' + esc(x.trim()) + '</div>'));
    const body = sanitize(d.body);
    const intro = d.showIntro ? (d.introText.trim() ? esc(d.introText.trim()).replace(/\n/g, '<br>') : esc(mbIntro(S))) : '';
    return headerHTML(S, E) + dateRowHTML(d) +
      (list.length ? '<div class="recipients">' + list.join('') + '</div>' : '') +
      '<div class="subject"><u>الموضوع:</u> ' + (d.subject ? '( ' + esc(d.subject) + ' )' : ph('(&nbsp;&nbsp;&nbsp;&nbsp;)')) + '</div>' +
      (intro ? '<div class="body dg-intro">' + intro + '</div>' : '') +
      (body ? '<div class="body dg-body">' + body + '</div>' : '') +
      dgTableHTML(memView(S), 'admins', d.tableTitle.trim() ? d.tableTitle : null) +
      (d.showClosing ? '<div class="closing">' + CLOSING + '</div>' : '<div style="height:' + 4 * E.sp + 'mm"></div>') +
      signHTML(S, 'members');
  }

  function mainHTML(S, type, E) { return type === 'members' ? membersHTML(S, E) : type === 'delegation' ? delegationHTML(S, E) : type === 'circular' ? circularHTML(S, E) : type === 'assign' ? assignHTML(S, E) : type === 'claim' ? claimHTML(S, E) : letterHTML(S, type, E); }

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
        // قياس الذيل وحده (سريع): يصغَّر الخط ثم المسافات حتى يبقى السطر واحدًا دون تجاوز عرض الصفحة
        const fm = () => {
          stage.innerHTML = '<div class="sheet measure" dir="rtl" style="' + styleOf(S, type, E) + '"><div class="sheet-main"></div><div class="sheet-foot">' + footerHTML(S, E).replace(/src="data:image\/[^"]*"/g, 'src="' + PIX + '"') + '</div></div>';
          return stage.firstChild.querySelector('.foot-info');
        };
        for (let i = 0; i < 16; i++) {
          const row = fm();
          if (!row || row.scrollWidth <= row.clientWidth + 1) break;
          const g = E.footGap == null ? S.footer.gap : E.footGap;
          if (g > 2) E.footGap = Math.max(2, g * 0.6);
          else if (E.footSize > 4.5) {
            const next = Math.max(4.5, Math.floor(E.footSize * (row.clientWidth / row.scrollWidth) * 4) / 4 - 0.25);
            E.footSize = next < E.footSize ? next : Math.max(4.5, E.footSize - 0.25);
          } else if (g > 0.5) {
            E.footGap = Math.max(0.5, g / 2);
          } else break;
        }
      }
      if (S.onePage || ((type === 'delegation' || type === 'members') && d.fitOne)) {
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
    return Object.assign({}, S, { bodySize: E.bodySize, lineHeight: E.lineHeight, logoSize: E.logoSize, bodyFont: E.bodyFont, footer: Object.assign({}, S.footer, { size: E.footSize, gap: E.footGap == null ? S.footer.gap : E.footGap }) });
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

  g.App = Object.assign(App, { SG, signGeom, signGeomSafe, signIssues, signers, boxLeft, shortDate, assignDateStr, DATE_DOC, variants, toLines, circTitle, attachPagesHTML, esc, sanitize, DG_KINDS, DG_COLS, dgCols, dgFit, durationParts, durationText, dgIntro, mbIntro, memView, dgCell, dgWidths, renderDoc, effective, withEffective, footItems, footExtra, footHasContent, iconSVG, fontCss, isLetter, TITLES, CONST: { MINISTRY, DEPT, CLUBS, CLUB_PREFIX, CLOSING, RECIPIENTS, DIRECTOR } });
})(window);
