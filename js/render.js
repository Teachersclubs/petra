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

  /* ---------- التوقيع والختم ---------- */
  function signImgs(S, p) {
    const o = S.people;
    const sg = p.showSign && p.signImg ? '<img class="sign-img" src="' + p.signImg + '" alt="" style="height:' + o.signSize + 'mm">' : '';
    const st = p.showStamp && p.stampImg ? '<img class="stamp-img" src="' + p.stampImg + '" alt="" style="height:' + o.stampSize + 'mm">' : '';
    if (!sg && !st) return '';
    // في الاتجاه RTL أول عنصر يظهر يمينًا
    return '<div class="sign-imgs">' + (o.stampPos === 'left' ? sg + st : st + sg) + '</div>';
  }
  function signBox(S, key) {
    const p = S.people[key], imgs = signImgs(S, p);
    return '<div class="sign-box"><div class="sign-title">' + TITLES[key] + '</div>' + imgs +
      '<div class="sign-name"' + (imgs ? ' style="margin-top:1mm"' : '') + '>' + esc(p.name) + '</div></div>';
  }

  /** رئيس الهيئة يمينًا ثم مدير النادي؛ لكل منهما خيار إظهار */
  function signHTML(S, d) {
    const who = [d.showHead ? 'head' : '', d.showManager ? 'manager' : ''].filter(Boolean);
    if (!who.length) return '';
    return '<div class="sign ' + (who.length === 2 ? 'sign-two' : 'sign-one') + '">' + who.map(k => signBox(S, k)).join('') + '</div>';
  }

  function dateLines(d) {
    const greg = App.formatGregorian(d.date);
    return d.showHijri ? { hijri: d.hijri, greg } : { hijri: '', greg };
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
    [d.to0, d.to1, d.to2].forEach((on, i) => { if (on) list.push('<div>' + RECIPIENTS[i] + '</div>'); });
    if (type === 'internal') {
      if (d.toDir) list.unshift('<div>' + DIRECTOR + ' الأكرم</div>' + (d.dirName.trim() ? '<div class="to-name">' + esc(d.dirName.trim()) + '</div>' : ''));
      if (d.toCustom.trim()) d.toCustom.split('\n').filter(x => x.trim()).forEach(x => list.push('<div>' + esc(x.trim()) + '</div>'));
    }
    const body = sanitize(d.body);
    return headerHTML(S, E) +
      '<div class="date-row">' + numCell + dateCells + '</div>' +
      (list.length ? '<div class="recipients">' + list.join('') + '</div>' : '') +
      '<div class="subject"><u>الموضوع:</u> ' + (d.subject ? '( ' + esc(d.subject) + ' )' : ph('(&nbsp;&nbsp;&nbsp;&nbsp;)')) + '</div>' +
      '<div class="body">' + (body || ph('نص الكتاب…')) + '</div>' +
      '<div class="closing">' + CLOSING + '</div>' +
      signHTML(S, d);
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
      signHTML(S, d);
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
  function mainHTML(S, type, E) { return type === 'circular' ? circularHTML(S, E) : letterHTML(S, type, E); }

  /* ---------- القياس التلقائي: صفحة واحدة + ضبط الذيل في سطر ---------- */
  let cache = { key: '', val: null };
  function sigKey(S, type) {
    const P = S.people, mark = v => (typeof v === 'string' && v.length > 200 ? v.length : v);
    return JSON.stringify([type, S[type], S.font, S.bodySize, S.lineHeight, S.logoSize, S.onePage, S.clubName, S.footer,
      S.logos.ministry.show, mark(S.logos.ministry.data), S.logos.club.show, mark(S.logos.club.data),
      [P.signSize, P.stampSize, P.stampPos], ['head', 'manager'].map(k => [P[k].name, P[k].showSign, P[k].showStamp, mark(P[k].signImg), mark(P[k].stampImg)]),
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
    const measure = () => {
      stage.innerHTML = '<div class="sheet measure" dir="rtl" style="' + styleOf(S, type, E) + '"><div class="sheet-main">' + mainHTML(S, type, E) + '</div><div class="sheet-foot">' + footerHTML(S, E) + '</div></div>';
      return stage.firstChild;
    };
    try {
      if (S.footer.show) {
        for (let i = 0; i < 4; i++) {
          const row = measure().querySelector('.foot-info');
          if (!row || row.scrollWidth <= row.clientWidth + 1) break;
          const next = Math.max(6, Math.floor(E.footSize * (row.clientWidth / row.scrollWidth) * 4) / 4 - 0.25);
          if (next >= E.footSize) break;
          E.footSize = next;
        }
      }
      if (S.onePage) {
        let s = 1;
        for (;;) {
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
  function renderDoc(S, type) {
    const E = effective(S, type);
    return { style: styleOf(S, type, E), main: mainHTML(S, type, E), foot: footerHTML(S, E), eff: E };
  }

  function attachPagesHTML(list) {
    return list.map(p => '<div class="att-page"><div class="att-cap">' + esc(p.cap) + '</div><img src="' + p.url + '" alt=""></div>');
  }

  g.App = Object.assign(App, { attachPagesHTML, esc, sanitize, renderDoc, effective, withEffective, footItems, footExtra, footHasContent, iconSVG, fontCss, isLetter, TITLES, CONST: { MINISTRY, DEPT, CLUBS, CLUB_PREFIX, CLOSING, RECIPIENTS, DIRECTOR } });
})(window);
