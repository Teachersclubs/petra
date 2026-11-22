/* شهادة شكر وتقدير: ورقة أفقية بإطار وزوايا ذهبية وشعار وعلامة مائية */
(function (g) {
  'use strict';
  const App = g.App || (g.App = {});
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* ---- الإطار: SVG بأبعاد 297×210 مم ---- */
  const corner = t => '<g transform="' + t + '"><polygon points="0,0 34,0 26,8 8,8" fill="#dfbd78"/><polygon points="0,0 8,8 8,26 0,34" fill="#c19a47"/>' +
    '<polygon points="8,8 26,8 22,11 11,11 11,22 8,26" fill="#a9822f" opacity=".55"/></g>';

  /* ---- القوالب العشرة: كل قالب = إطار SVG + ألوان افتراضية للنصوص ---- */
  const W = 297, H = 210;
  const svg = inner => '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 297 210" width="297" height="210">' + inner + '</svg>';
  const R = (x, y, w, h, a) => '<rect x="' + x + '" y="' + y + '" width="' + w + '" height="' + h + '" ' + a + '/>';
  const P = (pts, a) => '<polygon points="' + pts + '" ' + a + '/>';
  const frame = (x, c, w) => R(x, x, W - 2 * x, H - 2 * x, 'fill="none" stroke="' + c + '" stroke-width="' + w + '"');
  const four = f => f('') + f('translate(297,0) scale(-1,1)') + f('translate(0,210) scale(1,-1)') + f('translate(297,210) scale(-1,-1)');
  const grad = (id, a, b) => '<defs><linearGradient id="' + id + '" x1="0" y1="0" x2="1" y2="1"><stop offset="0" stop-color="' + a + '"/><stop offset=".5" stop-color="' + b + '"/><stop offset="1" stop-color="' + a + '"/></linearGradient></defs>';
  const diamonds = (c1, c2) => { let o = ''; for (let i = 0, x = 20; x < 280; x += 7, i++) { const c = i % 2 ? c1 : c2; o += P((x) + ',7 ' + (x + 2.6) + ',9.6 ' + x + ',12.2 ' + (x - 2.6) + ',9.6', 'fill="' + c + '"') + P(x + ',197.8 ' + (x + 2.6) + ',200.4 ' + x + ',203 ' + (x - 2.6) + ',200.4', 'fill="' + c + '"'); }
    for (let i = 0, y = 20; y < 195; y += 7, i++) { const c = i % 2 ? c1 : c2; o += P('9.6,' + (y - 2.6) + ' 12.2,' + y + ' 9.6,' + (y + 2.6) + ' 7,' + y, 'fill="' + c + '"') + P('287.4,' + (y - 2.6) + ' 290,' + y + ' 287.4,' + (y + 2.6) + ' 284.8,' + y, 'fill="' + c + '"'); } return o; };
  const TPLS = [
    { id: 'classic', name: 'كلاسيكي كحلي وذهبي', c: { club: '#1f2233', title: '#9a7425', text: '#1f2233' }, svg: svg(
      R(0, 0, W, H, 'fill="#3b3d52"') + R(11, 11, 275, 188, 'fill="#fff"') + corner('translate(11,11)') + corner('translate(286,11) scale(-1,1)') + corner('translate(11,199) scale(1,-1)') + corner('translate(286,199) scale(-1,-1)')) },
    { id: 'royal', name: 'ملكي ذهبي', c: { club: '#5b4210', title: '#9a7425', text: '#3a3426' }, svg: svg(
      R(0, 0, W, H, 'fill="#fbf6e9"') + frame(8, '#b8923a', 1.6) + frame(12, '#b8923a', .5) +
      four(t => '<g transform="' + t + '">' + R(6, 6, 14, 14, 'fill="#b8923a"') + P('13,8.5 17.5,13 13,17.5 8.5,13', 'fill="#fbf6e9"') + '</g>')) },
    { id: 'emerald', name: 'زمردي أنيق', c: { club: '#0f5132', title: '#9a7425', text: '#1d2b24' }, svg: svg(
      R(0, 0, W, H, 'fill="#0f5132"') + R(10, 10, 277, 190, 'fill="#fffdf6"') + frame(14, '#c9a24a', .8) +
      four(t => '<g transform="' + t + '"><circle cx="14" cy="14" r="4.2" fill="#c9a24a"/><circle cx="14" cy="14" r="2" fill="#0f5132"/></g>')) },
    { id: 'modern', name: 'عصري بزوايا مائلة', c: { club: '#1f2937', title: '#0e7490', text: '#1f2937' }, svg: svg(
      R(0, 0, W, H, 'fill="#fff"') + P('0,0 105,0 0,66', 'fill="#0e7490"') + P('0,0 66,0 0,40', 'fill="#f59e0b"') +
      P('297,210 252,210 297,180', 'fill="#0e7490"') + P('297,210 272,210 297,194', 'fill="#f59e0b"')) },
    { id: 'burgundy', name: 'عنابي فاخر', c: { club: '#6b1d2a', title: '#8a6a1f', text: '#2b2b2b' }, svg: svg(
      R(0, 0, W, H, 'fill="#6b1d2a"') + R(9, 9, 279, 192, 'fill="#fffaf5"') + frame(13, '#c9a24a', .7) +
      four(t => '<g transform="' + t + '">' + P('13,3 23,13 13,23 3,13', 'fill="#c9a24a"') + P('13,7 19,13 13,19 7,13', 'fill="#6b1d2a"') + '</g>')) },
    { id: 'waves', name: 'أمواج زرقاء', c: { club: '#1e3a8a', title: '#1d4ed8', text: '#1f2937' }, svg: svg(
      R(0, 0, W, H, 'fill="#fff"') +
      '<path d="M0 0H118C92 22 48 34 0 62Z" fill="#1d4ed8"/><path d="M0 0H84C62 16 30 26 0 42Z" fill="#60a5fa" opacity=".75"/>' +
      '<path d="M297 210H246C262 200 280 190 297 178Z" fill="#1d4ed8"/><path d="M297 210H272C282 204 290 200 297 194Z" fill="#60a5fa" opacity=".75"/>' +
      '<path d="M297 0H250C268 10 284 20 297 34Z" fill="#93c5fd" opacity=".7"/><path d="M0 210H47C29 200 13 190 0 176Z" fill="#93c5fd" opacity=".7"/>') },
    { id: 'minimal', name: 'بسيط أنيق', c: { club: '#111827', title: '#111827', text: '#374151' }, svg: svg(
      R(0, 0, W, H, 'fill="#fff"') + frame(10, '#111827', 1) + frame(14, '#111827', .3)) },
    { id: 'gold', name: 'ذهبي لامع', c: { club: '#4a3608', title: '#7a5a12', text: '#2f2a1c' }, svg: svg(
      grad('gg', '#e9c46a', '#a9822f') + R(0, 0, W, H, 'fill="url(#gg)"') + R(10, 10, 277, 190, 'fill="#fffbef"') + frame(13, '#1f2d4d', .6) + frame(16, '#b8923a', .4) +
      four(t => '<g transform="' + t + '"><circle cx="13" cy="13" r="3" fill="#1f2d4d"/></g>')) },
    { id: 'jordan', name: 'ألوان العلم الأردني', c: { club: '#007a3d', title: '#ce1126', text: '#1f2937' }, svg: svg(
      R(0, 0, W, H, 'fill="#fff"') + R(0, 0, W, 4, 'fill="#111"') + R(0, 4, W, 4, 'fill="#e5e7eb"') + R(0, 8, W, 4, 'fill="#007a3d"') +
      R(0, 198, W, 4, 'fill="#007a3d"') + R(0, 202, W, 4, 'fill="#e5e7eb"') + R(0, 206, W, 4, 'fill="#111"') +
      P('0,12 0,198 22,105', 'fill="#ce1126"') + P('297,12 297,198 275,105', 'fill="#ce1126"') + R(16, 18, 265, 174, 'fill="none" stroke="#007a3d" stroke-width=".6"')) },
    { id: 'geo', name: 'زخارف هندسية', c: { club: '#0f766e', title: '#9a7425', text: '#1f2937' }, svg: svg(
      R(0, 0, W, H, 'fill="#fdfaf3"') + frame(5, '#0f766e', .8) + frame(15, '#c9a24a', .6) + diamonds('#0f766e', '#c9a24a')) }
  ];
  const urlOf = t => 'data:image/svg+xml;base64,' + btoa(t.svg);
  TPLS.forEach(t => { t.url = urlOf(t); });
  const tplOf = S => TPLS.find(t => t.id === (S.cert && S.cert.tpl)) || TPLS[0];
  const SVG_URL = TPLS[0].url;

  function title(d) { return d.kind === 'excel' ? 'شهادة تفوق' : d.kind === 'custom' ? (d.customTitle || '').trim() : 'شكر وتقدير'; }
  function clubLine(S) {
    const d = S.cert;
    if ((d.clubText || '').trim()) return d.clubText.trim();
    const n = (S.clubName || '').trim();
    return n ? 'نادي معلمي ' + n : 'نادي المعلمين';
  }

  const fontCss = n => n ? "font-family:'" + n + "',Tahoma,Arial,sans-serif;" : '';
  const off = (x, y) => (+x || +y) ? 'left:' + (+x || 0) + 'mm;top:' + (+y || 0) + 'mm;' : '';
  /** الجملة التمهيدية: «يتقدّم نادي ... بالشكر والتقدير» */
  function leadText(S) { const d = S.cert; return (d.leadText || '').trim() || 'يتقدّم ' + clubLine(S) + ' بالشكر والتقدير إلى'; }

  /** نص الشهادة؛ variant = اسم في نسخة مكرّرة */
  function certHTML(S, variant) {
    const d = S.cert, ph = t => '<span class="ph">' + t + '</span>';
    const name = variant || d.name;
    const logo = S.logos && S.logos.club;
    const ttl = title(d);
    const wm = S.wmShow !== false ? '<img class="ct-wm" alt="" src="' + g.WM_LOGO + '" style="opacity:' + (Math.min(100, Math.max(2, +S.wmOpacity || 8)) / 100) + '">' : '';
    const lg = d.logoShow && logo && logo.data
      ? '<img class="ct-logo" alt="" src="' + logo.data + '" style="width:' + d.logoSize + 'mm;left:calc(50% + ' + d.logoX + 'mm);top:' + (14 + (+d.logoY || 0)) + 'mm">' : '';
    const lead = d.leadOn
      ? '<div class="ct-lead">' + esc(leadText(S)) + (d.honorific ? ' ' + esc(d.honorific) : '') + ' <b class="ct-nm">' + (name ? esc(name) : ph('اسم المكرَّم')) + '</b></div>' : '';
    const main = d.text.trim() ? '<div class="ct-main">' + esc(d.text.trim()).replace(/\n/g, '<br>') + '</div>' : (d.leadOn ? '' : '<div class="ct-main">' + ph('نص الشهادة') + '</div>');
    const tx = 'font-size:' + d.textSize + 'pt;color:' + esc(d.textColor) + ';text-align:' + d.textAlign + ';' + (d.textBold ? 'font-weight:700;' : '') + fontCss(d.textFont) + off(d.textX, d.textY);
    return '<div class="cert">' + wm + lg + '<div class="ct-body">' +
      '<div class="ct-club" style="font-size:' + d.clubSize + 'pt;color:' + esc(d.clubColor) + ';' + fontCss(d.clubFont) + off(d.clubX, d.clubY) + '">' + esc(clubLine(S)) + '</div>' +
      '<div class="ct-title" style="font-size:' + d.titleSize + 'pt;color:' + esc(d.titleColor) + ';' + fontCss(d.titleFont) + off(d.titleX, d.titleY) + '">' + (ttl ? esc(ttl) : ph('العنوان')) + '</div>' +
      '<div class="ct-text" style="' + tx + '">' + lead + main + '</div></div>' +
      '<div class="ct-sign">' + App.signHTML(S, 'cert') + '</div></div>';
  }

  /** خلفية الصفحة (إطار + علامة مائية) كصورة PNG لملف Word */
  function bgPng(S) {
    return new Promise((res, rej) => {
      const i = new Image();
      i.onerror = rej;
      i.onload = () => {
        const cv = document.createElement('canvas'); cv.width = 1754; cv.height = 1240;
        const cx = cv.getContext('2d'); cx.drawImage(i, 0, 0, cv.width, cv.height);
        const done = () => res(cv.toDataURL('image/png'));
        if (S.wmShow === false) return done();
        const w = new Image(); w.onload = () => {
          const h = cv.height * 0.62, ww = h * w.width / w.height;
          cx.globalAlpha = Math.min(100, Math.max(2, +S.wmOpacity || 8)) / 100;
          cx.drawImage(w, (cv.width - ww) / 2, cv.height * 0.54 - h / 2, ww, h); done();
        }; w.onerror = done; w.src = g.WM_LOGO;
      };
      i.src = tplOf(S).url;
    });
  }

  g.App = Object.assign(App, { certHTML, certBgUrl: SVG_URL, certBgUrlFor: S => tplOf(S).url, CERT_TPLS: TPLS, certBgPng: bgPng, certTitle: title, certClub: clubLine, certLead: leadText });
})(window);
