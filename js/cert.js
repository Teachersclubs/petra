/* شهادة شكر وتقدير: ورقة أفقية بإطار وزوايا ذهبية وشعار وعلامة مائية */
(function (g) {
  'use strict';
  const App = g.App || (g.App = {});
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));

  /* ---- الإطار: SVG بأبعاد 297×210 مم ---- */
  const corner = t => '<g transform="' + t + '"><polygon points="0,0 34,0 26,8 8,8" fill="#dfbd78"/><polygon points="0,0 8,8 8,26 0,34" fill="#c19a47"/>' +
    '<polygon points="8,8 26,8 22,11 11,11 11,22 8,26" fill="#a9822f" opacity=".55"/></g>';
  const SVG = '<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 297 210" width="297" height="210">' +
    '<rect width="297" height="210" fill="#3b3d52"/><rect x="11" y="11" width="275" height="188" fill="#fff"/>' +
    corner('translate(11,11)') + corner('translate(286,11) scale(-1,1)') + corner('translate(11,199) scale(1,-1)') + corner('translate(286,199) scale(-1,-1)') + '</svg>';
  const SVG_URL = 'data:image/svg+xml;base64,' + btoa(SVG);

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
      i.src = SVG_URL;
    });
  }

  g.App = Object.assign(App, { certHTML, certBgUrl: SVG_URL, certBgPng: bgPng, certTitle: title, certClub: clubLine, certLead: leadText });
})(window);
