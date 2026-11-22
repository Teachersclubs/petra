/* العلامة المائية: شعار شفاف في منتصف كل صفحة (معاينة/طباعة/PDF/Word) */
(function (g) {
  'use strict';
  const App = g.App || (g.App = {});
  const W = 794, H = 1123; // A4 عند 96dpi
  let img = null, tries = 0;
  const cache = { key: '', tile: '', logo: '' };

  function load() {
    if (img || !g.WM_LOGO) return;
    const i = new Image();
    i.onload = () => { img = i; g.dispatchEvent(new Event('wm-ready')); };
    i.src = g.WM_LOGO;
  }
  load();

  const num = (v, d) => { v = parseFloat(v); return isFinite(v) ? v : d; };
  const conf = S => ({ show: S.wmShow !== false, op: Math.min(100, Math.max(2, num(S.wmOpacity, 12))), size: Math.min(100, Math.max(20, num(S.wmSize, 60))) });
  const key = c => c.op + '|' + c.size;

  function build(S) {
    const c = conf(S); if (!c.show || !img) return null;
    if (cache.key !== key(c)) {
      const k = 1.25, cv = document.createElement('canvas'); cv.width = Math.round(W * k); cv.height = Math.round(H * k);
      const cx = cv.getContext('2d'); cx.globalAlpha = c.op / 100;
      const w = cv.width * c.size / 100, h = w * img.height / img.width;
      cx.drawImage(img, (cv.width - w) / 2, (cv.height - h) / 2, w, h);
      const lw = 520, lc = document.createElement('canvas'); lc.width = lw; lc.height = Math.round(lw * img.height / img.width);
      const lx = lc.getContext('2d'); lx.globalAlpha = c.op / 100; lx.drawImage(img, 0, 0, lc.width, lc.height);
      cache.key = key(c); cache.tile = cv.toDataURL('image/png'); cache.logo = lc.toDataURL('image/png');
    }
    return cache;
  }
  /** قيمة المتغير CSS للصفحة ('' إذا كانت العلامة مخفية) */
  App.wmVar = S => { const b = build(S); return b ? ';--wm-img:url(' + b.tile + ')' : ''; };
  /** صورة الصفحة الكاملة (للـPDF) */
  App.wmTileImage = S => new Promise(res => {
    const b = build(S); if (!b) return res(null);
    const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = b.tile;
  });
  /** الشعار الباهت وحده + حجمه بالمليمتر (لـWord) */
  App.wmLogo = S => { const b = build(S); if (!b) return null; const c = conf(S); return { url: b.logo, wMm: 210 * c.size / 100, ratio: img.height / img.width }; };
})(window);
