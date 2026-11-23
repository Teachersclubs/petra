/* قالب الشهادة الذهبي (مبني من ملف PowerPoint): كل عنصر له موضع وحجم وخط ولون مستقل */
(function (g) {
  'use strict';
  const App = g.App || (g.App = {});
  const esc = s => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
  const spec = () => App.T2_SPEC;
  const A = () => g.CERT_CANVA || {};

  const FEM = ['المعلمة', 'السيدة', 'الطالبة'];
  /** نص العنصر النهائي (بدون HTML) */
  function t2Text(S, id, variant) {
    const d = S.cert, v = k => d['q_' + id + '_' + k];
    if (id === 'lead') return (v('t') || '').trim() || 'يتقدم ' + App.certClub(S) + ' بخالص الشكر والتقدير إلى';
    if (id === 'name') {
      const nm = (variant || d.name || '').trim(), h = d.q_honor || '';
      if (!nm) return '';
      return (h ? h + ' ' : '') + nm + (d.q_suffix && h ? ' ' + (FEM.includes(h) ? 'المحترمة' : 'المحترم') + '،' : '');
    }
    const t = v('t') || '';
    if (id === 'body' && t === App.T2_TXT_DEF.body) return t.replace('نادي معلمي البترا', App.certClub(S));
    return t;
  }

  const font = n => "font-family:'" + n + "',Tahoma,Arial,sans-serif;";
  /** مستطيل العنصر بالمليمتر: [يسار، أعلى، عرض] */
  function box(S, e) {
    const d = S.cert, w = +d['q_' + e.id + '_w'] || e.w;
    return [e.cx - w / 2 + (+d['q_' + e.id + '_x'] || 0), e.y + (+d['q_' + e.id + '_y'] || 0), w];
  }
  const r2 = n => Math.round(n * 100) / 100;

  function imgSrc(S, id) {
    if (id === 'logo') { const L = S.logos && S.logos.club; return L && L.data ? L.data : ''; }
    return A()[id] || '';
  }

  function t2HTML(S, variant) {
    const d = S.cert, ph = t => '<span class="ph">' + t + '</span>';
    const wm = S.wmShow !== false ? '<img class="ct-wm" alt="" src="' + g.WM_LOGO + '" style="opacity:' + (Math.min(100, Math.max(2, +S.wmOpacity || 8)) / 100) + '">' : '';
    const out = spec().map(e => {
      if (!d['q_' + e.id + '_on']) return '';
      const [l, t, w] = box(S, e), pos = 'left:' + r2(l) + 'mm;top:' + r2(t) + 'mm;width:' + r2(w) + 'mm;';
      if (e.k === 'img') {
        const src = imgSrc(S, e.id); if (!src) return '';
        return '<img class="t2e i" data-el="' + e.id + '" alt="" src="' + src + '" style="' + pos + '">';
      }
      if (e.k === 'sign') return '<div class="t2e sg" data-el="sign" style="' + pos + '">' + App.signHTML(S, 'cert') + '</div>';
      if (e.k === 'line') return '<div class="t2e ln" data-el="' + e.id + '" style="' + pos + 'border-color:' + esc(d['q_' + e.id + '_c']) + '"></div>';
      let txt = t2Text(S, e.id, variant);
      let html = txt ? esc(txt).replace(/\n/g, '<br>') : (e.id === 'name' ? ph('اسم المكرَّم') : '');
      return '<div class="t2e tx" data-el="' + e.id + '" style="' + pos + 'font-size:' + (+d['q_' + e.id + '_s'] || e.s) + 'pt;color:' + esc(d['q_' + e.id + '_c']) + ';' +
        font(d['q_' + e.id + '_f'] || e.f) + 'font-weight:' + (d['q_' + e.id + '_b'] ? 700 : 400) + '">' + html + '</div>';
    }).join('');
    return '<div class="cert t2">' + wm + out + '</div>';
  }

  /** رسم الصور والخطوط فوق الخلفية على canvas (1754×1240) لملف Word */
  async function overlay(cx, cv, S) {
    const d = S.cert, k = cv.width / 297;
    const load = src => new Promise(res => { if (!src) return res(null); const i = new Image(); i.onload = () => res(i); i.onerror = () => res(null); i.src = src; });
    cx.globalAlpha = 1;
    const baked = [];
    for (const e of spec()) {
      if (!d['q_' + e.id + '_on']) continue;
      const [l, t, w] = box(S, e);
      if (baked.includes(e.id)) {
        const sz = (+d['q_' + e.id + '_s'] || e.s) * 0.3528 * k, txt = t2Text(S, e.id);
        cx.save(); cx.direction = 'rtl'; cx.textAlign = 'center'; cx.textBaseline = 'middle'; cx.fillStyle = d['q_' + e.id + '_c'];
        cx.font = (d['q_' + e.id + '_b'] ? '700 ' : '400 ') + sz + "px '" + (d['q_' + e.id + '_f'] || e.f) + "',Tahoma,Arial,sans-serif";
        cx.fillText(txt, (l + w / 2) * k, (t + 0.7 * (+d['q_' + e.id + '_s'] || e.s) * 0.3528) * k); cx.restore(); continue;
      }
      if (e.k === 'line') { cx.fillStyle = d['q_' + e.id + '_c']; cx.fillRect(l * k, t * k, w * k, Math.max(1, .3 * k)); }
      else if (e.k === 'img') {
        const im = await load(imgSrc(S, e.id)); if (!im) continue;
        const h = e.id === 'divider' ? w * e.ar : w * im.naturalHeight / im.naturalWidth;
        cx.drawImage(im, l * k, t * k, w * k, h * k);
      }
    }
  }

  Object.assign(App, { t2HTML, t2Text, t2Overlay: overlay, t2Box: box });
})(window);
