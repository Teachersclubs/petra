/* ===== الطباعة وتصدير PDF ===== */
(function (g) {
  'use strict';
  const App = g.App;
  const PAGE_W = 794, PAGE_H = 1123; // A4 عند 96dpi

  function sheetWrapper(S, type, forPrint, vi) {
    const r = App.renderDoc(S, type, vi);
    const w = document.createElement('div');
    w.setAttribute('style', r.style);
    w.dir = 'rtl';
    if (forPrint) {
      w.innerHTML =
        '<table class="pt"><thead><tr><td><div class="pt-gap"></div></td></tr></thead><tbody><tr><td><div class="sheet-main">' + r.main + '</div></td></tr></tbody>' +
        '<tfoot><tr><td><div class="sheet-foot">' + r.foot + '</div></td></tr></tfoot></table>';
    }
    return { wrap: w, r };
  }

  /* --- الطباعة --- */
  function printDoc(S, type) {
    // يظهر عنوان الصفحة كاسم افتراضي عند «حفظ كـ PDF» من نافذة الطباعة
    const oldTitle = document.title; document.title = App.fileName(S, type);
    const root = document.getElementById('printRoot');
    const wrap = document.createElement('div');
    App.variants(S, type).forEach((v, i) => {
      const w = sheetWrapper(S, type, true, i).wrap;
      if (i > 0) w.style.breakBefore = 'page';
      wrap.appendChild(w);
    });
    App.attachPagesHTML(App.attach.flat(type)).forEach(h => { const d = document.createElement('div'); d.innerHTML = h; wrap.appendChild(d.firstChild); });
    root.replaceChildren(wrap);
    const done = () => { document.title = oldTitle; root.replaceChildren(); window.removeEventListener('afterprint', done); };
    window.addEventListener('afterprint', done);
    // مهلة قصيرة ليحمّل المتصفح الصور قبل فتح نافذة الطباعة
    setTimeout(() => window.print(), 80);
  }

  /* --- PDF: التقاط صفحة الرأس/المحتوى والذيل ثم تقطيعهما إلى صفحات A4 --- */
  async function buildPdf(S, type) {
    const pdf = new g.jspdf.jsPDF({ orientation: 'p', unit: 'px', format: [PAGE_W, PAGE_H], compress: true, hotfixes: ['px_scaling'] });
    let first = true;
    const vars = App.variants(S, type);
    for (let vi = 0; vi < vars.length; vi++) { await addVariant(pdf, S, type, vi, () => { const f = first; first = false; return f; }); }
    return finishPdf(pdf, S, type, first);
  }

  async function addVariant(pdf, S, type, vi, takeFirst) {
    const { r } = sheetWrapper(S, type, false, vi);
    const stage = document.createElement('div');
    stage.style.cssText = 'position:fixed;left:-20000px;top:0;width:' + PAGE_W + 'px;direction:rtl;background:#fff;color:#000';
    stage.setAttribute('style', stage.getAttribute('style') + ';' + r.style);
    stage.innerHTML = '<div class="sheet-main" id="pdfMain" style="padding-bottom:0">' + r.main + '</div><div class="sheet-foot" id="pdfFoot">' + r.foot + '</div>';
    document.body.appendChild(stage);
    stage.querySelectorAll('.ph').forEach(e => { e.style.visibility = 'hidden'; });
    try {
      if (document.fonts && document.fonts.ready) await document.fonts.ready;
      const opts = { scale: 2, backgroundColor: '#ffffff', useCORS: true, logging: false };
      const mainEl = stage.querySelector('#pdfMain'), footEl = stage.querySelector('#pdfFoot');
      const mainH = mainEl.offsetHeight, footH = r.foot ? footEl.offsetHeight : 0;
      const mainC = await html2canvas(mainEl, Object.assign({ width: PAGE_W, height: mainH, windowWidth: PAGE_W }, opts));
      const footC = footH ? await html2canvas(footEl, Object.assign({ width: PAGE_W, height: footH, windowWidth: PAGE_W }, opts)) : null;

      // مواضع الأسطر والصور: لتفادي قطع سطر نصي بين صفحتين
      const base = mainEl.getBoundingClientRect().top;
      const boxes = [];
      const walker = document.createTreeWalker(mainEl, NodeFilter.SHOW_TEXT);
      for (let n = walker.nextNode(); n; n = walker.nextNode()) {
        if (!n.nodeValue.trim()) continue;
        const rg = document.createRange(); rg.selectNodeContents(n);
        Array.from(rg.getClientRects()).forEach(q => { if (q.height > 0) boxes.push([q.top - base, q.bottom - base]); });
      }
      mainEl.querySelectorAll('img').forEach(im => { const q = im.getBoundingClientRect(); boxes.push([q.top - base, q.bottom - base]); });
      const snap = cut => {
        for (let i = 0; i < 200; i++) {
          const hit = boxes.filter(b => b[0] + 1 < cut && cut < b[1] - 1);
          if (!hit.length) return cut;
          cut = Math.min.apply(null, hit.map(b => b[0]));
        }
        return cut;
      };

      const k = mainC.width / PAGE_W; // عامل التكبير الفعلي
      const topGap = 38, bottomGap = 6;

      let y = 0, page = 0;
      while (y < mainH - 1 || page === 0) {
        const offTop = page === 0 ? 0 : topGap;
        const avail = PAGE_H - footH - bottomGap - offTop;
        let sh = Math.min(avail, mainH - y);
        if (y + sh < mainH - 1) { const c = snap(y + sh); if (c - y > 60) sh = c - y; }
        const pc = document.createElement('canvas');
        pc.width = PAGE_W * k; pc.height = PAGE_H * k;
        const cx = pc.getContext('2d');
        cx.fillStyle = '#fff'; cx.fillRect(0, 0, pc.width, pc.height);
        if (sh > 0) cx.drawImage(mainC, 0, Math.round(y * k), mainC.width, Math.round(sh * k), 0, Math.round(offTop * k), mainC.width, Math.round(sh * k));
        if (footC) cx.drawImage(footC, 0, Math.round((PAGE_H - footH) * k));
        if (!(page === 0 && takeFirst())) pdf.addPage([PAGE_W, PAGE_H], 'p');
        pdf.addImage(pc.toDataURL('image/jpeg', 0.92), 'JPEG', 0, 0, PAGE_W, PAGE_H, undefined, 'FAST');
        if (App.setBusyMsg) App.setBusyMsg('جارٍ تجهيز PDF… صفحة ' + (pdf.getNumberOfPages()));
        await new Promise(r => setTimeout(r, 0));
        y += sh; page++;
        if (page > 30) break;
      }
    } finally {
      stage.remove();
    }
  }

  async function finishPdf(pdf, S, type) {
    {
      for (const p of App.attach.flat(type)) {
        pdf.addPage([PAGE_W, PAGE_H], 'p');
        const k2 = 2, pc = document.createElement('canvas');
        pc.width = PAGE_W * k2; pc.height = PAGE_H * k2;
        const cx = pc.getContext('2d');
        cx.fillStyle = '#fff'; cx.fillRect(0, 0, pc.width, pc.height);
        cx.fillStyle = '#444'; cx.font = (14 * k2) + 'px Tahoma, Arial, sans-serif'; cx.textAlign = 'center'; cx.direction = 'rtl';
        cx.fillText(p.cap, pc.width / 2, 48 * k2);
        const im = await new Promise((res, rej) => { const i = new Image(); i.onload = () => res(i); i.onerror = rej; i.src = p.url; });
        const bw = PAGE_W - 76, bh = PAGE_H - 120, kk = Math.min(bw / p.w, bh / p.h);
        const iw = p.w * kk, ih = p.h * kk;
        cx.drawImage(im, ((PAGE_W - iw) / 2) * k2, (62 + (bh - ih) / 2) * k2, iw * k2, ih * k2);
        pdf.addImage(pc.toDataURL('image/jpeg', 0.9), 'JPEG', 0, 0, PAGE_W, PAGE_H, undefined, 'FAST');
      }
      pdf.setProperties({ title: S[type].subject || App.fileName(S, type), creator: 'أندية المعلمين والعاملين' });
      return pdf.output('blob');
    }
  }

  g.App = Object.assign(App, { printDoc, buildPdf });
})(window);
