/* ===== واجهة التطبيق: الربط والتبويبات والمعاينة والتصدير ===== */
(function () {
  'use strict';
  const App = window.App;
  const $ = (s, r) => (r || document).querySelector(s);
  const $$ = (s, r) => Array.from((r || document).querySelectorAll(s));

  let S = App.build(null);
  const PAGE_H = 1122.52; // 297mm
  let zoom = 1, fitMode = true;

  /* ---------- تنبيه ---------- */
  let toastT;
  window.toast = function (msg) {
    const t = $('#toast'); t.textContent = msg; t.classList.add('on');
    clearTimeout(toastT); toastT = setTimeout(() => t.classList.remove('on'), 2800);
  };

  /* ---------- الحفظ والعرض ---------- */
  let raf = 0;
  function changed() {
    App.save(S);
    cancelAnimationFrame(raf);
    raf = requestAnimationFrame(renderPreview);
  }

  function renderPreview() {
    const sheet = $('#sheet');
    const r = App.renderDoc(S, S.activeDoc);
    App.lastEff = r.eff;
    sheet.setAttribute('style', r.style);
    sheet.innerHTML = '<div class="sheet-main">' + r.main + '</div><div class="sheet-foot">' + r.foot + '</div>';
    $('#copySheets').innerHTML = App.variants(S, S.activeDoc).slice(1).map((v, i) => {
      const c = App.renderDoc(S, S.activeDoc, i + 1);
      return '<div class="sheet copy" style="' + c.style.replace(/"/g, '&quot;') + '"><div class="sheet-main">' + c.main + '</div><div class="sheet-foot">' + c.foot + '</div></div>';
    }).join('');
    $('#attSheets').innerHTML = App.attachPagesHTML(App.attach.flat(S.activeDoc)).map(h => '<div class="sheet att">' + h + '</div>').join('');
    markSel();
    applyZoom(true);
  }

  function applyZoom(measure) {
    const sheet = $('#sheet'), holder = $('#sheetHolder');
    if (fitMode) {
      const avail = $('#previewScroll').clientWidth - 32;
      zoom = Math.max(0.3, Math.min(1, avail / 794));
    }
    if (measure) {
      holder.style.zoom = 1; sheet.style.minHeight = '';
      const h = sheet.offsetHeight;
      const pages = Math.max(1, Math.ceil((h - 2) / PAGE_H));
      sheet.style.minHeight = (pages * PAGE_H) + 'px';
      $('#pageInfo') && ($('#pageInfo').textContent = pages > 1 ? pages + ' صفحات' : '');
      const bar = $('.preview-bar > span');
      const eff = App.lastEff;
      bar.innerHTML = 'معاينة مباشرة · A4' + (pages > 1 ? ' · ' + pages + ' صفحات' : '') +
        (S.onePage && eff ? (eff.fits ? (eff.scale < 1 ? ' · <small>ضُبط على صفحة واحدة</small>' : '') : ' · <span class="warn-fit">تعذّر حصر المستند في صفحة واحدة</span>') : '');
    }
    holder.style.zoom = zoom;
    $('#zoomVal').textContent = Math.round(zoom * 100) + '%';
  }

  /* ---------- ربط الحقول ---------- */
  function readEl(el) {
    if (el.type === 'checkbox') return el.checked;
    if (el.type === 'number' || el.type === 'range') return el.value === '' ? 0 : +el.value;
    return el.value;
  }
  function writeEl(el, v) {
    if (el.type === 'checkbox') el.checked = !!v;
    else if (el !== document.activeElement || el.value !== String(v)) el.value = v == null ? '' : v;
  }

  function refreshControls(skip) {
    $$('[data-bind]').forEach(el => { if (el !== skip) writeEl(el, App.getPath(S, el.dataset.bind)); });
    $$('[data-doc-bind]').forEach(el => { if (el !== skip) writeEl(el, App.getPath(S, el.dataset.docBind)); });
    $$('[data-out]').forEach(o => { o.textContent = S[o.dataset.out]; });
    $$('[data-show-if]').forEach(el => { el.hidden = !App.getPath(S, el.dataset.showIf); });
    $$('[data-out-path]').forEach(o => { o.textContent = App.getPath(S, o.dataset.outPath); });
    $$('.people-card').forEach(card => {
      const P = S.people[card.dataset.person];
      $$('.img-slot', card).forEach(slot => {
        const url = slot.dataset.kind === 'sign' ? P.signImg : P.stampImg;
        const img = $('img', slot), em = $('em', slot);
        img.hidden = !url; em.hidden = !!url; if (url) img.src = url;
        $('[data-remove]', slot).hidden = !url;
        $('label.btn', slot).firstChild.nodeValue = (url ? 'استبدال ' : 'رفع ') + (slot.dataset.kind === 'sign' ? 'توقيع' : 'ختم');
      });
    });
    const ps = $('#persistState');
    if (ps) {
      ps.textContent = 'الحفظ الدائم مفعّل: تُحفظ البيانات تلقائيًا وتبقى بعد إغلاق الموقع' + (App.isPersisted() ? ' (تخزين محمي من المسح التلقائي).' : '.');
    }
    ['ministry', 'club'].forEach(k => {
      const card = $('.logo-card[data-logo="' + k + '"]');
      const L = S.logos[k], img = $('img', card), em = $('em', card);
      img.hidden = !L.data; em.hidden = !!L.data;
      if (L.data) img.src = L.data;
    });
    dgSync();
  }

  function onInput(el, path) {
    const v = readEl(el);
    App.setPath(S, path, v);
    if (path === 'assign.assignDate') S.assign.assignDateText = '';
    if (path === 'claim.claimDate') S.claim.claimDateText = '';
    const m = /^(letter|internal|thanks|circular|assign|claim|delegation)\.(date|hijri|showHijri)$/.exec(path);
    if (m) {
      const d = S[m[1]];
      if (m[2] === 'date') { if (!d.hijriManual) d.hijri = App.formatHijri(v); }
      else if (m[2] === 'hijri') { d.hijriManual = v.trim() !== ''; if (!d.hijriManual) d.hijri = App.formatHijri(d.date); }
      else if (v && !d.hijri) d.hijri = App.formatHijri(d.date);
    }
    if (el.hasAttribute && el.hasAttribute('data-geo') && !acceptGeo(path === 'people.protect')) { refreshControls(); return; }
    if (path === 'font' || path === 'footer.font') applyFonts();
    refreshControls(el);
    changed();
  }

  function bindNudge() {
    const lim = { 'thanks.toY': [-5, 80], 'thanks.toX': [-40, 40] };
    document.addEventListener('click', e => {
      const b = e.target.closest('[data-nudge]');
      if (b) {
        const p = b.dataset.nudge, st = +($('#posStep') && $('#posStep').value) || 1, [lo, hi] = lim[p];
        const v = Math.max(lo, Math.min(hi, Math.round(((+App.getPath(S, p) || 0) + (+b.dataset.d) * st) * 10) / 10));
        App.setPath(S, p, v); refreshControls(); changed(); return;
      }
      const rs = e.target.closest('[data-nudge-reset]');
      if (rs) { S.thanks.toX = 0; S.thanks.toY = 0; refreshControls(); changed(); }
    });
  }
  function bindAll() {
    bindNudge();
    document.addEventListener('input', e => {
      const el = e.target;
      if (el.dataset && el.dataset.bind) onInput(el, el.dataset.bind);
      else if (el.dataset && el.dataset.docBind) onInput(el, el.dataset.docBind);
    });
    document.addEventListener('change', e => {
      const el = e.target;
      if (el.type === 'checkbox' || el.tagName === 'SELECT') {
        if (el.dataset.bind) onInput(el, el.dataset.bind);
        else if (el.dataset.docBind) onInput(el, el.dataset.docBind);
      }
    });
  }

  /* ---------- محرر النص المنسق ---------- */
  const RTE_BTNS = [
    ['bold', '<b>B</b>', 'عريض'], ['italic', '<i style="font-family:serif">I</i>', 'مائل'], ['underline', '<u>U</u>', 'تسطير'], '|',
    ['insertUnorderedList', '•≡', 'قائمة نقطية'], ['insertOrderedList', '1.', 'قائمة رقمية'], '|',
    ['justifyRight', '⇥', 'محاذاة لليمين'], ['justifyCenter', '≡', 'توسيط'], ['justifyLeft', '⇤', 'محاذاة لليسار'], ['justifyFull', '☰', 'ضبط'], '|',
    ['undo', '↶', 'تراجع'], ['redo', '↷', 'إعادة'], '|',
    ['table', '▦ جدول', 'إدراج جدول'],
    ['image', '🖼 صورة', 'إدراج صورة داخل النص'], '|',
    ['kashida', 'ـــ كشيدة', 'مدّ الحرف (كشيدة): ضع المؤشر داخل الكلمة أو حدّد نصًا ثم اضغط'],
    ['unkashida', 'إزالة الكشيدة', 'إزالة الكشيدة من التحديد أو بجوار المؤشر']
  ];

  /* ---------- الكشيدة (التطويل) ---------- */
  const AR = /[\u0621-\u064A]/;
  const NOJOIN = /[\u0621\u0622\u0623\u0624\u0625\u0627\u062F\u0630\u0631\u0632\u0648\u0629\u0649]/; // حروف لا تتصل بما بعدها
  function kashidaText(t) {
    return t.replace(/[\u0621-\u064A\u0640\u064B-\u065F]{3,}/g, w => {
      const ch = Array.from(w); let best = -1, mid = ch.length / 2, bd = 1e9;
      for (let i = 0; i < ch.length - 1; i++) {
        const a = ch[i], b = ch[i + 1];
        if (!AR.test(a) || NOJOIN.test(a) || !AR.test(b) || b === '\u0640') continue;
        if (a === '\u0640' ) continue;
        const d = Math.abs(i + 0.5 - mid); if (d < bd) { bd = d; best = i; }
      }
      if (best < 0) return w;
      ch.splice(best + 1, 0, '\u0640');
      return ch.join('');
    });
  }
  function textNodesIn(range) {
    const out = [], root = range.commonAncestorContainer.nodeType === 3 ? range.commonAncestorContainer.parentNode : range.commonAncestorContainer;
    const w = document.createTreeWalker(root, NodeFilter.SHOW_TEXT);
    for (let n = w.nextNode(); n; n = w.nextNode()) if (range.intersectsNode(n)) out.push(n);
    return out;
  }
  function applyKashida(area, remove) {
    const sel = getSelection();
    if (!sel.rangeCount || !area.contains(sel.anchorNode)) { area.focus(); return false; }
    const r = sel.getRangeAt(0);
    if (r.collapsed) {
      if (remove) {
        const n = r.startContainer, o = r.startOffset;
        if (n.nodeType !== 3) return false;
        const v = n.nodeValue;
        if (v[o - 1] === '\u0640') { n.nodeValue = v.slice(0, o - 1) + v.slice(o); r.setStart(n, o - 1); r.collapse(true); }
        else if (v[o] === '\u0640') { n.nodeValue = v.slice(0, o) + v.slice(o + 1); }
        else return false;
      } else document.execCommand('insertText', false, '\u0640');
      return true;
    }
    let changed = false, first = null, last = null, so = 0, eo = 0;
    textNodesIn(r).forEach(n => {
      const s = n === r.startContainer ? r.startOffset : 0, e = n === r.endContainer ? r.endOffset : n.nodeValue.length;
      const v = n.nodeValue, mid = v.slice(s, e);
      const nm = remove ? mid.replace(/\u0640/g, '') : kashidaText(mid);
      if (nm !== mid) { n.nodeValue = v.slice(0, s) + nm + v.slice(e); changed = true; }
      if (!first) { first = n; so = s; }
      last = n; eo = s + nm.length;
    });
    if (first) { try { const nr = document.createRange(); nr.setStart(first, so); nr.setEnd(last, Math.min(eo, last.nodeValue.length)); sel.removeAllRanges(); sel.addRange(nr); } catch (x) { /* لا شيء */ } }
    return changed;
  }

  /* ---------- صورة داخل النص ---------- */
  function readImageFile(file) {
    return new Promise((res, rej) => {
      const fr = new FileReader();
      fr.onerror = rej;
      fr.onload = () => {
        const im = new Image();
        im.onerror = rej;
        im.onload = () => {
          const k = Math.min(1, 1400 / Math.max(im.naturalWidth, im.naturalHeight));
          const c = document.createElement('canvas'); c.width = Math.max(1, Math.round(im.naturalWidth * k)); c.height = Math.max(1, Math.round(im.naturalHeight * k));
          const x = c.getContext('2d');
          const png = /png|gif|webp|svg/.test(file.type);
          if (!png) { x.fillStyle = '#fff'; x.fillRect(0, 0, c.width, c.height); }
          x.drawImage(im, 0, 0, c.width, c.height);
          let url = png ? c.toDataURL('image/png') : c.toDataURL('image/jpeg', 0.88);
          if (png && url.length > 1.8e6) url = c.toDataURL('image/jpeg', 0.85);
          res({ url, ar: c.height / c.width });
        };
        im.src = fr.result;
      };
      fr.readAsDataURL(file);
    });
  }
  function cellOf(area) {
    const sel = getSelection();
    if (!sel.rangeCount) return null;
    let n = sel.anchorNode;
    if (!n || !area.contains(n)) return null;
    if (n.nodeType === 3) n = n.parentNode;
    return n.closest ? n.closest('td,th') : null;
  }
  function tableOp(area, op) {
    const c = cellOf(area); if (!c) return;
    const tr = c.parentElement, table = c.closest('table');
    if (op === 'del') { table.remove(); return; }
    if (op === 'row') {
      const nr = document.createElement('tr');
      Array.from(tr.children).forEach(() => { const td = document.createElement('td'); td.innerHTML = '<br>'; nr.appendChild(td); });
      tr.after(nr);
    } else {
      const idx = c.cellIndex;
      table.querySelectorAll('tr').forEach(r => { const td = document.createElement('td'); td.innerHTML = '<br>'; const ref = r.children[idx]; if (ref) ref.after(td); else r.appendChild(td); });
    }
  }
  function openTableDialog(area, sync) {
    const dlg = $('#tableDlg');
    const sel = getSelection();
    const saved = sel.rangeCount && area.contains(sel.anchorNode) ? sel.getRangeAt(0).cloneRange() : null;
    dlg.returnValue = '';
    dlg.onclose = () => {
      if (dlg.returnValue !== 'ok') return;
      const rows = Math.max(1, Math.min(30, +$('#tRows').value || 3)), cols = Math.max(1, Math.min(8, +$('#tCols').value || 3));
      const head = $('#tHead').checked;
      let h = '<table data-new="1"><tbody>';
      for (let r = 0; r < rows; r++) {
        h += '<tr>';
        for (let c = 0; c < cols; c++) h += (head && r === 0 ? '<th><br></th>' : '<td><br></td>');
        h += '</tr>';
      }
      h += '</tbody></table><p><br></p>';
      area.focus();
      // الإدراج دائمًا بعد الفقرة الحالية (وبعد الجدول إن كان المؤشر داخل جدول) لتفادي جداول متداخلة أو تقسيم الفقرة
      let blk = saved ? saved.startContainer : null;
      while (blk && blk.parentNode !== area) blk = blk.parentNode;
      const r2 = document.createRange();
      if (blk && blk.parentNode === area) r2.setStartAfter(blk); else { r2.selectNodeContents(area); r2.collapse(false); }
      r2.collapse(true);
      const s2 = getSelection(); s2.removeAllRanges(); s2.addRange(r2);
      let done = false;
      try { done = document.execCommand('insertHTML', false, h); } catch (e) { done = false; }
      let nt = area.querySelector('table[data-new]');
      if (!nt) { // احتياط: إدراج مباشر في DOM
        const tmp = document.createElement('div'); tmp.innerHTML = h;
        const ref = blk && blk.parentNode === area ? blk.nextSibling : null;
        Array.from(tmp.childNodes).forEach(n => area.insertBefore(n, ref));
        nt = area.querySelector('table[data-new]');
      }
      if (nt) {
        nt.removeAttribute('data-new');
        const first = nt.querySelector('th,td');
        const r3 = document.createRange(); r3.setStart(first, 0); r3.collapse(true);
        const s3 = getSelection(); s3.removeAllRanges(); s3.addRange(r3);
      }
      sync();
    };
    dlg.showModal();
  }

  function setupRte() {
    try { document.execCommand('defaultParagraphSeparator', false, 'p'); } catch (e) { /* غير مدعوم */ }
    $$('.rte').forEach(rte => {
      const bar = $('.rte-bar', rte), area = $('.rte-area', rte), path = area.dataset.docRte;
      RTE_BTNS.forEach(b => {
        if (b === '|') { const s = document.createElement('span'); s.className = 'sep'; bar.appendChild(s); return; }
        const btn = document.createElement('button');
        btn.type = 'button'; btn.innerHTML = b[1]; btn.title = b[2]; btn.setAttribute('aria-label', b[2]);
        btn.addEventListener('mousedown', e => e.preventDefault());
        if (/^(table|image|kashida|unkashida)$/.test(b[0])) btn.style.padding = '0 8px';
        btn.addEventListener('click', () => {
          if (b[0] === 'table') { openTableDialog(area, sync); return; }
          if (b[0] === 'image') { pickImage(); return; }
          if (b[0] === 'kashida' || b[0] === 'unkashida') { area.focus(); if (applyKashida(area, b[0] === 'unkashida')) sync(); else toast('ضع المؤشر داخل كلمة عربية أو حدّد نصًا أولًا'); return; }
          area.focus(); document.execCommand(b[0], false, null); sync();
        });
        bar.appendChild(btn);
      });
      const tools = document.createElement('span'); tools.className = 'tbl-tools';
      [['row', '+ صف'], ['col', '+ عمود'], ['del', 'حذف الجدول']].forEach(([k, t]) => {
        const tb = document.createElement('button'); tb.type = 'button'; tb.textContent = t;
        tb.addEventListener('mousedown', e => e.preventDefault());
        tb.addEventListener('click', () => { tableOp(area, k); sync(); });
        tools.appendChild(tb);
      });
      bar.appendChild(tools);
      document.addEventListener('selectionchange', () => { tools.classList.toggle('on', !!cellOf(area)); const s = getSelection(); if (s.rangeCount && area.contains(s.anchorNode)) lastRange = s.getRangeAt(0).cloneRange(); });
      // صورة داخل النص: اختيار، تكبير/تصغير، حذف
      let lastRange = null, selImg = null;
      const imgTools = document.createElement('span'); imgTools.className = 'tbl-tools img-tools';
      const setImgSel = im => { area.querySelectorAll('img.rte-sel').forEach(x => x.classList.remove('rte-sel')); selImg = im; if (im) im.classList.add('rte-sel'); imgTools.classList.toggle('on', !!im); };
      const resizeImg = d => {
        if (!selImg) return;
        const w = Math.max(10, Math.min(170, (parseFloat(selImg.getAttribute('data-w')) || 60) + d));
        selImg.setAttribute('data-w', w); selImg.style.width = w + 'mm'; sync();
      };
      [['−  تصغير', () => resizeImg(-10)], ['+  تكبير', () => resizeImg(10)], ['حذف الصورة', () => { if (selImg) { selImg.remove(); setImgSel(null); sync(); } }]].forEach(([t, fn]) => {
        const tb = document.createElement('button'); tb.type = 'button'; tb.textContent = t;
        tb.addEventListener('mousedown', e => e.preventDefault()); tb.addEventListener('click', fn); imgTools.appendChild(tb);
      });
      bar.appendChild(imgTools);
      area.addEventListener('click', e => setImgSel(e.target.tagName === 'IMG' ? e.target : null));
      area.addEventListener('keydown', e => { if (selImg && (e.key === 'Delete' || e.key === 'Backspace') && getSelection().isCollapsed === false) { /* حذف عادي */ } });
      const picker = document.createElement('input'); picker.type = 'file'; picker.accept = 'image/*'; picker.hidden = true; bar.appendChild(picker);
      const pickImage = () => picker.click();
      picker.addEventListener('change', async () => {
        const f = picker.files[0]; picker.value = '';
        if (!f) return;
        if (!/^image\//.test(f.type)) { toast('الملف المختار ليس صورة'); return; }
        try {
          const { url, ar } = await readImageFile(f);
          area.focus();
          const s = getSelection();
          if (lastRange) { s.removeAllRanges(); s.addRange(lastRange); }
          else { const r = document.createRange(); r.selectNodeContents(area); r.collapse(false); s.removeAllRanges(); s.addRange(r); }
          const html = '<img src="' + url + '" data-w="60" data-ar="' + ar.toFixed(3) + '" style="width:60mm;max-width:100%;height:auto">';
          if (!document.execCommand('insertHTML', false, html)) area.insertAdjacentHTML('beforeend', html);
          sync(); toast('تم إدراج الصورة — انقر عليها للتكبير أو التصغير');
        } catch (err) { toast('تعذّر قراءة الصورة'); }
      });
      const sync = () => { App.setPath(S, path, App.sanitize(area.innerHTML)); changed(); };
      area.addEventListener('input', sync);
      area.addEventListener('keydown', e => {
        if (e.key !== 'Tab') return;
        const c = cellOf(area); if (!c) return;
        e.preventDefault();
        const cells = Array.from(c.closest('table').querySelectorAll('th,td')); let i = cells.indexOf(c) + (e.shiftKey ? -1 : 1);
        if (i >= cells.length) { tableOp(area, 'row'); sync(); i = cells.indexOf(c) + 1; const all = Array.from(c.closest('table').querySelectorAll('th,td')); i = Math.min(i, all.length - 1); cells.splice(0, cells.length, ...all); }
        if (i < 0) return;
        const r = document.createRange(); r.selectNodeContents(cells[i]); r.collapse(true);
        const s2 = getSelection(); s2.removeAllRanges(); s2.addRange(r);
      });
      area.addEventListener('paste', e => {
        e.preventDefault();
        const txt = (e.clipboardData || window.clipboardData).getData('text/plain');
        const html = txt.replace(/\r/g, '').split('\n').map(l => '<p>' + (App.esc(l) || '<br>') + '</p>').join('');
        document.execCommand('insertHTML', false, html);
      });
      area.innerHTML = App.getPath(S, path) || '';
    });
  }

  /* ---------- الشعارات ---------- */
  function normalizeImage(file, keyWhite) {
    return new Promise((resolve, reject) => {
      const fr = new FileReader();
      fr.onerror = () => reject(new Error('read'));
      fr.onload = () => {
        const im = new Image();
        im.onerror = () => reject(new Error('img'));
        im.onload = () => {
          const w0 = im.naturalWidth || 300, h0 = im.naturalHeight || 300;
          const k = Math.min(1, 600 / Math.max(w0, h0));
          const c = document.createElement('canvas');
          c.width = Math.max(1, Math.round(w0 * k)); c.height = Math.max(1, Math.round(h0 * k));
          const cx = c.getContext('2d');
          cx.drawImage(im, 0, 0, c.width, c.height);
          if (keyWhite) {
            const id = cx.getImageData(0, 0, c.width, c.height), px = id.data;
            for (let i = 0; i < px.length; i += 4) {
              const mn = Math.min(px[i], px[i + 1], px[i + 2]);
              if (mn >= 240) px[i + 3] = 0;
              else if (mn >= 200) px[i + 3] = Math.round(px[i + 3] * (240 - mn) / 40);
            }
            cx.putImageData(id, 0, 0);
          }
          resolve(c.toDataURL('image/png'));
        };
        im.src = fr.result;
      };
      fr.readAsDataURL(file);
    });
  }
  function setupLogos() { /* الشعاران ثابتان: لا رفع ولا إزالة */ }

  /* ---------- التواقيع والأختام ---------- */
  function setupPeople() {
    $$('.people-card').forEach(card => {
      const key = card.dataset.person;
      $$('.img-slot', card).forEach(slot => {
        const field = slot.dataset.kind === 'sign' ? 'signImg' : 'stampImg', showF = slot.dataset.kind === 'sign' ? 'showSign' : 'showStamp';
        $('input[type=file]', slot).addEventListener('change', async e => {
          const f = e.target.files[0]; e.target.value = '';
          if (!f) return;
          if (!/^image\//.test(f.type)) { toast('الملف المختار ليس صورة'); return; }
          try { const url = await normalizeImage(f, S.people.removeBg); S.people[key][field] = url; S.people[key][field === 'signImg' ? 'signAR' : 'stampAR'] = Math.round((await probeAR(url)) * 1000) / 1000; S.people[key][showF] = true; refreshControls(); syncPos(); changed(); toast('تم الرفع'); }
          catch (err) { toast('تعذّر قراءة الصورة'); }
        });
        $('[data-remove]', slot).addEventListener('click', () => { S.people[key][field] = ''; S.people[key][field === 'signImg' ? 'signAR' : 'stampAR'] = 0; refreshControls(); syncPos(); changed(); });
      });
    });
  }

  /* ---------- الخطوط ---------- */
  function setupFonts() {
    $$('select[data-fonts]').forEach(sel => {
      sel.innerHTML = (sel.dataset.fonts === 'blank' ? '<option value="">الخط الافتراضي للمستند</option>' : '') +
        App.FONTS.map(f => '<option value="' + f[0] + '">' + f[1] + '</option>').join('');
    });
  }

  /** تحميل الخط المختار (للخطوط المضمّنة) ثم إعادة رسم المعاينة */
  function applyFonts() {
    if (!(document.fonts && document.fonts.load)) return;
    const names = [S.font, S.footer.font, S.circular.bodyFont].filter(Boolean);
    Promise.all(names.map(n => Promise.all(['400', '700'].map(w => document.fonts.load(w + " 14px '" + n + "'").catch(() => {}))))).then(() => { if (S) renderPreview(); });
  }

  /* ---------- كتاب داخلي: نسخة من نموذج كتاب الوزارة ---------- */
  function cloneLetterForm(type) {
    const src = $('.doc-form[data-form="letter"]');
    const c = src.cloneNode(true);
    c.dataset.form = type; c.hidden = true;
    c.querySelectorAll('*').forEach(el => {
      ['data-doc-bind', 'data-doc-rte', 'data-show-if'].forEach(a => { const v = el.getAttribute(a); if (v && v.indexOf('letter.') === 0) el.setAttribute(a, type + '.' + v.slice(7)); });
      if (el.dataset.rte === 'letter') el.dataset.rte = type;
    });
    return { src, c };
  }
  function buildThanksForm() {
    const { src, c } = cloneLetterForm('thanks');
    $('.panel-title', c).textContent = 'بيانات كتاب الشكر';
    // الجهات الموجه إليها: مربع كتابة بدل الخيارات الثابتة + تكرار الكتاب
    const old = $('[data-doc-bind="thanks.to0"]', c).closest('fieldset');
    const wrap = document.createElement('div');
    wrap.innerHTML =
      '<fieldset class="group"><legend>الجهة الموجه إليها الكتاب</legend>' +
      '<label class="field"><span>اكتب الجهة <small>(كل سطر يظهر في سطر مستقل)</small></span>' +
      '<textarea rows="3" data-doc-bind="thanks.toText" placeholder="مثال: عطوفة مدير التربية والتعليم الأكرم"></textarea></label></fieldset>' +
      '<fieldset class="group"><legend>تكرار الكتاب</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="thanks.copiesOn"> <span>تكرار الكتاب في صفحات جديدة لجهات أخرى</span></label>' +
      '<label class="field" data-show-if="thanks.copiesOn"><span>الجهات الإضافية <small>(اكتب كل جهة في سطر)</small></span>' +
      '<textarea rows="4" data-doc-bind="thanks.copiesText" placeholder="عطوفة مدير ... الأكرم&#10;عطوفة مدير ... الأكرم"></textarea>' +
      '<small class="hint" id="copiesCount"></small></label></fieldset>';
    old.replaceWith(...wrap.children);
    const posBox = document.createElement('fieldset'); posBox.className = 'group';
    posBox.innerHTML = '<legend>موضع صندوق «الجهة الموجه إليها»</legend>' +
      '<div class="pos-row"><b>تحريك الصندوق</b><div class="pos-pad">' +
      '<button type="button" data-nudge="thanks.toY" data-d="-1" title="للأعلى" aria-label="رفع الصندوق">▲</button>' +
      '<button type="button" data-nudge="thanks.toX" data-d="-1" title="يسار" aria-label="يسار">◀</button>' +
      '<button type="button" data-nudge="thanks.toX" data-d="1" title="يمين" aria-label="يمين">▶</button>' +
      '<button type="button" data-nudge="thanks.toY" data-d="1" title="للأسفل" aria-label="إنزال الصندوق">▼</button></div>' +
      '<label>رأسي <input type="number" min="-5" max="80" step="0.5" data-doc-bind="thanks.toY"></label>' +
      '<label>أفقي <input type="number" min="-40" max="40" step="0.5" data-doc-bind="thanks.toX"></label>' +
      '<button type="button" class="btn sm ghost" data-nudge-reset="thanks" title="إعادة ضبط">↺</button></div>' +
      '<p class="hint">الحركة بالمليمتر؛ إنزال الصندوق يزيح ما تحته، وتُطبَّق على المعاينة وWord وPDF والطباعة.</p>';
    wrap.children.length; c.querySelector('[data-doc-bind="thanks.toText"]').closest('fieldset').after(posBox);
    $('.doc-form[data-form="letter"]').after(c);
  }
  function buildInternalForm() {
    const { src, c } = cloneLetterForm('internal');
    $('.panel-title', c).textContent = 'بيانات الكتاب الداخلي';
    $('[data-recipients]', c).innerHTML = '<legend>الجهات الموجه إليها</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="internal.toDir"> <span>عطوفة مدير التربية والتعليم الأكرم</span></label>' +
      '<label class="field" data-show-if="internal.toDir"><span>اسم مدير التربية والتعليم</span><input type="text" data-doc-bind="internal.dirName" placeholder="مثال: الدكتور محمد أحمد"></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="internal.to0"> <span>معالي وزير التربية والتعليم وتنمية الموارد البشرية الأكرم</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="internal.to1"> <span>عطوفة مدير إدارة النشاطات التربوية الأكرم</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="internal.to2"> <span>عطوفة مدير أندية المعلمين والعاملين الأكرم</span></label>' +
      '<label class="field"><span>جهة أخرى <small>(كل سطر جهة)</small></span><textarea rows="2" data-doc-bind="internal.toCustom"></textarea></label>';
    src.after(c);
  }

  /* ---------- قرار التكليف والمطالبة المالية ---------- */
  const rteHtml = t => '<div class="field"><span>نص ' + (t === 'assign' ? 'القرار' : t === 'delegation' ? 'الكتاب' : 'المطالبة') + '</span><div class="rte" data-rte="' + t + '"><div class="rte-bar"></div><div class="rte-area" contenteditable="true" dir="rtl" data-doc-rte="' + t + '.body" spellcheck="true"></div></div></div>';
  const dateFields = (t, label) =>
    '<div class="row2"><label class="field"><span>' + label + '</span><input type="text" data-doc-bind="' + t + '.number" placeholder="مثال: 1/2026"></label>' +
    '<label class="field"><span>التاريخ (الموافق)</span><input type="date" data-doc-bind="' + t + '.date"></label></div>' +
    '<label class="check"><input type="checkbox" data-doc-bind="' + t + '.showHijri"> <span>إظهار التاريخ الهجري</span></label>' +
    '<label class="field" data-show-if="' + t + '.showHijri"><span>التاريخ الهجري <small>(يُحسب تلقائيًا ويمكن تعديله)</small></span><input type="text" data-doc-bind="' + t + '.hijri" placeholder="مثال: 30 جمادى الأولى 1448"></label>';
  function addForm(type, html) {
    const f = document.createElement('div');
    f.className = 'doc-form'; f.dataset.form = type; f.hidden = true; f.innerHTML = html;
    $('#sec-doc').insertBefore(f, $('#peopleBox'));
  }
  function buildAssignForm() {
    addForm('assign',
      '<h2 class="panel-title">بيانات قرار التكليف</h2>' +
      '<label class="field"><span>اسم النادي <small>(يظهر بعد كلمة «نادي»)</small></span><input type="text" data-bind="clubName" placeholder="مثال: البتراء"></label>' +
      dateFields('assign', 'رقم القرار') +
      '<fieldset class="group"><legend>المكلَّف</legend><div class="row2">' +
      '<label class="field"><span>الصفة</span><select data-doc-bind="assign.who"><option>السيدة</option><option>السيد</option><option>الآنسة</option><option>الأستاذ</option><option>الأستاذة</option></select></label>' +
      '<label class="field"><span>الاسم</span><input type="text" data-doc-bind="assign.assignee" placeholder="الاسم الكامل"></label></div></fieldset>' +
      '<fieldset class="group"><legend>تاريخ التكليف</legend>' +
      '<label class="field"><span>اختر التاريخ <small>(يدخل تلقائيًا في نص القرار)</small></span><input type="date" data-doc-bind="assign.assignDate"></label>' +
      '<label class="field"><span>أو اكتب التاريخ يدويًا <small>(اتركه فارغًا لاستخدام التاريخ المختار)</small></span><input type="text" data-doc-bind="assign.assignDateText" placeholder="مثال: 1/10/2026 أو 1 تشرين الأول 2026"></label>' +
      '<div class="inline-btns"><button type="button" class="btn sm" id="btnInsDate">إدراج التاريخ في النص عند المؤشر</button><button type="button" class="btn sm ghost" id="btnAssignToday">تاريخ اليوم</button></div>' +
      '<p class="hint">يظهر التاريخ في النص مكان الرمز <b>' + App.ASSIGN_TOKEN + '</b>؛ يمكنك نقله أو نسخه أو حذفه من النص، وتتحدّث الصفحة فورًا عند تغيير التاريخ.</p></fieldset>' +
      '<label class="field"><span>الموضوع <small>(اسم الملف عند التنزيل، ولا يُطبع)</small></span><input type="text" data-doc-bind="assign.subject" placeholder="مثال: تكليف بأعمال التنظيف"></label>' +
      rteHtml('assign') +
      '<fieldset class="group"><legend>التوقيع والاعتماد</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="assign.showHead"> <span>إظهار رئيس الهيئة الإدارية</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="assign.showManager"> <span>إظهار مدير النادي</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="assign.showFinance"> <span>إظهار صندوق «اللجنة المالية» أسفل رئيس الهيئة</span></label>' +
      '<label class="field" data-show-if="assign.showFinance"><span>اللجنة المالية <small>(اسم في كل سطر)</small></span><textarea rows="3" data-doc-bind="assign.financeText" placeholder="أسماء أعضاء اللجنة المالية"></textarea></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="assign.showClosing"> <span>إظهار عبارة «واقبلوا فائق الاحترام والتقدير»</span></label></fieldset>');
  }
  function buildClaimForm() {
    addForm('claim',
      '<h2 class="panel-title">بيانات المطالبة المالية</h2>' +
      '<label class="field"><span>اسم النادي <small>(يظهر بعد كلمة «نادي»)</small></span><input type="text" data-bind="clubName" placeholder="مثال: البتراء"></label>' +
      '<label class="field"><span>عنوان المطالبة</span><input type="text" data-doc-bind="claim.title"></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="claim.showDateRow"> <span>إظهار شريط الرقم والتاريخ أسفل الترويسة</span></label>' +
      '<div data-show-if="claim.showDateRow">' + dateFields('claim', 'رقم المطالبة') + '</div>' +
      '<label class="field"><span>الموضوع <small>(اسم الملف عند التنزيل، ولا يُطبع)</small></span><input type="text" data-doc-bind="claim.subject" placeholder="مثال: مطالبة مالية عن شهر آذار"></label>' +
      rteHtml('claim') +
      '<label class="check"><input type="checkbox" data-doc-bind="claim.showClosing"> <span>إظهار عبارة الختام</span></label>' +
      '<label class="field" data-show-if="claim.showClosing"><span>عبارة الختام</span><input type="text" data-doc-bind="claim.closing"></label>' +
      '<fieldset class="group"><legend>صاحب الاستحقاق</legend>' +
      '<div class="row2"><label class="field"><span>عنوان الحقل</span><input type="text" data-doc-bind="claim.lblOwner"></label><label class="field"><span>الاسم</span><input type="text" data-doc-bind="claim.owner" placeholder="اسم صاحب الاستحقاق"></label></div>' +
      '<label class="field"><span>عنوان حقل التوقيع</span><input type="text" data-doc-bind="claim.lblSign"></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="claim.showClaimDate"> <span>إظهار حقل التاريخ</span></label>' +
      '<div data-show-if="claim.showClaimDate"><label class="field"><span>عنوان الحقل</span><input type="text" data-doc-bind="claim.lblDate"></label>' +
      '<div class="row2"><label class="field"><span>التاريخ</span><input type="date" data-doc-bind="claim.claimDate"></label><label class="field"><span>أو يدويًا</span><input type="text" data-doc-bind="claim.claimDateText"></label></div></div></fieldset>' +
      '<fieldset class="group"><legend>اعتماد إضافي</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="claim.showHead"> <span>إظهار رئيس الهيئة الإدارية</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="claim.showManager"> <span>إظهار مدير النادي</span></label></fieldset>');
  }

  /* ---------- الوفد المشارك ---------- */
  function dgFieldsFor(kind) {
    const has = id => App.dgCols(S, kind).some(c => c[0] === id), d = S.delegation, out = [];
    if (has('name')) out.push(['name', 'الاسم', 'wide']);
    if (has('ministryNo')) out.push(['ministryNo', 'الرقم الوزاري']);
    if (has('memberNo')) out.push(['memberNo', 'رقم العضوية']);
    if (has('nationalNo')) out.push(['nationalNo', 'الرقم الوطني']);
    if (has('joinDate') || has('duration')) out.push(['joinDate', 'تاريخ الانتساب', '', 'date']);
    if (has('role')) out.push(['role', 'الصفة', '', 'role']);
    d[kind === 'admins' ? 'customAdmins' : 'customPlayers'].forEach(c => { if (has(c.id)) out.push(['x:' + c.id, c.label, '', 'x']); });
    return out;
  }
  const DG_TITLE = { admins: 'المرافقون الإداريون', players: 'اللاعبون المشاركون' };
  let dgSig = '';
  const dgSigOf = () => { const d = S.delegation; return JSON.stringify([d.admins, d.players, d.colsAdmins, d.colsPlayers, d.customAdmins, d.customPlayers]); };
  function dgEditorHTML(kind) {
    const rows = S.delegation[kind].map((r, i) => {
      const f = dgFieldsFor(kind).map(([k, lbl, cls, type]) => {
        const id = ' data-dg="' + kind + '.' + i + '.' + k + '"';
        const val = type === 'x' ? ((r.x || {})[k.slice(2)] || '') : r[k];
        const inp = type === 'role'
          ? '<select' + id + '><option value="">— الصفة —</option>' + App.DG_ROLES.map(x => '<option' + (r[k] === x ? ' selected' : '') + '>' + x + '</option>').join('') + '</select>'
          : '<input type="' + (type === 'date' ? 'date' : 'text') + '"' + id + ' value="' + App.esc(val) + '" placeholder="' + lbl + '" title="' + lbl + '"' + (type === 'date' ? '' : '') + '>';
        let extra = '';
        if (type === 'role') {
          const on = /^(لاعب|لاعبة)$/.test(r.role);
          extra = '<label class="wide dg-game"' + (on ? '' : ' hidden') + '><small>اللعبة <span>(بعد كلمة ' + (r.role || 'لاعب') + ')</span></small><input type="text" data-dg="' + kind + '.' + i + '.game" value="' + App.esc(r.game || '') + '" placeholder="اسم اللعبة، مثال: كرة قدم (اتركه فارغًا ليظهر «' + (r.role || 'لاعب') + '» فقط)"></label>';
        }
        return '<label class="' + (cls || '') + '"><small>' + lbl + '</small>' + inp + '</label>' + extra;
      }).join('');
      return '<div class="dg-row"><div class="dg-rh"><b>' + (i + 1) + '</b>' + (kind === 'players' ? '<span class="dg-dur" data-dg-dur="' + i + '"></span>' : '<span class="dg-dur"></span>') +
        '<button type="button" class="btn sm ghost" data-dg-act="del" data-k="' + kind + '" data-i="' + i + '" title="حذف الصف" aria-label="حذف الصف">✕ حذف</button></div><div class="dg-fields">' + f + '</div></div>';
    }).join('');
    return rows || '<p class="hint">لا توجد صفوف. استخدم «+ صف» لإضافة مشارك.</p>';
  }
  function dgRender() {
    ['admins', 'players'].forEach(k => {
      const el = $('#dg-' + k); if (el) el.innerHTML = dgEditorHTML(k);
      const cols = App.dgCols(S, k), ch = $('#dgc-' + k), sel = $('#dgcSel-' + k);
      if (ch) ch.innerHTML = cols.map(c => '<span class="dg-chip">' + App.esc(c[1]) + ' <button type="button" data-dg-act="delcol" data-k="' + k + '" data-id="' + c[0] + '" title="إزالة العمود" aria-label="إزالة عمود ' + App.esc(c[1]) + '">✕</button></span>').join('');
      if (sel) sel.innerHTML = '<option value="">عمود جديد (اكتب اسمه)</option>' + App.DG_COLS[k].filter(c => !cols.some(x => x[0] === c[0])).map(c => '<option value="' + c[0] + '">استرجاع: ' + c[1] + '</option>').join('');
    });
    dgSig = dgSigOf();
  }
  function dgSync() {
    if (!$('#dg-players')) return;
    if (dgSigOf() !== dgSig) dgRender();
    const d = S.delegation;
    S.delegation.players.forEach((r, i) => { const el = $('[data-dg-dur="' + i + '"]'); if (el) { const t = App.durationText(r.joinDate, d.date, d.showDays); el.textContent = t ? 'مدة الانتساب: ' + t : (r.joinDate ? 'تاريخ الانتساب بعد تاريخ الكتاب' : ''); } });
    $$('[data-dg-kind]').forEach(b => { const on = b.dataset.dgKind === d.kind; b.classList.toggle('primary', on); b.setAttribute('aria-pressed', on); });
    const lbl = $('#dgKindLbl'); if (lbl) lbl.textContent = App.DG_KINDS[d.kind][1];
    const at = $('#dgAuto'); if (at) at.textContent = App.dgIntro(S);
  }
  function buildDelegationForm() {
    const ed = k => '<fieldset class="group"><legend>' + DG_TITLE[k] + '</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.' + (k === 'admins' ? 'showAdmins' : 'showPlayers') + '"> <span>إظهار هذا الجدول في الكتاب <small>(ألغِ التأشير لإزالته)</small></span></label><div class="dg-tools">' +
      '<button type="button" class="btn sm primary" data-dg-act="add" data-k="' + k + '">+ صف</button>' +
      '<button type="button" class="btn sm" data-dg-act="add5" data-k="' + k + '">+ 5 صفوف</button>' +
      '<button type="button" class="btn sm ghost" data-dg-act="clear" data-k="' + k + '">حذف كل الصفوف</button></div>' +
      '<details class="dg-paste"><summary>الأعمدة: إضافة / إزالة</summary><div class="dg-chips" id="dgc-' + k + '"></div>' +
      '<div class="dg-addcol"><select id="dgcSel-' + k + '"></select><input type="text" id="dgcName-' + k + '" placeholder="اسم العمود الجديد"><button type="button" class="btn sm" data-dg-act="addcol" data-k="' + k + '">+ إضافة عمود</button></div></details>' +
      '<details class="dg-paste"><summary>إضافة أسماء دفعة واحدة</summary><textarea rows="4" id="dgPaste-' + k + '" placeholder="اكتب أو الصق الأسماء، اسم في كل سطر"></textarea>' +
      '<button type="button" class="btn sm" data-dg-act="paste" data-k="' + k + '">إضافة الأسماء كصفوف</button></details>' +
      '<div class="dg-ed" id="dg-' + k + '"></div></fieldset>';
    addForm('delegation',
      '<h2 class="panel-title">بيانات الوفد المشارك</h2>' +
      '<label class="field"><span>اسم النادي <small>(يظهر بعد كلمة «نادي» وفي نص الكتاب تلقائيًا)</small></span><input type="text" data-bind="clubName" placeholder="مثال: البتراء"></label>' +
      dateFields('delegation', 'رقم الكتاب') +
      '<fieldset class="group"><legend>الجهات الموجه إليها <small>(اختياري)</small></legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.to0"> <span>معالي وزير التربية والتعليم وتنمية الموارد البشرية الأكرم</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.to1"> <span>عطوفة مدير إدارة النشاطات التربوية الأكرم</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.to2"> <span>عطوفة مدير أندية المعلمين والعاملين الأكرم</span></label>' +
      '<label class="field"><span>جهة أخرى <small>(كل سطر جهة)</small></span><textarea rows="2" data-doc-bind="delegation.toCustom"></textarea></label></fieldset>' +
      '<label class="field"><span>الموضوع</span><input type="text" data-doc-bind="delegation.subject" placeholder="مثال: كشف أسماء الوفد المشارك (يُستخدم أيضًا اسمًا للملف)"></label>' +
      '<fieldset class="group"><legend>موضوع المشاركة</legend><div class="dg-seg">' +
      '<button type="button" class="btn sm" data-dg-kind="tournament">المشارك في بطولة</button>' +
      '<button type="button" class="btn sm" data-dg-kind="qualifiers">المشارك في تصفيات</button>' +
      '<button type="button" class="btn sm" data-dg-kind="game">المشارك في لعبة</button></div>' +
      '<label class="field"><span id="dgKindLbl">اسم التصفيات</span><input type="text" data-doc-bind="delegation.kindName"></label>' +
      '<label class="field"><span>مكان الإقامة</span><input type="text" data-doc-bind="delegation.venue" placeholder="مثال: مدينة العقبة"></label></fieldset>' +
      '<fieldset class="group"><legend>النص</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.showIntro"> <span>إظهار الفقرة تحت «الموضوع»</span></label>' +
      '<p class="hint">الفقرة التلقائية: <b id="dgAuto"></b></p>' +
      '<label class="field"><span>نص مستقل <small>(إن كتبت فيه يظهر مكان الفقرة التلقائية)</small></span><textarea rows="3" data-doc-bind="delegation.introText" placeholder="اتركه فارغًا لاستخدام الفقرة التلقائية"></textarea></label>' +
      rteHtml('delegation') +
      '<div class="inline-btns"><button type="button" class="btn sm ghost" id="dgToBody">نقل الفقرة التلقائية إلى نص الكتاب للتعديل عليها</button></div></fieldset>' +
      '<fieldset class="group"><legend>عناوين الجداول</legend>' +
      '<label class="field"><span>عنوان الجدول الأول</span><input type="text" data-doc-bind="delegation.titleAdmin"></label>' +
      '<label class="field"><span>عنوان الجدول الثاني</span><input type="text" data-doc-bind="delegation.titlePlayers"></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.showDays"> <span>إظهار الأيام في مدة الانتساب</span></label></fieldset>' +
      ed('admins') + ed('players') +
      '<fieldset class="group"><legend>الختام والتوقيع</legend>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.showClosing"> <span>إظهار عبارة «واقبلوا فائق الاحترام والتقدير»</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.showHead"> <span>إظهار رئيس الهيئة الإدارية</span></label>' +
      '<label class="check"><input type="checkbox" data-doc-bind="delegation.showManager"> <span>إظهار مدير النادي</span></label></fieldset>');
    const form = $('.doc-form[data-form="delegation"]');
    const upd = e => {
      const el = e.target.closest && e.target.closest('[data-dg]'); if (!el) return;
      const [k, i, f] = el.dataset.dg.split('.');
      const row = S.delegation[k][+i]; if (!row) return;
      if (f.startsWith('x:')) { row.x = row.x || {}; row.x[f.slice(2)] = el.value; } else row[f] = el.value;
      dgSig = dgSigOf();
      if (f === 'role') {
        const g = el.closest('.dg-fields').querySelector('.dg-game');
        if (g) { g.hidden = !/^(لاعب|لاعبة)$/.test(row.role); const sp = g.querySelector('small span'); if (sp) sp.textContent = '(بعد كلمة ' + (row.role || 'لاعب') + ')'; g.querySelector('input').placeholder = 'اسم اللعبة، مثال: كرة قدم (اتركه فارغًا ليظهر «' + (row.role || 'لاعب') + '» فقط)'; }
      }
      dgSync(); changed();
    };
    form.addEventListener('input', upd); form.addEventListener('change', upd);
    form.addEventListener('click', e => {
      const kb = e.target.closest('[data-dg-kind]');
      if (kb) { S.delegation.kind = kb.dataset.dgKind; dgSync(); changed(); return; }
      if (e.target.closest('#dgToBody')) {
        const area = $('[data-doc-rte="delegation.body"]');
        area.insertAdjacentHTML('beforeend', '<p>' + App.esc(App.dgIntro(S)) + '</p>');
        S.delegation.body = App.sanitize(area.innerHTML); S.delegation.showIntro = false; S.delegation.introText = '';
        refreshControls(); changed(); return;
      }
      const b = e.target.closest('[data-dg-act]'); if (!b) return;
      const k = b.dataset.k, mk = k === 'admins' ? () => App.dgAdmin('إداري') : () => App.dgPlayer('لاعب'), rows = S.delegation[k];
      const act = b.dataset.dgAct, d = S.delegation, ck = k === 'admins' ? 'colsAdmins' : 'colsPlayers', uk = k === 'admins' ? 'customAdmins' : 'customPlayers';
      if (act === 'delcol') {
        if (d[ck].length <= 1) { toast('يجب أن يبقى عمود واحد على الأقل'); return; }
        const id = b.dataset.id; d[ck] = d[ck].filter(x => x !== id);
        d[uk] = d[uk].filter(c => c.id !== id); rows.forEach(r => { if (r.x) delete r.x[id]; });
        dgRender(); dgSync(); changed(); return;
      }
      if (act === 'addcol') {
        const sel = $('#dgcSel-' + k).value, inp = $('#dgcName-' + k);
        if (sel) { const keep = new Set(d[ck].concat([sel])); d[ck] = App.DG_STD[k].filter(x => keep.has(x)).concat(d[ck].filter(x => !App.DG_STD[k].includes(x))); }
        else {
          const name = inp.value.trim(); if (!name) { toast('اكتب اسم العمود الجديد'); inp.focus(); return; }
          const id = 'c' + Date.now().toString(36) + Math.random().toString(36).slice(2, 5);
          d[uk].push({ id, label: name }); d[ck].push(id);
        }
        inp.value = ''; dgRender(); dgSync(); changed(); return;
      }
      if (act === 'add') rows.push(mk());
      else if (act === 'add5') for (let n = 0; n < 5; n++) rows.push(mk());
      else if (act === 'del') rows.splice(+b.dataset.i, 1);
      else if (act === 'clear') { if (rows.length && !confirm('حذف كل صفوف هذا الجدول؟')) return; rows.length = 0; }
      else if (act === 'paste') {
        const ta = $('#dgPaste-' + k), names = ta.value.split('\n').map(x => x.trim()).filter(Boolean);
        if (!names.length) { toast('اكتب اسمًا واحدًا على الأقل'); return; }
        names.forEach(n => { const r = mk(); r.name = n; rows.push(r); }); ta.value = '';
      }
      dgRender(); dgSync(); changed();
      if (act === 'add' || act === 'add5') { const el = $('[data-dg="' + k + '.' + (rows.length - (act === 'add' ? 1 : 5)) + '.name"]'); el && el.focus(); }
    });
    dgRender();
  }

  function setupAssign() {
    $('#btnInsDate').addEventListener('mousedown', e => e.preventDefault());
    $('#btnInsDate').addEventListener('click', () => {
      const area = $('[data-doc-rte="assign.body"]');
      const sel = getSelection();
      if (!(sel.rangeCount && area.contains(sel.anchorNode))) { area.focus(); const r = document.createRange(); r.selectNodeContents(area); r.collapse(false); sel.removeAllRanges(); sel.addRange(r); }
      document.execCommand('insertText', false, App.ASSIGN_TOKEN);
      S.assign.body = App.sanitize(area.innerHTML); changed();
    });
    $('#btnAssignToday').addEventListener('click', () => {
      S.assign.assignDate = App.todayISO(); S.assign.assignDateText = ''; refreshControls(); changed();
    });
  }

  /* ---------- مواضع الختم والتوقيع والأسماء ---------- */
  const POS_ELS = [
    ['head', 'sign', 'توقيع رئيس الهيئة'], ['head', 'stamp', 'ختم رئيس الهيئة'], ['head', 'name', 'الاسم + الصفة (رئيس الهيئة)'],
    ['manager', 'sign', 'توقيع المدير'], ['manager', 'stamp', 'ختم المدير'], ['manager', 'name', 'الاسم + الصفة (المدير)']
  ];
  let geoSnap = null;
  function snapPeople() {
    const P = S.people;
    return JSON.stringify({ a: [P.signSize, P.stampSize, P.stampPos, P.mode, P.gapName, P.gapImg, P.protect], h: P.head.pos, m: P.manager.pos });
  }
  function restorePeople(js) {
    const o = JSON.parse(js), P = S.people;
    [P.signSize, P.stampSize, P.stampPos, P.mode, P.gapName, P.gapImg, P.protect] = o.a;
    P.head.pos = o.h; P.manager.pos = o.m;
  }
  /** فحص مواضع الأختام/التواقيع/الأسماء بعد تغيير؛ يتراجع عنه إن سبّب تداخلًا أو خروجًا عن الصفحة */
  function issuesCount() {
    const both = { showHead: true, showManager: true };
    const cur = S[S.activeDoc] || both;
    const types = ['assign'];
    let n = App.signIssues(S, both, 'letter').length;
    if (cur !== both) n += App.signIssues(S, cur, S.activeDoc).length;
    if (S.activeDoc === 'assign') n += App.signIssues(S, Object.assign({}, S.assign), 'assign').length;
    return { n, msgs: App.signIssues(S, both, 'letter').concat(cur !== both ? App.signIssues(S, cur, S.activeDoc) : []) };
  }
  function acceptGeo(skipCheck) {
    if (skipCheck) { geoSnap = snapPeople(); return true; }
    const after = issuesCount();
    const prev = geoSnap; const keep = snapPeople();
    if (prev) {
      restorePeople(prev); const before = issuesCount(); restorePeople(keep);
      if (S.people.protect && after.n > before.n) {
        restorePeople(prev); toast(after.msgs[0] || 'لا يمكن تنفيذ هذه الحركة'); return false;
      }
    }
    geoSnap = keep; return true;
  }
  function setupPositions() {
    const list = $('#posList');
    list.innerHTML = POS_ELS.map(([who, el, label]) =>
      '<div class="pos-row" data-pos="' + who + '.' + el + '"><b>' + label + '</b>' +
      '<div class="pos-pad">' +
      '<button type="button" data-mv="0,-1" title="أعلى" aria-label="رفع ' + label + '">▲</button>' +
      '<button type="button" data-mv="-1,0" title="يسار" aria-label="تحريك ' + label + ' لليسار">◀</button>' +
      '<button type="button" data-mv="1,0" title="يمين" aria-label="تحريك ' + label + ' لليمين">▶</button>' +
      '<button type="button" data-mv="0,1" title="أسفل" aria-label="إنزال ' + label + '">▼</button></div>' +
      '<label>س <input type="number" step="0.5" min="-80" max="80" data-px="x" aria-label="الإزاحة الأفقية ' + label + '"></label>' +
      '<label>ص <input type="number" step="0.5" min="-80" max="80" data-px="y" aria-label="الإزاحة الرأسية ' + label + '"></label>' +
      '<button type="button" class="btn sm ghost" data-rst title="إعادة ضبط">↺</button></div>').join('');
    const q = row => { const [who, el] = row.dataset.pos.split('.'); return S.people[who].pos[el]; };
    const move = (row, dx, dy) => {
      const o = q(row), old = [o.x, o.y];
      o.x = Math.round((o.x + dx) * 10) / 10; o.y = Math.round((o.y + dy) * 10) / 10;
      if (!acceptGeo()) { o.x = old[0]; o.y = old[1]; }
      syncPos(); changed();
    };
    const step = () => +$('#posStep').value || 1;
    let timer = null, rep = null;
    const stop = () => { clearTimeout(timer); clearInterval(rep); };
    list.addEventListener('pointerdown', e => {
      const b = e.target.closest('[data-mv]'); if (!b) return;
      e.preventDefault();
      const [dx, dy] = b.dataset.mv.split(',').map(Number), row = b.closest('.pos-row');
      move(row, dx * step(), dy * step());
      timer = setTimeout(() => { rep = setInterval(() => move(row, dx * step(), dy * step()), 90); }, 400);
    });
    ['pointerup', 'pointerleave', 'pointercancel'].forEach(ev => list.addEventListener(ev, stop));
    list.addEventListener('keydown', e => { // تفعيل الأزرار بلوحة المفاتيح
      const b = e.target.closest('[data-mv]');
      if (b && (e.key === 'Enter' || e.key === ' ')) { e.preventDefault(); const [dx, dy] = b.dataset.mv.split(',').map(Number); move(b.closest('.pos-row'), dx * step(), dy * step()); }
    });
    list.addEventListener('change', e => {
      const inp = e.target.closest('[data-px]'); if (!inp) return;
      const row = inp.closest('.pos-row'), o = q(row), old = [o.x, o.y];
      const v = Math.max(-80, Math.min(80, +inp.value || 0));
      o[inp.dataset.px] = v;
      if (!acceptGeo()) { o.x = old[0]; o.y = old[1]; }
      syncPos(); changed();
    });
    list.addEventListener('click', e => {
      const b = e.target.closest('[data-rst]'); if (!b) return;
      const o = q(b.closest('.pos-row')); o.x = 0; o.y = 0; geoSnap = snapPeople(); syncPos(); changed();
    });
    $('#posReset').addEventListener('click', () => {
      ['head', 'manager'].forEach(k => { S.people[k].pos = { sign: { x: 0, y: 0 }, stamp: { x: 0, y: 0 }, name: { x: 0, y: 0 } }; });
      geoSnap = snapPeople(); syncPos(); changed(); toast('أُعيد ضبط جميع المواضع');
    });
    geoSnap = snapPeople();
  }
  /* ---------- سحب مباشر على المعاينة + تحديد بالنقر وتحريك بالأسهم ---------- */
  let sel = null; // { key, el }
  function markSel() {
    $$('#sheet .sg-mv').forEach(e => e.classList.toggle('sg-sel', !!sel && e.dataset.key === sel.key && e.dataset.el === sel.el));
  }
  function moveSel(dx, dy) {
    if (!sel) return;
    const o = S.people[sel.key].pos[sel.el], old = [o.x, o.y];
    o.x = Math.max(-80, Math.min(80, Math.round((o.x + dx) * 10) / 10)); o.y = Math.max(-80, Math.min(80, Math.round((o.y + dy) * 10) / 10));
    if (!acceptGeo()) { o.x = old[0]; o.y = old[1]; }
    syncPos(); changed();
  }
  function setupDrag() {
    const scroll = $('#previewScroll');
    let d = null;
    scroll.addEventListener('pointerdown', e => {
      const t = e.target.closest('.sg-mv');
      if (!t || !$('#sheet').contains(t)) { if (!e.target.closest('#sheet')) { sel = null; markSel(); } return; }
      e.preventDefault();
      const key = t.dataset.key, el = t.dataset.el, box = t.closest('.sign-box');
      sel = { key, el }; markSel();
      const o = S.people[key].pos[el];
      d = { key, el, x0: e.clientX, y0: e.clientY, ox: o.x, oy: o.y, k: box.getBoundingClientRect().width / App.SG.W, els: $$('.sg-mv[data-key="' + key + '"][data-el="' + el + '"]', box), moved: false, dx: 0, dy: 0 };
      d.els.forEach(n => n.classList.add('sg-drag'));
      try { t.setPointerCapture(e.pointerId); } catch (x) { /* لا شيء */ }
    });
    scroll.addEventListener('pointermove', e => {
      if (!d) return;
      let dx = (e.clientX - d.x0) / d.k, dy = (e.clientY - d.y0) / d.k;
      if (!d.moved && Math.abs(dx) + Math.abs(dy) < 0.6) return;
      d.moved = true;
      dx = Math.round(dx * 10) / 10; dy = Math.round(dy * 10) / 10;
      d.dx = dx; d.dy = dy;
      d.els.forEach(n => { n.style.transform = 'translate(' + dx + 'mm,' + dy + 'mm)'; });
    });
    const end = () => {
      if (!d) return;
      const c = d; d = null;
      c.els.forEach(n => { n.style.transform = ''; n.classList.remove('sg-drag'); });
      if (!c.moved) return;
      const o = S.people[c.key].pos[c.el];
      o.x = Math.max(-80, Math.min(80, Math.round((c.ox + c.dx) * 10) / 10)); o.y = Math.max(-80, Math.min(80, Math.round((c.oy + c.dy) * 10) / 10));
      if (!acceptGeo()) { o.x = c.ox; o.y = c.oy; }
      syncPos(); changed();
    };
    scroll.addEventListener('pointerup', end);
    scroll.addEventListener('pointercancel', end);
    document.addEventListener('keydown', e => {
      if (!sel) return;
      const a = e.target, tag = a && a.tagName;
      if (a && (a.isContentEditable || tag === 'INPUT' || tag === 'TEXTAREA' || tag === 'SELECT')) return;
      const st = (+$('#posStep').value || 1) * (e.shiftKey ? 5 : 1);
      const m = { ArrowLeft: [-st, 0], ArrowRight: [st, 0], ArrowUp: [0, -st], ArrowDown: [0, st] }[e.key];
      if (m) { e.preventDefault(); moveSel(m[0], m[1]); }
      else if (e.key === 'Escape') { sel = null; markSel(); }
    });
  }
  function syncPos() {
    $$('#posList .pos-row').forEach(row => {
      const [who, el] = row.dataset.pos.split('.'), o = S.people[who].pos[el];
      const p = S.people[who];
      $('[data-px=x]', row).value = o.x; $('[data-px=y]', row).value = o.y;
      const has = el === 'name' || (el === 'sign' ? p.signImg : p.stampImg);
      row.classList.toggle('dim', !has);
    });
  }

  /* نسبة أبعاد التوقيع والختم (لحساب المواضع بدقة) */
  function probeAR(url) {
    return new Promise(res => { const im = new Image(); im.onload = () => res((im.naturalWidth || 1) / (im.naturalHeight || 1)); im.onerror = () => res(0); im.src = url; });
  }
  async function fillAR() {
    let ch = false;
    for (const k of ['head', 'manager']) {
      const p = S.people[k];
      if (p.signImg && !p.signAR) { p.signAR = Math.round((await probeAR(p.signImg)) * 1000) / 1000; ch = ch || !!p.signAR; }
      if (p.stampImg && !p.stampAR) { p.stampAR = Math.round((await probeAR(p.stampImg)) * 1000) / 1000; ch = ch || !!p.stampAR; }
    }
    if (ch) changed();
  }

  /* ---------- التبويبات ---------- */
  function setDoc(type) {
    S.activeDoc = type;
    App.attach.renderList();
    $$('.doc-tab').forEach(b => { const on = b.dataset.doc === type; b.classList.toggle('active', on); b.setAttribute('aria-selected', on); });
    $$('.doc-form').forEach(f => { f.hidden = f.dataset.form !== type; });
    changed();
  }
  function setSection(sec) {
    $$('.side-tab').forEach(b => b.classList.toggle('active', b.dataset.sec === sec));
    $$('.panel').forEach(p => { p.hidden = p.id !== 'sec-' + sec; p.classList.toggle('active', !p.hidden); });
  }
  /* ---------- التصدير ---------- */
  const download = (blob, name) => App.saveFile(blob, name);

  /** تذكير الاعتماد قبل التنزيل أو الإرسال أو الطباعة */
  function approve(act) {
    const dlg = $('#apprDlg');
    $('#apAct').textContent = act;
    return new Promise(res => {
      const done = () => { dlg.removeEventListener('close', done); res(dlg.returnValue === 'yes'); };
      dlg.returnValue = 'no';
      dlg.addEventListener('close', done);
      dlg.showModal();
      setTimeout(() => { const y = $('#apYes'); y && y.focus(); }, 30);
    });
  }
  App.approve = approve;

  function setupActions() {
    const exportAs = fmt => approve('التنزيل').then(ok => ok && App.run(fmt === 'pdf' ? 'جارٍ تجهيز ملف PDF…' : 'جارٍ تجهيز ملف Word…', async () => {
      const f = await App.getFile(S, S.activeDoc, fmt);
      if (await App.saveFile(f.blob, f.name)) toast('تم تنزيل الملف: ' + f.name);
    }));
    $('#btnWord').addEventListener('click', () => exportAs('docx'));
    $('#btnPdf').addEventListener('click', () => exportAs('pdf'));
    $('#btnPrint').addEventListener('click', async () => { if (await approve('الطباعة')) App.printDoc(S, S.activeDoc); });

    $('#btnClearDoc').addEventListener('click', () => {
      if (!confirm('سيتم مسح بيانات هذا المستند فقط (الإعدادات والشعارات تبقى). متابعة؟')) return;
      const d = App.defaults()[S.activeDoc];
      if (d.date && d.hijri !== undefined) d.hijri = App.formatHijri(d.date);
      S[S.activeDoc] = d;
      $$('.rte-area').forEach(a => { a.innerHTML = App.getPath(S, a.dataset.docRte) || ''; });
      refreshControls(); changed();
    });

    $('#btnResetAll').addEventListener('click', async () => {
      if (!confirm('سيتم حذف جميع البيانات والشعارات والتواقيع والمرفقات المحفوظة في هذا المتصفح. متابعة؟')) return;
      await App.wipe();
      location.reload();
    });
    $('#btnExportSettings').addEventListener('click', () => {
      const copy = JSON.parse(JSON.stringify(S));
      download(new Blob([JSON.stringify(copy)], { type: 'application/json' }), 'teachers-clubs-settings.json');
    });
    $('#fileImportSettings').addEventListener('change', async e => {
      const f = e.target.files[0]; e.target.value = '';
      if (!f) return;
      try {
        const obj = JSON.parse(await f.text());
        if (!obj || typeof obj !== 'object' || !obj.v) throw new Error('bad');
        S = App.build(obj);
        $$('.rte-area').forEach(a => { a.innerHTML = App.getPath(S, a.dataset.docRte) || ''; });
        setDoc(S.activeDoc); refreshControls(); changed(); toast('تمت استعادة الإعدادات');
      } catch (err) { toast('ملف الإعدادات غير صالح'); }
    });
  }

  /* ---------- التكبير ---------- */
  function setupZoom() {
    $('#zoomIn').addEventListener('click', () => { fitMode = false; zoom = Math.min(1.6, zoom + 0.1); applyZoom(false); });
    $('#zoomOut').addEventListener('click', () => { fitMode = false; zoom = Math.max(0.3, zoom - 0.1); applyZoom(false); });
    $('#zoomFit').addEventListener('click', () => { fitMode = true; applyZoom(false); });
    window.addEventListener('resize', () => { if (fitMode) applyZoom(false); });
  }

  /* ---------- الوضع الليلي ---------- */
  function applyTheme() {
    const t = S.theme || (window.matchMedia && matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
    document.documentElement.dataset.theme = t;
  }
  function setupTheme() {
    $('#btnTheme').addEventListener('click', () => {
      S.theme = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
      applyTheme(); changed();
    });
  }

  /* ---------- البدء ---------- */
  function init() {
    buildInternalForm(); buildThanksForm(); buildAssignForm(); buildClaimForm(); buildDelegationForm(); setupAssign();
    bindAll(); setupRte(); setupLogos(); setupPeople(); setupFonts(); setupPositions(); setupDrag(); syncPos(); fillAR(); setupActions(); setupZoom(); setupTheme(); applyTheme();
    App.attach.init(() => S.activeDoc, changed); App.initSend(() => S); App.drive.init(() => S);
    $$('.doc-tab').forEach(b => b.addEventListener('click', () => setDoc(b.dataset.doc)));
    $$('.side-tab').forEach(b => b.addEventListener('click', () => setSection(b.dataset.sec)));
    applyFonts();
    refreshControls();
    setDoc(S.activeDoc);
    if (document.fonts && document.fonts.ready) document.fonts.ready.then(() => { renderPreview(); });
    window.__ready = true;
  }
  App.loadAsync().then(r => { S = r.state; App.attach.setAll(r.attachments); init(); });
  window.__state = () => S; // للاختبار
})();
