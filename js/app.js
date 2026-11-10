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
      $('[data-remove]', card).hidden = !L.data;
      $('label.btn', card).firstChild.nodeValue = L.data ? 'استبدال الشعار' : 'رفع شعار';
    });
  }

  function onInput(el, path) {
    const v = readEl(el);
    App.setPath(S, path, v);
    const m = /^(letter|internal|thanks|circular)\.(date|hijri|showHijri)$/.exec(path);
    if (m) {
      const d = S[m[1]];
      if (m[2] === 'date') { if (!d.hijriManual) d.hijri = App.formatHijri(v); }
      else if (m[2] === 'hijri') { d.hijriManual = v.trim() !== ''; if (!d.hijriManual) d.hijri = App.formatHijri(d.date); }
      else if (v && !d.hijri) d.hijri = App.formatHijri(d.date);
    }
    if (path === 'font' || path === 'footer.font') applyFonts();
    refreshControls(el);
    changed();
  }

  function bindAll() {
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
    ['table', '▦ جدول', 'إدراج جدول']
  ];
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
        if (b[0] === 'table') btn.style.padding = '0 8px';
        btn.addEventListener('click', () => {
          if (b[0] === 'table') { openTableDialog(area, sync); return; }
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
      document.addEventListener('selectionchange', () => { tools.classList.toggle('on', !!cellOf(area)); });
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
  function setupLogos() {
    $$('.logo-card').forEach(card => {
      const k = card.dataset.logo;
      $('input[type=file]', card).addEventListener('change', async e => {
        const f = e.target.files[0]; e.target.value = '';
        if (!f) return;
        if (!/^image\//.test(f.type)) { toast('الملف المختار ليس صورة'); return; }
        try {
          S.logos[k].data = await normalizeImage(f);
          S.logos[k].show = true;
          refreshControls(); changed(); toast('تم رفع الشعار');
        } catch (err) { toast('تعذّر قراءة الصورة، جرّب ملفًا آخر'); }
      });
      $('[data-remove]', card).addEventListener('click', () => {
        S.logos[k].data = ''; refreshControls(); changed();
      });
    });
  }

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
          try { S.people[key][field] = await normalizeImage(f, S.people.removeBg); S.people[key][showF] = true; refreshControls(); changed(); toast('تم الرفع'); }
          catch (err) { toast('تعذّر قراءة الصورة'); }
        });
        $('[data-remove]', slot).addEventListener('click', () => { S.people[key][field] = ''; refreshControls(); changed(); });
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
  function download(blob, name) {
    const a = document.createElement('a');
    a.href = URL.createObjectURL(blob); a.download = name;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 4000);
  }
  App.download = download;
  async function busy(btn, fn) {
    const old = btn.disabled; btn.disabled = true;
    try { await fn(); } catch (e) { console.error(e); toast('حدث خطأ أثناء التصدير: ' + (e.message || e)); }
    finally { btn.disabled = old; }
  }

  function setupActions() {
    $('#btnWord').addEventListener('click', e => busy(e.currentTarget, async () => {
      const blob = await App.buildDocx(S, S.activeDoc);
      download(blob, App.fileName(S, S.activeDoc) + '.docx'); toast('تم تنزيل ملف Word');
    }));
    $('#btnPdf').addEventListener('click', e => busy(e.currentTarget, async () => {
      toast('جارٍ تجهيز ملف PDF…');
      const blob = await App.buildPdf(S, S.activeDoc);
      download(blob, App.fileName(S, S.activeDoc) + '.pdf'); toast('تم تنزيل ملف PDF');
    }));
    $('#btnPrint').addEventListener('click', () => App.printDoc(S, S.activeDoc));

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
    buildInternalForm(); buildThanksForm();
    bindAll(); setupRte(); setupLogos(); setupPeople(); setupFonts(); setupActions(); setupZoom(); setupTheme(); applyTheme();
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
