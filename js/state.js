/* ===== الحالة العامة، الحفظ المحلي، أدوات التاريخ ===== */
(function (g) {
  'use strict';
  const KEY = 'teachers-clubs-editor:v2';

  const MONTHS = ['يناير', 'فبراير', 'مارس', 'أبريل', 'مايو', 'يونيو', 'يوليو', 'أغسطس', 'سبتمبر', 'أكتوبر', 'نوفمبر', 'ديسمبر'];
  const DAYS = ['الأحد', 'الاثنين', 'الثلاثاء', 'الأربعاء', 'الخميس', 'الجمعة', 'السبت'];

  function todayISO() {
    const d = new Date();
    const p = n => String(n).padStart(2, '0');
    return d.getFullYear() + '-' + p(d.getMonth() + 1) + '-' + p(d.getDate());
  }

  function parseISO(iso) {
    const m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(iso || '');
    if (!m) return null;
    return new Date(+m[1], +m[2] - 1, +m[3], 12);
  }

  /** مثال: الاثنين، 09 نوفمبر 2026 */
  function formatGregorian(iso) {
    const d = parseISO(iso);
    if (!d) return '';
    return DAYS[d.getDay()] + '، ' + String(d.getDate()).padStart(2, '0') + ' ' + MONTHS[d.getMonth()] + ' ' + d.getFullYear();
  }

  /** التاريخ الهجري (أم القرى) مثال: 30 جمادى الأولى 1448 */
  function formatHijri(iso) {
    const d = parseISO(iso);
    if (!d) return '';
    try {
      const parts = new Intl.DateTimeFormat('ar-u-ca-islamic-umalqura-nu-latn', { day: 'numeric', month: 'long', year: 'numeric' }).formatToParts(d);
      const get = t => (parts.find(p => p.type === t) || {}).value || '';
      return get('day') + ' ' + get('month') + ' ' + get('year');
    } catch (e) { return ''; }
  }

  const FONTS = [
    ['Simplified Arabic', 'Simplified Arabic (عربي مبسّط)'],
    ['Traditional Arabic', 'Traditional Arabic (عربي تقليدي)'],
    ['Sakkal Majalla', 'Sakkal Majalla (سكّال مجلّة)'],
    ['Arabic Typesetting', 'Arabic Typesetting (ويندوز)'],
    ['Andalus', 'Andalus (أندلس — ويندوز)'],
    ['Segoe UI', 'Segoe UI'],
    ['Tahoma', 'Tahoma'],
    ['Arial', 'Arial'],
    ['Times New Roman', 'Times New Roman'],
    ['Calibri', 'Calibri'],
    ['Cairo', 'Cairo (القاهرة — مضمَّن)'],
    ['Tajawal', 'Tajawal (تجوّل — مضمَّن)'],
    ['Almarai', 'Almarai (المراعي — مضمَّن)'],
    ['IBM Plex Sans Arabic', 'IBM Plex Sans Arabic (مضمَّن)'],
    ['Amiri', 'Amiri (أميري — نسخ تقليدي، مضمَّن)'],
    ['Noto Naskh Arabic', 'Noto Naskh Arabic (نسخ — مضمَّن)']
  ];

  function person() {
    return {
      name: '', signImg: '', stampImg: '', showSign: true, showStamp: true, signAR: 0, stampAR: 0,
      pos: { sign: { x: 0, y: 0 }, stamp: { x: 0, y: 0 }, name: { x: 0, y: 0 } }
    };
  }

  const ASSIGN_TOKEN = '[تاريخ التكليف]';
  const ASSIGN_BODY = '<p>قررت تكليفك بالقيام بأعمال ................ في نادي ................ وذلك اعتبارًا من تاريخ ' + ASSIGN_TOKEN + ' ولمدة (   ) يومًا، وبأجرة يومية مقدارها (   ) دنانير فقط لا غير.</p>';
  const CLAIM_BODY = '<p>يطلب لي من النادي مبلغ وقدره (   ) دينار، وذلك لقاء عملي بأعمال ................ في نادي ................ خلال الفترة من    /   /       لغاية    /   /       بواقع (   ) يومًا.</p>';

  function letterDefaults(internal) {
    return {
      number: '', date: todayISO(), showHijri: true, hijri: '', hijriManual: false,
      to0: !internal, to1: !internal, to2: !internal,
      toDir: !!internal, dirName: '', toCustom: '',
      subject: '', body: '', showHead: true, showManager: true
    };
  }

  /** الوفد المشارك: كشف أسماء الوفد (مرافقون إداريون + لاعبون) */
  const DG_ROLES = ['لاعب', 'مدرب', 'عضو هيئة إدارية', 'مدير نادي', 'إداري', 'أمين صندوق', 'سائق', 'معالج'];
  function dgAdmin(role) { return { name: '', ministryNo: '', nationalNo: '', role: role || '' }; }
  function dgPlayer(role) { return { name: '', ministryNo: '', memberNo: '', nationalNo: '', joinDate: '', role: role || '' }; }
  function delegationDefaults() {
    return {
      number: '', date: todayISO(), showHijri: true, hijri: '', hijriManual: false,
      to0: false, to1: false, to2: false, toCustom: '',
      subject: '',
      kind: 'qualifiers', kindName: 'كرة القدم للمربع الذهبي لدورة ميلاد القائد 23', venue: 'مدينة العقبة',
      showIntro: true, introText: '', body: '',
      showAdmins: true, titleAdmin: 'المرافقين الاداريين',
      showPlayers: true, titlePlayers: 'أسماء اللاعبين المشاركين',
      showDays: true,
      admins: Array.from({ length: 5 }, () => dgAdmin()),
      players: Array.from({ length: 8 }, () => dgPlayer()),
      showClosing: false, showHead: true, showManager: true
    };
  }

  function defaults() {
    return {
      v: 2,
      activeDoc: 'letter',
      clubName: '',
      font: 'Tajawal', fontV3: 1,
      bodySize: 14,
      lineHeight: 1.5,
      logoSize: 24,
      onePage: false,
      logoSeeded: false,
      ministrySeeded: false,
      logos: {
        ministry: { show: true, data: '' },
        club: { show: true, data: '' }
      },
      footer: {
        show: true, line: true, lineWidth: 1,
        name: '', nameSize: 12, nameAlign: 'center', nameBold: true,
        phone: '', whatsapp: '', email: '', facebook: '', website: '', address: '', fax: '', social: '', extra: '',
        font: 'Tajawal', size: 10, align: 'center', color: '#222222', bold: false,
        icons: { phone: true, whatsapp: true, email: true, facebook: true, location: true, web: true },
        iconSize: 10, iconPos: 'before', gap: 5
      },
      people: {
        head: person(), manager: person(),
        signSize: 16, stampSize: 22, stampPos: 'right', removeBg: true,
        mode: 'below', modeV2: 1, gapName: 2, gapImg: 2, protect: true
      },
      send: { email: '', whatsapp: '', format: 'pdf', subject: '', message: '', mailApp: 'gmail', useShare: false },
      theme: '',
      letter: letterDefaults(false),
      internal: letterDefaults(true),
      thanks: Object.assign(letterDefaults(false), { toText: '', copiesOn: false, copiesText: '', toX: 0, toY: 0 }),
      assign: Object.assign(letterDefaults(true), {
        who: 'السيدة', assignee: '', assignDate: todayISO(), assignDateText: '', body: ASSIGN_BODY,
        showClosing: false, showFinance: true, financeText: ''
      }),
      claim: {
        number: '', date: todayISO(), showHijri: false, hijri: '', hijriManual: false, showDateRow: false,
        title: 'مطالبة مالية', subject: '', body: CLAIM_BODY, showClosing: true, closing: 'وتفضلوا بقبول فائق الاحترام',
        owner: '', lblOwner: 'صاحب الاستحقاق:', lblSign: 'التوقيع:', lblDate: 'التاريخ:',
        showClaimDate: true, claimDate: '', claimDateText: '', showHead: false, showManager: false
      },
      delegation: delegationDefaults(),
      circular: {
        number: '', subject: '', body: '',
        date: todayISO(), showDate: true, showHijri: true, hijri: '', hijriManual: false,
        showNumUnder: false, showTitleNum: true,
        lineShow: true, lineWidth: 1.5, lineStyle: 'solid', lineColor: '#000000', lineBefore: 3, lineAfter: 3, linePct: 100,
        titleSize: 17, titleAlign: 'center', titleCustom: '',
        bodyFont: '', bodySize: 0, bodyAlign: 'justify',
        showHead: true, showManager: true
      }
    };
  }

  function merge(base, over) {
    if (over === null || typeof over !== 'object' || Array.isArray(over)) return base;
    Object.keys(base).forEach(k => {
      if (!(k in over)) return;
      if (base[k] !== null && typeof base[k] === 'object' && !Array.isArray(base[k])) base[k] = merge(base[k], over[k]);
      else if (typeof over[k] === typeof base[k]) base[k] = over[k];
    });
    return base;
  }

  /** ترحيل بيانات النسخة الأولى (v1) إلى البنية الجديدة */
  function migrate(raw) {
    if (!raw || typeof raw !== 'object') return {};
    if (raw.v >= 2) {
      const r = JSON.parse(JSON.stringify(raw));
      if (!r.fontV3) { r.font = 'Tajawal'; if (r.footer) r.footer.font = 'Tajawal'; r.fontV3 = 1; }
      if (r.people && !r.people.modeV2) { r.people.mode = 'below'; r.people.modeV2 = 1; }
      ['letter', 'internal', 'thanks', 'circular', 'assign', 'claim', 'delegation'].forEach(k => {
        const d = r[k]; if (!d || d.showHead !== undefined) return;
        const who = d.signer || d.signWho || 'manager';
        d.showHead = who === 'head' || who === 'both'; d.showManager = who === 'manager' || who === 'both';
      });
      return r;
    }
    const o = JSON.parse(JSON.stringify(raw));
    o.people = o.people || { head: {}, manager: {} };
    o.font = 'Tajawal'; if (o.footer) o.footer.font = 'Tajawal'; o.fontV3 = 1;
    if (!o.people.modeV2) { o.people.mode = 'below'; o.people.modeV2 = 1; }
    const L = o.letter || {}, C = o.circular || {};
    const key = L.signTitle === 'رئيس الهيئة الإدارية' ? 'head' : 'manager';
    if (L.signName) o.people[key].name = L.signName;
    if (L.signImg) { o.people[key].signImg = L.signImg; o.people[key].showSign = L.signShow !== false; }
    if (C.headName && !o.people.head.name) o.people.head.name = C.headName;
    if (C.managerName && !o.people.manager.name) o.people.manager.name = C.managerName;
    if (o.letter) { o.letter.showHead = key === 'head'; o.letter.showManager = key !== 'head'; }
    if (o.circular && o.circular.number != null) o.circular.date = App_todayISO();
    if (o.font === 'Tahoma') o.font = 'Simplified Arabic';
    if (o.footer && o.footer.font === 'Tahoma') o.footer.font = 'Simplified Arabic';
    if (o.bodySize === 13) o.bodySize = 14;
    if (o.footer && o.footer.size === 9) o.footer.size = 10;
    o.v = 2;
    return o;
  }
  function App_todayISO() { return todayISO(); }

  function build(raw) {
    const d = defaults();
    merge(d, migrate(raw));
    // شعار الأندية الافتراضي: يُضاف مرة واحدة فقط، فإن أزاله المستخدم لا يعود
    if (!d.logoSeeded) {
      if (!d.logos.club.data && g.DEFAULT_CLUB_LOGO) { d.logos.club.data = g.DEFAULT_CLUB_LOGO; d.logos.club.show = true; }
      d.logoSeeded = true;
    }
    // شعار الوزارة الافتراضي (مرة واحدة فقط)
    if (!d.ministrySeeded) {
      if (!d.logos.ministry.data && g.DEFAULT_MINISTRY_LOGO) { d.logos.ministry.data = g.DEFAULT_MINISTRY_LOGO; d.logos.ministry.show = true; }
      d.ministrySeeded = true;
    }
    // الشعاران ثابتان: لا يُستبدلان ولا يُزالان ولا يُخفيان؛ يتحكم المستخدم بحجمهما فقط
    if (g.DEFAULT_CLUB_LOGO) { d.logos.club.data = g.DEFAULT_CLUB_LOGO; }
    if (g.DEFAULT_MINISTRY_LOGO) { d.logos.ministry.data = g.DEFAULT_MINISTRY_LOGO; }
    d.logos.club.show = true; d.logos.ministry.show = true;
    const known = f => FONTS.some(x => x[0] === f);
    if (!known(d.font)) d.font = 'Tajawal';
    if (!known(d.footer.font)) d.footer.font = 'Tajawal';
    if (d.circular.bodyFont && !known(d.circular.bodyFont)) d.circular.bodyFont = '';
    if (!d.letter.hijri && !d.letter.hijriManual) d.letter.hijri = formatHijri(d.letter.date);
    if (!d.internal.hijri && !d.internal.hijriManual) d.internal.hijri = formatHijri(d.internal.date);
    if (!d.thanks.hijri && !d.thanks.hijriManual) d.thanks.hijri = formatHijri(d.thanks.date);
    if (!d.assign.hijri && !d.assign.hijriManual) d.assign.hijri = formatHijri(d.assign.date);
    if (!d.claim.hijri && !d.claim.hijriManual) d.claim.hijri = formatHijri(d.claim.date);
    ['admins', 'players'].forEach(k => {
      const mk = k === 'admins' ? dgAdmin : dgPlayer, a = d.delegation[k];
      d.delegation[k] = (Array.isArray(a) ? a : []).filter(r => r && typeof r === 'object').map(r => { const o = mk(); Object.keys(o).forEach(f => { if (typeof r[f] === 'string') o[f] = r[f]; }); return o; });
    });
    if (!d.delegation.hijri && !d.delegation.hijriManual) d.delegation.hijri = formatHijri(d.delegation.date);
    if (!d.circular.hijri && !d.circular.hijriManual) d.circular.hijri = formatHijri(d.circular.date);
    return d;
  }

  /* ---------- التخزين الدائم: IndexedDB مع احتياط localStorage ---------- */
  const DB = 'teachers-clubs-editor', STORE = 'kv';
  function openDb() {
    return new Promise((res, rej) => {
      if (!g.indexedDB) return rej(new Error('no idb'));
      const rq = indexedDB.open(DB, 1);
      rq.onupgradeneeded = () => rq.result.createObjectStore(STORE);
      rq.onsuccess = () => res(rq.result);
      rq.onerror = () => rej(rq.error);
    });
  }
  function idbGet(db, k) { return new Promise((res, rej) => { const r = db.transaction(STORE).objectStore(STORE).get(k); r.onsuccess = () => res(r.result); r.onerror = () => rej(r.error); }); }
  function idbSet(db, k, v) { return new Promise((res, rej) => { const t = db.transaction(STORE, 'readwrite'); t.objectStore(STORE).put(v, k); t.oncomplete = () => res(); t.onerror = () => rej(t.error); }); }

  let db = null;
  async function loadAsync() {
    let raw = null, att = null;
    try { db = await openDb(); raw = await idbGet(db, 'state'); att = await idbGet(db, 'attachments'); } catch (e) { db = null; }
    if (!raw) {
      try { const s = localStorage.getItem(KEY) || localStorage.getItem('teachers-clubs-editor:v1'); if (s) raw = JSON.parse(s); } catch (e) { /* تجاهل */ }
    }
    try { if (navigator.storage && navigator.storage.persist) App_persisted = await navigator.storage.persist(); } catch (e) { /* تجاهل */ }
    return { state: build(raw), attachments: att };
  }
  let App_persisted = false;

  let saveTimer = null, saveFailed = false;
  function save(state) {
    clearTimeout(saveTimer);
    saveTimer = setTimeout(() => flush(state), 250);
  }
  async function flush(state) {
    const att = g.App.attach && g.App.attach.store ? g.App.attach.store : null;
    try {
      if (db) {
        await idbSet(db, 'state', JSON.parse(JSON.stringify(state)));
        if (att) await idbSet(db, 'attachments', JSON.parse(JSON.stringify(att)));
      } else {
        localStorage.setItem(KEY, JSON.stringify(state));
      }
      saveFailed = false; g.App.lastSaved = Date.now();
    } catch (e) {
      if (!saveFailed && g.toast) g.toast('تعذّر الحفظ التلقائي (التخزين ممتلئ أو معطّل في المتصفح)');
      saveFailed = true;
    }
  }
  async function wipe() {
    try { localStorage.removeItem(KEY); localStorage.removeItem('teachers-clubs-editor:v1'); } catch (e) { /* تجاهل */ }
    if (db) { db.close(); await new Promise(r => { const q = indexedDB.deleteDatabase(DB); q.onsuccess = q.onerror = q.onblocked = () => r(); }); }
  }

  function getPath(obj, path) { return path.split('.').reduce((o, k) => (o == null ? o : o[k]), obj); }
  function setPath(obj, path, val) {
    const ks = path.split('.'); const last = ks.pop();
    const t = ks.reduce((o, k) => o[k], obj); t[last] = val;
  }

  g.App = Object.assign(g.App || {}, {
    KEY, FONTS, defaults, build, loadAsync, save, flush, wipe, getPath, setPath, migrate,
    isPersisted: () => App_persisted,
    todayISO, formatGregorian, formatHijri, DG_ROLES, dgAdmin, dgPlayer, ASSIGN_TOKEN, ASSIGN_BODY, CLAIM_BODY
  });
})(window);
