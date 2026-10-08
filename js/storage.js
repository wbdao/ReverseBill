/* ═══════════════════════════════════════════════════════════════════════
   js/storage.js — وحدة التخزين المحلي (Services)
   • قراءة/كتابة LocalStorage لكل كيان بمفتاح مستقل + تحميل أمان للاستثناءات
   • مفتاحان جاهزان للترحيل من النسخة القديمة (LEGACY_PRODUCTS)
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const KEYS = global.CONFIG.STORAGE_KEYS;

  const load = (key, fallback) => {
    try {
      const raw = localStorage.getItem(key);
      return raw === null ? fallback : JSON.parse(raw);
    } catch (err) {
      global.Utils && global.Utils.toast ? global.Utils.toast(`تعذّر قراءة بيانات "${key}" — قد تكون تالفة`, 'error') : console.error(err);
      return fallback;
    }
  };

  const isQuota = (err) => !!err && (err.name === 'QuotaExceededError' || err.code === 22 || err.code === 1014);
  // تنبيه واحد لكل رسالة في الجلسة — الحفظ يتكرر مع كل تحديث لستة فلا نُغرق الشاشة
  const warned = new Set();
  const warnOnce = (msg, err) => {
    if (warned.has(msg)) return;
    warned.add(msg);
    global.Utils && global.Utils.toast ? global.Utils.toast(msg, 'error') : console.error(err);
  };

  const save = (key, value) => {
    const text = JSON.stringify(value);
    try { localStorage.setItem(key, text); return true; }
    catch (err) {
      if (isQuota(err)) {
        // النسخة القديمة من نفس المفتاح قد تكون هي ما يملأ المساحة: نحذفها ونعيد المحاولة
        try { localStorage.removeItem(key); localStorage.setItem(key, text); return true; } catch (e) { /* ما زالت ممتلئة */ }
      }
      warnOnce(isQuota(err)
        ? 'مساحة التخزين المحلية ممتلئة — صدّر نسخة احتياطية من سجل الفواتير ثم امسح القديم منها'
        : 'فشل حفظ البيانات محلياً', err);
      return false;
    }
  };

  /* --- صيغة تخزين مضغوطة للستات (fmt 2) ---
     ~39 ألف سطر في 5 لستات بصيغة كائنات = ~5.6 مليون حرف، أكبر من حد localStorage (~5M).
     النصوص المتكررة (الاسم/الوحدة/الحجم/اللون) تُخزَّن مرة واحدة في جدول strings،
     وكل صنف مصفوفة [الكود, #الاسم, #الوحدة, #الحجم, #اللون, السعر] → ~1.3 مليون حرف.
     معرّفات الأصناف لا تُخزَّن (اللستات نسخة من الشيت)؛ تُولَّد عند القراءة. */
  function encodeLists(lists) {
    const strings = [];
    const pos = new Map();
    const ref = (s) => {
      const v = s == null ? '' : String(s);
      let i = pos.get(v);
      if (i === undefined) { i = strings.length; strings.push(v); pos.set(v, i); }
      return i;
    };
    const out = lists.map((l) => {
      const meta = {};
      for (const k of Object.keys(l)) if (k !== 'items' && k !== '__ver') meta[k] = l[k];
      meta.rows = (l.items || []).map((it) => [it.itemNumber || '', ref(it.name), ref(it.unit), ref(it.size), ref(it.color), Number(it.price) || 0]);
      return meta;
    });
    return { strings, lists: out };
  }
  function decodeLists(d) {
    const S = Array.isArray(d.strings) ? d.strings : [];
    return (Array.isArray(d.lists) ? d.lists : []).map((meta, li) => {
      const { rows, ...l } = meta;
      l.items = (Array.isArray(rows) ? rows : []).map((r, i) => ({
        id: `li_${li}_${i}`, itemNumber: r[0] || '', name: S[r[1]] || '', unit: S[r[2]] || '',
        size: S[r[3]] || '', color: S[r[4]] || '', price: Number(r[5]) || 0,
      }));
      return l;
    });
  }

  const Storage = {
    /* --- قوائم الأسعار (تُقرأ الصيغة المضغوطة fmt 2 أو القديمة بكائنات كاملة) --- */
    getPriceLists: () => {
      const d = load(KEYS.PRICE_LISTS, {});
      const lists = d.fmt === 2 ? decodeLists(d) : (Array.isArray(d.lists) ? d.lists : []);
      return { lists, activeId: d.activeId || null, version: Number(d.version) || 0 };
    },
    savePriceLists: (lists, activeId) => save(KEYS.PRICE_LISTS, { fmt: 2, ...encodeLists(lists), activeId, version: (global.CONFIG.LIST_DATA_VERSION || 1) }),

    /* --- (ترحيل) النسخة القديمة أحادية اللستة --- */
    getLegacyProducts: () => load(KEYS.LEGACY_PRODUCTS, null),

    /* --- سجل الفواتير --- */
    getInvoices: () => load(KEYS.INVOICES, []),
    saveInvoices: (invoices) => save(KEYS.INVOICES, invoices),

    /* --- إعدادات (اللستة النشطة + مفاتيح الربط السحابي) --- */
    getSettings: () => Object.assign({ activeListId: null, sheetApiKey: (global.CONFIG.SHEETS && global.CONFIG.SHEETS.API_KEY) || '', sheetUrl: '', lastSyncAt: null }, load(KEYS.SETTINGS, {})),
    saveSettings: (s) => save(KEYS.SETTINGS, s),

    /* --- مسودة الفاتورة الحالية --- */
    getDraft: () => load(KEYS.DRAFT, null),
    saveDraft: (d) => save(KEYS.DRAFT, d),
    clearDraft: () => { try { localStorage.removeItem(KEYS.DRAFT); } catch (e) {} },
  };

  global.Storage = Storage;
})(window);