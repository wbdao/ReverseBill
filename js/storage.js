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

  const save = (key, value) => {
    try { localStorage.setItem(key, JSON.stringify(value)); return true; }
    catch (err) {
      const msg = err && err.name === 'QuotaExceededError'
        ? 'مساحة التخزين المحلية ممتلئة — حاول تقليص حجم اللستات أو تصدير نسخة احتياطية'
        : 'فشل حفظ البيانات محلياً';
      global.Utils && global.Utils.toast ? global.Utils.toast(msg, 'error') : console.error(err);
      return false;
    }
  };

  const Storage = {
    /* --- قوائم الأسعار --- */
    getPriceLists: () => {
      const d = load(KEYS.PRICE_LISTS, {});
      return { lists: Array.isArray(d.lists) ? d.lists : [], activeId: d.activeId || null };
    },
    savePriceLists: (lists, activeId) => save(KEYS.PRICE_LISTS, { lists, activeId }),

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