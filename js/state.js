/* ═══════════════════════════════════════════════════════════════════════
   js/state.js — إدارة الحالة المركزية (State Management)
   • حقيبة حالة واحدة لكل الواجهة: البنود، التحرير، الفلاتر، الإعدادات
   • اشتراكات إعلام (pub/sub) حتى تعيد الواجهات رسم نفسها عند التغييرات
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const Storage = global.Storage;

  const State = {
    items: [],              // بنود الفاتورة الحالية
    editingItemId: null,
    editingInvoiceId: null, // فاتورة السجل المفتوحة حالياً في المحرر (null = فاتورة جديدة)
    productFilter: '',
    statusFilter: '',       // تصفية بنود الفاتورة حسب ملاحظة مراجعة السعر ('' | match | high | low | unknown)
    settings: Storage.getSettings(),
    syncBusy: false,
    lastSync: null,
    listPage: 0,            // صفحة جدول اللستة النشطة (ترقيم للبيانات الضخمة)
  };

  /* --- اشتراكات إعلام --- */
  const subs = new Set();
  const subscribe = (fn) => { subs.add(fn); return () => subs.delete(fn); };
  const publish = () => subs.forEach((fn) => { try { fn(); } catch (e) { /* تجاهل */ } });

  /* --- حفظ وتحديث بسيط للإعدادات --- */
  const persistSettings = () => { State.settings = Storage.getSettings(); Storage.saveSettings(State.settings); };
  Object.defineProperty(State, 'persistSettings', { value: persistSettings, enumerable: false });

  global.State = State;
  global.StateBus = { subscribe, publish };
})(window);