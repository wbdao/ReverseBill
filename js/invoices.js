/* ═══════════════════════════════════════════════════════════════════════
   js/invoices.js — وحدة سجل الفواتير والمسودة (Domain)
   • إنشاء فاتورة، سجل تلقائي بفرز الأحدث أولاً
   • تصدير/استيراد JSON (مصفوفة فواتير)، مسح، عينات
   • ربط كل فاتورة بلستة المقارنة المعتمدة وقت الحفظ (listId)
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const { genId, todayStr } = global.Utils;
  const Storage = global.Storage;
  const Lists = global.Lists;

  let invList = [];

  const init = () => {
    invList = Storage.getInvoices();
    if (!Array.isArray(invList)) invList = [];
  };
  const persist = () => Storage.saveInvoices(invList);

  const getAll = () => [...invList].sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  const getById = (id) => invList.find((inv) => inv.id === id) || null;
  const count = () => invList.length;

  function normalizeItem(it) {
    return {
      itemNumber: String(it.itemNumber || '').trim(),
      name: String(it.name || '').trim(),
      unit: String(it.unit || '').trim(),
      quantity: Number(it.quantity) || 0,
      unitPrice: Number(it.unitPrice) || 0,
      discountValue: Number(it.discountValue) || 0,
      discountPct: Number(it.discountPct) || 0,
    };
  }

  // قراءة متوافقة: تحويلة `customer` الجديدة مع دعم البيانات القديمة المحفوظة تحت `vendor`
  const customerOf = (d) => {
    const v = d && typeof d === 'object' ? (d.customer !== undefined && d.customer !== null ? d.customer : d.vendor) : '';
    return String(v || '').trim();
  };

  function create(data) {
    const invoice = {
      id: genId('inv'),
      customer: customerOf(data),
      invoiceNo: String(data.invoiceNo || '').trim(),
      date: data.date || todayStr(),
      notes: String(data.notes || '').trim(),
      listId: data.listId || Lists.activeIdOf() || null,
      invoiceDiscountPct: Number(data.invoiceDiscountPct) || 0,
      items: (data.items || []).map(normalizeItem),
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString(),
    };
    invList.push(invoice);
    persist();
    return invoice;
  }

  /** تحديث فاتورة موجودة في مكانها (يحافظ على id و createdAt، يُحدّث updatedAt). */
  function update(id, data) {
    const idx = invList.findIndex((inv) => inv.id === id);
    if (idx === -1) return null;
    const prev = invList[idx];
    const next = {
      ...prev,
      customer: customerOf(data),
      invoiceNo: String(data.invoiceNo || '').trim(),
      date: data.date || todayStr(),
      notes: String(data.notes || '').trim(),
      listId: data.listId || prev.listId || Lists.activeIdOf() || null,
      invoiceDiscountPct: Number(data.invoiceDiscountPct) || 0,
      items: (data.items || []).map(normalizeItem),
      updatedAt: new Date().toISOString(),
    };
    invList[idx] = next;
    persist();
    return next;
  }

  /** Upsert: تحديث إن وُجدت بالمعرّف، وإلا إنشاء جديد — يمنع تكرار الفاتورة المعروضة. */
  function save(data, id) {
    return id && getById(id) ? update(id, data) : create(data);
  }

  const addSample = () => create({
    customer: 'شركة النور للتجارة — عميل',
    invoiceNo: 'INV-1001',
    date: todayStr(),
    notes: 'فاتورة تجريبية لمراجعة أسعار بيع العميل مقابل القائمة المعتمدة',
    listId: Lists.activeIdOf(),
    items: [
      { itemNumber: 'T-1001', name: 'سكر 1 كجم', unit: 'كيس', quantity: 10, unitPrice: 30 },          // 🟢 مطابق
      { itemNumber: 'T-1003', name: 'زيت عباد الشمس 1 لتر', unit: 'عبوة', quantity: 5, unitPrice: 75 }, // 🔺 زيادة
      { itemNumber: 'T-1006', name: 'لبن 1 لتر', unit: 'عبوة', quantity: 8, unitPrice: 24 },            // 🔻 انخفاض خطير
      { itemNumber: 'T-1005', name: 'شاي 500 جم', unit: 'علبة', quantity: 4, unitPrice: 100 },          // 🟢 مطابق
      { itemNumber: 'T-1011', name: 'قهوة تركية 250 جم', unit: 'عبوة', quantity: 3, unitPrice: 88 },     // ⚠️ غير مسجل
    ],
  });

  const remove = (id) => { const b = invList.length; invList = invList.filter((i) => i.id !== id); persist(); return invList.length !== b; };
  const clearAll = () => { invList = []; persist(); };

  const exportJson = () => JSON.stringify(invList, null, 2);

  function importJson(text, replace = false) {
    const parsed = JSON.parse(text);
    if (!Array.isArray(parsed)) throw new Error('صيغة غير صحيحة: يجب أن يكون الملف مصفوفة فواتير.');
    if (replace) invList = [];
    const existing = new Set(invList.map((inv) => inv.id));
    let added = 0;
    for (const raw of parsed) {
      if (!raw || typeof raw !== 'object' || existing.has(raw.id)) continue;
      const invoice = {
        id: raw.id || genId('inv'),
        customer: customerOf(raw),
        invoiceNo: String(raw.invoiceNo || '').trim(),
        date: raw.date || todayStr(),
        notes: String(raw.notes || '').trim(),
        listId: raw.listId || null,
        invoiceDiscountPct: Number(raw.invoiceDiscountPct) || 0,
        items: Array.isArray(raw.items) ? raw.items.map(normalizeItem) : [],
        createdAt: raw.createdAt || new Date().toISOString(),
      };
      invList.push(invoice);
      existing.add(invoice.id);
      added++;
    }
    persist();
    return { added };
  }

  /* --- مسودة الفاتورة الحالية --- */
  const saveDraft = (d) => Storage.saveDraft(d);
  const loadDraft = () => Storage.getDraft();
  const clearDraft = () => Storage.clearDraft();

  global.Invoices = { init, getAll, getById, count, create, update, save, addSample, remove, clearAll, exportJson, importJson, saveDraft, loadDraft, clearDraft };
})(window);