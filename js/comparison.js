/* ═══════════════════════════════════════════════════════════════════════
   js/comparison.js — محرك المقارنة الذكية (Domain)
   • يشغّل المطابقة عبر Lists.matchInList (فهرس Map سريع)
   • المؤشرات: مطابق 🟢، زيادة 🔺، سعر منخفض خطير 🔻، غير مسجل ⚠️
   • تلخيص مالي: إجمالي الفاتورة/المتوقع/الزيادة/الانخفاض/صافي الفرق/الانحراف
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const { num } = global.Utils;
  const Lists = global.Lists;

  const STATUS = { MATCH: 'match', HIGH: 'high', LOW: 'low', UNKNOWN: 'unknown' };

  /** صافي سعر الوحدة بعد الخصومات:
      discountPct مُخزَّن ككسر (1% = 0.01) — الإدخال اليدوي يقسم قيمته على 100
      ليتطابق مع سحب السيستم (الذي يأتي بالفعل ككسر). القيمة تُحدّ من الأسفل بصفر. */
  function netUnitOf(item) {
    const unit = num(item.unitPrice);
    const pct = num(item.discountPct);
    const value = num(item.discountValue);
    return Math.max(0, unit * (1 - pct) - value);
  }

  /** تحليل بند واحد مقابل لستة معتمدة (كائن لستة أو null للاستعانة بالنشطة)
      المقارنة تتم على صافي السعر بعد الخصم (وهو المبلغ الفعلي الذي يُدفع للعميل). */
  function analyzeItem(item, listOverride) {
    const qty = num(item.quantity);
    const unit = num(item.unitPrice);
    const netUnit = netUnitOf(item);
    const invoiceTotal = qty * netUnit;
    // لستة المقارنة: المعطاة صراحة → وإلا النشطة عبر مخزن الذاكرة المركزي window.appLists
    let listItem = null;
    if (!listOverride) {
      const act = Lists.active();
      if (act && act.cfgSlug) listItem = Lists.findInStore(act.cfgSlug, item.itemNumber, item.name, item.unit);
      else listItem = act ? Lists.matchInList(act, item.itemNumber, item.name, item.unit) : null;
    } else {
      listItem = Lists.matchInList(listOverride, item.itemNumber, item.name, item.unit);
    }
    const listPrice = listItem !== null ? num(listItem.price) : null;

    const unitPrice = num(item.unitPrice);

    if (listItem === null || listPrice === null) {
      return { item, listItem: null, status: STATUS.UNKNOWN, listPrice: null, unitPrice, unitDiff: 0, unitDiffBefore: 0, totalDiff: 0, expectedTotal: invoiceTotal, invoiceTotal, netUnit };
    }

    const unitDiff = netUnit - listPrice;
    const unitDiffBefore = unitPrice - listPrice;   // فرق السعر قبل الخصم عن السعر المعتمد
    const totalDiff = unitDiff * qty;
    let status;
    if (Math.abs(unitDiff) < 1e-9) status = STATUS.MATCH;      // 🟢
    else if (unitDiff > 0) status = STATUS.HIGH;               // 🔺 زيادة في السعر (لصالح الشركة)
    else status = STATUS.LOW;                                  // 🔻 سعر أقل من المعتمد (خسارة / انخفاض خطير)

    return { item, listItem, status, listPrice, unitPrice, unitDiff, unitDiffBefore, totalDiff, expectedTotal: qty * listPrice, invoiceTotal, netUnit };
  }

  /** تحليل مصفوفة بنود دفعة واحدة */
  const analyzeBatch = (items, listOverride) => items.map((it) => analyzeItem(it, listOverride));

  /** تلخيص مالي للبنود */
  function summarize(items, listOverride) {
    const rows = analyzeBatch(items, listOverride);
    const totals = { invoiceTotal: 0, expectedTotal: 0, knownInvoiceTotal: 0, highTotal: 0, lowTotal: 0, unknownCount: 0, unknownTotal: 0, matchCount: 0 };
    for (const r of rows) {
      totals.invoiceTotal += r.invoiceTotal;
      if (r.status === STATUS.UNKNOWN) { totals.unknownCount++; totals.unknownTotal += r.invoiceTotal; continue; }
      totals.expectedTotal += r.expectedTotal;
      totals.knownInvoiceTotal += r.invoiceTotal;
      if (r.status === STATUS.HIGH) totals.highTotal += r.totalDiff;
      // البيع دون السعر المعتمد = خسارة/تجاوز: يُجمع كقيمة موجبة للتنبيه (لا يُطلق عليه "وفر")
      if (r.status === STATUS.LOW) totals.lowTotal += Math.abs(r.totalDiff);
      if (r.status === STATUS.MATCH) totals.matchCount++;
    }
    const netDiff = totals.knownInvoiceTotal - totals.expectedTotal;
    const deviationPct = totals.expectedTotal > 0 ? (netDiff / totals.expectedTotal) * 100 : 0;
    return { rows, totals, netDiff, deviationPct };
  }

  const statusLabel = (status) => {
    switch (status) {
      case STATUS.MATCH: return '🟢 مطابق';
      case STATUS.HIGH: return '🔺 زيادة في السعر';
      case STATUS.LOW: return '🔻 سعر منخفض خطير';
      default: return '⚠️ غير مسجل بالقائمة';
    }
  };
  const statusText = (status) => {
    switch (status) {
      case STATUS.MATCH: return 'مطابق';
      case STATUS.HIGH: return 'زيادة في السعر';
      case STATUS.LOW: return 'سعر منخفض خطير';
      default: return 'غير مسجل بالقائمة';
    }
  };

  global.Comparison = { STATUS, analyzeItem, analyzeBatch, summarize, statusLabel, statusText, netUnitOf };

  /* ================= المقارنة الشاملة بين كل اللستات =================
     يقرأ الأصناف مباشرة من مخزن الذاكرة المركزي window.appLists (بالترتيب حسب
     CONFIG.LISTS) لضمان سرعة قصوى دون لمس localStorage. يوحّد الأصناف عبر جميع
     اللستات بمفتاح {رقم الصنف+الوحدة | الاسم+الوحدة} ثم يُخرج مصفوفة:
     {itemNumber, name, unit, prices:[سعر لكل لستة], min, max, diff, presentCount, status}.
     status: 'match' (متطابق في كل اللستات) / 'diff' (متفاوت) / 'single' (مسجّل في لستة واحدة). */
  function compareAllLists(storeOrLists) {
    const cfg = (global.CONFIG && global.CONFIG.LISTS) || [];
    const store = global.appLists || {};
    const width = cfg.length || (Array.isArray(storeOrLists) ? storeOrLists.length : 0);
    const keyOf = (it) => {
      const nc = global.Utils.normalizeNum(it.itemNumber);
      const nm = global.Utils.normalizeName(it.name);
      return (nc ? 'n:' + nc : 'm:' + nm) + '\u0000' + global.Utils.normalizeUnit(it.unit);
    };
    const map = new Map();
    for (let li = 0; li < width; li++) {
      // مسار المخزن المركزي أولاً (أفضل أداء وتزامن)، ثم رجوع للصفيف الممرَّر (متوافق مع الوضع القديم)
      let its;
      if (cfg[li] && store[cfg[li].slug]) its = store[cfg[li].slug];
      else if (Array.isArray(storeOrLists) && storeOrLists[li]) its = storeOrLists[li].items || storeOrLists[li];
      else its = [];
      for (const it of its) {
        if (!it.name && !it.itemNumber) continue;
        const k = keyOf(it);
        let row = map.get(k);
        if (!row) {
          row = { itemNumber: it.itemNumber, name: it.name, unit: it.unit, prices: new Array(width).fill(null), present: new Set() };
          map.set(k, row);
        }
        if (row.present.has(li)) continue;
        row.prices[li] = num(it.price);
        row.present.add(li);
      }
    }
    const rows = Array.from(map.values()).map((r) => {
      const vals = r.prices.filter((p) => p !== null && p !== undefined);
      const min = vals.length ? Math.min.apply(null, vals) : null;
      const max = vals.length ? Math.max.apply(null, vals) : null;
      return {
        itemNumber: r.itemNumber, name: r.name, unit: r.unit, prices: r.prices,
        min, max,
        diff: vals.length > 1 ? (max - min) : 0,
        presentCount: r.present.size,
        status: vals.length <= 1 ? 'single' : (Math.abs(max - min) < 1e-9 ? 'match' : 'diff'),
      };
    });
    rows.sort((a, b) => (b.diff - a.diff) || (a.presentCount - b.presentCount) || String(a.name).localeCompare(String(b.name), 'ar'));
    return rows;
  }

  global.Comparison = { STATUS, analyzeItem, analyzeBatch, summarize, statusLabel, statusText, compareAllLists, netUnitOf };
})(window);