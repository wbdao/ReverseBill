/* ═══════════════════════════════════════════════════════════════════════
   js/parser.js — طبقة المعالجة (Services)
   • تحليل نصوص TSV/CSV المنسوخة من Excel أو ملفات السيستم
   • كشف رؤوس الأعمدة (عربي/إنجليزي) وتحديد دور كل عمود
   • استخراج بنود الفاتورة / بنود اللستة من أي شكل جدولي
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const { normKey } = global.Utils;
  const HINTS = global.CONFIG.HEADER_HINTS;

  /* --- تقسيم صف واحد إلى خلايا (تبويبات أولاً، ثم فواصل شائعة) --- */
  function splitTSVRow(row) {
    if (row.includes('\t')) return row.split('\t');
    const comma = row.split(/[،,;؛]/);
    if (comma.length > 1) return comma;
    return row.split(/\s{2,}/);
  }

  /* --- تحويل نص (TSV/CSV) إلى صفوف خلايا --- */
  function parseRows(text) {
    const rows = [];
    String(text || '').split(/\r\n|\r|\n/).forEach((line, i) => {
      const cells = splitTSVRow(line).map((c) => (i === 0 ? c.replace(/^\uFEFF/, '') : c).trim());
      if (cells.length === 1 && cells[0] === '') return;
      rows.push({ cells });
    });
    return rows;
  }

  /* --- تحليل CSV حقيقي (مقتبس بـ ") من روابط Google Sheets المنشورة ---
     محلل أحادي المسار: يُقص الخلايا ويزيل BOM ويهمل الأسطر الفارغة أثناء القراءة،
     فلا يُنشئ مصفوفة صفوف وسيطة ولا خريطة كاملة ثانية (أضعف الذاكرة في المسارات الضخمة). */
  function parseCSV(text) {
    const rows = [];
    let row = [], field = '', inQuotes = false;
    let firstRow = true, firstCell = true;
    const pushField = () => {
      let f = field.trim();
      if (firstCell && firstRow) f = f.replace(/^\uFEFF/, '');
      row.push(f);
      field = '';
      firstCell = false;
    };
    const pushRow = () => {
      // تُهمل الأسطر الفارغة فقط (خلية واحدة فارغة) كما في السلوك السابق
      // يظل firstRow على حاله حتى يُدفع أول سطر فعلي (حتى يُزال BOM من الرأس الحقيقي)
      if (row.length !== 1 || row[0] !== '') { rows.push({ cells: row }); firstRow = false; }
      row = [];
      firstCell = true;
    };
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (inQuotes) {
        if (ch === '"') { if (text[i + 1] === '"') { field += '"'; i++; } else inQuotes = false; }
        else field += ch;
      } else if (ch === '"') inQuotes = true;
      else if (ch === ',' || ch === '،') pushField();
      else if (ch === '\n') { pushField(); pushRow(); }
      else if (ch === '\r') { /* تجاهل */ }
      else field += ch;
    }
    if (field !== '' || row.length) { pushField(); pushRow(); }
    return rows;
  }

  /* --- كشف دور عمود من نص الرأس --- */
  function bestRole(raw) {
    if (!raw) return null;
    const key = normKey(raw);
    let best = null, bestScore = -1;
    for (const [role, hints] of Object.entries(HINTS)) {
      for (const h of hints) {
        const nk = normKey(h);
        let score = -1;
        if (nk === key) score = 1000 + nk.length;
        else if (key.length >= 4 && key.length >= nk.length && key.includes(nk)) score = 500 + key.length;
        else if (nk.length >= 5 && nk.includes(key)) score = 100 + nk.length;
        if (score > bestScore) { best = role; bestScore = score; }
      }
    }
    return best;
  }

  /* --- كشف رؤوس كل الأعمدة في صف ما --- */
  function detectHeaderRoles(cells) {
    const roles = {};
    let matched = 0;
    cells.forEach((cell, idx) => {
      const role = bestRole(cell);
      if (role && roles[role] === undefined) { roles[role] = idx; matched++; }
    });
    return matched > 0 ? roles : null;
  }

  const cv = (c, idx) => (idx !== undefined && idx !== null && c[idx] !== undefined && c[idx] !== null ? (c[idx] || '').trim() : '');

  /* --- بناء خريطة أعمدة الفاتورة (إن وُجدت رؤوس) --- */
  function buildInvoiceMapping(cellsRows) {
    const first = cellsRows[0] && cellsRows[0];
    const roles = first ? detectHeaderRoles(first) : null;
    if (roles && roles.qty !== undefined && (roles.name !== undefined || roles.itemNo !== undefined)) {
      return {
        header: true,
        m: {
          itemNo: roles.itemNo, name: roles.name, unit: roles.unit, size: roles.size, color: roles.color, qty: roles.qty,
          unitPrice: roles.unitPrice !== undefined ? roles.unitPrice : roles.price,
          amount: roles.amount,
          discountValue: roles.discountValue, discountPct: roles.discountPct,
        },
      };
    }
    return { header: false };
  }
  /* --- خريطة الأعمدة الموضعية للفواتير بدون رؤوس --- */
  function invoicePosMapping(cells) {
    const n = cells.length;
    if (n <= 2) return { name: 0, qty: 1 };
    if (n === 3) return { name: 0, qty: 1, unitPrice: 2 };
    if (n === 4) return { itemNo: 0, name: 1, qty: 2, unitPrice: 3 };
    if (n === 5) return { itemNo: 0, name: 1, unit: 2, qty: 3, unitPrice: 4 };
    if (n === 6) return { itemNo: 0, name: 1, unit: 2, qty: 3, unitPrice: 4, discountValue: 5 };
    return { itemNo: 0, name: 1, unit: 2, qty: 3, unitPrice: 4, discountValue: 5, discountPct: 6 };
  }

  /* --- بناء خريطة أعمدة اللستة (إن وُجدت رؤوس) --- */
  function buildListMapping(cellsRows) {
    const first = cellsRows[0] && cellsRows[0];
    const roles = first ? detectHeaderRoles(first) : null;
    if (roles && (roles.price !== undefined || roles.unitPrice !== undefined) && (roles.name !== undefined || roles.itemNo !== undefined)) {
      return {
        header: true,
        m: { itemNo: roles.itemNo, name: roles.name, unit: roles.unit, size: roles.size, color: roles.color, price: roles.price !== undefined ? roles.price : roles.unitPrice },
      };
    }
    return { header: false };
  }

  /**
   * استخراج بنود فاتورة من صفوف خلايا (من اللصق/Excel/Sheets).
   * السعر الغائب لا يُملأ من اللستة افتراضياً (autoPrice=false): البند يبقى بسعر 0
   * فيظهر «أقل من المعتمد» للمراجعة بدل أن يُخفى كـ «مطابق».
   * @returns {{ items:Array, ignored:number, autofilled:number, unknown:number, noPrice:number }}
   */
  function extractInvoiceItems(cellsRows, { autoPrice = false } = {}) {
    const { parseNum, isNumeric } = global.Utils;
    const map = buildInvoiceMapping(cellsRows);
    const items = [];
    let ignored = 0, autofilled = 0, unknown = 0, noPrice = 0;
    cellsRows.forEach((c, i) => {
      if (map.header && i === 0) return;
      const m = map.header ? map.m : invoicePosMapping(c);
      const itemNo = cv(c, m.itemNo);
      const name = cv(c, m.name);
      const unit = cv(c, m.unit);
      const size = cv(c, m.size);
      const color = cv(c, m.color);
      const qtyRaw = cv(c, m.qty);
      const prRaw = cv(c, m.unitPrice);
      const dvRaw = cv(c, m.discountValue);
      const dpRaw = cv(c, m.discountPct);
      const key = name || itemNo;
      if (!key) { ignored++; return; }
      // كمية بلا أي رقم (نص/رأس مكرر) = صف متجاهل؛ الكمية 0 تبقى 0 ولا تتحول إلى 1
      if (!isNumeric(qtyRaw)) { ignored++; return; }
      const qty = parseNum(qtyRaw);

      let price = isNumeric(prRaw) ? parseNum(prRaw) : 0;
      if (!(price > 0)) {
        const li = autoPrice && global.Lists ? global.Lists.findItem(itemNo, name, unit, size, color) : null;
        if (li) { price = global.Utils.num(li.price); autofilled++; }
        else { noPrice++; if (autoPrice) unknown++; }
      }
      // خصم القيمة: مبلغ ثابت لكل وحدة. خصم النسبة: الكسر المتسق مع سحب السيستم.
      // القيم ≥ 1 تُعتبر نسباً مئوية صحيحة (1 = 1%) فتُقسَّم على 100 لتوحيد التخزين.
      const discountValue = isNumeric(dvRaw) ? Math.max(0, parseNum(dvRaw)) : 0;
      let discountPct = isNumeric(dpRaw) ? Math.max(0, parseNum(dpRaw)) : 0;
      if (discountPct >= 1) discountPct = discountPct / 100;
      items.push({ itemNumber: itemNo, name: name || itemNo, unit, size, color, quantity: qty, unitPrice: price, discountValue, discountPct });
    });
    return { items, ignored, autofilled, unknown, noPrice };
  }

  /**
   * استخراج بنود لستة أسعار من صفوف خلايا.
   * @returns {{ rows:Array<{itemNumber,name,unit,size,color,price}>, ignored:number }}
   */
  function extractListRows(cellsRows) {
    const map = buildListMapping(cellsRows);
    const out = [];
    let ignored = 0;
    cellsRows.forEach((c, i) => {
      if (map.header && i === 0) return;
      let itemNo, name, unit, priceRaw, size = '', color = '';
      if (map.header) {
        itemNo = cv(c, map.m.itemNo);
        name = cv(c, map.m.name);
        unit = cv(c, map.m.unit);
        priceRaw = cv(c, map.m.price);
        size = cv(c, map.m.size);
        color = cv(c, map.m.color);
      } else {
        const n = c.length;
        if (n <= 1) { ignored++; return; }
        if (n === 2) { itemNo = ''; name = cv(c, 0); unit = ''; priceRaw = cv(c, 1); }
        else if (n === 3) { itemNo = cv(c, 0); unit = cv(c, 1); priceRaw = cv(c, 2); name = itemNo; }
        else { itemNo = cv(c, 0); name = cv(c, 1); unit = cv(c, 2); priceRaw = cv(c, 3); }
      }
      if (!name && !itemNo) { ignored++; return; }
      if (priceRaw === '' || !(global.Utils.parseNum(priceRaw) > 0)) { ignored++; return; }
      out.push({ itemNumber: itemNo.trim(), name: (name || itemNo).trim(), unit, size, color, price: global.Utils.parseNum(priceRaw) });
    });
    return { rows: out, ignored };
  }

  /* --- تحويل صفوف الخلايا إلى نص TSV (لعرضها في مربعات اللصق) --- */
  function tableToTSV(cellsRows, delimiter = '\t') {
    return cellsRows.map((r) => r.join(delimiter)).join('\n');
  }

  global.Parser = {
    splitTSVRow, parseRows, parseCSV, bestRole, detectHeaderRoles,
    buildInvoiceMapping, invoicePosMapping, buildListMapping,
    extractInvoiceItems, extractListRows, tableToTSV,
  };
})(window);