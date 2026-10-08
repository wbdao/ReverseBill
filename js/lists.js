/* ═══════════════════════════════════════════════════════════════════════
   js/lists.js — قوائم الأسعار المتعددة (State/Domain)
   • إدارة لستات مستقلة (إنشاء/تسمية/حذف/تفعيل/استيراد/تصدير)
   • فهرس Map مدمج للبحث السريع (O(1) تقريباً) — يدعم عشرات آلاف الأصناف
     index = { byCode: Map, byName: Map, exact: Map }  (exact: الكود/الاسم + الوحدة + الحجم + اللون)
   • المطابقة: رقم الصنف+الوحدة أولاً ثم الاسم+الوحدة — لا يُخلط سعر وحدة بغيرها
   • مخزن ذاكرة مركزي window.appLists = { [slug]: items[] } — مصدر الأصناف
     في الذاكرة فور الجلب (يتمزامن تلقائياً مع كل تغيير عبر persist) وتقرأ منه
     دوال البحث/السعر/المقارنة في app.js دون أي قراءة متكررة من localStorage
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const { fmtNum, genId, normalizeNum, normalizeName, normalizeUnit } = global.Utils;
  const Storage = global.Storage;

  let lists = [];
  let activeId = null;
  const indexCache = new Map(); // listId -> { ver, idx }
  // مخزن الذاكرة المركزي: { slug: items[] } — مرجع مباشر لصفيف أصناف اللستة (لا نسخ، يبقى متزامناً)
  const store = (global.appLists = {});

  /* --- مزامنة المخزن المركزي مع اللستات الحالية (تُستدعى من persist على كل تغيير) --- */
  function syncMemory() {
    for (const l of lists) { if (l.cfgSlug) store[l.cfgSlug] = l.items; }
    for (const k of Object.keys(store)) { if (!lists.some((l) => l.cfgSlug === k)) delete store[k]; }
    return store;
  }
  const storeGet = (slug) => (slug && store[slug]) || [];

  /* --- بحث مفهرس سريع من المخزن المركزي بمعرّف اللستة (slug) ---
     يفضّل مسار الفهرس Map عبر كائن اللستة، ويعتمد على المصفوفة المباشرة من المخزن كبديل. */
  function findInStore(slug, itemNumber, name, unit, size, color) {
    return resolveInStore(slug, itemNumber, name, unit, size, color).item;
  }
  function resolveInStore(slug, itemNumber, name, unit, size, color) {
    const l = slug ? lists.find((x) => x.cfgSlug === slug) : null;
    return resolveInList(l || storeGet(slug), itemNumber, name, unit, size, color);
  }

  const normName = (s) => String(s || '').trim();
  const nameKey = (s) => normName(s).toLocaleLowerCase('ar-EG');

  function mkList(name) {
    const l = { id: genId('lst'), name: normName(name) || 'لستة جديدة', items: [], createdAt: new Date().toISOString() };
    Object.defineProperty(l, '__ver', { value: 0, writable: true, enumerable: false, configurable: true });
    return l;
  }
  const bump = (l) => { if (l) { l.__ver = (l.__ver || 0) + 1; } };

  // الحجم واللون يُطبَّعان كالوحدة (حروف موحّدة + بلا حساسية حالة): "Gold" = "gold"
  function itemNorm(item) {
    return {
      nc: normalizeNum(item.itemNumber),
      nm: normalizeName(item.name),
      nu: normalizeUnit(item.unit),
      ns: normalizeUnit(item.size),
      ncl: normalizeUnit(item.color),
    };
  }
  /* --- مفتاح هوية الصنف في اللستة: (الكود أو الاسم) + الوحدة + الحجم + اللون ---
     نفس الكود والوحدة بحجم/لون مختلف = سطر مستقل بسعره (لا يُدمج). */
  function exactKey(item) {
    const { nc, nm, nu, ns, ncl } = itemNorm(item);
    return (nc ? 'n:' + nc : 'm:' + nm) + '\u0000' + nu + '\u0000' + ns + '\u0000' + ncl;
  }

  /* --- إضافة صنف إلى فهرس Map (يُستخدم عند البناء الضخم وأثناء الحشو) --- */
  function indexAdd(idx, it) {
    const { nc, nm } = itemNorm(it);
    if (nc) {
      if (!idx.byCode.has(nc)) idx.byCode.set(nc, []);
      idx.byCode.get(nc).push(it);
    }
    if (nm) {
      if (!idx.byName.has(nm)) idx.byName.set(nm, []);
      idx.byName.get(nm).push(it);
    }
    if (nc || nm) {
      const key = exactKey(it);
      if (!idx.exact.has(key)) idx.exact.set(key, it);
    }
  }

  /* --- بناء فهرس Map للبحث السريع على لستة كبيرة --- */
  function buildIndex(items) {
    const idx = { byCode: new Map(), byName: new Map(), exact: new Map() };
    if (!items || !items.length) return idx;
    for (const it of items) indexAdd(idx, it);
    return idx;
  }

  /* --- بيانات اللستة التي ستُبنى على أساسها المطابقة --- */
  function dataOf(list) {
    if (!list) return null;
    const items = Array.isArray(list) ? list : list.items;
    if (!Array.isArray(items)) return null;
    const isArray = Array.isArray(list);
    if (isArray) return { items, idx: null, isArray };
    const cached = indexCache.get(list.id);
    if (cached && cached.ver === list.__ver) return { items, idx: cached.idx, isArray };
    const idx = buildIndex(items);
    indexCache.set(list.id, { ver: list.__ver, idx });
    return { items, idx, isArray };
  }

  /* --- تضييق المرشحين بالحجم/اللون ---
     • البعد (حجم أو لون) يُطبَّق فقط إن كان مسجلاً لبعض المرشحين — صنف بلا متغيرات
       في اللستة يُطابَق مهما كتبت الفاتورة في الحجم/اللون.
     • حجم/لون معطى ولا يطابق أي متغير مسجل = غير مسجل (لا نسحب سعر متغير آخر).
     • لم يُعطَ حجم/لون والمتغيرات المتبقية بأسعار مختلفة = «يحتاج تحديد» (ambiguous)؛
       إن تساوت أسعارها فلا تعارض ونطابق أولها. */
  function pickVariant(cands, ns, ncl) {
    let c = cands;
    const narrow = (val, field) => {
      if (!c.some((it) => normalizeUnit(it[field]))) return; // بُعد غير مستخدم لهذا الصنف
      if (val) c = c.filter((it) => normalizeUnit(it[field]) === val);
    };
    narrow(ns, 'size');
    narrow(ncl, 'color');
    if (!c.length) return { item: null, ambiguous: false };
    if (c.length === 1) return { item: c[0], ambiguous: false };
    const p0 = Number(c[0].price) || 0;
    if (c.every((it) => Math.abs((Number(it.price) || 0) - p0) < 1e-9)) return { item: c[0], ambiguous: false };
    return { item: null, ambiguous: true, variants: c };
  }

  /**
   * المطابقة الدقيقة ضمن لستة (كائن لستة أو مصفوفة items) — تُرجع { item, ambiguous }.
   * المفتاح: رقم الصنف (أو الاسم عند غياب الكود) → الوحدة → الحجم/اللون.
   * • الكود غير موجود باللستة = «غير مسجل» (لا رجوع للاسم: كود قديم/جديد بنفس الاسم).
   * • وحدة معطاة بلا تطابق: لا يُستعمل سعر وحدة أخرى أبداً، مع السماح بإدخال بلا وحدة.
   * • بلا وحدة في البند: نطابق وحدة أول إدخال (السلوك السابق).
   */
  function resolveInList(list, itemNumber, name, unit, size, color) {
    const none = { item: null, ambiguous: false };
    const D = dataOf(list);
    if (!D || !D.items.length) return none;
    const nc = normalizeNum(itemNumber);
    const nm = normalizeName(name);
    const nu = normalizeUnit(unit);

    let cands;
    if (nc) {
      cands = D.idx ? (D.idx.byCode.get(nc) || []) : D.items.filter((it) => normalizeNum(it.itemNumber) === nc);
    } else if (nm) {
      cands = D.idx ? (D.idx.byName.get(nm) || []) : D.items.filter((it) => normalizeName(it.name) === nm);
    } else return none;
    if (!cands.length) return none;

    if (nu) {
      const byU = cands.filter((it) => normalizeUnit(it.unit) === nu);
      cands = byU.length ? byU : cands.filter((it) => !normalizeUnit(it.unit));
    } else {
      const u0 = normalizeUnit(cands[0].unit);
      cands = cands.filter((it) => normalizeUnit(it.unit) === u0);
    }
    if (!cands.length) return none;
    return pickVariant(cands, normalizeUnit(size), normalizeUnit(color));
  }
  const matchInList = (list, itemNumber, name, unit, size, color) => resolveInList(list, itemNumber, name, unit, size, color).item;

  /* ================= إدارة اللستات ================= */
  function init() {
    const data = Storage.getPriceLists();
    lists = data.lists;
    activeId = data.activeId;

    // بادرة إصدار: عند تغيّر بنية البيانات المخزنة (اندماج وحدات قديم مثلاً) تُجبر
    // كل اللستات على إعادة المزامنة من المصدر مرة واحدة — تُحدَّث القيمة في persist التالي.
    const expectVer = global.CONFIG.LIST_DATA_VERSION || 1;
    if (!lists.length || data.version !== expectVer) {
      for (const l of lists) { l.syncedAt = null; }
    }

    if (!lists.length) {
      const legacy = Storage.getLegacyProducts();
      if (Array.isArray(legacy) && legacy.length) {
        const l = mkList('قائمة الأسعار الرئيسية');
        l.items = legacy.map((p) => ({
          id: genId('li'),
          itemNumber: normName(p.itemNumber ?? p.code ?? ''),
          name: normName(p.name ?? ''),
          unit: normName(p.unit ?? ''),
          price: Number(p.basePrice ?? p.price ?? 0) || 0,
        }));
        lists.push(l);
      }
    }
    if (!lists.length) seed();
    if (!activeId || !getList(activeId)) activeId = lists.length ? lists[0].id : null;
    indexCache.clear();
    lists.forEach(bump);
    persist();
    return { seeded: lists.length > 0 };
  }

  function persist() {
    syncMemory(); // المخزن في الذاكرة يتمزامن مع كل كتابة (بدون أي قراءة من التخزين هنا)
    Storage.savePriceLists(lists, activeId);
  }

  const all = () => lists;
  const count = () => lists.length;
  const getList = (id) => lists.find((l) => l.id === id) || null;
  const active = () => getList(activeId);
  const activeIdOf = () => activeId;
  const itemsOf = (id) => { const l = getList(id); return l ? l.items : []; };

  function create(name) {
    const l = mkList(name);
    lists.push(l);
    activeId = l.id;
    persist();
    return l;
  }
  function rename(id, name) {
    const l = getList(id);
    if (!l) return;
    const n = normName(name);
    if (n && !lists.some((x) => x.id !== id && nameKey(x.name) === nameKey(n))) { l.name = n; }
    persist();
  }
  /** إيجاد لستة بالاسم أو إنشاؤها دون تغيير اللستة النشطة (للتخزين السحابي) */
  function ensureList(name) {
    const n = normName(name);
    const found = lists.find((l) => nameKey(l.name) === nameKey(n));
    if (found) return { list: found, created: false };
    const l = mkList(n);
    lists.push(l);
    bump(l);
    persist();
    return { list: l, created: true };
  }
  function remove(id) {
    lists = lists.filter((l) => l.id !== id);
    indexCache.delete(id);
    if (activeId === id) activeId = lists.length ? lists[0].id : null;
    persist();
  }
  function setActive(id) { if (getList(id)) { activeId = id; persist(); } }

  /* === النموذج "اللستات سحابية فقط" ===
     يفرض قائمة ثابتة من اللستات (من CONFIG.LISTS) بترتيب محدد:
     يحتفظ بمعرّف وأصناف أي لستة موجودة تطابق الاسم حتى تبقى إشارات الفواتير صالحة،
     ويلحق بيانات الربط السحابي (url/tab والأيقونة)، ويحذف أي لستة خارجة عن القائمة. */
  function reconcile(templates) {
    const next = [];
    for (const t of templates || []) {
      const found = lists.find((l) => l.cfgSlug === t.slug || nameKey(l.name) === nameKey(t.name));
      const l = found || mkList(t.name);
      l.cfgSlug = t.slug;
      l.cfgName = t.name;
      l.cfgUrl = t.url || '';
      l.cfgTab = t.tab || '';
      l.cfgColor = t.color || 'indigo';
      l.cfgIcon = t.icon || 'fa-list';
      next.push(l);
    }
    lists = next;
    if (!activeId || !getList(activeId)) activeId = lists.length ? lists[0].id : null;
    indexCache.clear();
    lists.forEach(bump);
    persist();
    return next;
  }

  /* ================= صنف واحد ================= */
  function buildItem(data) {
    return {
      id: genId('li'),
      itemNumber: normName(data.itemNumber ?? ''),
      name: normName(data.name ?? ''),
      unit: normName(data.unit ?? ''),
      size: normName(data.size ?? ''),
      color: normName(data.color ?? ''),
      price: Number(data.price) || 0,
    };
  }

  // نفس الصنف = نفس (الكود أو الاسم) + الوحدة + الحجم + اللون
  const sameKey = (a, b) => exactKey(a) === exactKey(b);

  function upsertItem(data) {
    const l = active();
    if (!l) throw new Error('لا توجد لستة نشطة.');
    const item = buildItem(data);
    if (!item.name && !item.itemNumber) throw new Error('مطلوب اسم منتج أو رقم صنف.');
    const dup = l.items.find((x) => sameKey(x, item));
    if (dup) {
      Object.assign(dup, item, { id: dup.id });
      bump(l); persist();
      return { item: dup, updated: true };
    }
    l.items.push(item);
    bump(l); persist();
    return { item, updated: false };
  }

  function updateItem(id, data) {
    const l = active();
    if (!l) throw new Error('لا توجد لستة نشطة.');
    const it = l.items.find((x) => x.id === id);
    if (!it) throw new Error('الصنف غير موجود.');
    Object.assign(it, buildItem({ ...it, ...data }));
    bump(l); persist();
  }

  function removeItem(id) {
    const l = active();
    if (!l) return;
    l.items = l.items.filter((x) => x.id !== id);
    bump(l); persist();
  }
  function clearList(id) {
    const l = getList(id);
    if (l) { l.items = []; bump(l); persist(); }
  }
  function clearAllByReset() {
    lists = [mkList('لستة جديدة')];
    activeId = lists[0].id;
    indexCache.clear();
    persist();
  }

  const findItem = (itemNumber, name, unit, size, color) => {
    const l = active();
    const slug = l && l.cfgSlug;
    // القراءة من المخزن المركزي (window.appLists) مع مسار الفهرس السريع
    return slug ? findInStore(slug, itemNumber, name, unit, size, color) : matchInList(l, itemNumber, name, unit, size, color);
  };

  /* ================= إدخال / استيراد ضخم (بأجزاء) ================= */
  /* --- المكرر أثناء الحشو = نفس مفتاح الهوية تماماً (الكود/الاسم + الوحدة + الحجم + اللون) ---
     كود غير موجود أو متغير حجم/لون جديد → بند مستقل بسعره، فلا يُدمج كود قديم/جديد
     أو متغيران بسعرين مختلفين في سطر واحد. */
  function lookupItem(idx, items, item) {
    if (idx) return idx.exact.get(exactKey(item)) || null;
    const key = exactKey(item);
    return items.find((it) => exactKey(it) === key) || null;
  }

  /* --- قلب الاستيراد: يحشر شريحة rows على نفس الفهرس المبني مرة واحدة --- */
  function importRowsCore(l, rows, start, end) {
    const D = dataOf(l);
    const idx = D.idx;
    let added = 0, updated = 0;
    for (let i = start; i < end; i++) {
      const item = buildItem(rows[i]);
      if (!item.name && !item.itemNumber) continue;
      const dup = lookupItem(idx, D.items, item);
      if (dup) { Object.assign(dup, item, { id: dup.id }); updated++; }
      else { l.items.push(item); if (idx) indexAdd(idx, item); added++; }
    }
    return { added, updated };
  }

  /**
   * استيراد صفوف {itemNumber,name,unit,price} إلى لستة محددة مع تحديث المكررات.
   * يستخدم فهرس المطابقة لتفادي البحث الخطي أثناء الحشو (بدون إعادة بناء لكل صف).
   */
  function importRows(listId, rows, { replace = false } = {}) {
    const l = getList(listId);
    if (!l) throw new Error('لستة الهدف غير موجودة.');
    if (replace) return replaceFrom(l, rows, (h) => { importRowsCore(h, rows, 0, rows.length); });
    const res = importRowsCore(l, rows, 0, rows.length);
    if (res.added || res.updated) { bump(l); persist(); }
    return { added: res.added, updated: res.updated, total: l.items.length };
  }

  /* --- استبدال أصناف اللستة كاملة بنسخة المصدر (مزامنة سحابية) ---
     يُبنى في حاوية مؤقتة ثم يُبدَّل دفعة واحدة: الأصناف المحذوفة من الشيت تختفي،
     ولا تُقرأ لستة نصف ممتلئة أثناء الحشو. مصدر فارغ لا يمسح اللستة القائمة. */
  async function replaceFrom(l, rows, fill) {
    if (!rows.length) throw new Error('المصدر لم يُرجع أي صنف صالح — أُبقيت اللستة كما هي.');
    const h = mkList(l.name);
    h.id = l.id + '__incoming';
    await fill(h);
    indexCache.delete(h.id);
    const previous = l.items.length;
    l.items = h.items;
    bump(l); persist();
    return { added: 0, updated: 0, total: l.items.length, previous, replaced: true };
  }

  /**
   * استيراد ضخم على أجزاء متقاطعة مع الخيط الرئيسي (عدم حجب الواجهة أثناء تحميل الخلفية).
   * يقسم الحشو إلى شرائح ويُفرّغ الخيط بينها عبر requestIdleCallback (السقوط لـ setTimeout).
   */
  async function importRowsChunked(listId, rows, { chunkSize = 400, idle = true, replace = false } = {}) {
    const l = getList(listId);
    if (!l) throw new Error('لستة الهدف غير موجودة.');
    let added = 0, updated = 0;
    const yieldNext = () => new Promise((res) => {
      const ric = (typeof self !== 'undefined' ? self.requestIdleCallback : null)
        || (typeof window !== 'undefined' ? window.requestIdleCallback : null);
      if (idle && ric) ric(res, { timeout: 50 });
      else setTimeout(res, 0);
    });
    const fillChunked = async (target) => {
      let i = 0;
      while (i < rows.length) {
        const res = importRowsCore(target, rows, i, Math.min(rows.length, i + chunkSize));
        added += res.added; updated += res.updated;
        i += chunkSize;
        if (i < rows.length) await yieldNext();
      }
    };
    if (replace) return replaceFrom(l, rows, fillChunked);
    await fillChunked(l);
    if (added || updated) { bump(l); persist(); }
    return { added, updated, total: l.items.length };
  }

  /* ================= تصدير / استيراد JSON وTSV ================= */
  const exportAllJson = () => JSON.stringify(lists, null, 2);
  const exportListJson = (id) => JSON.stringify({ name: (getList(id) || {}).name || '', items: itemsOf(id) }, null, 2);
  const listToTSV = (id) => {
    const l = getList(id);
    const header = 'رقم الصنف\tاسم المنتج\tالوحدة\tالحجم\tاللون\tالسعر المعتمد';
    if (!l || !l.items.length) return header;
    return [header, ...l.items.slice(0, 20001).map((it) => `${it.itemNumber || ''}\t${it.name}\t${it.unit || ''}\t${it.size || ''}\t${it.color || ''}\t${fmtNum(it.price)}`)].join('\n');
  };

  function importAllJson(text, replace = false) {
    const raw = JSON.parse(text);
    const arr = Array.isArray(raw) ? raw
      : (Array.isArray(raw.lists) ? raw.lists
      : (raw && typeof raw === 'object' && Array.isArray(raw.items) ? [raw] : null));
    if (!arr) throw new Error('صيغة JSON غير صحيحة: يجب أن تحتوي على مصفوفة لستات.');
    if (replace) { lists = []; indexCache.clear(); }
    let added = 0;
    for (const rl of arr) {
      if (!rl || typeof rl !== 'object') continue;
      let list = lists.find((l) => l.id === rl.id) || (rl.name && lists.find((l) => nameKey(l.name) === nameKey(rl.name)));
      if (!list) {
        list = mkList(rl.name || 'لستة مستوردة');
        if (rl.id) list.id = rl.id;
        lists.push(list);
        added++;
      }
      const seen = new Set();
      for (const rit of Array.isArray(rl.items) ? rl.items : []) {
        const item = {
          id: rit.id || genId('li'),
          itemNumber: normName(rit.itemNumber ?? ''),
          name: normName(rit.name ?? ''),
          unit: normName(rit.unit ?? ''),
          size: normName(rit.size ?? ''),
          color: normName(rit.color ?? ''),
          price: Number(rit.price ?? rit.basePrice ?? 0) || 0,
        };
        if (!item.name && !item.itemNumber) continue;
        const dup = list.items.find((x) => x.id === item.id) || list.items.find((x) => !seen.has(x.id) && sameKey(x, item));
        if (dup) { Object.assign(dup, item); seen.add(dup.id); }
        else { list.items.push(item); seen.add(item.id); }
      }
      bump(list);
    }
    if (!lists.length) seed();
    if (!activeId || !getList(activeId)) activeId = lists.length ? lists[0].id : null;
    lists.forEach(bump);
    indexCache.clear();
    persist();
    return { added };
  }

  /* ================= البيانات التجريبية ================= */
  function seed() {
    lists = global.CONFIG.SEED_LISTS.map((s) => {
      const l = mkList(s.name);
      l.items = s.rows.map(([itemNumber, pname, unit, price]) => ({ id: genId('li'), itemNumber, name: pname, unit, price: Number(price) || 0 }));
      return l;
    });
    activeId = lists.length ? lists[0].id : null;
    indexCache.clear();
    persist();
  }

  global.Lists = {
    init, all, count, getList, active, activeIdOf, create, rename, remove, setActive,
    itemsOf, ensureList, upsertItem, updateItem, removeItem, clearList, clearAllByReset, findItem, matchInList,
    reconcile,
    syncMemory, storeGet, findInStore, resolveInList, resolveInStore, exactKey,
    importRows, importRowsChunked, exportAllJson, exportListJson, listToTSV, importAllJson, seed, buildIndex,
  };
})(window);