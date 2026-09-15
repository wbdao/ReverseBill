/* ═══════════════════════════════════════════════════════════════════════
   js/lists.js — قوائم الأسعار المتعددة (State/Domain)
   • إدارة لستات مستقلة (إنشاء/تسمية/حذف/تفعيل/استيراد/تصدير)
   • فهرس Map مدمج للبحث السريع (O(1) تقريباً) — يدعم عشرات آلاف الأصناف
     index = { byCode: Map, byName: Map, codeUnit: Map, nameUnit: Map }
   • المطابقة الدقيقة: رقم الصنف+الوحدة أولاً ثم الاسم+الوحدة
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
  function findInStore(slug, itemNumber, name, unit) {
    const l = slug ? lists.find((x) => x.cfgSlug === slug) : null;
    return l ? matchInList(l, itemNumber, name, unit) : matchInList(storeGet(slug), itemNumber, name, unit);
  }

  const normName = (s) => String(s || '').trim();
  const nameKey = (s) => normName(s).toLocaleLowerCase('ar-EG');

  function mkList(name) {
    const l = { id: genId('lst'), name: normName(name) || 'لستة جديدة', items: [], createdAt: new Date().toISOString() };
    Object.defineProperty(l, '__ver', { value: 0, writable: true, enumerable: false, configurable: true });
    return l;
  }
  const bump = (l) => { if (l) { l.__ver = (l.__ver || 0) + 1; } };

  function itemNorm(item) {
    return {
      nc: normalizeNum(item.itemNumber),
      nm: normalizeName(item.name),
      nu: normalizeUnit(item.unit),
    };
  }

  /* --- بناء فهرس Map للبحث السريع على لستة كبيرة --- */
  function buildIndex(items) {
    const byCode = new Map(), byName = new Map(), codeUnit = new Map(), nameUnit = new Map();
    if (!items || !items.length) return { byCode, byName, codeUnit, nameUnit };
    for (const it of items) {
      const { nc, nm, nu } = itemNorm(it);
      if (nc) {
        if (!byCode.has(nc)) byCode.set(nc, []);
        byCode.get(nc).push(it);
        const key = nc + '\u0000' + nu;
        if (!codeUnit.has(key)) codeUnit.set(key, it);
      }
      if (nm) {
        if (!byName.has(nm)) byName.set(nm, []);
        byName.get(nm).push(it);
        const key = nm + '\u0000' + nu;
        if (!nameUnit.has(key)) nameUnit.set(key, it);
      }
    }
    return { byCode, byName, codeUnit, nameUnit };
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

  /**
   * المطابقة الدقيقة ضمن لستة (كائن لستة أو مصفوفة items).
   * الأولوية: رقم الصنف+الوحدة → الاسم+الوحدة → رقم الصنف → الاسم.
   */
  function matchInList(list, itemNumber, name, unit) {
    const D = dataOf(list);
    if (!D || !D.items.length) return null;
    const nc = normalizeNum(itemNumber);
    const nm = normalizeName(name);
    const nu = normalizeUnit(unit);

    if (D.idx) {
      // == المسار السريع (فهرس Map) — مناسب للبيانات الضخمة ==
      if (nc) {
        let arr = D.idx.byCode.get(nc);
        if (nu && arr && arr.length) {
          const byU = arr.filter((it) => normalizeUnit(it.unit) === nu);
          if (byU.length) return byU[0];
          return arr[0]; // رقم مطابق لكن الوحدة مختلفة: نفضّل إرجاع الصنف بدل اعتباره غير مسجل
        }
        if (arr && arr.length) return arr[0];
        if (nm) { const nmArr = D.idx.byName.get(nm); if (nmArr && nmArr.length) return nmArr[0]; }
        return null;
      }
      if (nm) {
        if (nu) {
          const nHit = D.idx.nameUnit.get(nm + '\u0000' + nu);
          if (nHit) return nHit;
          const nmArr = D.idx.byName.get(nm);
          if (nmArr && nmArr.length) return nmArr[0];
          return null;
        }
        const arr = D.idx.byName.get(nm);
        if (arr && arr.length) return arr[0];
        return null;
      }
      return null;
    }

    // == المسار الاحتياطي (مصفوفة مباشرة) ==
    const items = D.items;
    if (nc) {
      const byNum = items.filter((it) => normalizeNum(it.itemNumber) === nc);
      if (byNum.length) {
        if (nu) { const byU = byNum.filter((it) => normalizeUnit(it.unit) === nu); if (byU.length) return byU[0]; }
        return byNum[0];
      }
    }
    if (nm) {
      const byName = items.filter((it) => normalizeName(it.name) === nm);
      if (byName.length) {
        if (nu) { const byU = byName.filter((it) => normalizeUnit(it.unit) === nu); if (byU.length) return byU[0]; }
        return byName[0];
      }
    }
    return null;
  }

  /* ================= إدارة اللستات ================= */
  function init() {
    const data = Storage.getPriceLists();
    lists = data.lists;
    activeId = data.activeId;

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
      price: Number(data.price) || 0,
    };
  }

  function sameKey(a, b) {
    const nkA = normalizeNum(a.itemNumber), nkB = normalizeNum(b.itemNumber);
    if (nkA && nkB && nkA === nkB) {
      if (!a.unit && !b.unit) return true;
      return normalizeUnit(a.unit) === normalizeUnit(b.unit);
    }
    const nmA = normalizeName(a.name), nmB = normalizeName(b.name);
    if (nmA && nmB && nmA === nmB) {
      if (!a.unit && !b.unit) return true;
      return normalizeUnit(a.unit) === normalizeUnit(b.unit);
    }
    return false;
  }

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

  const findItem = (itemNumber, name, unit) => {
    const l = active();
    const slug = l && l.cfgSlug;
    // القراءة من المخزن المركزي (window.appLists) مع مسار الفهرس السريع
    return slug ? findInStore(slug, itemNumber, name, unit) : matchInList(l, itemNumber, name, unit);
  };

  /* ================= إدخال / استيراد ضخم (بأجزاء) ================= */
  /**
   * استيراد صفوف {itemNumber,name,unit,price} إلى لستة محددة مع تحديث المكررات.
   * يستخدم فهرس المطابقة لتفادي البحث الخطي أثناء الحشو الكبير.
   */
  function importRows(listId, rows) {
    const l = getList(listId);
    if (!l) throw new Error('لستة الهدف غير موجودة.');
    let added = 0, updated = 0;
    for (const r of rows) {
      const item = buildItem(r);
      if (!item.name && !item.itemNumber) continue;
      const dup = matchInList(l, item.itemNumber, item.name, item.unit);
      if (dup) { Object.assign(dup, item, { id: dup.id }); updated++; }
      else { l.items.push(item); added++; }
      bump(l);
    }
    if (added || updated) { bump(l); persist(); }
    return { added, updated, total: l.items.length };
  }

  /* ================= تصدير / استيراد JSON وTSV ================= */
  const exportAllJson = () => JSON.stringify(lists, null, 2);
  const exportListJson = (id) => JSON.stringify({ name: (getList(id) || {}).name || '', items: itemsOf(id) }, null, 2);
  const listToTSV = (id) => {
    const l = getList(id);
    const header = 'رقم الصنف\tاسم المنتج\tالوحدة\tالسعر المعتمد';
    if (!l || !l.items.length) return header;
    return [header, ...l.items.slice(0, 20001).map((it) => `${it.itemNumber || ''}\t${it.name}\t${it.unit || ''}\t${fmtNum(it.price)}`)].join('\n');
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
    syncMemory, storeGet, findInStore,
    importRows, exportAllJson, exportListJson, listToTSV, importAllJson, seed, buildIndex,
  };
})(window);