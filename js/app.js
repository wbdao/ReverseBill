/* ═══════════════════════════════════════════════════════════════════════
   js/app.js — طبقة الواجهة والتحكم (UI) — تُحمَّل أخيراً
   • العرض والأحداث لكل الشاشات + الإقلاع
   • التبويبات، جداول الفاتورة/اللستة/السجل، التقرير، الطباعة
   • السحب السحابي (Google Sheets) + استيراد Excel محلي
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const {
    $, $$, fmtNum, money, signedMoney, signedNum, parseNum, num, esc,
    todayStr, genId, toast, confirmBox, copyText, downloadJSON, downloadTSV, importFile,
  } = global.Utils;
  const CONFIG = global.CONFIG;
  const Storage = global.Storage;
  const Parser = global.Parser;
  const Lists = global.Lists;
  const Invoices = global.Invoices;
  const Comparison = global.Comparison;
  const Sheets = global.Sheets;
  const Excel = global.Excel;
  const { State } = global;

  /* ─────────────────── تنقل وتبويبات ─────────────────── */
  function switchTab(tab) {
    $$('[data-view]').forEach((v) => v.classList.add('hidden'));
    $$('.tab-btn').forEach((b) => b.classList.toggle('active', b.dataset.tab === tab));
    const view = $(`#view-${tab}`);
    if (view) view.classList.remove('hidden');
  }
  const activeListName = () => { const l = Lists.active(); return l ? l.name : '—'; };
  const activeSlugOf = () => { const l = Lists.active(); return l ? l.cfgSlug || '' : ''; };

  /* ─────────────────── قوائم الأسماء / شرائح اللستات ─────────────────── */
  const renderDatalist = () => {
    // القراءة من مخزن الذاكرة المركزي window.appLists لاقتراحات البحث الفورية
    const l = Lists.active();
    const items = (l && l.cfgSlug && global.appLists && global.appLists[l.cfgSlug])
      ? global.appLists[l.cfgSlug]
      : (l ? l.items : []);
    $('#product-names').innerHTML = items.slice(0, 500)
      .map((p) => `<option value="${esc(p.name)}"></option>`).join('');
  };

  function renderQuickList() {
    $('#quick-list').innerHTML = Lists.all()
      .map((l) => `<option value="${esc(l.id)}"${l.id === Lists.activeIdOf() ? ' selected' : ''}>${esc(l.name)}</option>`)
      .join('');
    const el = $('#active-list-name');
    if (el) el.textContent = activeListName();
  }

  /* ─────────────────── اللستات السحابية: بطاقات + الجلب التلقائي ─────────────────── */
  const cloudBusy = {};
  const CLOUD_LIST_HINT = 'رابط اللستة يجب أن يكون «منشوراً على الويب» في Google Sheets (روابط التعديل /edit أو العرض /view تُنظَّف تلقائياً إلى نشر CSV، لكن الجدول نفسه يجب أن يكون منشوراً). عيّن الرابط الصحيح في js/config.js ضمن CONFIG.LISTS.';

  const LIST_COLORS = {
    indigo: { border: 'border-indigo-200', bg: 'bg-indigo-50', text: 'text-indigo-600' },
    emerald: { border: 'border-emerald-200', bg: 'bg-emerald-50', text: 'text-emerald-600' },
    sky: { border: 'border-sky-200', bg: 'bg-sky-50', text: 'text-sky-600' },
    amber: { border: 'border-amber-200', bg: 'bg-amber-50', text: 'text-amber-600' },
    rose: { border: 'border-rose-200', bg: 'bg-rose-50', text: 'text-rose-600' },
  };

  function listBySlug(slug) {
    const cfgs = CONFIG.LISTS || [];
    const i = cfgs.findIndex((c) => c.slug === slug);
    return { cfg: cfgs[i] || null, list: Lists.all()[i] || null };
  }

  /* --- كشف gid لكل تبويب من قائمة التبويبات المنشورة (استدعاء واحد pubhtml لكل الجلسة) ---
     كل اللستات في نفس المستند: نجلب قائمة التبويبات مرة واحدة ونطابقها بأسماء cfg.tab ثم
     نخزن النتيجة على cfg.gid (بالذاكرة) حتى تستخدم كل بطاقة رابطاً دقيقاً بالـ gid. */
  async function resolveCloudGids() {
    const missing = (CONFIG.LISTS || []).filter((c) => c.url && !c.gid);
    if (!missing.length) return;
    let menu = [];
    try { menu = await Sheets.listTabs(missing[0].url); } catch (e) { menu = []; }
    if (!menu.length) return;
    const keyOf = (s) => Sheets.normSheet(s);
    for (const cfg of missing) {
      const want = keyOf(cfg.tab || cfg.slug || '');
      if (!want) continue;
      const found = menu.find((m) => keyOf(m.name) === want) || menu.find((m) => keyOf(m.name).includes(want));
      if (found && found.gid) cfg.gid = found.gid;
    }
  }

  function renderCloudLists() {
    const grid = $('#cloud-lists-grid');
    if (!grid) return;
    const arr = Lists.all();
    grid.innerHTML = (CONFIG.LISTS || []).map((cfg, i) => {
      const l = arr[i];
      const active = !!(l && l.id === Lists.activeIdOf());
      const c = LIST_COLORS[cfg.color] || LIST_COLORS.indigo;
      const cnt = l ? l.items.length : 0;
      const syncedAt = l && l.syncedAt ? new Date(l.syncedAt).toLocaleString('ar-EG') : null;
      const busy = !!cloudBusy[cfg.slug];
      let statusHTML;
      if (!cfg.url) statusHTML = '<span class="text-amber-600"><i class="fa-solid fa-triangle-exclamation"></i> لم يُحدد الرابط — عيّنه في js/config.js</span>';
      else if (busy) statusHTML = '<span class="text-indigo-600"><i class="fa-solid fa-spinner fa-spin"></i> جارٍ الجلب...</span>';
      else if (l && l.lastError) statusHTML = `<span class="text-rose-600"><i class="fa-solid fa-circle-xmark"></i> ${esc(l.lastError)}</span>`;
      else statusHTML = `<span class="text-slate-600"><i class="fa-solid fa-check-circle text-emerald-600"></i> ${cnt} صنف${syncedAt ? ` · آخر تحديث ${syncedAt}` : ' · لم تُجلب بعد'}</span>`;
      return `
        <div data-list-id="${l ? esc(l.id) : ''}" data-slug="${esc(cfg.slug)}"
             class="rounded-2xl border-2 ${active ? c.border + ' ring-2 ring-indigo-200' : 'border-slate-200'} bg-white shadow-sm overflow-hidden cursor-pointer hover:shadow-md transition"
             title="انقر لاختيار هذه اللستة كلستة مقارنة نشطة">
          <div class="flex items-center justify-between px-4 py-3 ${c.bg}">
            <div class="flex items-center gap-3 min-w-0">
              <span class="w-10 h-10 shrink-0 rounded-xl ${c.bg} ${c.text} grid place-items-center text-lg"><i class="fa-solid ${esc(cfg.icon || 'fa-list')}"></i></span>
              <div class="min-w-0">
                <p class="font-extrabold text-slate-800 text-sm truncate">${esc(cfg.name)}</p>
                <p class="text-[11px] text-slate-400 truncate" dir="ltr">${esc(cfg.url || 'رابط غير مضبوط بعد')}${cfg.gid ? ` · gid=${esc(cfg.gid)}` : ''}</p>
              </div>
            </div>
            <span class="status-badge ${active ? 'status-match' : 'bg-slate-100 text-slate-500'}">${active ? 'النشطة' : 'لستة'}</span>
          </div>
          <div class="px-4 py-3 space-y-2">
            <p class="text-xs font-bold text-slate-600 leading-5">${statusHTML}</p>
            <button data-refresh="${esc(cfg.slug)}" class="btn btn-ghost w-full text-xs py-2" ${busy ? 'disabled' : ''}>
              <i class="fa-solid fa-rotate"></i> تحديث هذه اللستة
            </button>
          </div>
        </div>`;
    }).join('') || '<p class="text-sm text-slate-400">لا توجد لستات مضبوطة في CONFIG.LISTS.</p>';
  }

  function activateList(id) {
    if (!id) return;
    Lists.setActive(id);
    State.settings.activeListId = Lists.activeIdOf();
    Storage.saveSettings(State.settings);
    State.listPage = 0;
    renderQuickList(); renderDatalist(); renderCloudLists(); renderListItems(); updateAllRowsAndSummary();
  }

  async function fetchListFromConfig(slug, { silent = false, defer = false } = {}) {
    const { cfg, list } = listBySlug(slug);
    if (!cfg || !list) return;
    if (!cfg.url) {
      list.lastError = 'لم يُحدد الرابط في js/config.js (CONFIG.LISTS)';
      Storage.savePriceLists(Lists.all(), Lists.activeIdOf());
      renderCloudLists();
      if (!silent) toast(`«${cfg.name}»: ${list.lastError}`, 'error');
      return;
    }
    cloudBusy[slug] = true;
    renderCloudLists();
    try {
      await resolveCloudGids();
      const gid = cfg.gid || '';
      // تبويب محدد بدقة عبر معامل gid (نفس المستند، خمس تبويبات مختلفة)
      const rowsMatrix = gid
        ? await Sheets.fetchTab(cfg.url, { gid })
        : await Sheets.fetchTab(cfg.url, { name: cfg.tab || cfg.slug });
      const { rows } = Parser.extractListRows(rowsMatrix);
      // الخلفية تُدخل بأجزاء متقاطعة مع الخيط (idle)؛ النشطة فورياً وكلها مرّة واحدة
      const res = defer
        ? await Lists.importRowsChunked(list.id, rows)
        : Lists.importRows(list.id, rows);
      list.syncedAt = new Date().toISOString();
      list.lastError = null;
      Storage.savePriceLists(Lists.all(), Lists.activeIdOf());
      renderCloudLists();
      // إعادة الرسم الكاملة للمسار النشط الفوري فقط؛ الخلفية تكتفي بحالة البطاقات
      if (!defer) { renderListItems(); renderDatalist(); renderQuickList(); updateAllRowsAndSummary(); }
      if (!silent) toast(`✓ «${cfg.name}»: ${res.total} صنف${res.added ? ` · ${res.added} جديد` : ''}${res.updated ? ` · ${res.updated} تحديث` : ''}`);
    } catch (err) {
      list.lastError = (err && err.message) ? err.message : 'فشل الجلب';
      Storage.savePriceLists(Lists.all(), Lists.activeIdOf());
      renderCloudLists();
      toast(`✗ «${cfg.name}»: ${list.lastError} — ${CLOUD_LIST_HINT}`, 'error');
    } finally {
      delete cloudBusy[slug];
      renderCloudLists();
      if (!defer) runAllListCompareSilent();
    }
  }

  /* --- يحمّل اللستة النشطة فوراً ثم يجدول الباقي في الخلفية دون حجب الواجهة --- */
  function scheduleCloudFetch(cfgs, onAllDone) {
    const activeSlug = activeSlugOf();
    const sorted = cfgs.slice().sort((a, b) => (a.slug === activeSlug ? -1 : 0) - (b.slug === activeSlug ? -1 : 0));
    if (!sorted.length) return Promise.resolve();
    const [first, ...rest] = sorted;
    return (async () => {
      await fetchListFromConfig(first.slug, { silent: true });
      let done = 1;
      const total = sorted.length;
      const settle = () => {
        done++;
        if (done >= total) { runAllListCompareSilent(); if (onAllDone) onAllDone(); }
      };
      for (const cfg of rest) fetchListFromConfig(cfg.slug, { silent: true, defer: true }).then(settle, settle);
      if (!rest.length) settle();
    })();
  }

  async function refreshAllCloudLists() {
    const cfgs = (CONFIG.LISTS || []).filter((c) => c.url);
    if (!cfgs.length) { toast('لا توجد روابط مضبوطة بعد — افتح js/config.js', 'error'); return; }
    toast(`جارٍ تحديث ${cfgs.length} لستة...`, 'info');
    await scheduleCloudFetch(cfgs, () => toast('اكتمل تحديث اللستات السحابية'));
  }

  async function autoFetchCloudLists() {
    const hours = CONFIG.LISTS_REFRESH_HOURS || 6;
    const th = hours * 3600 * 1000;
    const need = [];
    for (const cfg of CONFIG.LISTS || []) {
      const { list } = listBySlug(cfg.slug);
      if (!list || !cfg.url) continue;
      const stale = !list.syncedAt || (Date.now() - new Date(list.syncedAt).getTime()) > th;
      if (!stale && list.items.length) continue;
      need.push(cfg);
    }
    if (need.length) await scheduleCloudFetch(need);
  }

  /* ─────────────────── المقارنة الشاملة بين كل اللستات ─────────────────── */
  const cmpState = { rows: [] };
  const cmpNames = () => (CONFIG.LISTS || []).map((c) => c.name);

  function runAllListCompare() {
    // المقارنة الشاملة تقرأ مباشرة من مخزن الذاكرة المركزي window.appLists
    cmpState.rows = Comparison.compareAllLists(global.appLists);
    renderAllListCompare();
    const el = $('#cmp-status');
    if (el) el.textContent = `${cmpState.rows.length} صنف موحّد عبر ${(CONFIG.LISTS || []).length} لستة`;
  }

  function runAllListCompareSilent() {
    if (!$('#cmp-body')) return;
    runAllListCompare();
  }

  function renderAllListCompare() {
    const names = cmpNames();
    const head = $('#cmp-head');
    if (head) {
      head.innerHTML = `<tr class="text-slate-500 text-xs bg-white">
        <th class="text-right py-2 px-3">رقم الصنف</th>
        <th class="text-right py-2 px-3">اسم المنتج</th>
        <th class="text-right py-2 px-3">الوحدة</th>
        ${names.map((n) => `<th class="text-right py-2 px-3">${esc(n)}</th>`).join('')}
        <th class="text-right py-2 px-3">فرق السعر</th>
        <th class="text-right py-2 px-3">الملاحظة</th>
      </tr>`;
    }

    const diffOnly = $('#cmp-diff-only') ? $('#cmp-diff-only').checked : false;
    let rows = cmpState.rows || [];
    if (diffOnly) rows = rows.filter((r) => r.status === 'diff');
    const st = $('#cmp-status');
    if (st) {
      st.textContent = `${cmpState.rows.length} صنف موحّد · ${cmpState.rows.filter((r) => r.status === 'match').length} مطابق في كل اللستات · ${cmpState.rows.filter((r) => r.status === 'diff').length} بسعر متفاوت · ${cmpState.rows.filter((r) => r.status === 'single').length} مسجّل في لستة واحدة · معروض: ${rows.length}`;
    }

    const body = $('#cmp-body');
    if (!body) return;
    body.innerHTML = rows.slice(0, 5000).map((r) => {
      const cells = names.map((_, li) => {
        const p = r.prices[li];
        if (p === null || p === undefined) return '<td class="py-1.5 px-3 text-center"><span class="clr-neutral">—</span></td>';
        const isMin = r.min !== null && Math.abs(p - r.min) < 1e-9;
        const isMax = r.max !== null && Math.abs(p - r.max) < 1e-9 && r.diff > 1e-9;
        const cls = isMin ? 'text-emerald-700 font-extrabold' : (isMax ? 'text-rose-600 font-bold' : 'text-slate-700');
        return `<td class="py-1.5 px-3 text-center whitespace-nowrap"><span class="${cls}">${fmtNum(p)}</span></td>`;
      }).join('');
      const note = r.status === 'single'
        ? '<span class="status-badge status-unknown">في لستة واحدة</span>'
        : r.status === 'match'
          ? '<span class="status-badge status-match">متطابق</span>'
          : '<span class="status-badge status-high">متفاوت ⚠️</span>';
      return `<tr class="border-t border-slate-100 hover:bg-slate-50/70">
        <td class="py-2 px-3 text-xs" dir="ltr">${esc(r.itemNumber || '—')}</td>
        <td class="py-2 px-3 font-bold text-slate-800">${esc(r.name)}</td>
        <td class="py-2 px-3 text-xs text-slate-500">${esc(r.unit || '—')}</td>
        ${cells}
        <td class="py-2 px-3 font-extrabold whitespace-nowrap ${r.diff > 1e-9 ? 'text-rose-600' : 'text-slate-400'}">${r.diff > 1e-9 ? fmtNum(r.diff) : '—'}</td>
        <td class="py-2 px-3">${note}</td>
      </tr>`;
    }).join('') || `<tr><td colspan="${names.length + 5}" class="py-8 text-center text-slate-400">لا توجد نتائج — حدّث اللستات أولاً.</td></tr>`;
  }

  function allListCmpTSV() {
    const names = cmpNames();
    const noteOf = (r) => (r.status === 'diff' ? 'متفاوت' : r.status === 'match' ? 'متطابق' : 'في لستة واحدة');
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', ...names, 'فرق السعر', 'الملاحظة'].join('\t');
    const body = (cmpState.rows || []).map((r) => [
      r.itemNumber || '', r.name, r.unit || '',
      ...names.map((_, li) => (r.prices[li] == null ? '' : fmtNum(r.prices[li]))),
      r.diff > 1e-9 ? fmtNum(r.diff) : '', noteOf(r),
    ].join('\t'));
    return [header, ...body].join('\n');
  }

  function allListCmpAOA() {
    const names = cmpNames();
    const noteOf = (r) => (r.status === 'diff' ? 'متفاوت' : r.status === 'match' ? 'متطابق' : 'في لستة واحدة');
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', ...names, 'فرق السعر', 'الملاحظة'];
    const body = (cmpState.rows || []).map((r) => [
      r.itemNumber || '', r.name, r.unit || '',
      ...names.map((_, li) => (r.prices[li] == null ? '' : r.prices[li])),
      r.diff > 1e-9 ? r.diff : '', noteOf(r),
    ]);
    return [header, ...body];
  }

  /* ─────────────────── جدول بنود الفاتورة ─────────────────── */
  function statusBadgeHTML(r) {
    return `<span class="status-badge status-${r.status}">${Comparison.statusLabel(r.status)}</span>`;
  }
  function diffHTML(value, status) {
    if (status === Comparison.STATUS.UNKNOWN) return '<span class="clr-neutral">—</span>';
    if (Math.abs(value) < 1e-9) return '<span class="clr-neutral">بدون فرق</span>';
    const cls = status === Comparison.STATUS.HIGH ? 'clr-high' : status === Comparison.STATUS.LOW ? 'clr-low' : 'clr-neutral';
    return `<span class="${cls}">${signedMoney(value)}</span>`;
  }
  function itemRowHTML(item, index) {
    const r = Comparison.analyzeItem(item);
    return `
      <tr data-id="${item.id}" class="border-t border-slate-100 hover:bg-slate-50/70">
        <td class="py-2 px-3 text-slate-400 text-xs font-bold">${index + 1}</td>
        <td class="py-2 px-3 w-28">
          <input type="text" value="${esc(item.itemNumber || '')}" placeholder="—"
                 data-field="itemNumber" data-id="${item.id}" autocomplete="off"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 min-w-44">
          <input type="text" value="${esc(item.name)}" list="product-names"
                 data-field="name" data-id="${item.id}" autocomplete="off"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 w-24">
          <input type="text" value="${esc(item.unit || '')}" placeholder="الوحدة"
                 data-field="unit" data-id="${item.id}" autocomplete="off"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 w-20">
          <input type="number" min="1" step="any" value="${item.quantity}"
                 data-field="quantity" data-id="${item.id}"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 w-32">
          <input type="number" min="0" step="any" value="${item.unitPrice}"
                 data-field="unitPrice" data-id="${item.id}"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 font-bold cell-total">${money(r.invoiceTotal)}</td>
        <td class="py-2 px-3 text-slate-600 cell-list">${r.listPrice === null ? '<span class="clr-neutral">—</span>' : money(r.listPrice)}</td>
        <td class="py-2 px-3 cell-udiff">${diffHTML(r.unitDiff, r.status)}</td>
        <td class="py-2 px-3 cell-tdiff">${diffHTML(r.totalDiff, r.status)}</td>
        <td class="py-2 px-3 cell-status">${statusBadgeHTML(r)}</td>
        <td class="py-2 px-3">
          <button class="remove-item w-8 h-8 grid place-items-center rounded-lg text-rose-500 hover:bg-rose-50" data-id="${item.id}" title="حذف البند"><i class="fa-solid fa-xmark"></i></button>
        </td>
      </tr>`;
  }

  function renderInvoiceItems() {
    $('#items-table-body').innerHTML = State.items.map(itemRowHTML).join('');
    const hasItems = State.items.length > 0;
    $('#empty-items').classList.toggle('hidden', hasItems);
    $('#items-table-wrap').classList.toggle('hidden', !hasItems);
    $('#items-count').textContent = hasItems
      ? `عدد البنود: ${State.items.length} · اللستة المختارة: ${activeListName()}`
      : `اللستة المختارة: ${activeListName()}`;
    updateSummary();
  }

  function syncCell(id, field, value) {
    const input = document.querySelector(`#items-table-body tr[data-id="${id}"] [data-field="${field}"]`);
    if (input) input.value = value;
  }

  function updateRowUI(id) {
    const item = State.items.find((i) => i.id === id);
    const tr = document.querySelector(`#items-table-body tr[data-id="${id}"]`);
    if (!item || !tr) return;
    const r = Comparison.analyzeItem(item);
    tr.querySelector('.cell-total').innerHTML = money(r.invoiceTotal);
    tr.querySelector('.cell-list').innerHTML = r.listPrice === null ? '<span class="clr-neutral">—</span>' : money(r.listPrice);
    tr.querySelector('.cell-status').innerHTML = statusBadgeHTML(r);
    tr.querySelector('.cell-udiff').innerHTML = diffHTML(r.unitDiff, r.status);
    tr.querySelector('.cell-tdiff').innerHTML = diffHTML(r.totalDiff, r.status);
  }

  const updateAllRowsAndSummary = () => { State.items.forEach((it) => updateRowUI(it.id)); updateSummary(); };

  function updateItemField(id, field, rawValue) {
    const item = State.items.find((i) => i.id === id);
    if (!item) return;
    if (field === 'itemNumber') {
      item.itemNumber = rawValue;
      const li = Lists.findInStore(activeSlugOf(), item.itemNumber, item.name, item.unit);
      if (li) {
        if (!item.name) { item.name = li.name; syncCell(id, 'name', li.name); }
        if (!item.unit) { item.unit = li.unit || ''; syncCell(id, 'unit', item.unit); }
        if (!parseNum(item.unitPrice)) { item.unitPrice = num(li.price); syncCell(id, 'unitPrice', item.unitPrice); }
      }
    } else if (field === 'name') {
      item.name = rawValue;
      if (!parseNum(item.unitPrice)) {
        const li = Lists.findInStore(activeSlugOf(), item.itemNumber, item.name, item.unit);
        if (li) { item.unitPrice = num(li.price); syncCell(id, 'unitPrice', item.unitPrice); }
      }
    } else if (field === 'unit') {
      item.unit = rawValue;
      if (!parseNum(item.unitPrice)) {
        const li = Lists.findInStore(activeSlugOf(), item.itemNumber, item.name, item.unit);
        if (li) { item.unitPrice = num(li.price); syncCell(id, 'unitPrice', item.unitPrice); }
      }
    } else if (field === 'quantity') {
      item.quantity = Math.max(0, parseNum(rawValue));
    } else if (field === 'unitPrice') {
      item.unitPrice = Math.max(0, parseNum(rawValue));
    }
    updateRowUI(id);
    updateSummary();
    saveDraft();
  }

  function removeItem(id) {
    State.items = State.items.filter((i) => i.id !== id);
    renderInvoiceItems();
    saveDraft();
  }

  function addItem() {
    const itemNumber = $('#item-number').value.trim();
    const name = $('#item-name').value.trim();
    const unit = $('#item-unit').value.trim();
    const qty = parseNum($('#item-qty').value) || 1;
    let unitPrice = parseNum($('#item-price').value);
    const priceEmpty = $('#item-price').value.trim() === '';

    const key = name || itemNumber;
    if (!key) { toast('أدخل اسم المنتج أو رقم الصنف أولاً', 'error'); $('#item-name').focus(); return; }

    if (!unitPrice && priceEmpty) {
      const li = Lists.findInStore(activeSlugOf(), itemNumber, name, unit);
      if (li) unitPrice = num(li.price);
    }
    if (!unitPrice && priceEmpty) {
      toast('اكتب السعر المدرج بالفاتورة، أو اختر صنفاً من اللستة المختارة', 'error');
      $('#item-price').focus();
      return;
    }

    State.items.push({ id: genId('it'), itemNumber, name: name || itemNumber, unit, quantity: qty, unitPrice });
    renderInvoiceItems();
    saveDraft();
    $('#item-qty').value = '1';
    $('#item-price').value = '';
    $('#item-unit').value = $('#item-number').value = '';
    $('#item-name').value = '';
    $('#item-name').focus();
  }

  /* ─────────────────── ملخص الفاتورة ─────────────────── */
  function updateSummary() {
    const s = Comparison.summarize(State.items);
    const t = s.totals;
    let netLabel = 'صافي الفرق';
    let netDot = 'bg-slate-400';
    if (s.netDiff > 0.004) { netLabel = 'زيادة صافية عن المعتمد'; netDot = 'bg-rose-500'; }
    else if (s.netDiff < -0.004) { netLabel = 'انخفاض صافٍ عن المعتمد'; netDot = 'bg-amber-500'; }
    const cards = [
      { label: 'إجمالي الفاتورة', value: money(t.invoiceTotal), dot: 'bg-indigo-500', sub: `${State.items.length} بند · ${money(s.expectedTotal)} معتمد` },
      { label: 'الإجمالي المعتمد (القائمة المختارة)', value: money(s.expectedTotal), dot: 'bg-slate-400', sub: `«${activeListName()}» للبنود المسجلة فقط` },
      { label: 'إجمالي زيادة الأسعار', value: money(t.highTotal), dot: 'bg-rose-500', sub: t.highTotal > 0.004 ? 'أعلى من السعر المعتمد' : 'لا توجد زيادات' },
      { label: 'إجمالي الانخفاض عن المعتمد', value: money(t.lowTotal), dot: 'bg-amber-500', sub: t.lowTotal > 0.004 ? 'بيع دون السعر المعتمد — مراجعة عاجلة' : 'لا يوجد انخفاض', valueColor: t.lowTotal > 0.004 ? 'text-rose-600' : '' },
      { label: netLabel, value: signedMoney(s.netDiff), dot: netDot, sub: `انحراف ${fmtNum(s.deviationPct)}% عن المعتمد`, valueColor: s.netDiff > 0.004 ? 'text-rose-600' : s.netDiff < -0.004 ? 'text-amber-600' : '' },
      { label: 'بنود غير مسجلة بالقائمة', value: `${t.unknownCount} بند`, dot: 'bg-amber-500', sub: money(t.unknownTotal) },
    ];
    $('#summary-grid').innerHTML = cards.map((c) => `
      <div class="bg-white rounded-2xl border border-slate-200 shadow-sm p-4 flex flex-col justify-between min-h-28">
        <div class="flex items-center gap-2 text-xs font-bold text-slate-500">
          <span class="w-2.5 h-2.5 rounded-full ${c.dot}"></span>${c.label}
        </div>
        <div class="mt-2 text-xl font-extrabold ${c.valueColor || 'text-slate-900'}">${c.value}</div>
        <div class="text-[11px] text-slate-400">${c.sub}</div>
      </div>`).join('');
  }

  /* ─────────────────── الفاتورة: حفظ / مسودة / جديد ─────────────────── */
  const readMeta = () => ({ customer: $('#inv-customer').value, invoiceNo: $('#inv-no').value, date: $('#inv-date').value, notes: $('#inv-notes').value });
  const saveDraft = () => Invoices.saveDraft({ ...readMeta(), listId: Lists.activeIdOf(), invoiceId: State.editingInvoiceId || null, items: State.items, savedAt: new Date().toISOString() });

  // زر الحفظ يتغيّر إلى "تحديث" عند العمل على فاتورة محفوظة مسبقاً
  function updateSaveButton() {
    const btn = $('#btn-save-invoice');
    const editing = !!State.editingInvoiceId;
    btn.innerHTML = editing
      ? '<i class="fa-solid fa-pen-to-square"></i> تحديث الفاتورة في السجل'
      : '<i class="fa-solid fa-floppy-disk"></i> حفظ الفاتورة في السجل';
    btn.title = editing ? 'يحدّث الفاتورة الحالية مكانها — بدون تكرار' : 'يحفظ الفاتورة كسجل جديد';
  }

  function loadDraftIfAny() {
    const draft = Invoices.loadDraft();
    if (!draft) return;
    $('#inv-customer').value = draft.customer !== undefined && draft.customer !== null ? draft.customer : draft.vendor || '';
    $('#inv-no').value = draft.invoiceNo || '';
    $('#inv-date').value = draft.date || todayStr();
    $('#inv-notes').value = draft.notes || '';
    State.items = Array.isArray(draft.items) ? draft.items.map((it) => ({
      id: genId('it'),
      itemNumber: String(it.itemNumber || '').trim(),
      name: String(it.name || '').trim(),
      unit: String(it.unit || '').trim(),
      quantity: it.quantity ?? 1,
      unitPrice: Number(it.unitPrice) || 0,
    })) : [];
    // استعادة الفاتورة قيد التعديل إن كانت لا تزال موجودة في السجل (لا نعيدها لفاتورة محذوفة)
    State.editingInvoiceId = draft.invoiceId && Invoices.getById(draft.invoiceId) ? draft.invoiceId : null;
    updateSaveButton();
  }

  function saveInvoice() {
    if (!State.items.length) { toast('أضف بنداً واحداً على الأقل قبل الحفظ', 'error'); switchTab('invoice'); return; }
    const editing = State.editingInvoiceId && Invoices.getById(State.editingInvoiceId);
    const inv = Invoices.save({ ...readMeta(), items: State.items }, State.editingInvoiceId);
    Invoices.clearDraft();
    const no = inv.invoiceNo ? `"${inv.invoiceNo}" ` : '';
    toast(editing ? `تم تحديث الفاتورة ${no}في السجل — لا تكرار` : `تم حفظ الفاتورة ${no}في السجل بنجاح`);
    renderHistory();
    resetCurrentInvoice(true);
  }

  function resetCurrentInvoice(silent) {
    State.items = [];
    State.editingInvoiceId = null;
    updateSaveButton();
    $('#inv-customer').value = $('#inv-no').value = $('#inv-notes').value = '';
    $('#inv-date').value = todayStr();
    $('#item-name').value = $('#item-price').value = $('#item-unit').value = $('#item-number').value = '';
    $('#item-qty').value = '1';
    Invoices.clearDraft();
    renderInvoiceItems();
    if (!silent) toast('بدأت فاتورة جديدة فارغة', 'info');
  }

  function newInvoiceClick() {
    const touched = State.items.length > 0 || $('#inv-customer').value.trim();
    if (!touched) { resetCurrentInvoice(); return; }
    confirmBox({ title: 'فاتورة جديدة', message: 'سيتم مسح البنود الحالية. هل تريد البدء من جديد؟', confirmLabel: 'نعم، امسح ثم ابدأ', danger: true, onConfirm: () => resetCurrentInvoice() });
  }

  function loadInvoiceIntoEditor(id) {
    const inv = Invoices.getById(id);
    if (!inv) return;
    const proceed = () => {
      State.items = inv.items.map((it) => ({ id: genId('it'), itemNumber: it.itemNumber || '', name: it.name, unit: it.unit || '', quantity: it.quantity, unitPrice: it.unitPrice }));
      if (inv.listId && Lists.getList(inv.listId)) { Lists.setActive(inv.listId); renderQuickList(); renderCloudLists(); renderDatalist(); }
      State.editingInvoiceId = inv.id;
      updateSaveButton();
      $('#inv-customer').value = inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || '';
      $('#inv-no').value = inv.invoiceNo || '';
      $('#inv-date').value = inv.date || todayStr();
      $('#inv-notes').value = inv.notes || '';
      saveDraft();
      renderInvoiceItems();
      switchTab('invoice');
      toast(`تم تحميل الفاتورة ${inv.invoiceNo ? `"${inv.invoiceNo}" ` : ''}للعمل عليها — اضغط "تحديث" لحفظ التعديلات`);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    };
    if (State.items.length) {
      confirmBox({ title: 'استبدال الفاتورة الحالية؟', message: 'البنود الحالية ستحل محلها فاتورة السجل. هل تريد المتابعة؟', confirmLabel: 'تحميل الفاتورة', onConfirm: proceed });
    } else { proceed(); }
  }

  /* ─────────────────── قوائم الأسعار: الواجهة (+ ترقيم للكبير) ─────────────────── */
  function renderListItems() {
    const l = Lists.active();
    const q = State.productFilter.trim().toLocaleLowerCase('ar-EG');
    const all = l ? l.items : [];
    const filtered = !q ? all : all.filter((p) =>
      p.name.toLocaleLowerCase('ar-EG').includes(q)
      || String(p.itemNumber || '').toLocaleLowerCase('ar-EG').includes(q)
      || String(p.unit || '').toLocaleLowerCase('ar-EG').includes(q));
    const pageSize = CONFIG.LIST_PAGE_SIZE;
    const pageCount = Math.max(1, Math.ceil(filtered.length / pageSize));
    if (State.listPage >= pageCount) State.listPage = pageCount - 1;
    const slice = filtered.slice(State.listPage * pageSize, (State.listPage + 1) * pageSize);

    $('#prod-table-body').innerHTML = slice.length
      ? slice.map((p) => `
          <tr class="border-t border-slate-100 hover:bg-slate-50/70">
            <td class="py-2.5 px-3 font-bold text-slate-700 dir-ltr">${esc(p.itemNumber || '—')}</td>
            <td class="py-2.5 px-3 font-bold text-slate-800">${esc(p.name)}</td>
            <td class="py-2.5 px-3 text-xs text-slate-500">${esc(p.unit || '—')}</td>
            <td class="py-2.5 px-3 font-extrabold text-indigo-700">${money(p.price)}</td>
          </tr>`).join('')
      : '<tr><td colspan="4" class="py-10 text-center text-slate-400">لا توجد أصناف مطابقة في اللستة النشطة.</td></tr>';

    $('#prod-stats').textContent = `أصناف اللستة «${activeListName()}»: ${all.length} · معروض: ${filtered.length}`;
    renderPager(pageCount, filtered.length);
  }

  function renderPager(pageCount, totalResult) {
    const wrap = $('#prod-pager');
    if (!wrap) return;
    if (pageCount <= 1 && totalResult <= CONFIG.LIST_PAGE_SIZE) { wrap.innerHTML = ''; return; }
    const page = State.listPage;
    const visible = [];
    for (let p = 0; p < pageCount; p++) {
      if (p === 0 || p === pageCount - 1 || Math.abs(p - page) <= 2) visible.push(p);
      else if (visible[visible.length - 1] !== '…') visible.push('…');
    }
    wrap.innerHTML = `
      <button data-page="${page - 1}" class="pager-btn btn btn-ghost text-xs px-3 py-1.5" ${page === 0 ? 'disabled' : ''}><i class="fa-solid fa-chevron-right"></i> السابق</button>
      ${visible.map((p) => p === '…'
        ? '<span class="px-1 text-slate-400 select-none">…</span>'
        : `<button data-page="${p}" class="pager-btn btn ${p === page ? 'btn-primary' : 'btn-ghost'} text-xs px-3 py-1.5">${p + 1}</button>`).join('')}
      <button data-page="${page + 1}" class="pager-btn btn btn-ghost text-xs px-3 py-1.5" ${page >= pageCount - 1 ? 'disabled' : ''}>التالي <i class="fa-solid fa-chevron-left"></i></button>`;
    wrap.querySelectorAll('[data-page]').forEach((b) => {
      b.addEventListener('click', () => {
        const p = Math.max(0, Math.min(pageCount - 1, parseInt(b.dataset.page, 10)));
        State.listPage = p;
        renderListItems();
      });
    });
  }

  /* ─────────────────── اللصق السريع (فاتورة فقط) ─────────────────── */
  function applyPasteInvoice() {
    const rows = Parser.parseRows($('#paste-invoice').value);
    const { items, ignored, autofilled, unknown } = Parser.extractInvoiceItems(rows.map((r) => r.cells), { autoPrice: true });
    if (!items.length) { toast('لم تُعثر على بنود صالحة — تأكد من وجود أعمدة: اسم/رقم الصنف، الكمية، والسعر.', 'error'); return; }
    State.items.push(...items.map((it) => ({ ...it, id: genId('it') })));
    renderInvoiceItems();
    saveDraft();
    const extra = autofilled ? ` · اتُملئ سعر ${autofilled} بند تلقائياً من اللستة` : '';
    const unk = unknown ? ` · ${unknown} بند بلا سعر مؤكد (يُعرض كغير مسجل) — أكمل سعره يدوياً` : '';
    toast(`تم لصق ${items.length} بند في الفاتورة · ${ignored} صفّ متجاهل${extra}${unk}`);
    $('#paste-invoice').value = '';
  }

  /* ─────────────────── استيراد Excel محلي (للفاتورة فقط) ─────────────────── */
  function handleXlsxFile(fileInputId, targetTextareaId, targetTab, label) {
    const file = $(fileInputId).files[0];
    if (!file) return;
    toast('جارٍ قراءة ملف Excel...', 'info');
    Excel.readFirstSheet(file)
      .then((sheet) => {
        if (!sheet || !sheet.rows.length) { toast('الملف فارغ أو بلا أوراق', 'error'); return; }
        $(targetTextareaId).value = Excel.sheetToTSV(sheet);
        switchTab(targetTab);
        toast(`${label} من «${sheet.name}»: ${sheet.rows.length} صف — راجع ثم اضغط زر الاستيراد`);
      })
      .catch((err) => toast(err.message || 'فشل قراءة ملف Excel', 'error'))
      .finally(() => { $(fileInputId).value = ''; });
  }

  /* ─────────────────── سجل الفواتير ─────────────────── */
  function renderHistory() {
    const list = Invoices.getAll();
    $('#empty-history').classList.toggle('hidden', list.length > 0);
    $('#history-list').innerHTML = list.map((inv) => {
      const refList = inv.listId ? Lists.getList(inv.listId) : Lists.active();
      const s = Comparison.summarize(inv.items, refList);
      let badge = '<span class="status-badge status-match">متوازنة</span>';
      let value = signedMoney(s.netDiff);
      let vc = 'text-slate-500';
      if (s.netDiff > 0.004) { badge = '<span class="status-badge status-high">🔺 زيادة صافية</span>'; vc = 'text-rose-600'; }
      else if (s.netDiff < -0.004) { badge = '<span class="status-badge status-low">🔻 انخفاض صافٍ عن المعتمد</span>'; vc = 'text-amber-600'; }
      return `
        <div class="card flex flex-wrap items-center justify-between gap-3">
          <div class="min-w-52">
            <p class="font-extrabold text-slate-800">${esc(inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || '') || 'بدون عميل'}
              ${inv.invoiceNo ? `<span class="ms-2 text-xs font-bold text-indigo-600 bg-indigo-50 rounded-full px-2 py-0.5">${esc(inv.invoiceNo)}</span>` : ''}
              <span class="ms-1 text-xs font-bold text-slate-400">${esc(inv.date) || ''}</span>
            </p>
            <p class="text-xs text-slate-500 mt-1">${inv.items.length} بند · اللستة: ${esc(refList ? refList.name : '—')} · حُفظت في ${new Date(inv.createdAt).toLocaleString('ar-EG')}</p>
          </div>
          <div class="flex flex-wrap items-center gap-4">
            <div class="text-center">
              <div class="text-[11px] text-slate-400 font-bold">إجمالي الفاتورة</div>
              <div class="font-extrabold text-slate-800">${money(s.totals.invoiceTotal)}</div>
            </div>
            <div class="text-center">
              <div class="text-[11px] text-slate-400 font-bold">صافي الفرق</div>
              <div class="font-extrabold ${vc}">${value}</div>
            </div>
            ${badge}
            <div class="flex gap-2">
              <button data-open="${inv.id}" class="action-btn rounded-lg px-3 py-1.5 text-xs font-bold bg-indigo-600 text-white hover:bg-indigo-700 transition"><i class="fa-solid fa-eye"></i> عرض</button>
              <button data-exp="${inv.id}" class="action-btn rounded-lg px-3 py-1.5 text-xs font-bold bg-slate-100 text-slate-700 hover:bg-slate-200 transition"><i class="fa-solid fa-download"></i> تصدير</button>
              <button data-del="${inv.id}" class="action-btn rounded-lg px-3 py-1.5 text-xs font-bold bg-rose-50 text-rose-600 hover:bg-rose-100 transition"><i class="fa-solid fa-trash"></i> حذف</button>
            </div>
          </div>
        </div>`;
    }).join('');
  }

  /* ─────────────────── التقارير (TSV / Excel / طباعة) ─────────────────── */
  function reportToTSV() {
    const s = Comparison.summarize(State.items);
    const t = s.totals;
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'إجمالي المبلغ', 'السعر المعتمد', 'فرق الوحدة', 'إجمالي الفرق', 'مؤشر المراجعة'].join('\t');
    const rows = State.items.map((it) => {
      const r = Comparison.analyzeItem(it);
      const ud = r.status === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.unitDiff) < 1e-9 ? '0.00' : signedNum(r.unitDiff));
      const td = r.status === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.totalDiff) < 1e-9 ? '0.00' : signedNum(r.totalDiff));
      return [
        it.itemNumber || '—', it.name, it.unit || '—',
        fmtNum(it.quantity), fmtNum(it.unitPrice), fmtNum(r.invoiceTotal),
        r.listPrice === null ? '—' : fmtNum(r.listPrice),
        ud, td, Comparison.statusText(r.status),
      ].join('\t');
    });
    const summary = [
      '',
      ['', '', '', '', 'إجمالي الفاتورة', fmtNum(t.invoiceTotal), 'الإجمالي المعتمد', fmtNum(t.expectedTotal), '', ''].join('\t'),
      ['', '', '', '', 'إجمالي الزيادة', fmtNum(t.highTotal), 'إجمالي الانخفاض', fmtNum(t.lowTotal), '', ''].join('\t'),
      ['', '', '', '', 'صافي الفرق', signedNum(s.netDiff), 'بنود غير مسجلة', String(t.unknownCount), '', ''].join('\t'),
    ];
    return [header, ...rows, ...summary].join('\n');
  }

  /** مصفوفة خلايا لتصدير Excel مباشرة (أرقام حقيقية قابلة للحساب) */
  function reportToAOACells() {
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'إجمالي المبلغ', 'السعر المعتمد', 'فرق الوحدة', 'إجمالي الفرق', 'مؤشر المراجعة'];
    const body = State.items.map((it) => {
      const r = Comparison.analyzeItem(it);
      return [
        it.itemNumber || '', it.name, it.unit || '',
        it.quantity, it.unitPrice, r.invoiceTotal,
        r.listPrice === null ? '' : r.listPrice,
        r.status === Comparison.STATUS.UNKNOWN ? '' : (Math.abs(r.unitDiff) < 1e-9 ? 0 : r.unitDiff),
        r.status === Comparison.STATUS.UNKNOWN ? '' : (Math.abs(r.totalDiff) < 1e-9 ? 0 : r.totalDiff),
        Comparison.statusText(r.status),
      ];
    });
    return [header, ...body];
  }

  function invoicesToTSV() {
    const header = ['العميل', 'رقم الفاتورة', 'التاريخ', 'رقم الصنف', 'المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'إجمالي البند'].join('\t');
    const rows = Invoices.getAll().flatMap((inv) =>
      inv.items.map((it) => [
        inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || '', inv.invoiceNo || '', inv.date || '',
        it.itemNumber || '', it.name || '', it.unit || '',
        fmtNum(it.quantity), fmtNum(it.unitPrice), fmtNum((it.quantity || 0) * (it.unitPrice || 0)),
      ].join('\t')));
    return [header, ...rows].join('\n');
  }

  function exportXlsxReport() {
    if (!State.items.length) { toast('لا توجد بنود لتصديرها', 'error'); return; }
    Excel.exportXLSX('تقرير المراجعة', reportToAOACells(), `تقرير-مراجعة-${todayStr()}.xlsx`)
      .then(() => toast('تم تنزيل تقرير Excel ✓'))
      .catch((err) => toast(err.message || 'تعذّر تصدير Excel', 'error'));
  }

  /* ─────────────────── تقرير الطباعة ─────────────────── */
  function buildPrintHTML() {
    const meta = readMeta();
    const s = Comparison.summarize(State.items); // يقرأ مصفوفة البنود الحالية فوراً (لا نسخ مخزنة)
    const t = s.totals;
    const listName = activeListName();
    const qtySum = State.items.reduce((a, it) => a + (Number(it.quantity) || 0), 0);

    // توليد صفوف <tr> ديناميكياً لكل بند من مصفوفة البنود الحالية (لا قالب ثابت)
    // لكل بند: رقم الصنف، الاسم، الوحدة، الكمية، سعر الوحدة، إجمالي المبلغ،
    // السعر المعتمد، فرق الوحدة، إجمالي الفرق، حالة المراجعة
    const rows = State.items.map((it, i) => {
      try {
        const r = Comparison.analyzeItem(it);
        const st = r.status;
        const ud = st === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.unitDiff) < 1e-9 ? '—' : (r.unitDiff > 0 ? `+${fmtNum(r.unitDiff)}` : fmtNum(r.unitDiff)));
        const td = st === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.totalDiff) < 1e-9 ? '—' : (r.totalDiff > 0 ? `+${fmtNum(r.totalDiff)}` : fmtNum(r.totalDiff)));
        return `<tr>
          <td>${i + 1}</td>
          <td>${esc(it.itemNumber || '—')}</td>
          <td>${esc(it.name || '—')}</td>
          <td>${esc(it.unit || '—')}</td>
          <td>${fmtNum(it.quantity)}</td>
          <td>${fmtNum(it.unitPrice)}</td>
          <td>${fmtNum(r.invoiceTotal)}</td>
          <td>${r.listPrice == null ? '—' : fmtNum(r.listPrice)}</td>
          <td>${ud}</td>
          <td>${td}</td>
          <td>${Comparison.statusText(st)}</td>
        </tr>`;
      } catch (e) {
        // بند معطوب لا يفرّغ الجدول كاملاً — صف احتياطي يُبقي التقرير مقروءاً
        return `<tr>
          <td>${i + 1}</td>
          <td>${esc(it.itemNumber || '—')}</td>
          <td>${esc(it.name || '—')}</td>
          <td>${esc(it.unit || '—')}</td>
          <td>${fmtNum(it.quantity)}</td>
          <td>${fmtNum(it.unitPrice)}</td>
          <td colspan="5">بيانات غير مكتملة</td>
        </tr>`;
      }
    }).join('');;

    const isHigh = s.netDiff > 0.004, isLow = s.netDiff < -0.004;
    const netBadge = isHigh
      ? `<span class="eq-high">زيادة صافية عن المعتمد: ${money(s.netDiff)}</span>`
      : isLow
        ? `<span class="eq-low">انخفاض صافٍ عن المعتمد: ${money(Math.abs(s.netDiff))}</span>`
        : `<span class="eq-ok">الفاتورة متوازنة مع المعتمد</span>`;

    // بطاقات الملخص النهائي (إجمالي الفاتورة / الزيادة / الانخفاض / عدد البنود) — الانخفاض بتلوين تحذيري
    const sumCards = [
      { label: 'إجمالي الفاتورة', value: money(t.invoiceTotal) },
      { label: 'إجمالي زيادة الأسعار', value: money(t.highTotal) },
      { label: 'إجمالي الانخفاض عن المعتمد', value: money(t.lowTotal), warn: true },
      { label: 'عدد البنود', value: `${State.items.length} بند` },
    ];
    const sumHTML = sumCards.map((c) => `<div class="pr-card"><span>${c.label}</span><b${c.warn ? ' style="color:#c2410c"' : ''}>${c.value}</b></div>`).join('');

    const customer = meta.customer !== undefined && meta.customer !== null ? meta.customer : meta.vendor || '—';
    return `
      <div class="pr-title">
        <div><h1>تقرير مراجعة فاتورة العميل ومقارنة أسعار البيع</h1></div>
        <div class="pr-meta">تاريخ الطباعة: ${new Date().toLocaleString('ar-EG')}</div>
      </div>
      <div class="pr-meta">
        <div><b>العميل:</b> ${esc(customer)} &nbsp;|&nbsp; <b>رقم الفاتورة:</b> ${esc(meta.invoiceNo || '—')}</div>
        <div><b>تاريخ الفاتورة:</b> ${esc(meta.date || todayStr())} &nbsp;|&nbsp; <b>القائمة المعتمدة:</b> ${esc(listName)}</div>
      </div>
      <table>
        <thead>
          <tr>
            <th>#</th><th>رقم الصنف</th><th>اسم المنتج</th><th>الوحدة</th><th>الكمية</th><th>سعر الوحدة</th>
            <th>إجمالي المبلغ</th><th>السعر المعتمد</th><th>فرق الوحدة</th><th>إجمالي الفرق</th><th>المراجعة</th>
          </tr>
        </thead>
        <tbody>${rows || '<tr><td colspan="11" style="text-align:center">لا توجد بنود</td></tr>'}</tbody>
        <tfoot>
          <tr>
            <td colspan="4">إجمالي البنود (${State.items.length})</td>
            <td>${fmtNum(qtySum)}</td>
            <td></td>
            <td>${fmtNum(t.invoiceTotal)}</td>
            <td>${fmtNum(t.expectedTotal)}</td>
            <td colspan="2">${signedMoney(s.netDiff)}</td>
            <td>${isHigh ? 'زيادة صافية' : isLow ? 'انخفاض صافٍ' : 'متوازن'}</td>
          </tr>
        </tfoot>
      </table>
      <div class="pr-sum">${sumHTML}</div>
      <div class="pr-equation">
        الإجمالي المعتمد: <b>${money(t.expectedTotal)}</b> &nbsp;·&nbsp;
        صافي الفرق عن المعتمد: <b>${signedMoney(s.netDiff)}</b> &nbsp;·&nbsp;
        الانحراف: <b>${fmtNum(s.deviationPct)}%</b> &nbsp;·&nbsp;
        البنود غير المسجلة بالقائمة: <b>${t.unknownCount}</b> (${money(t.unknownTotal)})<br/>
        ${netBadge}
      </div>
      ${meta.notes ? `<div class="pr-note"><b>ملاحظات:</b> ${esc(meta.notes)}</div>` : ''}
      <div class="pr-footer">
        <div class="sign">توقيع المراجع<div class="line"></div></div>
        <div class="sign">خاتم الشركة<div class="line"></div></div>
        <div class="sign">توقيع المستلم<div class="line"></div></div>
      </div>`;
  }

  function printReport() {
    if (!State.items.length) { toast('لا توجد بنود للطباعة بعد', 'error'); return; }
    const host = $('#print-area');
    if (!host) { toast('عنصر منطقة الطباعة مفقود', 'error'); return; }
    // 1) حقن قالب التقرير (ترويسة + إجماليات + صفوف البنود) في الحاوية المخصصة
    host.innerHTML = buildPrintHTML();
    // 2) إجبار إعادة الحساب والتخطيط بعد الإدراج
    void host.offsetHeight;
    // 3) تحقق قبل الطباعة: يجب أن يكون <tbody> قد امتلأ بصفوف فعلية
    const tbody = host.querySelector('tbody');
    const injected = tbody ? tbody.children.length : 0;
    if (!injected) {
      toast(`تعذّر توليد صفوف البنود (0 صف من ${State.items.length} بند)`, 'error');
      return;
    }
    console.log(`طباعة: ${injected} صف بنود من ${State.items.length} بند`);
    // 4) فتح نافذة الطباعة بعد اكتمال الرسم (لا window.print قبل الحقن أبداً)
    requestAnimationFrame(() => requestAnimationFrame(() => setTimeout(() => window.print(), 20)));
  }

  /* ─────────────────── مقارنة سحابية: موزع ضد شركة ─────────────────── */
  const CLOUD = CONFIG.CLOUD_COMPARE;
  let cloudTabs = null;        // { tabs: [{name,gid,rows,count}] } من fetchSheetsData
  let cloudComparePage = 0;

  const cloudLabel = (name) => (CLOUD.LABELS && CLOUD.LABELS[name]) || name;

  function fillCloudTabSelects() {
    const opts = (CLOUD.TABS || []).map((t) => `<option value="${esc(t)}">${esc(cloudLabel(t))}</option>`).join('');
    const a = $('#cloud-sheet-agent'); if (a) a.innerHTML = opts;
    const r = $('#cloud-sheet-ref'); if (r) r.innerHTML = opts;
  }

  function seedCloudPanel() {
    const urlEl = $('#cloud-url');
    if (urlEl && !urlEl.value) urlEl.value = CLOUD.DEFAULT_URL || '';
    fillCloudTabSelects();
    if (State.settings.cloudCompareUrl && urlEl) urlEl.value = State.settings.cloudCompareUrl;
    const a = $('#cloud-sheet-agent'); if (a && State.settings.cloudAgentTab && a.value !== State.settings.cloudAgentTab) a.value = State.settings.cloudAgentTab;
    const r = $('#cloud-sheet-ref'); if (r && State.settings.cloudRefTab && r.value !== State.settings.cloudRefTab) r.value = State.settings.cloudRefTab;
  }

  async function handleCloudFetch() {
    const url = $('#cloud-url').value.trim();
    if (!url) { toast('أدخل رابط الشيت المنشور أولاً', 'error'); $('#cloud-url').focus(); return; }
    const apiKey = $('#gs-api-key').value.trim();
    const agent = $('#cloud-sheet-agent').value;
    const ref = $('#cloud-sheet-ref').value;
    State.settings.cloudCompareUrl = url;
    State.settings.cloudAgentTab = agent;
    State.settings.cloudRefTab = ref;
    Storage.saveSettings(State.settings);

    const btn = $('#btn-cloud-fetch');
    const orig = btn.innerHTML;
    btn.disabled = true;
    btn.innerHTML = '<i class="fa-solid fa-spinner fa-spin"></i> جارٍ جلب التبويبتين...';
    const statusEl = $('#cloud-status');
    statusEl.textContent = 'جارٍ الاتصال بـ Google Sheets...';
    statusEl.className = 'text-xs font-bold text-slate-500';
    $('#cloud-result').classList.add('hidden');
    $('#cloud-compare-wrap').classList.add('hidden');
    $('#btn-cloud-run-compare').disabled = true;

    try {
      const data = await Sheets.fetchSheetsData(url, {
        tabs: [agent, ref],
        apiKey,
        onTab: ({ tab, count, done, total }) => {
          statusEl.textContent = `تم جلب «${cloudLabel(tab)}» (${count} صف) — ${done}/${total}`;
        },
      });
      cloudTabs = data;
      State.settings.lastSyncAt = new Date().toISOString();
      Storage.saveSettings(State.settings);
      statusEl.textContent = `✓ جُلب ${data.tabs.length} تبويب عبر ${data.source === 'api' ? 'Google Sheets API' : 'رابط CSV منشور'}`;
      statusEl.className = 'text-xs font-bold text-emerald-600';
      $('#cloud-result').classList.remove('hidden');
      $('#btn-cloud-run-compare').disabled = false;
      renderCloudPreview();
      runCloudCompare();
      $('#cloud-compare-wrap').classList.remove('hidden');
      toast('تم جلب اللستتين من الشيت بنجاح');
    } catch (err) {
      cloudTabs = null;
      State.cloudCompareRows = [];
      State.cloudCompareStats = null;
      statusEl.textContent = 'فشل الجلب — تحقق من الرابط أو جرب وضع API بمفتاح';
      statusEl.className = 'text-xs font-bold text-rose-600';
      toast(err && err.message ? err.message : 'فشل الاتصال بجداول Google', 'error');
    } finally {
      btn.disabled = false;
      btn.innerHTML = orig;
    }
  }

  function renderCloudPreview() {
    if (!cloudTabs || !$('#cloud-preview-grid')) return;
    $('#cloud-preview-grid').innerHTML = cloudTabs.tabs.map((t) => {
      const parsed = Parser.extractListRows(t.rows);
      const sample = parsed.rows.slice(0, CLOUD.PREVIEW_ROWS || 8);
      return `
        <div class="rounded-xl border border-slate-200 bg-white overflow-hidden">
          <div class="px-3 py-2 bg-slate-50 border-b font-extrabold text-sm text-slate-700 flex justify-between">
            <span>${esc(cloudLabel(t.name))}</span>
            <span class="text-xs font-bold text-slate-400">${t.count} صف · ${parsed.rows.length} بند</span>
          </div>
          <div class="overflow-x-auto max-h-64">
            <table class="w-full text-xs">
              <thead>
                <tr class="bg-slate-100 text-slate-500">
                  <th class="px-2 py-1 text-start">الكود</th><th class="px-2 py-1 text-start">الوحدة</th>
                  <th class="px-2 py-1 text-start">السعر</th><th class="px-2 py-1 text-start">اسم المنتج</th>
                </tr>
              </thead>
              <tbody>
                ${sample.map((r) => `<tr class="border-t border-slate-100">
                  <td class="px-2 py-1" dir="ltr">${esc(r.itemNumber || '—')}</td>
                  <td class="px-2 py-1">${esc(r.unit || '—')}</td>
                  <td class="px-2 py-1">${fmtNum(r.price)}</td>
                  <td class="px-2 py-1">${esc(r.name)}</td>
                </tr>`).join('')}
              </tbody>
            </table>
          </div>
        </div>`;
    }).join('');
  }

  function runCloudCompare() {
    if (!cloudTabs) { toast('اجلب البيانات أولاً', 'error'); return; }
    const agentTab = $('#cloud-sheet-agent').value;
    const refTab = $('#cloud-sheet-ref').value;
    const parsed = {};
    for (const t of cloudTabs.tabs) parsed[t.name] = Parser.extractListRows(t.rows).rows;
    const agentRows = parsed[agentTab] || parsed[Object.keys(parsed)[0]] || [];
    const refRows = parsed[refTab] || parsed[Object.keys(parsed)[1]] || [];
    if (!agentRows.length || !refRows.length) { toast('لا توجد بيانات كافية للمقارنة', 'error'); return; }

    // كائن لستة مُفهرَس بخرائط Map للمرجع (مطابقة سريعة للقوائم الضخمة)
    const refList = { id: 'clcmp_' + Date.now().toString(36), name: refTab, items: refRows.map((r) => ({ ...r })) };
    Object.defineProperty(refList, '__ver', { value: 1, writable: true, enumerable: false, configurable: true });

    const strictUnit = !$('#cloud-ignore-unit').checked;
    const rows = agentRows.map((r) => {
      const probe = strictUnit ? r : { ...r, unit: '' };
      return Comparison.analyzeItem({ ...probe, quantity: 1, unitPrice: r.price }, refList);
    });

    const stats = { total: rows.length, matched: 0, high: 0, highTotal: 0, save: 0, saveTotal: 0, unknown: 0 };
    for (const x of rows) {
      if (x.status === Comparison.STATUS.MATCH) stats.matched++;
      else if (x.status === Comparison.STATUS.HIGH) { stats.high++; stats.highTotal += x.unitDiff; }
      else if (x.status === Comparison.STATUS.LOW) { stats.save++; stats.saveTotal += Math.abs(x.unitDiff); }
      else stats.unknown++;
    }
    State.cloudCompareRows = rows;
    State.cloudCompareStats = stats;
    cloudComparePage = 0;
    renderCloudCompare();
  }

  function renderCloudCompare() {
    const rows = State.cloudCompareRows || [];
    if (!$('#cloud-compare-body')) return;
    if (!rows.length) { $('#cloud-compare-body').innerHTML = '<tr><td colspan="8" class="p-4 text-center text-slate-400 text-sm">لا توجد نتائج مقارنة.</td></tr>'; return; }

    const diffOnly = $('#cloud-diff-only').checked;
    let visible = diffOnly ? rows.filter((x) => x.status !== Comparison.STATUS.MATCH) : rows;
    const pageSize = CLOUD.COMPARE_PAGE_SIZE || 200;
    const pages = Math.max(1, Math.ceil(visible.length / pageSize));
    cloudComparePage = Math.max(0, Math.min(cloudComparePage, pages - 1));
    const slice = visible.slice(cloudComparePage * pageSize, cloudComparePage * pageSize + pageSize);

    const st = State.cloudCompareStats || {};
    const setTxt = (id, v) => { const el = $(id); if (el) el.textContent = v; };
    setTxt('#cloud-sum-match', st.matched || 0);
    setTxt('#cloud-sum-high', `${st.high || 0} · ${signedMoney(st.highTotal || 0)}`);
    setTxt('#cloud-sum-save', `${st.save || 0} · −${money(st.saveTotal || 0)}`);
    setTxt('#cloud-sum-unknown', st.unknown || 0);
    setTxt('#cloud-total', `${st.total || 0} بند`);

    $('#cloud-compare-body').innerHTML = slice.map((x, i) => {
      const n = cloudComparePage * pageSize + i + 1;
      const dx = x.status === Comparison.STATUS.UNKNOWN ? '<span class="clr-neutral">—</span>'
        : Math.abs(x.unitDiff) < 1e-9 ? '<span class="clr-neutral">بدون فرق</span>'
        : `<span class="${x.status === Comparison.STATUS.HIGH ? 'clr-high' : 'clr-low'}">${signedMoney(x.unitDiff)}</span>`;
      return `<tr class="border-t border-slate-100 hover:bg-slate-50/70">
        <td class="py-2 px-3 text-slate-400 text-xs font-bold">${n}</td>
        <td class="py-2 px-3 text-xs" dir="ltr">${esc(x.item.itemNumber || '—')}</td>
        <td class="py-2 px-3 text-sm">${esc(x.item.name)}</td>
        <td class="py-2 px-3 text-xs">${esc(x.item.unit || '—')}</td>
        <td class="py-2 px-3 text-xs">${money(x.invoiceTotal)}</td>
        <td class="py-2 px-3 text-xs">${x.listPrice === null ? '—' : money(x.listPrice)}</td>
        <td class="py-2 px-3 text-xs">${dx}</td>
        <td class="py-2 px-3">${statusBadgeHTML(x.status)}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="8" class="p-4 text-center text-slate-400 text-sm">لا توجد نتائج مطابقة للفلتر.</td></tr>';

    setTxt('#cloud-page-label', `${cloudComparePage + 1} / ${pages}`);
    $('#cloud-prev').disabled = cloudComparePage === 0;
    $('#cloud-next').disabled = cloudComparePage >= pages - 1;
    const hasRows = rows.length > 0;
    const copyBtn = $('#btn-cloud-copy'); if (copyBtn) copyBtn.disabled = !hasRows;
    const xlsxBtn = $('#btn-cloud-export-xlsx'); if (xlsxBtn) xlsxBtn.disabled = !hasRows;
  }

  function cloudCompareTSV() {
    const rows = State.cloudCompareRows || [];
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'سعر الموزع', 'سعر الشركة', 'الفرق', 'مؤشر المراجعة'].join('\t');
    const body = rows.map((x, i) => {
      const dx = x.status === Comparison.STATUS.UNKNOWN ? '—'
        : Math.abs(x.unitDiff) < 1e-9 ? '0.00' : signedNum(x.unitDiff);
      return [i + 1, x.item.itemNumber || '', x.item.name, x.item.unit || '',
        fmtNum(x.invoiceTotal), x.listPrice === null ? '—' : fmtNum(x.listPrice),
        dx, Comparison.statusText(x.status)].join('\t');
    });
    const st = State.cloudCompareStats || {};
    return [header, ...body,
      '',
      ['مطابق', String(st.matched || 0), '', 'زيادة', `${st.high || 0} (${signedNum(st.highTotal || 0)})`, 'انخفاض', `${st.save || 0} (-${fmtNum(st.saveTotal || 0)})`, 'غير مسجل', String(st.unknown || 0)].join('\t'),
    ].join('\n');
  }

  function exportCloudCompareXlsx() {
    const rows = State.cloudCompareRows || [];
    if (!rows.length) { toast('لا توجد نتائج لتصديرها', 'error'); return; }
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'سعر الموزع', 'سعر الشركة', 'الفرق', 'مؤشر المراجعة'];
    const body = rows.map((x) => [
      x.item.itemNumber || '', x.item.name, x.item.unit || '',
      x.invoiceTotal, x.listPrice === null ? '' : x.listPrice,
      x.status === Comparison.STATUS.UNKNOWN ? '' : (Math.abs(x.unitDiff) < 1e-9 ? 0 : x.unitDiff),
      Comparison.statusText(x.status),
    ]);
    Excel.exportXLSX('مقارنة موزع-شركة', [header, ...body], `مقارنة-${todayStr()}.xlsx`)
      .then(() => toast('تم تنزيل تقرير المقارنة Excel ✓'))
      .catch((err) => toast(err && err.message ? err.message : 'تعذّر تصدير Excel', 'error'));
  }

  /* ─────────────────── ربط الأحداث ─────────────────── */
  function bindEvents() {
    // التبويبات
    $$('.tab-btn').forEach((btn) => btn.addEventListener('click', () => switchTab(btn.dataset.tab)));

    // تبديل اللستة المختارة
    $('#quick-list').addEventListener('change', (e) => {
      activateList(e.target.value);
      toast(`تم التبديل إلى «${activeListName()}»`);
    });

    // الإدخال اليدوي للبنود
    $('#btn-add-item').addEventListener('click', addItem);
    ['#item-name', '#item-qty', '#item-price', '#item-unit', '#item-number'].forEach((sel) => {
      $(sel).addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addItem(); } });
    });

    // تعديل/حذف البنود داخل الجدول (تفويض)
    $('#items-table-body').addEventListener('input', (e) => {
      const input = e.target.closest('.cell-input');
      if (input) updateItemField(input.dataset.id, input.dataset.field, input.value);
    });
    $('#items-table-body').addEventListener('change', (e) => {
      const input = e.target.closest('.cell-input');
      if (input && (input.dataset.field === 'quantity' || input.dataset.field === 'unitPrice')) {
        const item = State.items.find((i) => i.id === input.dataset.id);
        if (item) input.value = input.dataset.field === 'quantity' ? item.quantity : item.unitPrice;
      }
    });
    $('#items-table-body').addEventListener('click', (e) => {
      const btn = e.target.closest('.remove-item');
      if (btn) removeItem(btn.dataset.id);
    });

    // حفظ تلقائي للمسودة
    ['#inv-customer', '#inv-no', '#inv-date', '#inv-notes'].forEach((sel) => $(sel).addEventListener('input', saveDraft));

    // أزرار الفاتورة
    $('#btn-save-invoice').addEventListener('click', saveInvoice);
    $('#btn-copy-invoice-json').addEventListener('click', () => {
      if (!State.items.length) { toast('لا توجد بنود للنسخ', 'error'); return; }
      copyText(JSON.stringify({ ...readMeta(), listId: Lists.activeIdOf(), items: State.items }, null, 2));
    });
    $('#btn-copy-report-tsv').addEventListener('click', () => {
      if (!State.items.length) { toast('لا توجد بنود لتوليد التقرير', 'error'); return; }
      copyText(reportToTSV());
    });
    $('#btn-export-xlsx-report').addEventListener('click', exportXlsxReport);
    $('#btn-print').addEventListener('click', printReport);
    $('#btn-new-invoice').addEventListener('click', newInvoiceClick);

    // اللصق السريع — الفاتورة
    $('#btn-apply-paste-invoice').addEventListener('click', applyPasteInvoice);
    $('#btn-clear-paste-invoice').addEventListener('click', () => { $('#paste-invoice').value = ''; });
    $('#btn-paste-sample-invoice').addEventListener('click', () => {
      $('#paste-invoice').value =
        'T-1001\tسكر 1 كجم\tكيس\t10\t30\t300\n' +
        'T-1003\tزيت عباد الشمس 1 لتر\tعبوة\t5\t75\t375\n' +
        'T-1006\tلبن 1 لتر\tعبوة\t8\t24\t192\n' +
        'T-1005\tشاي 500 جم\tعلبة\t4\t100\t400\n' +
        'T-1011\tقهوة تركية 250 جم\tعبوة\t3\t88\t264';
    });
    $('#btn-xlsx-invoice').addEventListener('click', () => $('#xlsx-invoice-file').click());
    $('#xlsx-invoice-file').addEventListener('change', () => handleXlsxFile('#xlsx-invoice-file', '#paste-invoice', 'invoice', 'رتّبت الفاتورة'));

    // اللستات السحابية (بطاقات + تحديث فردي/جماعي)
    $('#cloud-lists-grid').addEventListener('click', (e) => {
      const refresh = e.target.closest('[data-refresh]');
      if (refresh) { e.stopPropagation(); fetchListFromConfig(refresh.dataset.refresh); return; }
      const card = e.target.closest('[data-list-id]');
      if (card && card.dataset.listId) activateList(card.dataset.listId);
    });
    $('#btn-refresh-all-lists').addEventListener('click', refreshAllCloudLists);

    // لستة النشطة (جدول قراءة فقط) + نسخ
    $('#btn-copy-list-tsv').addEventListener('click', () => {
      if (!Lists.active() || !Lists.active().items.length) { toast('اللستة فارغة', 'error'); return; }
      copyText(Lists.listToTSV(Lists.activeIdOf()));
    });
    let searchTimer = null;
    $('#prod-search').addEventListener('input', (e) => {
      clearTimeout(searchTimer);
      searchTimer = setTimeout(() => { State.productFilter = e.target.value; State.listPage = 0; renderListItems(); }, 150);
    });

    // المقارنة الشاملة بين كل اللستات
    $('#btn-cmp-run').addEventListener('click', runAllListCompare);
    $('#cmp-diff-only').addEventListener('change', renderAllListCompare);
    $('#btn-cmp-copy').addEventListener('click', () => {
      if (!cmpState.rows.length) { toast('شغّل المقارنة أولاً', 'error'); return; }
      copyText(allListCmpTSV());
    });
    $('#btn-cmp-export-xlsx').addEventListener('click', () => {
      if (!cmpState.rows.length) { toast('شغّل المقارنة أولاً', 'error'); return; }
      Excel.exportXLSX('مقارنة اللستات', allListCmpAOA(), `مقارنة-لستات-${todayStr()}.xlsx`)
        .then(() => toast('تم تنزيل تقرير المقارنة Excel ✓'))
        .catch((err) => toast(err && err.message ? err.message : 'تعذّر تصدير Excel', 'error'));
    });

    // سجل الفواتير
    $('#btn-copy-history-json').addEventListener('click', () => {
      if (!Invoices.count()) { toast('السجل فارغ', 'error'); return; }
      copyText(Invoices.exportJson());
    });
    $('#btn-copy-history-tsv').addEventListener('click', () => {
      if (!Invoices.count()) { toast('السجل فارغ', 'error'); return; }
      copyText(invoicesToTSV());
    });
    $('#btn-export-invoices').addEventListener('click', () => downloadJSON(Invoices.getAll(), 'سجل-الفواتير.json'));
    $('#btn-import-invoices').addEventListener('click', () => $('#import-invoices-file').click());
    $('#import-invoices-file').addEventListener('change', (e) => {
      importFile(e.target.files[0], (text) => {
        const res = Invoices.importJson(text, false);
        renderHistory();
        toast(`تم الاستيراد: ${res.added} فاتورة جديدة`);
      });
      e.target.value = '';
    });
    $('#btn-sample-invoice').addEventListener('click', () => { const inv = Invoices.addSample(); renderHistory(); toast(`أُضيفت فاتورة تجريبية ${inv.invoiceNo}`); });
    $('#btn-clear-invoices').addEventListener('click', () => {
      confirmBox({
        title: 'مسح سجل الفواتير',
        message: 'سيتم حذف جميع الفواتير المحفوظة نهائياً (المسودة الحالية تبقى).',
        confirmLabel: 'نعم، امسح السجل',
        danger: true,
        onConfirm: () => { Invoices.clearAll(); renderHistory(); toast('تم مسح السجل'); },
      });
    });
    $('#history-list').addEventListener('click', (e) => {
      const open = e.target.closest('[data-open]');
      const exp = e.target.closest('[data-exp]');
      const del = e.target.closest('[data-del]');
      if (open) { loadInvoiceIntoEditor(open.dataset.open); return; }
      if (exp) { const inv = Invoices.getById(exp.dataset.exp); downloadJSON({ invoice: inv }, `فاتورة-${inv.invoiceNo || inv.date}.json`); return; }
      if (del) {
        const inv = Invoices.getById(del.dataset.del);
        if (!inv) return;
        confirmBox({ title: 'حذف فاتورة', message: `هل تريد حذف فاتورة ${inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || 'بدون عميل'}؟`, confirmLabel: 'حذف', danger: true, onConfirm: () => { Invoices.remove(inv.id); renderHistory(); toast('تم حذف الفاتورة'); } });
      }
    });

    // أزرار Excel من شاشة البيانات (للفاتورة فقط)
    $('#btn-xlsx-invoice-2').addEventListener('click', () => $('#xlsx-invoice-file-2').click());
    $('#xlsx-invoice-file-2').addEventListener('change', () => handleXlsxFile('#xlsx-invoice-file-2', '#paste-invoice', 'invoice', 'رتّبت الفاتورة'));

    // تخزين مفتاح API وقيمتها عند الكتابة (للمقارنة السحابية)
    $('#gs-api-key').addEventListener('input', (e) => {
      State.settings.sheetApiKey = e.target.value.trim();
      Storage.saveSettings(State.settings);
    });

    // لوحة المقارنة السحابية (موزع ضد شركة)
    const cloudReady = () => !!(cloudTabs && State.cloudCompareRows && State.cloudCompareRows.length);
    $('#btn-cloud-fetch').addEventListener('click', handleCloudFetch);
    $('#cloud-url').addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); handleCloudFetch(); } });
    $('#btn-cloud-run-compare').addEventListener('click', runCloudCompare);
    $('#cloud-ignore-unit').addEventListener('change', () => { if (cloudTabs) runCloudCompare(); });
    $('#cloud-diff-only').addEventListener('change', () => { cloudComparePage = 0; renderCloudCompare(); });
    $('#btn-cloud-copy').addEventListener('click', () => {
      if (!cloudReady()) { toast('شغّل المقارنة أولاً', 'error'); return; }
      copyText(cloudCompareTSV());
    });
    $('#btn-cloud-export-xlsx').addEventListener('click', exportCloudCompareXlsx);
    $('#cloud-prev').addEventListener('click', () => { if (cloudComparePage > 0) { cloudComparePage--; renderCloudCompare(); } });
    $('#cloud-next').addEventListener('click', () => { cloudComparePage++; renderCloudCompare(); });
  }

  /* ─────────────────── الإقلاع ─────────────────── */
  function init() {
    Lists.init();
    Invoices.init();
    // نموذج «اللستات سحابية فقط»: فرض اللستات الخمس من config (يحتفظ بالمعرّفات القائمة)
    if ((CONFIG.LISTS || []).length) Lists.reconcile(CONFIG.LISTS);
    State.settings = Storage.getSettings();
    if (Lists.active()) {
      State.settings.activeListId = Lists.activeIdOf();
      Storage.saveSettings(State.settings);
    }

    if (!$('#inv-date').value) $('#inv-date').value = todayStr();

    const apiKeyEl = $('#gs-api-key');
    if (apiKeyEl) apiKeyEl.value = State.settings.sheetApiKey || '';

    loadDraftIfAny();
    seedCloudPanel();
    renderQuickList();
    renderCloudLists();
    renderDatalist();
    renderListItems();
    renderInvoiceItems();
    renderHistory();
    bindEvents();
    switchTab('invoice');
    autoFetchCloudLists(); // جلب تلقائي للست الناقص/القديم (خلفية، لا يمنع العمل)
  }

  document.addEventListener('DOMContentLoaded', init);

  global.App = { switchTab, renderHistory, updateAllRowsAndSummary, reportToTSV, invoicesToTSV };
})(window);