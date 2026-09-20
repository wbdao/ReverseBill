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
                <p class="text-[11px] text-slate-400 truncate" dir="ltr">${esc(cfg.slug || '')}</p>
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
    cmpAreaShown = true;
    const wa = $('#cmp-area');
    if (wa) wa.classList.remove('hidden');
    cmpState.rows = Comparison.compareAllLists(global.appLists);
    renderAllListCompare();
    const el = $('#cmp-status');
    if (el) el.textContent = `${cmpState.rows.length} صنف موحّد عبر ${(CONFIG.LISTS || []).length} لستة`;
  }

  function runAllListCompareSilent() {
    if (!cmpAreaShown) return; // خفيف: لا تُحسب المقارنة في الخلفية إلا إذا فتح المستخدم قسمها
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
    return `<span class="status-badge status-${r.status}">${Comparison.statusLabel(r.status, r.invoiceDiscountPct)}</span>`;
  }
  function diffHTML(value, status) {
    if (status === Comparison.STATUS.UNKNOWN) return '<span class="clr-neutral">—</span>';
    if (Math.abs(value) < 1e-9) return '<span class="clr-neutral">بدون فرق</span>';
    const cls = status === Comparison.STATUS.HIGH ? 'clr-high' : status === Comparison.STATUS.LOW ? 'clr-low' : 'clr-neutral';
    return `<span class="${cls}">${signedMoney(value)}</span>`;
  }
  // فرق السعر قبل الخصم عن السعر المعتمد: أعلى من المعتمد (تحذير) / أدنى من المعتمد (جيد)
  function prediffHTML(r) {
    if (r.listPrice === null) return '<span class="clr-neutral">—</span>';
    if (Math.abs(r.unitDiffBefore) < 1e-9) return '<span class="clr-neutral">يساوي المعتمد</span>';
    const cls = r.unitDiffBefore > 0 ? 'clr-abv' : 'clr-bel';
    return `<span class="${cls}">${signedMoney(r.unitDiffBefore)}</span>`;
  }
  // عرض نسبة الخصم كرقم بدون علامة % (1% تظهر 1) مع قص الفضلة العشرية المتكررة
  const fmtPct = (p) => { const x = Math.round((num(p) * 100) * 100) / 100; return String(x); };
  function itemRowHTML(item, index) {
    const r = curAnalyze(item);
    return `
      <tr data-id="${item.id}" class="border-t border-slate-100 hover:bg-slate-50/70">
        <td class="py-2 px-3 text-slate-400 text-xs font-bold">${index + 1}</td>
        <td class="py-2 px-3 w-28">
          <input type="text" value="${esc(item.itemNumber || '')}" placeholder="—"
                 data-field="itemNumber" data-id="${item.id}" autocomplete="off"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 min-w-56">
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
        <td class="py-2 px-3 w-28">
          <input type="number" min="0" step="any" value="${item.discountValue ? fmtNum(item.discountValue) : ''}"
                 data-field="discountValue" data-id="${item.id}" placeholder="0.00"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 w-24">
          <input type="number" min="0" step="any" value="${item.discountPct ? fmtPct(item.discountPct) : ''}"
                 data-field="discountPct" data-id="${item.id}" placeholder="0"
                 class="cell-input w-full rounded-lg border border-slate-200 px-2.5 py-1.5 text-sm outline-none focus:border-indigo-400" />
        </td>
        <td class="py-2 px-3 font-bold cell-net">${money(r.netUnit)}</td>
        <td class="py-2 px-3 font-bold cell-total">${money(r.invoiceTotal)}</td>
        <td class="py-2 px-3 text-slate-600 cell-list">${r.listPrice === null ? '<span class="clr-neutral">—</span>' : money(r.listPrice)}</td>
        <td class="py-2 px-3 cell-prediff">${prediffHTML(r)}</td>
        <td class="py-2 px-3 cell-udiff">${diffHTML(r.unitDiff, r.status)}</td>
        <td class="py-2 px-3 cell-tdiff">${diffHTML(r.totalDiff, r.status)}</td>
        <td class="py-2 px-3 cell-status">${statusBadgeHTML(r)}</td>
        <td class="py-2 px-3">
          <button class="remove-item w-8 h-8 grid place-items-center rounded-lg text-rose-500 hover:bg-rose-50" data-id="${item.id}" title="حذف البند"><i class="fa-solid fa-xmark"></i></button>
        </td>
      </tr>`;
  }

  function renderInvoiceItems() {
    // تحليل البنود دفعة واحدة: عدّ الملاحظات وتصفية الصفوف حسب الملاحظة المختارة
    const counts = { match: 0, high: 0, low: 0, unknown: 0 };
    const statusById = new Map();
    for (const it of State.items) {
      const st = curAnalyze(it).status;
      statusById.set(it.id, st);
      if (counts[st] !== undefined) counts[st]++;
    }
    const f = State.statusFilter;
    const filtered = f ? State.items.filter((it) => statusById.get(it.id) === f) : State.items;
    const hasItems = State.items.length > 0;

    $('#items-table-body').innerHTML = filtered.map(itemRowHTML).join('')
      || (hasItems ? '<tr><td colspan="15" class="py-10 text-center text-slate-400 text-sm">لا توجد بنود مطابقة لهذه الملاحظة.</td></tr>' : '');
    $('#empty-items').classList.toggle('hidden', hasItems);
    $('#items-table-wrap').classList.toggle('hidden', !hasItems);
    $('#items-count').textContent = hasItems
      ? `عدد البنود: ${State.items.length}${f ? ` · المعروض: ${filtered.length}` : ''} · اللستة المختارة: ${activeListName()}`
      : `اللستة المختارة: ${activeListName()}`;

    const countsEl = $('#status-counts');
    if (countsEl) {
      countsEl.innerHTML = hasItems
        ? `🟢 ${counts.match} · 🔺 ${counts.high} · 🔻 ${counts.low} · ⚠️ ${counts.unknown}`
        : '';
    }
    $$('#items-filter-bar [data-status-filter]').forEach((b) => b.classList.toggle('active', (b.dataset.statusFilter || '') === f));
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
    const r = curAnalyze(item);
    tr.querySelector('.cell-net').innerHTML = money(r.netUnit);
    tr.querySelector('.cell-total').innerHTML = money(r.invoiceTotal);
    tr.querySelector('.cell-list').innerHTML = r.listPrice === null ? '<span class="clr-neutral">—</span>' : money(r.listPrice);
    tr.querySelector('.cell-prediff').innerHTML = prediffHTML(r);
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
    } else if (field === 'discountValue') {
      item.discountValue = Math.max(0, parseNum(rawValue));
    } else if (field === 'discountPct') {
      // الادخال اليدوي بالنسب (1 = 1%) يُخزَّن ككسر (0.01) للاتساق مع سحب السيستم
      item.discountPct = Math.max(0, parseNum(rawValue) / 100);
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

    const discountValue = Math.max(0, parseNum($('#item-discount-value').value));
    const discountPctRaw = $('#item-discount-pct').value.trim();
    // خصم النسبة يُدخل كنسبة (1 = 1%) ويُخزن ككسر (0.01) للاتساق مع سحب السيستم
    const discountPct = discountPctRaw === '' ? 0 : Math.max(0, parseNum(discountPctRaw) / 100);

    State.items.push({ id: genId('it'), itemNumber, name: name || itemNumber, unit, quantity: qty, unitPrice, discountValue, discountPct });
    renderInvoiceItems();
    saveDraft();
    $('#item-qty').value = '1';
    $('#item-price').value = '';
    $('#item-discount-value').value = '';
    $('#item-discount-pct').value = '';
    $('#item-unit').value = $('#item-number').value = '';
    $('#item-name').value = '';
    $('#item-name').focus();
  }

  /* ─────────────────── ملخص الفاتورة ─────────────────── */
  function updateSummary() {
    const s = curSummarize();
    const t = s.totals;
    let netLabel = 'صافي الفرق';
    let netDot = 'bg-slate-400';
    if (s.netDiff > 0.004) { netLabel = 'زيادة صافية عن المعتمد'; netDot = 'bg-rose-500'; }
    else if (s.netDiff < -0.004) { netLabel = 'انخفاض صافٍ عن المعتمد'; netDot = 'bg-amber-500'; }
    const cards = [
      { label: 'إجمالي الفاتورة', value: money(t.invoiceTotal), dot: 'bg-indigo-500', sub: `${State.items.length} بند · ${money(s.expectedTotal)} معتمد` },
      { label: 'خصم الاتفاقية على الأصناف', value: State.invoiceDiscountPct ? `${fmtPct(State.invoiceDiscountPct)}%` : 'لا يوجد', dot: 'bg-emerald-500', sub: State.invoiceDiscountPct ? 'مطبّق على كل البنود في المقارنة' : 'أدخل النسبة في «خصم الاتفاقية %» لتطبيقها' },
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
  // نسبة مئوية مُدخلة (مثل 5) → كسر (0.05). القيم تُقيّد بين 0 و100.
  const pctToFraction = (raw) => { const n = parseNum(raw); return isNaN(n) ? 0 : Math.max(0, Math.min(100, n)) / 100; };
  // تحليل بنود الفاتورة الحالية مع خصم الاتفاقية المطبّق على كل الأصناف
  const curAnalyze = (it) => Comparison.analyzeItem(it, undefined, State.invoiceDiscountPct);
  const curSummarize = () => Comparison.summarize(State.items, undefined, State.invoiceDiscountPct);
  const readMeta = () => ({ customer: $('#inv-customer').value, invoiceNo: $('#inv-no').value, date: $('#inv-date').value, notes: $('#inv-notes').value, invoiceDiscountPct: pctToFraction($('#inv-discount-pct').value) });
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
    State.invoiceDiscountPct = Number(draft.invoiceDiscountPct) || 0;
    $('#inv-discount-pct').value = State.invoiceDiscountPct ? fmtPct(State.invoiceDiscountPct) : '';
    State.items = Array.isArray(draft.items) ? draft.items.map((it) => ({
      id: genId('it'),
      itemNumber: String(it.itemNumber || '').trim(),
      name: String(it.name || '').trim(),
      unit: String(it.unit || '').trim(),
      quantity: it.quantity ?? 1,
      unitPrice: Number(it.unitPrice) || 0,
      discountValue: Number(it.discountValue) || 0,
      discountPct: Number(it.discountPct) || 0,
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
    State.invoiceDiscountPct = 0;
    updateSaveButton();
    $('#inv-customer').value = $('#inv-no').value = $('#inv-notes').value = '';
    $('#inv-discount-pct').value = '';
    $('#inv-date').value = todayStr();
    $('#item-name').value = $('#item-price').value = $('#item-unit').value = $('#item-number').value = '';
    $('#item-discount-value').value = $('#item-discount-pct').value = '';
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
      State.items = inv.items.map((it) => ({ id: genId('it'), itemNumber: it.itemNumber || '', name: it.name, unit: it.unit || '', quantity: it.quantity, unitPrice: it.unitPrice, discountValue: Number(it.discountValue) || 0, discountPct: Number(it.discountPct) || 0 }));
      if (inv.listId && Lists.getList(inv.listId)) { Lists.setActive(inv.listId); renderQuickList(); renderCloudLists(); renderDatalist(); }
      State.editingInvoiceId = inv.id;
      updateSaveButton();
      $('#inv-customer').value = inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || '';
      $('#inv-no').value = inv.invoiceNo || '';
      $('#inv-date').value = inv.date || todayStr();
      $('#inv-notes').value = inv.notes || '';
      State.invoiceDiscountPct = Number(inv.invoiceDiscountPct) || 0;
      $('#inv-discount-pct').value = State.invoiceDiscountPct ? fmtPct(State.invoiceDiscountPct) : '';
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

  /* ─────────────────── العرض عند الطلب (تخفيف الموقع) ───────────────────
     جداول اللستة والمقارنة الشاملة تُبنى فقط بعد ضغط المستخدم لزر «عرض...» */
  let listAreaShown = false;
  let cmpAreaShown = false;
  const showListArea = () => {
    listAreaShown = true;
    const area = $('#list-table-area');
    if (area) area.classList.remove('hidden');
    renderListItems();
  };
  const showCmpArea = () => {
    cmpAreaShown = true;
    const area = $('#cmp-area');
    if (area) area.classList.remove('hidden');
    if (cmpState.rows.length) renderAllListCompare();
    else runAllListCompare();
  };

  /* ─────────────────── قوائم الأسعار: الواجهة (+ ترقيم للكبير) ─────────────────── */
  function renderListItems() {
    if (!listAreaShown) return; // لا يُبنى الجدول في الخلفية — يُعرض عند الطلب
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
      const s = Comparison.summarize(inv.items, refList, Number(inv.invoiceDiscountPct) || 0);
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
    const s = curSummarize();
    const t = s.totals;
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'خصم القيمة', 'خصم النسبة', 'صافي السعر', 'إجمالي المبلغ', 'السعر المعتمد', 'فرق قبل الخصم', 'فرق الوحدة', 'إجمالي الفرق', 'مؤشر المراجعة'].join('\t');
    const rows = State.items.map((it) => {
      const r = curAnalyze(it);
      const ud = r.status === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.unitDiff) < 1e-9 ? '0.00' : signedNum(r.unitDiff));
      const td = r.status === Comparison.STATUS.UNKNOWN ? '—' : (Math.abs(r.totalDiff) < 1e-9 ? '0.00' : signedNum(r.totalDiff));
      return [
        it.itemNumber || '—', it.name, it.unit || '—',
        fmtNum(it.quantity), fmtNum(it.unitPrice),
        fmtNum(it.discountValue || 0), fmtPct(it.discountPct || 0), fmtNum(r.netUnit),
        fmtNum(r.invoiceTotal),
        r.listPrice === null ? '—' : fmtNum(r.listPrice),
        r.listPrice === null ? '—' : (Math.abs(r.unitDiffBefore) < 1e-9 ? '0.00' : signedNum(r.unitDiffBefore)),
        ud, td, Comparison.statusText(r.status, r.invoiceDiscountPct),
      ].join('\t');
    });
    const summary = [
      '',
      ['', '', '', '', 'إجمالي الفاتورة', '', '', '', fmtNum(t.invoiceTotal), 'الإجمالي المعتمد', '', '', fmtNum(t.expectedTotal)].join('\t'),
      ['', '', '', '', 'إجمالي الزيادة', '', '', '', fmtNum(t.highTotal), 'إجمالي الانخفاض', '', '', fmtNum(t.lowTotal)].join('\t'),
      ['', '', '', '', 'صافي الفرق', '', '', '', signedNum(s.netDiff), 'بنود غير مسجلة', '', '', String(t.unknownCount)].join('\t'),
    ];
    return [header, ...rows, ...summary].join('\n');
  }

  /** مصفوفة خلايا لتصدير Excel مباشرة (أرقام حقيقية قابلة للحساب) */
  function reportToAOACells() {
    const header = ['رقم الصنف', 'اسم المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'خصم القيمة', 'خصم النسبة', 'صافي السعر', 'إجمالي المبلغ', 'السعر المعتمد', 'فرق قبل الخصم', 'فرق الوحدة', 'إجمالي الفرق', 'مؤشر المراجعة'];
    const body = State.items.map((it) => {
      const r = curAnalyze(it);
      return [
        it.itemNumber || '', it.name, it.unit || '',
        it.quantity, it.unitPrice, it.discountValue || 0, it.discountPct || 0, r.netUnit, r.invoiceTotal,
        r.listPrice === null ? '' : r.listPrice,
        r.listPrice === null ? '' : (Math.abs(r.unitDiffBefore) < 1e-9 ? 0 : r.unitDiffBefore),
        r.status === Comparison.STATUS.UNKNOWN ? '' : (Math.abs(r.unitDiff) < 1e-9 ? 0 : r.unitDiff),
        r.status === Comparison.STATUS.UNKNOWN ? '' : (Math.abs(r.totalDiff) < 1e-9 ? 0 : r.totalDiff),
        Comparison.statusText(r.status, r.invoiceDiscountPct),
      ];
    });
    return [header, ...body];
  }

  function invoicesToTSV() {
    const header = ['العميل', 'رقم الفاتورة', 'التاريخ', 'رقم الصنف', 'المنتج', 'الوحدة', 'الكمية', 'سعر الوحدة', 'خصم القيمة', 'خصم النسبة', 'صافي السعر', 'إجمالي البند'].join('\t');
    const rows = Invoices.getAll().flatMap((inv) =>
      inv.items.map((it) => {
        const r = Comparison.analyzeItem(it, undefined, Number(inv.invoiceDiscountPct) || 0);
        return [
          inv.customer !== undefined && inv.customer !== null ? inv.customer : inv.vendor || '', inv.invoiceNo || '', inv.date || '',
          it.itemNumber || '', it.name || '', it.unit || '',
          fmtNum(it.quantity), fmtNum(it.unitPrice), fmtNum(it.discountValue || 0), fmtPct(it.discountPct || 0), fmtNum(r.netUnit), fmtNum(r.invoiceTotal),
        ].join('\t');
      }));
    return [header, ...rows].join('\n');
  }

  function exportXlsxReport() {
    if (!State.items.length) { toast('لا توجد بنود لتصديرها', 'error'); return; }
    Excel.exportXLSX('تقرير المراجعة', reportToAOACells(), `تقرير-مراجعة-${todayStr()}.xlsx`)
      .then(() => toast('تم تنزيل تقرير Excel ✓'))
      .catch((err) => toast(err.message || 'تعذّر تصدير Excel', 'error'));
  }

  /* ═══════════ معمارية الطباعة المعزولة (نافذة مؤقتة ذاتية الدعم) ═══════════
     مستند HTML كامل ومستقل يُبنى في iframe خفي بلا أي CSS من واجهة التطبيق،
     فتُفرض ألوان سوداء صريحة في كل النصوص وخلفيات بيضاء في الترويسة —
     يستحيل معها ظهور "أبيض على أبيض". كل عناصر التحكم في الواجهة خارج هذه
     النافذة فلا تُطبع إطلاقاً. الطباعة لا تُستدعى قبل التحقق من امتلاء <tbody>. */

  /* 1) توليد صف <tr> لكل بند من مصفوفة البنود الحالية عبر محرك المقارنة */
  function printRowHTML(it, i) {
    try {
      const r = curAnalyze(it);
      const st = r.status;
      const fmtDiff = (v) => {
        if (st === Comparison.STATUS.UNKNOWN || Math.abs(v) < 1e-9) return '—';
        return v > 0 ? `+${fmtNum(v)}` : fmtNum(v);
      };
      return `<tr>
        <td>${i + 1}</td>
        <td>${esc(it.itemNumber || '—')}</td>
        <td>${esc(it.name || '—')}</td>
        <td>${esc(it.unit || '—')}</td>
        <td>${fmtNum(it.quantity)}</td>
        <td>${fmtNum(it.unitPrice)}</td>
        <td>${fmtNum(it.discountValue || 0)}</td>
        <td>${fmtPct(it.discountPct || 0)}</td>
        <td>${fmtNum(r.netUnit)}</td>
        <td>${fmtNum(r.invoiceTotal)}</td>
        <td>${r.listPrice == null ? '—' : fmtNum(r.listPrice)}</td>
        <td>${fmtDiff(r.unitDiff)}</td>
        <td>${fmtDiff(r.totalDiff)}</td>
        <td>${Comparison.statusText(st, r.invoiceDiscountPct)}</td>
      </tr>`;
    } catch (e) {
      // بند معطوب لا يُفرّغ الجدول كاملاً — صف احتياطي يُبقي التقرير مقروءاً
      return `<tr>
        <td>${i + 1}</td>
        <td>${esc(it.itemNumber || '—')}</td>
        <td>${esc(it.name || '—')}</td>
        <td>${esc(it.unit || '—')}</td>
        <td>${fmtNum(it.quantity)}</td>
        <td>${fmtNum(it.unitPrice)}</td>
        <td colspan="8">بيانات غير مكتملة</td>
      </tr>`;
    }
  }

  /* 2) بطاقات الملخص + معادلة الصافي (داخل قالب الطباعة مباشرة تحت الجدول) */
  function buildPrintTotalsHTML() {
    const s = curSummarize();
    const t = s.totals;
    const isHigh = s.netDiff > 0.004, isLow = s.netDiff < -0.004;
    const cards = [
      { label: 'إجمالي الفاتورة', value: money(t.invoiceTotal) },
      { label: 'إجمالي زيادة الأسعار', value: money(t.highTotal) },
      { label: 'إجمالي الانخفاض عن المعتمد', value: money(t.lowTotal) },
      { label: 'صافي الفرق', value: signedMoney(s.netDiff), cls: isHigh ? 'eq-high' : (isLow ? 'eq-low' : 'eq-ok') },
    ];
    return `
      <div class="sum-cards">
        ${cards.map((c) => `<div class="sum-card"><span>${c.label}</span><b class="${c.cls || ''}">${c.value}</b></div>`).join('')}
      </div>
      <div class="equation">
        الإجمالي المعتمد: <b>${money(t.expectedTotal)}</b> &nbsp;·&nbsp;
        صافي الفرق عن المعتمد: <b>${signedMoney(s.netDiff)}</b> &nbsp;·&nbsp;
        الانحراف: <b>${fmtNum(s.deviationPct)}%</b> &nbsp;·&nbsp;
        البنود غير المسجلة بالقائمة: <b>${t.unknownCount}</b> (${money(t.unknownTotal)})<br/>
        ${State.invoiceDiscountPct ? `خصم الاتفاقية المطبّق على كل الأصناف: <b>${fmtPct(State.invoiceDiscountPct)}%</b><br/>` : ''}
        ${isHigh ? '<span class="eq-high">زيادة صافية عن المعتمد</span>' : isLow ? '<span class="eq-low">انخفاض صافٍ عن المعتمد</span>' : '<span class="eq-ok">الفاتورة متوازنة مع المعتمد</span>'}
      </div>`;
  }

  /* 3) بناء مستند الطباعة الكامل — هيكل نظامي (ترويسة/بيانات/جدول/ملخص/تذييل)
        بكل تنسيقاته inline: خط أسود صريح + خلفية بيضاء على كل خلية ترويسة */
  function buildPrintDocumentHTML() {
    const meta = readMeta();
    const s = curSummarize();
    const t = s.totals;
    const customer = (meta.customer !== undefined && meta.customer !== null) ? meta.customer : meta.vendor || '—';
    const qtySum = State.items.reduce((a, it) => a + (Number(it.quantity) || 0), 0);
    const isHigh = s.netDiff > 0.004, isLow = s.netDiff < -0.004;
    const rowsHTML = State.items.map(printRowHTML).join('');

    return `<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
<meta charset="UTF-8">
<title>تقرير مراجعة فاتورة — ${esc(meta.invoiceNo || '')}</title>
<style>
  * { box-sizing: border-box; }
  @page { size: A4 portrait; margin: 12mm 10mm; }
  html, body { margin: 0; padding: 0; background: #fff; font-family: 'Tajawal','Segoe UI',Arial,sans-serif; color: #000000; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
  h1 { font-size: 20px; margin: 0 0 4px; color: #000000; }
  .title { display: flex; justify-content: space-between; align-items: center; border-bottom: 3px solid #000000; padding-bottom: 10px; margin-bottom: 12px; page-break-after: avoid; }
  .meta { font-size: 12px; color: #000000; line-height: 1.8; }
  .meta b { color: #000000; }
  table { width: 100%; border-collapse: collapse; font-size: 12px; margin: 12px 0; color: #000000; }
  thead { display: table-header-group; }
  th, td { border: 1px solid #000000; padding: 7px 9px; text-align: right; vertical-align: top; color: #000000 !important; }
  thead th { background: #ffffff !important; color: #000000 !important; font-weight: 800; font-size: 11px; }
  tbody tr { page-break-inside: avoid; break-inside: avoid; }
  tfoot td { font-weight: 800; border-top: 2px solid #000000; background: #ffffff !important; color: #000000 !important; }
  .sum-cards { display: grid; grid-template-columns: repeat(4, 1fr); gap: 10px; margin: 14px 0; page-break-inside: avoid; }
  .sum-card { border: 1px solid #000000; border-radius: 8px; padding: 9px 11px; font-size: 11px; color: #000000; background: #ffffff !important; }
  .sum-card b { display: block; font-size: 15px; margin-top: 2px; color: #000000 !important; }
  .equation { font-size: 12px; line-height: 2; border-top: 2px solid #000000; padding-top: 10px; margin-top: 8px; page-break-inside: avoid; color: #000000; }
  .eq-high { color: #b91c1c !important; font-weight: 700; }
  .eq-low  { color: #b45309 !important; font-weight: 700; }
  .eq-ok   { color: #047857 !important; font-weight: 700; }
  .note { border-top: 1px dashed #6b7280; margin-top: 16px; padding-top: 10px; font-size: 12px; color: #000000; page-break-inside: avoid; }
  .footer { margin-top: 34px; display: flex; justify-content: space-between; align-items: flex-end; font-size: 12px; color: #000000; page-break-inside: avoid; }
  .footer .sign { text-align: center; }
  .footer .line { width: 150px; border-top: 1px solid #374151; margin-top: 4px; }
</style>
</head>
<body>
  <div class="title">
    <div><h1>تقرير مراجعة فاتورة العميل ومقارنة أسعار البيع</h1></div>
    <div class="meta">تاريخ الطباعة: ${new Date().toLocaleString('ar-EG')}</div>
  </div>
  <div class="meta">
    <div><b>العميل:</b> ${esc(customer)} &nbsp;|&nbsp; <b>رقم الفاتورة:</b> ${esc(meta.invoiceNo || '—')}</div>
    <div><b>تاريخ الفاتورة:</b> ${esc(meta.date || todayStr())} &nbsp;|&nbsp; <b>القائمة المعتمدة:</b> ${esc(activeListName())}${meta.invoiceDiscountPct ? ` &nbsp;|&nbsp; <b>خصم الاتفاقية:</b> ${fmtPct(meta.invoiceDiscountPct)}%` : ''}</div>
  </div>
  <table>
    <thead>
      <tr>
        <th>#</th><th>رقم الصنف</th><th>اسم المنتج</th><th>الوحدة</th><th>الكمية</th><th>سعر الوحدة</th>
        <th>خصم القيمة</th><th>خصم النسبة</th><th>صافي السعر</th><th>إجمالي المبلغ</th><th>السعر المعتمد</th><th>فرق الوحدة</th><th>إجمالي الفرق</th><th>المراجعة</th>
      </tr>
    </thead>
    <tbody>${rowsHTML || '<tr><td colspan="14" style="text-align:center">لا توجد بنود</td></tr>'}</tbody>
    <tfoot>
      <tr>
        <td colspan="4">إجمالي البنود (${State.items.length})</td>
        <td>${fmtNum(qtySum)}</td>
        <td></td><td></td><td></td><td></td>
        <td>${fmtNum(t.invoiceTotal)}</td>
        <td>${fmtNum(t.expectedTotal)}</td>
        <td colspan="2">${signedMoney(s.netDiff)}</td>
        <td>${isHigh ? 'زيادة صافية' : isLow ? 'انخفاض صافٍ' : 'متوازن'}</td>
      </tr>
    </tfoot>
  </table>
  ${buildPrintTotalsHTML()}
  ${meta.notes ? `<div class="note"><b>ملاحظات:</b> ${esc(meta.notes)}</div>` : ''}
  <div class="footer">
    <div class="sign">توقيع المراجع<div class="line"></div></div>
    <div class="sign">خاتم الشركة<div class="line"></div></div>
    <div class="sign">توقيع المستلم<div class="line"></div></div>
  </div>
</body>
</html>`;
  }

  /* 4) نافذة طباعة احتياطية (لا تُستخدم إلا إذا حُجبت النوافذ المنبثقة) */
  let _printFrame = null;
  function getPrintFrame() {
    if (_printFrame && _printFrame.contentDocument) return _printFrame;
    _printFrame = document.createElement('iframe');
    _printFrame.style.cssText = 'position: fixed; right: -10000px; top: 0; width: 0; height: 0; border: 0; visibility: hidden;';
    _printFrame.setAttribute('aria-hidden', 'true');
    _printFrame.setAttribute('title', 'منطقة طباعة تقرير الفاتورة');
    document.body.appendChild(_printFrame);
    return _printFrame;
  }

  /* 5) الدالة الأم — الطباعة عبر نافذة منبثقة معزولة بالكامل (الأضمن عبر المتصفحات:
        Chrome/Safari ينبو عن طباعة iframe بخاصية display:none فيطبع صفحة فارغة).
        لا تُستدعى window.print إلا بعد التحقق من امتلاء <tbody> بصفوف فعلية. */
  async function printInvoiceReport() {
    if (!State.items.length) { toast('لا توجد بنود للطباعة بعد', 'error'); return; }
    const html = buildPrintDocumentHTML();               // (أ) الهيكل الكامل مع الصفوف
    const rowsIn = (doc) => {                            // (ب) عدّاد صفوف الحقن
      if (!doc) return 0;
      const tb = doc.querySelector('table tbody');
      return tb ? tb.querySelectorAll('tr').length : 0;
    };

    // — المسار الأساسي: نافذة منبثقة within user-gesture (بدون iframe) —
    let win = null;
    try { win = window.open('', '_blank'); } catch (e) { win = null; }
    if (win && win.document) {
      const doc = win.document;
      doc.open();
      doc.write(html);
      doc.close();
      if (doc.readyState !== 'complete') {
        await new Promise((r) => { win.onload = r; setTimeout(r, 400); });
      }
      const injected = rowsIn(doc);
      if (!injected) {
        try { win.close(); } catch (e) { /* يتجاهل */ }
        toast(`تعذّر حقن صفوف البنود (0 من ${State.items.length})`, 'error');
        return;
      }
      console.log(`طباعة: ${injected} صف بنود من ${State.items.length} بند`);
      win.focus();
      // مهلة قصيرة لاستقرار التخطيط ثم الطباعة وإغلاق النافذة لاحقاً تلقائياً
      setTimeout(() => {
        try {
          win.print();
          win.onafterprint = () => { try { win.close(); } catch (e) { /* يتجاهل */ } };
        } catch (e) {
          toast('تعذّر فتح نافذة الطباعة — اسمح بالنوافذ المنبثقة لهذا الموقع', 'error');
          try { win.close(); } catch (e2) { /* يتجاهل */ }
        }
      }, 60);
      return;
    }

    // — السقوط: iframe خفي إن مُنعت النوافذ المنبثقة —
    const frame = getPrintFrame();
    const doc = frame.contentDocument;
    doc.open();
    doc.write(html);
    doc.close();
    await new Promise((r) => { if (doc.readyState === 'complete') r(); else { frame.onload = r; setTimeout(r, 300); } });
    const injected = rowsIn(doc);
    if (!injected) { toast(`تعذّر حقن صفوف البنود (0 من ${State.items.length})`, 'error'); return; }
    console.log(`طباعة (iframe): ${injected} صف بنود من ${State.items.length} بند`);
    await new Promise((r) => requestAnimationFrame(() => requestAnimationFrame(() => r())));
    frame.contentWindow.focus();
    frame.contentWindow.print();
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
    ['#item-name', '#item-qty', '#item-price', '#item-unit', '#item-number', '#item-discount-value', '#item-discount-pct'].forEach((sel) => {
      $(sel).addEventListener('keydown', (e) => { if (e.key === 'Enter') { e.preventDefault(); addItem(); } });
    });

    // تعديل/حذف البنود داخل الجدول (تفويض)
    $('#items-table-body').addEventListener('input', (e) => {
      const input = e.target.closest('.cell-input');
      if (input) updateItemField(input.dataset.id, input.dataset.field, input.value);
    });
    $('#items-table-body').addEventListener('change', (e) => {
      const input = e.target.closest('.cell-input');
      if (input && ['quantity', 'unitPrice', 'discountValue', 'discountPct'].includes(input.dataset.field)) {
        const item = State.items.find((i) => i.id === input.dataset.id);
        if (item) {
          const f = input.dataset.field;
          input.value = f === 'quantity' ? item.quantity
            : f === 'unitPrice' ? item.unitPrice
            : f === 'discountValue' ? fmtNum(item.discountValue)
            : fmtPct(item.discountPct);
        }
      }
    });
    $('#items-table-body').addEventListener('click', (e) => {
      const btn = e.target.closest('.remove-item');
      if (btn) removeItem(btn.dataset.id);
    });

    // تصفية بنود الفاتورة حسب ملاحظة مراجعة السعر
    $('#items-filter-bar').addEventListener('click', (e) => {
      const chip = e.target.closest('[data-status-filter]');
      if (!chip) return;
      State.statusFilter = chip.dataset.statusFilter || '';
      renderInvoiceItems();
    });

    // حفظ تلقائي للمسودة
    ['#inv-customer', '#inv-no', '#inv-date', '#inv-notes'].forEach((sel) => $(sel).addEventListener('input', saveDraft));

    // خصم الاتفاقية على كل الأصناف: يحدّث الحالة والتحليل والمسودة فوراً
    $('#inv-discount-pct').addEventListener('input', () => {
      State.invoiceDiscountPct = pctToFraction($('#inv-discount-pct').value);
      updateAllRowsAndSummary();
      saveDraft();
    });

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
    $('#btn-print').addEventListener('click', printInvoiceReport);
    $('#btn-new-invoice').addEventListener('click', newInvoiceClick);

    // اللصق السريع — الفاتورة
    $('#btn-apply-paste-invoice').addEventListener('click', applyPasteInvoice);
    $('#btn-clear-paste-invoice').addEventListener('click', () => { $('#paste-invoice').value = ''; });
    $('#btn-paste-sample-invoice').addEventListener('click', () => {
      $('#paste-invoice').value =
        'T-1001\tسكر 1 كجم\tكيس\t10\t30\t2\t5\n' +
        'T-1003\tزيت عباد الشمس 1 لتر\tعبوة\t5\t75\t0\t0\n' +
        'T-1006\tلبن 1 لتر\tعبوة\t8\t24\t1\t0\n' +
        'T-1005\tشاي 500 جم\tعلبة\t4\t100\t0\t10\n' +
        'T-1011\tقهوة تركية 250 جم\tعبوة\t3\t88\t0\t0';
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

    // جداول اللستة والمقارنة تُعرض عند الطلب (تخفيف الموقع)
    $('#btn-show-list').addEventListener('click', showListArea);
    $('#btn-show-cmp').addEventListener('click', showCmpArea);

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
  }

  /* ─────────────────── الإقلاع ───────────────────
     لا يُقلَع التطبيق فعلياً إلا بعد نجاح تسجيل الدخول:
     • عند حمل الصفحة بجلسة صالحة يُقلَع فوراً بعد فتح بوابة الدخول (auth.js أولاً).
     • عند تسجيل الدخول الآن تُطلق auth.js حدث 'bills:auth-ok' ثم يتقلع هنا.
     • الستات المبنية والمقارنة والبطاقات محكومة بـ CONFIG.LISTS بعد فلترة الدور. */
  let booted = false;
  function boot() {
    if (booted) return;
    booted = true;
    init();
  }

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

    loadDraftIfAny();
    renderQuickList();
    renderCloudLists();
    renderDatalist();
    renderInvoiceItems();
    renderHistory();
    bindEvents();
    switchTab('invoice');
    autoFetchCloudLists(); // جلب تلقائي للست الناقص/القديم (خلفية، لا يمنع العمل)
  }

  document.addEventListener('DOMContentLoaded', () => {
    // لا يُقلع التطبيق قبل جلسة صالحة — يُبقى مغلقاً خلف بوابة الدخول
    if (!global.Auth || global.Auth.isAuthenticated()) boot();
  });
  document.addEventListener('bills:auth-ok', boot);

  global.App = {
  switchTab, renderHistory, updateAllRowsAndSummary, reportToTSV, invoicesToTSV,
  printInvoiceReport, buildPrintDocumentHTML, buildPrintTotalsHTML, printRowHTML,
  boot,
};
})(window);