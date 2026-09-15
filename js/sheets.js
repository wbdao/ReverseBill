/* ═══════════════════════════════════════════════════════════════════════
   js/sheets.js — ربط سحابي مع Google Sheets (Services)
   • استخراج {id/pubId} و{gid} من أي رابط Spreadsheet (عادي أو منشور pubhtml)
   • enumerate تبويبات الشيت المنشور (listTabs) بقراءة قائمة الأوراق من pubhtml
   • جلب تبويب على حدة (fetchTab) عبر نقاط CSV/GViz المتوافقة مع CORS
   • جلب عدة تبويبات والتفكيك (fetchSheetsData) — جاهزة لمحرك Parser
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const CFG = global.CONFIG.SHEETS || {};

  /* --- تنظيف اسم تبويب من تشكيل/علامات الفهارس المنطقية (لسانيات صفحات النشر) --- */
  const normSheet = (s) => String(s || '')
    .replace(/[\u064B-\u065F\u0670\u200B-\u200F\u202A-\u202E\uFEFF]/g, '')
    .replace(/\s+/g, ' ').trim().toLowerCase();

  const decodeEscapes = (s) => String(s || '')
    .replace(/\\\//g, '/')
    .replace(/\\u([0-9a-fA-F]{4})/g, (_, h) => String.fromCharCode(parseInt(h, 16)));

  /* --- استخراج طبقة الجدول من أي رابط --- */
  function parseUrl(url) {
    if (!url) return null;
    const pubMatch = url.match(/\/spreadsheets\/d\/e\/([a-zA-Z0-9\-_]+)/);
    if (pubMatch) {
      const pubId = pubMatch[1];
      const gidMatch = url.match(/[#?&]gid=(-?\d+)/);
      const sheetMatch = url.match(/[?&]sheet=([^&#]+)/) || url.match(/[#?&]tab=([^&#]+)/);
      return {
        type: 'publish',
        pubId,
        id: pubId,
        gid: gidMatch ? gidMatch[1] : '0',
        sheetName: sheetMatch ? decodeURIComponent(sheetMatch[1].replace(/\+/g, ' ')) : '',
      };
    }
    const idMatch = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9\-_]+)/) || url.match(/[?&]id=([a-zA-Z0-9\-_]+)/);
    const id = idMatch ? idMatch[1] : (url.match(/^([a-zA-Z0-9\-_]{20,})$/) ? url : null);
    if (!id) return null;
    const gidMatch = url.match(/[#?&]gid=(-?\d+)/);
    const sheetMatch = url.match(/[?&]sheet=([^&#]+)/) || url.match(/[#?&]tab=([^&#]+)/);
    return {
      type: 'classic',
      id,
      gid: gidMatch ? gidMatch[1] : '0',
      sheetName: sheetMatch ? decodeURIComponent(sheetMatch[1].replace(/\+/g, ' ')) : '',
    };
  }

  /* --- الرابط الأساسي لنقاط التصدير حسب نوع الرابط --- */
  function baseUrl(p) {
    return p.type === 'publish'
      ? `https://docs.google.com/spreadsheets/d/e/${p.pubId}`
      : `https://docs.google.com/spreadsheets/d/${p.id}`;
  }

  /* --- بناء رابط CSV منشور من الرابط المدخل (pub أو export أو gviz) --- */
  function buildCsvUrl(url, { gid = '', sheetName = '', preferred = 'csv' } = {}) {
    const p = parseUrl(url);
    if (!p) return url;
    const g = gid || p.gid || '0';
    if (preferred === 'gviz') {
      const span = sheetName || g;
      return `${baseUrl(p)}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(span)}`;
    }
    return p.type === 'publish'
      ? `${baseUrl(p)}/pub?gid=${g}&single=true&output=csv`
      : `${baseUrl(p)}/export?format=csv&gid=${g}`;
  }

  /* --- خطأ يحمل حالة HTTP حقيقية (للتمييز بين رفض حقيقي وحظر CORS) --- */
  function makeHttpError(status, url) {
    const e = new Error(`فشل الجلب (HTTP ${status})`);
    e.status = status;
    e.url = url;
    return e;
  }

  /* --- كشف أجسام أخطاء بعض الوسطاء (JSON فيه "error" مع حالة 200) --- */
  const isLikelyErrorBody = (text) => {
    const s = String(text || '').trim();
    return s.startsWith('{') && /"error"\s*[:=]/i.test(s);
  };

  const PROXIES = Array.isArray(CFG.PROXIES) ? CFG.PROXIES.filter(Boolean) : [];

  /**
   * جلب نص مع سلسلة محاولات: اتصال مباشر أولاً ثم وساطات CORS تباعاً.
   * الأخطاء النصية (HTTP 404/403…) لا يُعاد تجربتها عبر الوسيط — الوسيط لن يصلحها؛
   * فقط أخطاء الشبكة/حظر CORS (TypeError) تدفع إلى الوسيط التالي في القائمة.
   * @param {string} url
   * @param {{useProxy?:boolean}} opts
   * @returns {Promise<{text:string,status:number,viaProxy:boolean,finalUrl:string}>}
   */
  async function fetchTextRes(url, { useProxy = true } = {}) {
    const attempts = [{ url, viaProxy: false }];
    if (useProxy) PROXIES.forEach((proxy) => attempts.push({ url: proxy + encodeURIComponent(url), viaProxy: true }));

    let lastErr = null;
    for (const it of attempts) {
      try {
        const res = await fetch(it.url, { method: 'GET', credentials: 'omit', redirect: 'follow' });
        if (!res.ok) throw makeHttpError(res.status, it.url);
        const text = await res.text();
        // مصدر CSV حقيقي لا يُعيد JSON خطأ — نرفضه مهما جاء (مباشرة أو عبر وسيط)
        if (isLikelyErrorBody(text)) throw new Error('المصدر رجع بخطأ — تحقق من الرابط الأصلي');
        return { text, status: res.status, viaProxy: it.viaProxy, finalUrl: it.url };
      } catch (err) {
        lastErr = err;
        // رفض HTTP حقيقي (404/мисс…) لا يصلح له وسيط — نُنهي مباشرة
        if (!it.viaProxy && err && typeof err.status === 'number') throw err;
      }
    }
    throw lastErr || new Error('تعذّر الاتصال بمصدر البيانات');
  }

  /* --- جلب نص متوافق مع الوضع السابق (يُرجع النص فقط) --- */
  async function fetchText(url, opts) {
    const r = await fetchTextRes(url, opts);
    return r.text;
  }

  /** إلحاق `output=csv` تلقائياً لرابط شيت منشور (pubhtml/pub) يفتقر إليها */
  function ensureOutputCsv(url) {
    const p = parseUrl(url);
    if (!p || p.type !== 'publish') return url || '';
    if (/(output=csv)/i.test(url)) return url;
    const g = p.gid || '0';
    return `${baseUrl(p)}/pub?gid=${g}&single=true&output=csv`;
  }

  /**
   * تنظيف رابط Google Sheets قبل أي طلب fetch (يمنع أخطاء 401/404 من روابط التعديل):
   * 1) أي رابط تعديل/عرض (/edit أو /view — كلاسيكي أو المنشور d/e/…) يتحول تلقائياً
   *    إلى صيغة النشر CSV `pub?output=csv` مع الحفاظ على gid لو كان مذكوراً (#gid=X أو ?gid=X).
   * 2) روابط النشر (pubhtml/pub) يضمن لها معامل output=csv إن لم يكن موجوداً.
   * 3) روابط export/gviz الجاهزة تُعاد كما هي.
   * @param {string} url
   * @returns {string} رابط منظف جاهز للجلب
   */
  function cleanSheetUrl(url) {
    const raw = String(url || '').trim();
    const p = parseUrl(raw);
    if (!p) return raw;
    const base = baseUrl(p);

    // 1) روابط التعديل/العرض → نشر CSV (المستخدم يلصق /edit?usp=sharing غالباً)
    if (/(^|[\/?#&])(edit|view)([\/?#]|$)/i.test(raw) || /\/edit\b|\/view\b/i.test(raw)) {
      const g = (p.gid && p.gid !== '0') ? p.gid : '';
      return g ? `${base}/pub?gid=${g}&single=true&output=csv` : `${base}/pub?output=csv`;
    }

    // 2) روابط النشر التي تفتقر إلى output=csv
    if (p.type === 'publish') return ensureOutputCsv(raw);

    // 3) روابط export/gviz جاهزة
    return raw;
  }

  /**
   * اختبار اتصال برابط الجدول وإرجاع نتيجة صريحة بدل صمت الفشل:
   * - وضع API بمفتاح: يفحص نطاقاً صغيراً (A1:Z5) ويتأكد من قبول المفتاح ووجود الورقة.
   * - وضع CSV منشور: يبني رابط CSV (output=csv) ويجلبه مباشرة أو عبر وسيط CORS.
   * @param {string} url رابط الجدول/المجموعة/الروابط
   * @param {{apiKey?:string,sheetName?:string}} opts
   * @returns {Promise<{ok:boolean,code:string,status?:number,message?:string,rows?:any[],count?:number,viaProxy?:boolean,source?:string}>}
   */
  async function testSheetConnection(url, { apiKey = '', sheetName = '' } = {}) {
    const p = parseUrl(url);
    if (!p) {
      return { ok: false, code: 'INVALID', message: 'رابط الجدول غير صالح — الصق رابط Google Sheets (مثل .../edit أو .../pubhtml) وسيُنظَّف تلقائياً لصيغة نشر CSV.' };
    }

    try {
      // وضع API: فحص المفتاح والنطاق
      if (apiKey) {
        const span = (sheetName || p.sheetName || '').trim();
        const range = span ? `${escapeSheetName(span)}!A1:Z5` : 'A1:Z5';
        const apiUrl = `${CFG.API_BASE}/${p.id}/values/${range}?valueRenderOption=UNFORMATTED_VALUE&key=${encodeURIComponent(apiKey)}`;
        const res = await fetch(apiUrl, { method: 'GET', credentials: 'omit' });
        if (!res.ok) return { ok: false, code: 'HTTP', status: res.status, message: `مفتاح API رفض الوصول (HTTP ${res.status})` };
        const json = await res.json();
        if (json.error) return { ok: false, code: 'API', status: json.error.code, message: json.error.message || 'خطأ من Google API' };
        const rows = Array.isArray(json.values) ? json.values : [];
        return { ok: true, source: 'api', rows, count: rows.length, status: res.status, viaProxy: false };
      }

      // وضع CSV منشور: نختبر الرابط المنظف حصراً (avoid 401 من روابط /edit أو /view)
      const target = cleanSheetUrl(url);
      const r = await fetchTextRes(target);
      const rows = global.Parser.parseCSV(r.text).map((x) => x.cells);
      if (!rows.length) {
        return { ok: false, code: 'EMPTY', status: r.status, viaProxy: r.viaProxy,
          message: 'الاستجابة سليمة لكن لا توجد بيانات CSV صالحة — تأكد من نشر الورقة بصيغة CSV.' };
      }
      return { ok: true, source: 'csv', rows, count: rows.length, status: r.status, viaProxy: r.viaProxy };
    } catch (err) {
      return { ok: false, code: 'FAIL', status: err && err.status, message: err && err.message ? err.message : 'خطأ غير متوقع أثناء الفحص.' };
    }
  }

  /**
   * سرد تبويبات شيت منشور (pubhtml) بقراءة `items.push(...)` من الصفحة.
   * يعمل مع روابط `d/e/{pubId}/pubhtml` وكذلك العادية عند النشر.
   * @returns {Promise<Array<{name:string, gid:string}>>}
   */
  function listTabs(url) {
    const p = parseUrl(url);
    if (!p) return Promise.resolve([]);
    const metaUrl = `${baseUrl(p)}/pubhtml`;
    return fetchText(metaUrl).then((html) => {
      const out = [];
      const re = /items\.push\(\{name:\s*"((?:\\.|[^"\\])*)",\s*pageUrl:\s*"((?:\\.|[^"\\])*)",\s*gid:\s*"(-?\d+)"/g;
      let m;
      while ((m = re.exec(html))) {
        const name = normSheet(decodeEscapes(m[1]));
        const gid = m[3];
        if (name && !out.some((t) => t.gid === gid)) out.push({ name, gid });
      }
      return out;
    }).catch(() => []);
  }

  function escapeSheetName(name) {
    return String(name || '').replace(/[!'"#]/g, (c) => "'" + c).replace(/'/g, "''");
  }

  /* --- جلب تبويب واحد (CSV منشور / GViz / API) --- */
  async function fetchTab(url, { name = '', gid = '', apiKey = '' } = {}) {
    url = cleanSheetUrl(url); // تنظيف (تحويل /edit و/view إلى نشر CSV) قبل أي طلب
    const p = parseUrl(url);
    if (!p) throw new Error('رابط جدول غير صالح');
    const base = baseUrl(p);
    const g = gid || p.gid || '0';

    // 1) وضع API بمفتاح (للأجهزة المقيدة) — يفسّر اسم الورقة مباشرة
    if (apiKey) {
      const span = (name || p.sheetName || '').trim();
      const range = span ? `${escapeSheetName(span)}!A1:ZZ100000` : 'A1:ZZ100000';
      const apiUrl = `${CFG.API_BASE}/${p.id}/values/${range}?valueRenderOption=UNFORMATTED_VALUE&majorDimension=ROWS&key=${encodeURIComponent(apiKey)}`;
      const res = await fetch(apiUrl, { credentials: 'omit' });
      if (!res.ok) throw new Error(`فشل اتصال API (${res.status}) — تحقق من المفتاح وصلاحية النشر`);
      const json = await res.json();
      if (json.error) throw new Error(json.error.message || 'خطأ في الاستجابة من API');
      const matrix = Array.isArray(json.values) ? json.values : [];
      return matrix.map((r) => r.map((v) => (typeof v === 'number' ? String(v) : String(v ?? ''))));
    }

    // 2) نقاط تصدير عامة (تسامح مع الشبكات المقيدة)
    const candidates = [
      base + (p.type === 'publish'
          ? `/pub?gid=${g}&single=true&output=csv`
          : `/export?format=csv&gid=${g}`),
      `${base}/gviz/tq?tqx=out:csv&gid=${g}`,
    ];
    if (name) candidates.push(`${base}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(name)}`);

    let lastErr = null;
    for (const u of candidates) {
      try {
        const text = await fetchText(u);
        return global.Parser.parseCSV(text).map((r) => r.cells);
      } catch (err) { lastErr = err; }
    }
    throw lastErr || new Error('تعذّر جلب محتوى الورقة');
  }

  /**
   * الجلب الرئيسي متعدد التبويبات (AgentDist و Company وغيرها).
   * @param {string} url رابط الشيت (منشور pubhtml يُفضَّل).
   * @param {Object} opts { tabs:Array<string> (أسماء التبويبات), apiKey:string, onTab:Function }
   * @returns {Promise<{ tabs:Array<{name,gid,rows,count}>, source:'csv'|'api' }>}
   */
  async function fetchSheetsData(url, { tabs = [], apiKey = '', onTab = null } = {}) {
    if (!url) throw new Error('أدخل رابط جدول Google Sheets أولاً');
    if (!parseUrl(url)) throw new Error('رابط جدول غير صالح');

    // نحدد "رابط CSV صريح" قبل التنظيف حتى لا ندمج العملية متعددة التبويبات
    const isExplicitPublish = /docs\.google\.com\/spreadsheets/.test(url)
      && (/(export|gviz\/tq|output=csv)/.test(url) || /\/pub\?[^#]*single=true/i.test(url));

    url = cleanSheetUrl(url); // تنظيف (تحويل /edit و/view إلى نشر CSV)
    const p = parseUrl(url);

    // 0) إذا كان الرابط نفسه نقطة تصدير/نشر CSV صريحة نجلبه مرة واحدة (مع ضمان output=csv)
    if (isExplicitPublish) {
      const text = await fetchText(url);
      return {
        tabs: [{ name: p.sheetName || 'الورقة', gid: p.gid || '0', rows: global.Parser.parseCSV(text).map((r) => r.cells) }],
        source: 'csv',
      };
    }

    // 1) كشف تبويبات الشيت المنشور (الأسماء + gid)
    let menu = [];
    try { menu = await listTabs(url); } catch (e) { menu = []; }

    let targets;
    if (tabs && tabs.length) {
      targets = tabs.map((t) => {
        const found = menu.find((m) => normSheet(m.name) === normSheet(t));
        return { name: t, gid: found ? found.gid : '' };
      });
    } else if (menu.length) {
      targets = menu.map((m) => ({ name: m.name, gid: m.gid }));
    } else {
      targets = [{ name: p.sheetName || 'الورقة الأولى', gid: p.gid || '0' }];
    }

    // 2) جلب كل تبويب تباعاً (لمراقبة التقدم وتفادي حدود التردد)
    const results = [];
    for (const t of targets) {
      const rows = await fetchTab(url, { name: t.name, gid: t.gid, apiKey });
      results.push({ name: t.name, gid: t.gid, rows, count: rows.length });
      if (onTab) onTab({ tab: t.name, count: rows.length, done: results.length, total: targets.length });
    }
    return { tabs: results, source: apiKey ? 'api' : 'csv' };
  }

  /* --- الجلب الأحادي (متوافق مع الوضع السابق) --- */
  async function fetchSheet(url, { apiKey = '', gid = '', sheetName = '' } = {}) {
    if (!url) throw new Error('أدخل رابط جدول Google Sheets أولاً');

    // نحدد "رابط CSV صريح" قبل التنظيف
    const isExplicit = /docs\.google\.com\/spreadsheets/.test(url)
      && (/(export|gviz|tq|output=csv)/.test(url) || /\/pub\?[^#]*single=true/i.test(url));

    url = cleanSheetUrl(url); // تنظيف (تحويل /edit و/view إلى نشر CSV)

    // إذا كان الرابط نفسه رابط CSV منشور جاهز نجلبه مباشرة (مع ضمان output=csv)
    if (isExplicit) {
      const text = await fetchText(url);
      return { rows: global.Parser.parseCSV(text).map((r) => r.cells), source: 'csv' };
    }

    const rows = await fetchTab(url, { name: sheetName, gid, apiKey });
    return { rows, source: apiKey ? 'api' : 'csv' };
  }

  global.Sheets = {
    parseUrl, buildCsvUrl, cleanSheetUrl, listTabs, fetchTab, fetchSheetsData, fetchSheet,
    fetchText, fetchTextRes, ensureOutputCsv, testSheetConnection, normSheet,
  };
})(window);