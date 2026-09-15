/* ═══════════════════════════════════════════════════════════════════════
   js/utils.js — أدوات عامة وتجريد واجهة المستخدم الصغيرة
   • تنسيق الأرقام والمال، تطبيع النصوص، التعامل مع الحافظة والتنزيلات
   • حوافز Toast + نافذة التأكيد + اختصارات DOM
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const CONFIG = global.CONFIG;

  /* --- اختصارات DOM --- */
  const $  = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  /* --- تنسيق الأرقام --- */
  const fmtNum = (n) => (!isFinite(n) ? '0.00' : n.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }));
  const money = (n) => `${fmtNum(n)} ${CONFIG.CURRENCY}`;
  const signedMoney = (n) => (n > 0 ? '+' : '') + money(n);
  const signedNum = (n) => (n > 0 ? '+' : '') + fmtNum(n);

  /* --- تحويل الأرقام العربية والهندية إلى لاتينية ثم رقم فعلي --- */
  const toWestern = (s) => String(s)
    .replace(/[٠-٩]/g, (d) => '٠١٢٣٤٥٦٧٨٩'.indexOf(d))
    .replace(/[۰-۹]/g, (d) => '۰۱۲۳۴۵۶۷۸۹'.indexOf(d));
  const parseNum = (s) => {
    const n = parseFloat(toWestern(s).replace(/[^\d.-]/g, ''));
    return isNaN(n) ? 0 : n;
  };
  const num = (v) => { const n = Number(v); return Number.isFinite(n) ? n : 0; };

  /* --- تطبيع النصوص --- */
  // للمطابقة الدقيقة: يزيل المسافات/الرموز ويحوّل للصغار
  const normKey = (s) => String(s || '').replace(/[\s_\-.:/\\]/g, '').toLowerCase();
  // لرمز الصنف (T-Code): يزيل الفواصل والمسافات الداخلية والأصفار البادئة —
  // يطابق T-1001 مع T 1001 أو t1001 أو T.1001 على كلا الطرفين (الفهرس والمطابقة نفس التطبيع)
  const normalizeNum = (s) => String(s || '').trim().toLowerCase().replace(/\s+/g, '').replace(/[_\-.:/\\]/g, '').replace(/^0+/, '');
  // توحيد الأحرف المتشابهة: الهمزات → ا، ؤ → و، ئ → ي، ة → ه، ى → ي + إزالة التشكيل ومسافات التنسيق
  const _normLetters = (s) => String(s || '')
    .replace(/[\u064B-\u065F\u0670\u200B-\u200F\u202A-\u202E\uFEFF]/g, '')
    .replace(/[أإآٱ]/g, 'ا')
    .replace(/ؤ/g, 'و')
    .replace(/ئ/g, 'ي')
    .replace(/ة/g, 'ه')
    .replace(/ى/g, 'ي');
  // للاسم والوحدة: ضغط المسافات المتتالية + توحيد الحروف (مطابقة مرنة وثابتة الطرفين)
  const normalizeName = (s) => _normLetters(s).replace(/\s+/g, ' ').trim().toLocaleLowerCase('ar-EG');
  const normalizeUnit = (s) => _normLetters(s).replace(/\s+/g, ' ').trim().toLocaleLowerCase('ar-EG');

  /* --- عام --- */
  const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const todayStr = () => new Date().toISOString().slice(0, 10);
  const genId = (prefix) => prefix + '_' + Date.now().toString(36) + Math.random().toString(36).slice(2, 8);

  /* --- الحافظة مع بديل للبروتوكول غير الآمن (file://) --- */
  function fallbackCopy(text) {
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.opacity = '0';
    document.body.appendChild(ta);
    ta.select();
    let ok = false;
    try { ok = document.execCommand('copy'); } catch (err) { /* تجاهل */ }
    ta.remove();
    return ok;
  }
  function copyText(text) {
    return new Promise((resolve) => {
      const done = (ok) => { toast(ok ? 'تم النسخ إلى الحافظة ✓' : 'تعذّر النسخ تلقائياً — استخدم Ctrl+C يدوياً', ok ? 'success' : 'error'); resolve(ok); };
      if (navigator.clipboard && window.isSecureContext) {
        navigator.clipboard.writeText(text).then(() => done(true)).catch(() => { const ok = fallbackCopy(text); done(ok); });
        return;
      }
      done(fallbackCopy(text));
    });
  }

  /* --- تنزيل نصوص / ملفات --- */
  function downloadText(text, filename, mime = 'text/plain') {
    const blob = new Blob([text], { type: mime });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = filename;
    document.body.appendChild(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }
  const downloadJSON = (obj, filename) => downloadText(JSON.stringify(obj, null, 2), filename, 'application/json');
  const downloadTSV = (text, filename) => downloadText(text, filename, 'text/tab-separated-values');

  /* --- قراءة ملف نصي محلي ثم تمريره لدالة --- */
  function importFile(file, callback, onError) {
    if (!file) return;
    const reader = new FileReader();
    reader.onerror = () => (onError || toast)('تعذّر قراءة الملف', 'error');
    reader.onload = () => {
      try { callback(reader.result); }
      catch (err) { (onError || toast)(err && err.message ? err.message : 'ملف غير صالح', 'error'); }
    };
    reader.readAsText(file);
  }

  /* --- إشعارات Toast --- */
  const toast = (message, type = 'success') => {
    const wrap = $('#toast-wrap');
    if (!wrap) { alert(message); return; }
    const el = document.createElement('div');
    el.className = `toast toast-${type}`;
    el.textContent = message;
    wrap.appendChild(el);
    setTimeout(() => el.remove(), 3400);
  };

  /* --- نافذة تأكيد نمطية --- */
  const confirmBox = ({ title, message, confirmLabel = 'تأكيد', danger = false, onConfirm }) => {
    const root = $('#modal-root');
    if (!root) { if (window.confirm(`\n${message}`) && onConfirm) onConfirm(); return; }
    const overlay = document.createElement('div');
    overlay.className = 'fixed inset-0 z-40 bg-slate-900/40 grid place-items-center p-4';
    overlay.innerHTML = `
      <div class="bg-white rounded-2xl shadow-xl p-5 w-full max-w-sm text-center">
        <div class="text-3xl mb-2">${danger ? '⚠️' : '❓'}</div>
        <h3 class="font-extrabold text-lg mb-1">${esc(title)}</h3>
        <p class="text-sm text-slate-500 mb-4">${esc(message)}</p>
        <div class="flex gap-2 justify-center">
          <button data-act="no" class="btn btn-ghost">إلغاء</button>
          <button data-act="yes" class="btn ${danger ? 'bg-rose-600 text-white hover:bg-rose-700' : 'btn-primary'}">${esc(confirmLabel)}</button>
        </div>
      </div>`;
    overlay.addEventListener('click', (e) => {
      const act = e.target.closest('button')?.dataset.act;
      if (act === 'no' || act === 'yes') {
        overlay.remove();
        if (act === 'yes' && onConfirm) onConfirm();
      }
    });
    root.appendChild(overlay);
  };

  /* --- مزامنة معالجات كثيفة بأجزاء (لتجنب تهنيج المتصفح مع آلاف الأصناف) --- */
  function batchProcess(items, fn, chunkSize = 2000, onProgress, onDone) {
    let i = 0;
    const step = () => {
      const end = Math.min(i + chunkSize, items.length);
      for (; i < end; i++) fn(items[i], i);
      if (onProgress) onProgress(i, items.length);
      if (i < items.length) setTimeout(step, 0);
      else if (onDone) onDone(i);
    };
    step();
  }

  global.Utils = {
    $, $$, fmtNum, money, signedMoney, signedNum, toWestern, parseNum, num,
    normKey, normalizeNum, normalizeName, normalizeUnit,
    esc, todayStr, genId, copyText, downloadText, downloadJSON, downloadTSV,
    importFile, toast, confirmBox, batchProcess,
  };
})(window);