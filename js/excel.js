/* ═══════════════════════════════════════════════════════════════════════
   js/excel.js — التعامل مع ملفات Excel المحلية (Services)
   • تحميل بنية SheetJS (xlsx) من CDN حسب الطلب فقط
   • قراءة ملف xlsx/xls إلى صفوف خلايا + قائمة الأوراق
   • تصدير مصفوفة خلايا إلى ملف xlsx (بديل التنزيل بهيئة إكسل)
   ملاحظة: للاستخدام دون إنترنت أضِف vendor/xlsx.full.min.js وغيّر XLSX_CDN
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const XLSX_CDN = global.CONFIG.XLSX_CDN;

  /* --- تحميل XLSX بشكل كسول (يُستدعى مرة واحدة) --- */
  let loadingPromise = null;
  function load() {
    if (global.XLSX) return Promise.resolve(global.XLSX);
    if (loadingPromise) return loadingPromise;
    loadingPromise = new Promise((resolve, reject) => {
      const s = document.createElement('script');
      s.src = XLSX_CDN;
      s.async = true;
      s.onload = () => resolve(global.XLSX);
      s.onerror = () => { loadingPromise = null; reject(new Error('تعذّر تحميل مكتبة Excel من CDN — تحقق من الاتصال أو أضِف نسخة محلية')); };
      document.head.appendChild(s);
    });
    return loadingPromise;
  }

  /* --- تحويل خلية xlsx إلى نص للتحليل (الأرقام نَصيَّة) --- */
  function cellText(cell) {
    if (!cell) return '';
    if (cell.t === 'n') { const v = cell.v; return Number.isInteger(v) ? String(v) : String(Math.round(v * 1000000) / 1000000); }
    if (cell.t === 'd') return cell.v && cell.v.toISOString ? cell.v.toISOString().slice(0, 10) : String(cell.v || '');
    if (cell.t === 'b') return cell.v ? 'نعم' : '';
    if (cell.f && cell.w && cell.t === 's') return cell.w;
    return String(cell.v ?? cell.w ?? '');
  }

  function sheetToCells(sheet) {
    const rows = [];
    const range = sheet['!ref'];
    if (!range) return rows;
    const ref = XLSX.utils.decode_range(range);
    for (let R = ref.s.r; R <= ref.e.r; R++) {
      const row = [];
      for (let C = ref.s.c; C <= ref.e.c; C++) {
        const addr = XLSX.utils.encode_cell({ r: R, c: C });
        row.push(cellText(sheet[addr]));
      }
      rows.push(row);
    }
    return rows;
  }

  /**
   * قراءة ملف إكسل محلي.
   * @returns {Promise<Array<{name, rows}>>
   */
  async function readFile(file) {
    const XLSXLib = await load();
    const buf = await file.arrayBuffer();
    const wb = XLSXLib.read(buf, { type: 'array' });
    return wb.SheetNames.map((name) => ({ name, rows: sheetToCells(wb.Sheets[name]) }));
  }

  /** قراءة ورقة أولى فقط (معظم الاستخدامات) */
  async function readFirstSheet(file) {
    const sheets = await readFile(file);
    return sheets[0] || null;
  }

  /* --- تصدير مصفوفة خلايا إلى ملف xlsx حقيقي --- */
  function exportXLSX(sheetName, cellsRows, filename) {
    return load().then((x) => {
      const wb = x.utils.book_new();
      const aoa = cellsRows.map((r) => r.slice());
      const ws = x.utils.aoa_to_sheet(aoa);
      x.utils.book_append_sheet(wb, ws, (sheetName || 'ورقة1').slice(0, 31));
      const out = x.write(wb, { bookType: 'xlsx', type: 'array' });
      const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const url = URL.createObjectURL(blob);
      const a = document.createElement('a');
      a.href = url;
      a.download = filename || 'تقرير.xlsx';
      document.body.appendChild(a);
      a.click();
      a.remove();
      setTimeout(() => URL.revokeObjectURL(url), 2000);
    });
  }

  /* --- تحويل ورقة إلى TSV (لمربعات اللصق) --- */
  function sheetToTSV(sheetData) {
    return global.Parser.tableToTSV(sheetData.rows);
  }

  global.Excel = { load, readFile, readFirstSheet, exportXLSX, sheetToTSV, cellText };
})(window);