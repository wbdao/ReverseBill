/* ═══════════════════════════════════════════════════════════════════════
   js/config.js — إعدادات التطبيق المركزية
   • مفاتيح التخزين، بيانات الرؤوس، نماذج اللستات التجريبية، روابط الخدمات.
   • يُحمَّل أولاً ويُكشَف للباقي عبر window.CONFIG
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

// المستند المنشور الوحيد الذي تستضيف اللستات الخمس جميعاً (تبويب لكل لستة)
const SHEET_PUBLISHED_BASE = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vSLFu_tyMm4wGKbX60I0wqxed4OMHQDIIcd1omvVGCJyn1fqWn67lAVZeAeGIG1Dq_1L6gUfiFqLbIT/pub?output=csv';

window.CONFIG = {
  /* --- عام --- */
  APP_NAME: 'مراجع الفواتير',
  CURRENCY: 'ج.م',
  LIST_PAGE_SIZE: 200,            // عدد الصفوف الظاهرة بصفحة جدول اللستة (تجنب تهنيج البيانات الضخمة)

  /* --- مفاتيح LocalStorage (كل كيان بنطاق مستقل) --- */
  STORAGE_KEYS: {
    PRICE_LISTS: 'invoicereviewer_pricelists',   // قوائم الأسعار + معرف اللستة النشطة
    LEGACY_PRODUCTS: 'invoicereviewer_products', // (ترحيل) النسخة القديمة أحادية اللستة
    INVOICES: 'invoicereviewer_invoices',
    SETTINGS: 'invoicereviewer_settings',
    DRAFT: 'invoicereviewer_draft',
  },

  /* --- كشف رؤوس الأعمدة تلقائياً (لملفات السيستم / Excel / Sheets) ---
     كل هينت يُقارن بعد تطبيع المسافات والرموز وعدم حساسية الحالة. */
  HEADER_HINTS: {
    docNo: ['invoice no', 'invoiceno', 'bill no', 'billno', 'رقم الفاتورة', 'الفاتورة', 'فاتورة', 'doc no'],
    date: ['date', 'التاريخ', 'تاريخ'],
    tax: ['vat', 'tax', 'الضريبة', 'ضريبة', 'discount', 'الخصم', 'خصم'],
    itemNo: ['item #', 'item no', 'itemnumber', 'item number', 'item', 'id code', 'code', 'sku', 'رقم الصنف', 'رقم البند', 'كود الصنف', 'كود', 'رمز', 'رقم', '#'],
    name: ['product name', 'productname', 'product', 'item description', 'item name', 'description', 'اسم المنتج', 'اسم الصنف', 'المنتج', 'الصنف', 'الوصف', 'البيان', 'البند', 'اسم'],
    unit: ['unit', 'uom', 'الوحدة', 'وحدة'],
    qty: ['quantity', 'qty', 'الكمية', 'كمية', 'العدد', 'عدد'],
    unitPrice: ['unit price', 'unitprice', 'unit cost', 'سعر الوحدة', 'سعر الشراء', 'السعر', 'سعر'],
    discountValue: ['discount value', 'discountvalue', 'disc value', 'discount amount', 'discountamount', 'خصم القيمة', 'قيمة الخصم', 'قيمه الخصم', 'الخصم القيمي'],
    discountPct: ['discount rate', 'discountrate', 'discount %', 'discount pct', 'discountpercent', 'discount percentage', 'خصم النسبة', 'نسبة الخصم', 'نسبه الخصم', 'الخصم النسبي'],
    amount: ['net amount', 'netamount', 'amount', 'total', 'net', 'الإجمالي', 'المبلغ', 'القيمة', 'إجمالي', 'قيمة'],
    price: ['price', 'السعر', 'سعر', 'سعر البيع'],
  },

  /* --- قوائم الأسعار التجريبية (تُستدعى عند أول تشغيل فقط أو بالزر) ---
     كل صف: [رقم الصنف, الاسم, الوحدة, السعر] */
  SEED_LISTS: [
    {
      name: 'لستة التجاري',
      rows: [
        ['T-1001', 'سكر 1 كجم', 'كيس', 30],
        ['T-1002', 'أرز مصري 1 كجم', 'كيس', 40],
        ['T-1003', 'زيت عباد الشمس 1 لتر', 'عبوة', 70],
        ['T-1004', 'دقيق فاخر 1 كجم', 'كيس', 22],
        ['T-1005', 'شاي 500 جم', 'علبة', 100],
        ['T-1006', 'لبن 1 لتر', 'عبوة', 26],
        ['T-1007', 'مكرونة 400 جم', 'كيس', 12],
        ['T-1008', 'طماطم معلبة 400 جم', 'علبة', 18],
        ['T-1009', 'صابون سائل غسيل 1 لتر', 'عبوة', 45],
        ['T-1010', 'عيش فينو', 'قطعة', 2],
      ],
    },
    {
      name: 'لستة الشركات',
      rows: [
        ['C-2001', 'سكر 1 كجم', 'كيس', 28.5],
        ['C-2002', 'زيت عباد الشمس 1 لتر', 'عبوة', 64],
        ['C-2003', 'شاي 500 جم', 'علبة', 95],
        ['C-2004', 'لبن 1 لتر', 'عبوة', 24.5],
        ['C-2005', 'دقيق فاخر 1 كجم', 'كيس', 21],
        ['C-2006', 'مكرونة 400 جم', 'كيس', 11],
      ],
    },
    {
      name: 'لستة الريتيل',
      rows: [
        ['R-3001', 'لبن 1 لتر', 'عبوة', 25],
        ['R-3002', 'سكر 1 كجم', 'كيس', 31],
        ['R-3003', 'زيت عباد الشمس 1 لتر', 'عبوة', 72],
        ['R-3004', 'شاي 500 جم', 'علبة', 102],
        ['R-3005', 'عيش فينو', 'قطعة', 2.5],
      ],
    },
  ],

  /* --- اللستات السحابية الثابتة (نفس المستند المنشور، تبويب مستقل لكل لستة) ---
     • url: الرابط الأساسي للمستند — كل الستات في نفس الملف، والتبويب يُحدَّد عبر معامل gid.
     • gid: معرّف التبويب الدقيق (&gid=XXXXX). '0' = التبويب الأول. القيمة الفارغة "" تعني
       اكتشافاً تلقائياً من قائمة التبويبات المنشورة حسب الاسم tab عند أول جلب/تحديث،
       ويُخزَّن مؤقتاً على cfg.gid خلال الجلسة. لتثبيت يدوي: افتح التبويب وانسخ الرقم من
       شريط العنوان (...#gid=XXXXX) ودعه هنا لإلغاء الكشف التلقائي.
     • tab: اسم التبويب داخل المستند — دليل الكشف الآلي وخيار النسخ الاحتياطي بالاسم.
     • url فارغ "" = رسالة توجيه داخل البطاقة. color/icon لألوان البطاقة. */
  LISTS: [
    { slug: 'AgentDist', name: 'قائمة الموزع (AgentDist)', gid: '0',          tab: 'AgentDist', color: 'indigo', icon: 'fa-truck-fast',     url: SHEET_PUBLISHED_BASE },
    { slug: 'Company',   name: 'قائمة الشركة (Company)',   gid: '',           tab: 'Company',   color: 'emerald', icon: 'fa-building',       url: SHEET_PUBLISHED_BASE },
    { slug: 'online',    name: 'قائمة أونلاين (online)',    gid: '',           tab: 'online',    color: 'sky',     icon: 'fa-cart-shopping', url: SHEET_PUBLISHED_BASE },
    { slug: 'Retail',    name: 'قائمة الريتيل (Retail)',    gid: '',           tab: 'Retail',    color: 'amber',   icon: 'fa-store',         url: SHEET_PUBLISHED_BASE },
    { slug: 'Shaheen',   name: 'قائمة شاهين (Shaheen)',     gid: '',           tab: 'Shaheen',   color: 'rose',    icon: 'fa-bolt',         url: SHEET_PUBLISHED_BASE },
  ],
  LISTS_REFRESH_HOURS: 6, // إعادة جلب تلقائية عند الإقلاع إذا مضى هذا القدر من الساعات على آخر مزامنة
  // إصدار بنية بيانات اللستات المخزنة — عند تغييره تُعاد مزامنة كل اللستات مرة واحدة
  // (الإقلاع للتغلب على بيانات قديمة اندمجت فيها وحدات مختلفة في سطر واحد قبل إصلاح المطابقة بالوحدة)
  LIST_DATA_VERSION: 2,

  /* --- خدمات ربط Google Sheets --- */
  SHEETS: {
    // رابط CSV المنشور المباشر بعد استخراج {id} و{gid} (يعمل بدون API key ومعظم البيئات)
    CSV_EXPORT: 'https://docs.google.com/spreadsheets/d/{id}/export?format=csv&gid={gid}',
    // نقطة GViz = بديل أكثر توافقاً مع CORS في بعض الشبكات
    GVIZ_EXPORT: 'https://docs.google.com/spreadsheets/d/{id}/gviz/tq?tqx=out:csv&sheet={sheet}',
    // واجهة Google Sheets v4 — تُستخدم فقط عند تزويد مفتاح API (للأجهزة المقيدة)
    API_BASE: 'https://sheets.googleapis.com/v4/spreadsheets',
    API_KEY: '', // املأه من شاشة "البيانات" وستُحفظ في الإعدادات
    // وساطات CORS مجانية — تُجرَّب تباعاً عند فشل الاتصال المباشر (ملف محلي file:// أو شبكات مقيدة)
    PROXIES: [
      'https://api.allorigins.win/raw?url=',
      'https://corsproxy.io/?url=',
    ],
  },

  /* --- مكتبة SheetJS (للتعامل مع ملفات Excel محلياً) --- */
  XLSX_CDN: 'https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js',
};