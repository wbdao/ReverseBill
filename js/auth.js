/* ═══════════════════════════════════════════════════════════════════════
   js/auth.js — بوابة تسجيل الدخول وصلاحيات الوصول (Security)
   • بوابة أمامية تُعرض قبل التطبيق وتُظهره فقط بعد تسجيل دخول ناجح
   • حسابان بصلاحيات وصول محددة للستات:
        BranchAccount → الستات: AgentDist · Company · online
        MainAccount   → جميع الستات
   • كلمات المرور مخزّنة بصيغة SHA-256 (مع ملح ثابت) وليست نصاً صريحاً
   • جلسة محفوظة محلياً (localStorage) مع إعادة بناء صلاحيات الست عند كل تحميل
   • عند نجاح تسجيل الدخول يُفلتَر CONFIG.LISTS حسب الدور قبل إقلاع (app.js)
     فتُبنى اللستات والبطاقات والمقارنة تلقائياً ضمن الصلاحية المخوّلة فقط
   ═══════════════════════════════════════════════════════════════════════ */
'use strict';

(function (global) {
  const AUTH_KEY = 'invoicereviewer_auth';
  const SALT = 'invoicereviewer::';

  /* --- الحسابات المخوّلة ---
     passHash = sha256(SALT + password) — محسوبة ولا تُخزَّن كلمة المرور نصاً. */
  const ACCOUNTS = {
    BranchAccount: {
      label: 'حساب الفرع',
      desc: 'الوصول للستات: الموزع (AgentDist) · الشركة (Company) · أونلاين (online)',
      access: ['AgentDist', 'Company', 'online'],
      passHash: '90043cc3f6e3f8432dada8d71bb93afc4cd8ec38878e4a30db8fbff152d07020',
    },
    MainAccount: {
      label: 'الحساب الرئيسي',
      desc: 'الوصول لجميع الستات (All)',
      access: 'all',
      passHash: '57e27ca8edb8e2c0efa290273563c08ea044a9d36f9884ca3121443f1d7fa34f',
    },
  };

  /* ---------- SHA-256 (تطبيق خفيف، يعمل على نص ASCII) ---------- */
  function sha256(ascii) {
    const rightRotate = (value, amount) => (value >>> amount) | (value << (32 - amount));
    const mathPow = Math.pow;
    const maxWord = mathPow(2, 32);
    const lengthProperty = 'length';
    let i, j;
    let result = '';
    const words = [];
    const asciiBitLength = ascii[lengthProperty] * 8;
    let hash = sha256.h = sha256.h || [];
    const k = sha256.k = sha256.k || [];
    let primeCounter = k[lengthProperty];
    const isComposite = {};
    for (let candidate = 2; primeCounter < 64; candidate++) {
      if (!isComposite[candidate]) {
        for (i = 0; i < 313; i += candidate) isComposite[i] = candidate;
        hash[primeCounter] = (mathPow(candidate, 0.5) * maxWord) | 0;
        k[primeCounter++] = (mathPow(candidate, 1 / 3) * maxWord) | 0;
      }
    }
    ascii += '\x80';
    while (ascii[lengthProperty] % 64 - 56) ascii += '\x00';
    for (i = 0; i < ascii[lengthProperty]; i++) {
      j = ascii.charCodeAt(i);
      if (j >> 8) return '';
      words[i >> 2] |= j << ((3 - i) % 4) * 8;
    }
    words[words[lengthProperty]] = ((asciiBitLength / maxWord) | 0);
    words[words[lengthProperty]] = asciiBitLength;
    for (j = 0; j < words[lengthProperty];) {
      const w = words.slice(j, j += 16);
      const oldHash = hash;
      hash = hash.slice(0, 8);
      for (i = 0; i < 64; i++) {
        const w15 = w[i - 15], w2 = w[i - 2];
        const a = hash[0], e = hash[4];
        const temp1 = hash[7]
          + (rightRotate(e, 6) ^ rightRotate(e, 11) ^ rightRotate(e, 25))
          + ((e & hash[5]) ^ ((~e) & hash[6]))
          + k[i]
          + (w[i] = (i < 16) ? w[i] : (
            w[i - 16]
            + (rightRotate(w15, 7) ^ rightRotate(w15, 18) ^ (w15 >>> 3))
            + w[i - 7]
            + (rightRotate(w2, 17) ^ rightRotate(w2, 19) ^ (w2 >>> 10))
          ) | 0);
        const temp2 = (rightRotate(a, 2) ^ rightRotate(a, 13) ^ rightRotate(a, 22))
          + ((a & hash[1]) ^ (a & hash[2]) ^ (hash[1] & hash[2]));
        hash = [(temp1 + temp2) | 0].concat(hash);
        hash[4] = (hash[4] + temp1) | 0;
      }
      for (i = 0; i < 8; i++) hash[i] = (hash[i] + oldHash[i]) | 0;
    }
    for (i = 0; i < 8; i++) {
      for (j = 3; j + 1; j--) {
        const b = (hash[i] >> (j * 8)) & 255;
        result += ((b < 16) ? 0 : '') + b.toString(16);
      }
    }
    return result;
  }

  const hashOf = (pw) => sha256(SALT + String(pw || ''));

  /* ---------- الجلسة ---------- */
  function getSession() {
    try {
      const raw = localStorage.getItem(AUTH_KEY);
      if (!raw) return null;
      const s = JSON.parse(raw);
      return (s && s.username && ACCOUNTS[s.username]) ? s : null;
    } catch (err) { return null; }
  }
  const isAuthenticated = () => !!getSession();
  const currentUser = () => {
    const s = getSession();
    if (!s) return null;
    const acc = ACCOUNTS[s.username];
    return { username: s.username, label: acc.label, desc: acc.desc, access: acc.access };
  };

  function setSession(username) {
    const s = { username, loginAt: new Date().toISOString() };
    try { localStorage.setItem(AUTH_KEY, JSON.stringify(s)); } catch (err) { /* تجاهل */ }
    return s;
  }
  function clearSession() {
    try { localStorage.removeItem(AUTH_KEY); } catch (err) { /* تجاهل */ }
  }

  /* ---------- تقييد الستات حسب الدور ----------
     يُفلتَر CONFIG.LISTS في مكانه (نفس الكائن المُشار إليه في app.js)
     قبل الإقلاع: فاللستات والبطاقات وقوائم الأسماء والمقارنة تُبنى
     تلقائياً ضمن الستات المخوّل بها الحساب فقط. */
  function applyListsAccess(username) {
    if (!global.CONFIG || !Array.isArray(global.CONFIG.LISTS)) return false;
    const acc = ACCOUNTS[username];
    if (!acc) return false;
    if (acc.access === 'all') return true;
    const allowed = new Set(acc.access);
    global.CONFIG.LISTS = [...global.CONFIG.LISTS].filter((c) => allowed.has(c.slug));
    return true;
  }

  /* ---------- منطق الواجهة ---------- */
  function el(id) { return document.getElementById(id); }

  function showLoginError(msg) {
    const text = el('login-error-text');
    if (text) text.textContent = msg;
    const box = el('login-error');
    if (box) box.classList.remove('hidden');
  }
  function clearLoginError() {
    const box = el('login-error');
    if (box) box.classList.add('hidden');
  }

  function renderSessionUI(username) {
    const user = currentUser();
    const chip = el('user-chip');
    if (chip && user) {
      chip.innerHTML = `<i class="fa-solid fa-user-shield"></i> <span>${user.label}</span><span class="text-[10px] opacity-80">(${user.username})</span>`;
      chip.classList.remove('hidden');
      if (!chip.classList.contains('inline-flex')) chip.classList.add('inline-flex');
    }
    const area = el('session-area');
    if (area) area.classList.remove('hidden');
    const logoutBtn = el('btn-logout');
    if (logoutBtn) logoutBtn.classList.remove('hidden');
    if (username) { const loginName = el('login-username'); if (loginName) loginName.value = username; }
  }

  function showShell() {
    const gate = el('login-gate');
    const shell = el('app-shell');
    if (gate) gate.classList.add('hidden');
    if (shell) shell.classList.remove('hidden');
  }
  function showGate() {
    const gate = el('login-gate');
    const shell = el('app-shell');
    if (gate) gate.classList.remove('hidden');
    if (shell) shell.classList.add('hidden');
  }

  function unlock(username) {
    applyListsAccess(username);
    showShell();
    renderSessionUI(username);
    clearLoginError();
  }

  /* ---------- تسجيل الدخول / الخروج ---------- */
  function login(username, password) {
    const u = String(username || '').trim();
    const acc = ACCOUNTS[u];
    if (!acc) return { ok: false, error: 'اسم المستخدم غير موجود' };
    if (hashOf(password) !== acc.passHash) return { ok: false, error: 'كلمة المرور غير صحيحة' };
    setSession(u);
    unlock(u);
    // إقلاع التطبيق فقط بعد كشف الواجهة (listener مُسجَّل في app.js)
    document.dispatchEvent(new CustomEvent('bills:auth-ok'));
    return { ok: true, user: { username: u, label: acc.label } };
  }

  function logout() {
    clearSession();
    window.location.reload();
  }

  function bindForm() {
    const form = el('login-form');
    if (!form) return;
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      const username = el('login-username').value.trim();
      const password = el('login-password').value;
      const result = login(username, password);
      if (!result.ok) showLoginError(result.error);
    });

    const toggle = el('btn-toggle-pass');
    if (toggle) {
      toggle.addEventListener('click', () => {
        const input = el('login-password');
        const showing = input.type === 'text';
        input.type = showing ? 'password' : 'text';
        toggle.innerHTML = showing
          ? '<i class="fa-solid fa-eye"></i>'
          : '<i class="fa-solid fa-eye-slash"></i>';
      });
    }

    const hint = el('btn-toggle-hint');
    if (hint) {
      hint.addEventListener('click', () => {
        const box = el('login-hint');
        if (box) box.classList.toggle('hidden');
      });
    }

    const logoutBtn = el('btn-logout');
    if (logoutBtn) logoutBtn.addEventListener('click', logout);
  }

  /* ---------- الإقلاع: التحقق من الجلسة قبل عرض التطبيق ---------- */
  document.addEventListener('DOMContentLoaded', () => {
    bindForm();
    const session = getSession();
    if (session) {
      unlock(session.username);
    } else {
      showGate();
    }
  });

  /* ---------- الواجهة العامة ---------- */
  global.Auth = {
    login, logout,
    getSession, isAuthenticated, currentUser,
    accounts: ACCOUNTS,
  };
})(window);