/* ITCB yearly budget – client */
(() => {
  const VERSION = "1.3.1";
  const ROLE_LABEL = { admin: "מנהל", editor: "עורך", viewer: "צופה" };
  const MONTHS = ["", "ינואר", "פברואר", "מרץ", "אפריל", "מאי", "יוני", "יולי", "אוגוסט", "ספטמבר", "אוקטובר", "נובמבר", "דצמבר"];
  const MSHORT = ["", "ינו׳", "פבר׳", "מרץ", "אפר׳", "מאי", "יוני", "יולי", "אוג׳", "ספט׳", "אוק׳", "נוב׳", "דצמ׳"];
  const DONE = "בוצע", FC = "תחזית", NOTHAPPEN = "לא התקיים";

  // column definitions – order and Hebrew headers follow the original Excel file
  const INC = [
    { k: "category", l: "קטגוריה", t: "cat", req: 1 },
    { k: "status", l: "תחזית/בוצע", t: "cat", opts: [FC, DONE], req: 1 },
    { k: "body", l: "גוף", t: "cat" },
    { k: "level", l: "רמת הסמכה/סיווג", t: "cat" },
    { k: "receipt", l: "מספר קבלה", t: "text" },
    { k: "month", l: "חודש", t: "month" },
    { k: "amount", l: 'סכום בש"ח', t: "money" },
    { k: "detail", l: "פירוט", t: "long" },
    { k: "examinees", l: "כמות נבחנים שנגשו", t: "num" },
    { k: "examDate", l: "תאריך בחינה", t: "date", hint: "התאריך שבו התקיימה הבחינה בפועל" },
    { k: "training", l: "הדרכה", t: "cat" },
    { k: "price", l: "מחיר", t: "cat" },
    { k: "sector", l: "מגזר", t: "cat" },
    { k: "client", l: "לקוח", t: "cat" },
    { k: "date", l: "תאריך", t: "date" },
  ];
  const EXP = [
    { k: "serial", l: "מס׳ סידורי", t: "num" },
    { k: "year", l: "שנה", t: "num" },
    { k: "ref", l: "חשבונית/קבלה/אסמכתא", t: "text" },
    { k: "amount", l: 'סכום בש"ח', t: "money" },
    { k: "detail", l: "פירוט", t: "long" },
    { k: "category", l: "קטגוריה", t: "cat", req: 1 },
    { k: "classification", l: "סיווג", t: "cat" },
    { k: "month", l: "חודש", t: "month" },
    { k: "status", l: "תחזית/בוצע/לא התקיים", t: "cat", opts: [FC, DONE, NOTHAPPEN], req: 1 },
    { k: "priority", l: "עדיפות", t: "cat" },
    { k: "manager", l: "מנהל תקציב", t: "cat" },
    { k: "impl", l: "יישום", t: "cat", hint: "1 - אושר לביצוע, 2 - אושר עקרונית תלוי באישור תכנית, 3 - בקשה חדשה שטרם אושרה" },
    { k: "emergency", l: "עדכון חירום", t: "cat" },
    { k: "notes", l: "הערות", t: "long" },
  ];
  const COLS = { income: INC, expenses: EXP };
  const KIND_LABEL = { income: "הכנסה", expenses: "הוצאה" };

  const VIEWS = [
    ["income", "הכנסות"],
    ["incSum", "ריכוז הכנסות"],
    ["incAct", "ריכוז הכנסות בפועל"],
    ["expSum", "ריכוז הוצאות"],
    ["expAct", "ריכוז הוצאות בפועל"],
    ["expenses", "הוצאות"],
    ["exam", "כמות נבחנים"],
    ["pl", "רווח והפסד"],
    ["fc", "הכנסות – תחזית מול ביצוע"],
    ["srcAct", "מקור הכנסה – בפועל"],
    ["srcAll", "מקור הכנסה – תחזית"],
    ["invoices", "חשבוניות"],
  ];
  const EXP_FILTER_KEYS = [["status", "תחזית/בוצע/לא התקיים"], ["priority", "עדיפות"], ["impl", "יישום"], ["emergency", "עדכון חירום"]];

  const $app = document.getElementById("app");
  const store = {
    get: (k) => { try { return localStorage.getItem(k); } catch { return null; } },
    set: (k, v) => { try { localStorage.setItem(k, v); } catch {} },
    del: (k) => { try { localStorage.removeItem(k); } catch {} },
  };
  const S = {
    token: store.get("itcb_budget_token"), user: null, years: [], year: null, view: "income", docs: {},
    filters: { income: {}, expenses: {} }, q: { income: "", expenses: "" }, sort: { income: null, expenses: null },
    collapsed: {}, pv: { incStatus: "", examStatus: "", fcCat: "", srcActCat: "", srcAllCat: "" }, users: [],
  };

  // ---------- utils ----------
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const num = (v) => {
    if (typeof v === "number") return Number.isFinite(v) ? v : 0;
    const t = String(v ?? "").replace(/[,\s₪]/g, "");
    return /^-?\d+(\.\d+)?$/.test(t) ? parseFloat(t) : 0;
  };
  const isNum = (v) => v !== null && v !== undefined && v !== "" && /^-?\d+(\.\d+)?$/.test(String(v).replace(/[,\s₪]/g, ""));
  const fmtN = (n) => Math.round(n).toLocaleString("he-IL");
  const shekel = (n) => `<bdi dir="ltr">${n < 0 ? "−" : ""}₪${Math.abs(Math.round(n)).toLocaleString("he-IL")}</bdi>`;
  const sum = (rows, k = "amount") => rows.reduce((a, r) => a + num(r[k]), 0);
  const str = (v) => (v === null || v === undefined ? "" : String(v).trim());
  const heSort = (a, b) => {
    if (isNum(a) && isNum(b)) return num(a) - num(b);
    return String(a).localeCompare(String(b), "he");
  };
  function fmtDate(iso) {
    if (!iso) return "";
    const d = new Date(iso);
    if (isNaN(d)) return esc(iso);
    return d.toLocaleDateString("he-IL") + (String(iso).length > 10 ? " " + d.toLocaleTimeString("he-IL", { hour: "2-digit", minute: "2-digit" }) : "");
  }
  function toast(msg, bad) {
    const t = document.getElementById("toast");
    t.textContent = msg; t.className = "show" + (bad ? " bad" : "");
    clearTimeout(toast._t); toast._t = setTimeout(() => (t.className = ""), 2800);
  }
  const canEdit = () => S.user && (S.user.role === "admin" || S.user.role === "editor");
  const isAdmin = () => S.user && S.user.role === "admin";
  const doc = () => S.docs[S.year];
  const statusBadge = (s) => {
    if (!s) return "";
    const c = s === DONE ? "b-done" : s === NOTHAPPEN ? "b-no" : s === FC ? "b-fc" : "b-fc";
    return `<span class="badge ${c}">${esc(s)}</span>`;
  };
  const EYE = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M1 12s4-8 11-8 11 8 11 8-4 8-11 8-11-8-11-8z"/><circle cx="12" cy="12" r="3"/></svg>`;
  const EYE_OFF = `<svg width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><path d="M17.94 17.94A10.07 10.07 0 0 1 12 20c-7 0-11-8-11-8a18.45 18.45 0 0 1 5.06-5.94"/><path d="M9.9 4.24A9.12 9.12 0 0 1 12 4c7 0 11 8 11 8a18.5 18.5 0 0 1-2.16 3.19"/><path d="M14.12 14.12a3 3 0 1 1-4.24-4.24"/><line x1="1" y1="1" x2="23" y2="23"/></svg>`;
  function pwField(id, label, ac, extra = "") {
    return `<div class="field"><label for="${id}">${label}</label><div class="pw-wrap">
      <input class="input" id="${id}" type="password" autocomplete="${ac}" dir="ltr" ${extra}>
      <button type="button" class="pw-eye" data-eye="${id}" aria-label="הצג סיסמה" title="הצג סיסמה">${EYE}</button></div></div>`;
  }
  function bindEyes(root = document) {
    root.querySelectorAll("[data-eye]").forEach((b) => {
      b.onclick = () => {
        const inp = document.getElementById(b.dataset.eye);
        const show = inp.type === "password";
        inp.type = show ? "text" : "password";
        b.innerHTML = show ? EYE_OFF : EYE;
        b.setAttribute("aria-label", show ? "הסתר סיסמה" : "הצג סיסמה");
        b.title = show ? "הסתר סיסמה" : "הצג סיסמה";
      };
    });
  }
  const genPw = () => {
    const a = "abcdefghjkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789";
    const r = crypto.getRandomValues(new Uint32Array(10));
    return Array.from(r, (x) => a[x % a.length]).join("");
  };

  async function api(path, opts = {}) {
    const res = await fetch("/api/" + path, {
      method: opts.method || "GET",
      headers: { "content-type": "application/json", ...(S.token ? { authorization: "Bearer " + S.token } : {}) },
      body: opts.body ? JSON.stringify(opts.body) : undefined,
    });
    const data = await res.json().catch(() => ({}));
    if (res.status === 401 && !["login", "password"].includes(path)) { logout(); throw new Error(data.error || "נדרשת התחברות"); }
    if (!res.ok) throw new Error(data.error || "שגיאה");
    return data;
  }

  // ---------- auth screens ----------
  function loginShell(inner) {
    $app.innerHTML = `<div class="login-wrap"><form class="login-card" id="authForm">
      <img src="/logo.png" alt="ITCB">${inner}</form></div>`;
    bindEyes();
    return document.getElementById("authForm");
  }
  function renderLogin() {
    const f = loginShell(`
      <h1>תוכנית תקציב שנתית</h1>
      <p class="sub">כניסה למערכת ניהול התקציב של ITCB</p>
      <div class="field"><label for="u">מייל</label><input class="input" id="u" type="email" autocomplete="username" dir="ltr" required></div>
      ${pwField("p", "סיסמה", "current-password", "required")}
      <div class="err" id="err"></div>
      <button class="btn primary block" type="submit">כניסה</button>`);
    f.onsubmit = async (e) => {
      e.preventDefault();
      const btn = f.querySelector("button[type=submit]"); btn.disabled = true;
      try {
        const pw = document.getElementById("p").value;
        const r = await api("login", { method: "POST", body: { email: document.getElementById("u").value, password: pw } });
        S.token = r.token; S.user = r.user; store.set("itcb_budget_token", r.token);
        if (r.user.mustChange) return renderFirstPassword(pw);
        await boot();
      } catch (er) { document.getElementById("err").textContent = er.message; btn.disabled = false; }
    };
  }
  function renderSetup() {
    const f = loginShell(`
      <h1>הגדרה ראשונית</h1>
      <p class="sub">יצירת חשבון המנהל הראשון של המערכת</p>
      <div class="field"><label for="n">שם</label><input class="input" id="n" required></div>
      <div class="field"><label for="u">מייל</label><input class="input" id="u" type="email" dir="ltr" autocomplete="username" required></div>
      ${pwField("p", "סיסמה (8 תווים לפחות)", "new-password", "required minlength=8")}
      ${pwField("p2", "אימות סיסמה", "new-password", "required minlength=8")}
      <p class="sub" style="margin:-4px 0 10px;text-align:right">טיפ: לחיצה על סמל העין מציגה את הסיסמה. ודא שהמקלדת באנגלית.</p>
      <div class="err" id="err"></div>
      <button class="btn primary block" type="submit">יצירת מנהל</button>`);
    f.onsubmit = async (e) => {
      e.preventDefault();
      if (p.value !== p2.value) return (document.getElementById("err").textContent = "הסיסמאות אינן תואמות");
      try {
        await api("setup", { method: "POST", body: { name: n.value, email: u.value, password: p.value } });
        toast("החשבון נוצר – אפשר להתחבר"); renderLogin();
      } catch (er) { document.getElementById("err").textContent = er.message; }
    };
  }
  function renderFirstPassword(temp) {
    const f = loginShell(`
      <h1>בחירת סיסמה אישית</h1>
      <p class="sub">זו הכניסה הראשונה שלך – יש לבחור סיסמה חדשה</p>
      ${pwField("p1", "סיסמה חדשה (8 תווים לפחות)", "new-password", "required minlength=8")}
      ${pwField("p2", "אימות סיסמה", "new-password", "required minlength=8")}
      <div class="err" id="err"></div>
      <button class="btn primary block" type="submit">שמירה והמשך</button>`);
    f.onsubmit = async (e) => {
      e.preventDefault();
      const err = document.getElementById("err");
      if (p1.value !== p2.value) return (err.textContent = "הסיסמאות אינן תואמות");
      try {
        const r = await api("password", { method: "POST", body: { current: temp, password: p1.value } });
        S.token = r.token; S.user = r.user; store.set("itcb_budget_token", r.token);
        toast("הסיסמה עודכנה"); await boot();
      } catch (er) { err.textContent = er.message; }
    };
  }
  function logout() { S.token = null; S.user = null; S.docs = {}; store.del("itcb_budget_token"); renderLogin(); }

  // ---------- routing ----------
  function setHash() {
    const h = S.view === "users" ? "#/users" : `#/${S.year || ""}/${S.view}`;
    if (location.hash !== h) history.replaceState(null, "", h);
  }
  function readHash() {
    const [, a, b] = location.hash.split("/");
    if (a === "users") return { view: "users" };
    return { year: a, view: b };
  }

  async function boot() {
    $app.innerHTML = `<div class="loading">טוען…</div>`;
    try {
      if (!S.user) S.user = (await api("me")).user;
      if (!S.user.mustChange) S.config = (await api("config").catch(() => ({ config: {} }))).config || {};
      if (S.user.mustChange) return renderFirstPassword("");
      S.years = (await api("years")).years;
      const h = readHash();
      if (h.view === "users" && isAdmin()) S.view = "users";
      else {
        S.year = S.years.includes(h.year) ? h.year : (S.years.includes(String(new Date().getFullYear())) ? String(new Date().getFullYear()) : S.years[S.years.length - 1] || null);
        S.view = VIEWS.some((v) => v[0] === h.view) ? h.view : "income";
      }
      if (S.year) await loadYear(S.year);
      render();
    } catch (e) { if (S.token) { $app.innerHTML = `<div class="loading">${esc(e.message)}</div>`; } }
  }
  async function loadYear(y, force) {
    if (!S.docs[y] || force) S.docs[y] = (await api("years/" + y)).doc;
    return S.docs[y];
  }

  // ---------- shell ----------
  function render() {
    setHash();
    const initials = (S.user.name || S.user.email).trim().slice(0, 1).toUpperCase();
    const yearTabs = S.years.map((y) => `<button data-year="${y}" class="${S.view !== "users" && S.year === y ? "active" : ""}">${y}</button>`).join("");
    $app.innerHTML = `
      <header class="top">
        <div class="top-inner">
          <button class="burger" id="burger" aria-label="תפריט" aria-expanded="false"><span></span><span></span><span></span></button>
          <img src="/logo.png" alt="ITCB">
          <div class="top-title"><b>תוכנית תקציב שנתית</b><span>עמותת הבודקים – ITCB</span><small class="m-ctx">${S.view === "users" ? "ניהול משתמשים" : S.year ? `${S.year} · ${esc((VIEWS.find((v) => v[0] === S.view) || ["", ""])[1])}` : ""}</small></div>
          <div class="spacer"></div>
          <div class="user-chip" id="userChip"><span class="uname">${esc(S.user.name)}</span><span class="role-tag">${ROLE_LABEL[S.user.role]}</span><span class="avatar">${esc(initials)}</span>
            <div class="menu hidden" id="userMenu">
              <button data-act="pw">שינוי סיסמה</button>
              ${isAdmin() ? `<button data-act="users">ניהול משתמשים</button>` : ""}
              <hr><button data-act="logout">יציאה</button>
            </div>
          </div>
        </div>
        <nav class="tabs" aria-label="שנים">${yearTabs}
          ${canEdit() ? `<button class="add-year" data-act="newYear">+ שנה חדשה</button>` : ""}
          ${isAdmin() ? `<span class="sep"></span><button data-act="users" class="${S.view === "users" ? "active" : ""}">ניהול משתמשים</button>` : ""}
        </nav>
        ${S.view !== "users" && S.year ? `<nav class="subtabs" aria-label="לשוניות">${VIEWS.map(([k, l], i) => `<button data-view="${k}" class="${S.view === k ? "active" : ""}"><span class="num">${i + 1}</span>${l}</button>`).join("")}</nav>` : ""}
      </header>
      <main id="main"></main>
      <footer class="ver">ITCB · ניהול תקציב · גרסה ${VERSION}</footer>
      <div class="side-ov hidden" id="sideOv"></div>
      <aside class="side-nav" id="sideNav" aria-label="תפריט ראשי" aria-hidden="true">
        <div class="side-head"><img src="/logo.png" alt="ITCB"><button class="btn ghost" id="sideClose" aria-label="סגירת התפריט">✕</button></div>
        <div class="side-user"><span class="avatar">${esc(initials)}</span><div><b>${esc(S.user.name)}</b><div class="muted" style="font-size:13px">${ROLE_LABEL[S.user.role]} · <span dir="ltr">${esc(S.user.email)}</span></div></div></div>
        ${S.years.length ? `<div class="side-group">שנת תקציב</div><div class="side-years">${S.years.map((y) => `<button data-year="${y}" class="${S.view !== "users" && S.year === y ? "active" : ""}">${y}</button>`).join("")}${canEdit() ? `<button class="add-year" data-act="newYear">+ חדשה</button>` : ""}</div>` : canEdit() ? `<button data-act="newYear"><span class="si">+</span>שנה חדשה</button>` : ""}
        ${S.year ? `<div class="side-group">לשוניות ${S.year}</div>${VIEWS.map(([k, l], i) => `<button data-view="${k}" class="${S.view === k ? "active" : ""}"><span class="si num-badge">${i + 1}</span>${l}</button>`).join("")}` : ""}
        <div class="side-group">חשבון</div>
        ${isAdmin() ? `<button data-act="users" class="${S.view === "users" ? "active" : ""}"><span class="si">👥</span>ניהול משתמשים</button>` : ""}
        <button data-act="pw"><span class="si">🔑</span>שינוי סיסמה</button>
        <button data-act="logout"><span class="si">↩</span>יציאה</button>
        <div class="side-ver">גרסה ${VERSION}</div>
      </aside>`;
    $app.querySelectorAll("[data-year]").forEach((b) => (b.onclick = async () => { S.year = b.dataset.year; if (S.view === "users") S.view = "income"; await loadYear(S.year); render(); }));
    $app.querySelectorAll("[data-view]").forEach((b) => (b.onclick = () => { S.view = b.dataset.view; render(); window.scrollTo(0, 0); }));
    const chip = document.getElementById("userChip"), menu = document.getElementById("userMenu");
    chip.onclick = (e) => { if (!e.target.closest(".menu")) menu.classList.toggle("hidden"); };
    document.onclick = (e) => { if (!e.target.closest("#userChip")) menu.classList.add("hidden"); };
    // mobile side menu
    const nav = document.getElementById("sideNav"), ov = document.getElementById("sideOv"), burger = document.getElementById("burger");
    const setNav = (open) => {
      nav.classList.toggle("open", open); ov.classList.toggle("hidden", !open);
      nav.setAttribute("aria-hidden", String(!open)); burger.setAttribute("aria-expanded", String(open));
      document.body.classList.toggle("no-scroll", open);
    };
    burger.onclick = () => setNav(true);
    ov.onclick = () => setNav(false);
    document.getElementById("sideClose").onclick = () => setNav(false);
    document.body.classList.remove("no-scroll");
    $app.querySelectorAll("[data-act]").forEach((b) => (b.onclick = (e) => {
      e.stopPropagation(); menu.classList.add("hidden"); setNav(false);
      const a = b.dataset.act;
      if (a === "logout") logout();
      if (a === "pw") openChangePassword();
      if (a === "users") { S.view = "users"; render(); }
      if (a === "newYear") openNewYear();
    }));
    renderMain();
  }

  function renderMain() {
    const $m = document.getElementById("main");
    if (S.view === "users") return renderUsers($m);
    if (!S.year) {
      $m.innerHTML = `<div class="panel"><h2>עדיין אין שנים במערכת</h2><p class="hint">אפשר ליצור שנה חדשה ריקה או לייבא קובץ אקסל של תקציב קיים.</p>
        ${canEdit() ? `<button class="btn primary" id="ny">+ שנה חדשה</button> <button class="btn" id="imp">ייבוא קובץ</button>` : ""}</div>`;
      if (canEdit()) { ny.onclick = openNewYear; imp.onclick = () => openImport(); }
      return;
    }
    const fn = { income: () => renderList($m, "income"), expenses: () => renderList($m, "expenses"), incSum: () => renderIncPivot($m, false), incAct: () => renderIncPivot($m, true),
      expSum: () => renderExpPivot($m, false), expAct: () => renderExpPivot($m, true), exam: () => renderExam($m), pl: () => renderPL($m), fc: () => renderFc($m),
      srcAct: () => renderSrc($m, true), srcAll: () => renderSrc($m, false), invoices: () => renderInvoices($m) }[S.view];
    fn();
  }

  // page header with year actions (import/export)
  function pageHead(title, sub, extra = "") {
    const d = doc();
    return `<div class="page-head"><div><h1>${title} · ${S.year}</h1><p>${sub || ""}</p></div><div class="spacer"></div>${extra}
      <div class="user-chip" style="padding:0;border:0;background:none" id="yearMenuWrap">
        <button class="btn" id="yearMenuBtn">ייבוא / ייצוא ▾</button>
        <div class="menu hidden" id="yearMenu" style="left:0;right:auto">
          <button data-y="xlsx">ייצוא שנת ${S.year} לאקסל</button>
          <button data-y="json">גיבוי שנת ${S.year} (JSON)</button>
          <button data-y="jsonAll">גיבוי כל השנים (JSON)</button>
          ${canEdit() ? `<hr><button data-y="import">ייבוא קובץ (אקסל / גיבוי)…</button>` : ""}
          ${isAdmin() ? `<hr><button data-y="del" style="color:var(--danger)">מחיקת שנת ${S.year}…</button>` : ""}
        </div></div></div>
      <div class="audit" style="margin:-8px 0 12px">עודכן לאחרונה ${fmtDate(d.updatedAt)}${d.updatedBy ? " · " + esc(d.updatedBy) : ""}</div>`;
  }
  function bindPageHead() {
    const btn = document.getElementById("yearMenuBtn"), m = document.getElementById("yearMenu");
    if (!btn) return;
    btn.onclick = (e) => { e.stopPropagation(); m.classList.toggle("hidden"); };
    document.addEventListener("click", () => m.classList.add("hidden"), { once: true });
    m.querySelectorAll("[data-y]").forEach((b) => (b.onclick = (e) => {
      e.stopPropagation(); m.classList.add("hidden");
      const a = b.dataset.y;
      if (a === "xlsx") exportXlsx(doc());
      if (a === "json") downloadJSON({ app: "itcb-budget", version: VERSION, exported: new Date().toISOString(), years: { [S.year]: stripDoc(doc()) } }, `ITCB_Budget_${S.year}.json`);
      if (a === "jsonAll") exportAllJSON();
      if (a === "import") openImport();
      if (a === "del") confirmDeleteYear();
    }));
  }

  // ================= LIST VIEWS (income / expenses) =================
  function distinct(rows, k) {
    const set = new Set(rows.map((r) => str(r[k])));
    return [...set].sort(heSort);
  }
  function filteredRows(kind) {
    const cols = COLS[kind], f = S.filters[kind], q = S.q[kind].trim().toLowerCase();
    let rows = doc()[kind].filter((r) => {
      for (const c of cols) {
        const fv = f[c.k];
        if (fv === undefined || fv === "") continue;
        const v = str(r[c.k]);
        if (["cat", "month", "num"].includes(c.t)) { if (v !== fv) return false; }
        else if (c.t === "money") {
          const m = /^\s*([<>]=?|=)?\s*(-?[\d,.]+)\s*$/.exec(fv);
          if (m) {
            const n = num(r[c.k]), x = num(m[2]), op = m[1] || "=";
            if (!({ "<": n < x, "<=": n <= x, ">": n > x, ">=": n >= x, "=": n === x }[op])) return false;
          } else if (!v.includes(fv)) return false;
        } else if (c.t === "date") { if (!fmtDate(v).includes(fv) && !v.includes(fv)) return false; }
        else if (!v.toLowerCase().includes(fv.toLowerCase())) return false;
      }
      if (q && !cols.some((c) => str(r[c.k]).toLowerCase().includes(q))) return false;
      return true;
    });
    const s = S.sort[kind];
    if (s) {
      const c = cols.find((x) => x.k === s.k);
      rows = rows.slice().sort((a, b) => {
        const va = a[s.k], vb = b[s.k];
        let d;
        if (["money", "num", "month"].includes(c.t) || (isNum(va) && isNum(vb))) {
          const ea = !isNum(va), eb = !isNum(vb);
          d = ea && eb ? 0 : ea ? 1 : eb ? -1 : num(va) - num(vb);
          if (ea !== eb) return d; // blanks last always
        } else d = str(va).localeCompare(str(vb), "he");
        return s.dir * d;
      });
    }
    return rows;
  }
  function cellHtml(c, v) {
    if (v === null || v === undefined || v === "") return "";
    if (c.t === "money") return shekel(num(v));
    if (c.k === "status") return statusBadge(String(v));
    if (c.t === "month") return isNum(v) ? `${v} <span class="muted">${MSHORT[num(v)] || ""}</span>` : esc(v);
    if (c.t === "date") return fmtDate(v);
    if (c.t === "long") return `<div class="clamp" title="${esc(v)}">${esc(v)}</div>`;
    if (c.k === "priority" && isNum(v)) return `<span class="prio ${num(v) === 1 ? "p1" : num(v) === 2 ? "p2" : ""}">${esc(v)}</span>`;
    return esc(v);
  }

  function renderList($m, kind) {
    const cols = COLS[kind], d = doc(), all = d[kind], f = S.filters[kind];
    const isInc = kind === "income";
    const activeF = Object.entries(f).filter(([, v]) => v !== "" && v !== undefined);
    $m.innerHTML = `
      ${pageHead(isInc ? "הכנסות" : "הוצאות", isInc ? "רשימת כל ההכנסות – תחזית וביצוע" : "רשימת כל ההוצאות – תחזית, ביצוע ופריטים שלא התקיימו",
        canEdit() ? `<button class="btn accent" id="addRow">+ ${KIND_LABEL[kind]} חדשה</button>` : "")}
      <div class="kpis k4" id="kpis"></div>
      <div class="toolbar">
        <input class="input search" id="q" type="search" placeholder="חיפוש חופשי בכל העמודות…" value="${esc(S.q[kind])}">
        ${activeF.length ? `<button class="btn sm" id="clearF">ניקוי ${activeF.length} סינונים</button>` : ""}
        ${S.sort[kind] ? `<button class="btn sm" id="clearS">ביטול מיון</button>` : ""}
        <span class="meta-line" id="resLine"></span>
      </div>
      <div class="dt-wrap"><table class="dt" id="dt">
        <thead><tr>${cols.map((c) => {
          const s = S.sort[kind];
          const arr = s && s.k === c.k ? `<span class="arr">${s.dir > 0 ? "▲" : "▼"}</span>` : "";
          return `<th class="sortable" data-sort="${c.k}" title="${esc(c.hint || "לחיצה למיון")}">${esc(c.l)}${arr}</th>`;
        }).join("")}<th class="col-actions"></th></tr>
        <tr class="frow">${cols.map((c) => {
          const v = f[c.k] ?? "";
          if (["cat", "month", "num"].includes(c.t)) {
            const vals = distinct(all, c.k);
            return `<th><select class="input ${v !== "" ? "on" : ""}" data-f="${c.k}"><option value="">הכל</option>${vals.map((x) => `<option value="${esc(x)}" ${x === v ? "selected" : ""}>${x === "" ? "(ריק)" : esc(x)}</option>`).join("")}</select></th>`;
          }
          return `<th><input class="input ${v !== "" ? "on" : ""}" data-f="${c.k}" value="${esc(v)}" placeholder="${c.t === "money" ? "למשל >1000" : "סינון…"}"></th>`;
        }).join("")}<th class="col-actions"></th></tr></thead>
        <tbody id="tb"></tbody><tfoot id="tf"></tfoot></table></div>`;
    bindPageHead();
    const draw = () => {
      const rows = filteredRows(kind);
      const tb = document.getElementById("tb");
      tb.innerHTML = rows.length ? rows.map((r) => `<tr data-id="${r.id}">${cols.map((c) => `<td class="${c.t === "money" ? "money" : c.t === "num" ? "numc" : c.t === "long" ? "long" : ""}">${cellHtml(c, r[c.k])}</td>`).join("")}
        <td class="col-actions"><div class="row-actions">${canEdit() ? `<button class="icon-btn" data-edit="${r.id}" title="עריכה / העברה לסעיף אחר">עריכה</button><button class="icon-btn del" data-del="${r.id}" title="מחיקה">מחיקה</button>` : `<button class="icon-btn" data-edit="${r.id}">צפייה</button>`}</div></td></tr>`).join("")
        : `<tr><td colspan="${cols.length + 1}"><div class="empty">לא נמצאו שורות התואמות לסינון</div></td></tr>`;
      document.getElementById("tf").innerHTML = `<tr>${cols.map((c, i) => `<td class="${c.t === "money" || c.t === "num" ? "money" : ""}">${i === 0 ? `סה״כ (${rows.length})` : c.t === "money" ? shekel(sum(rows, c.k)) : c.k === "examinees" ? fmtN(sum(rows, c.k)) : ""}</td>`).join("")}<td class="col-actions"></td></tr>`;
      document.getElementById("resLine").textContent = rows.length === all.length ? `${all.length} שורות` : `מוצגות ${rows.length} מתוך ${all.length} שורות`;
      const done = rows.filter((r) => r.status === DONE), fc = rows.filter((r) => r.status === FC);
      document.getElementById("kpis").innerHTML = isInc ? `
        <div class="kpi c-inc"><div class="lbl">סה״כ הכנסות (תחזית + בוצע)</div><div class="val">${shekel(sum(rows))}</div><div class="sub">${rows.length} שורות</div></div>
        <div class="kpi c-done"><div class="lbl">בוצע בפועל</div><div class="val">${shekel(sum(done))}</div><div class="sub">${done.length} שורות · ${sum(rows) ? Math.round((sum(done) / sum(rows)) * 100) : 0}% מהסך</div></div>
        <div class="kpi c-fc"><div class="lbl">תחזית (טרם בוצע)</div><div class="val">${shekel(sum(fc))}</div><div class="sub">${fc.length} שורות</div></div>
        <div class="kpi c-n"><div class="lbl">כמות נבחנים</div><div class="val">${fmtN(sum(rows, "examinees"))}</div><div class="sub">בוצע: ${fmtN(sum(done, "examinees"))}</div></div>` : `
        <div class="kpi c-exp"><div class="lbl">סה״כ הוצאות (ללא "לא התקיים")</div><div class="val">${shekel(sum(rows.filter((r) => r.status !== NOTHAPPEN)))}</div><div class="sub">${rows.length} שורות ברשימה</div></div>
        <div class="kpi c-done"><div class="lbl">בוצע בפועל</div><div class="val">${shekel(sum(done))}</div><div class="sub">${done.length} שורות</div></div>
        <div class="kpi c-fc"><div class="lbl">תחזית (טרם בוצע)</div><div class="val">${shekel(sum(fc))}</div><div class="sub">${fc.length} שורות</div></div>
        <div class="kpi c-n"><div class="lbl">לא התקיים</div><div class="val">${shekel(sum(rows.filter((r) => r.status === NOTHAPPEN)))}</div><div class="sub">${rows.filter((r) => r.status === NOTHAPPEN).length} שורות</div></div>`;
      tb.querySelectorAll("[data-edit]").forEach((b) => (b.onclick = (e) => { e.stopPropagation(); openRow(kind, all.find((r) => r.id === b.dataset.edit)); }));
      tb.querySelectorAll("[data-del]").forEach((b) => (b.onclick = (e) => { e.stopPropagation(); confirmDelete(kind, all.find((r) => r.id === b.dataset.del)); }));
      tb.querySelectorAll("tr[data-id]").forEach((tr) => (tr.ondblclick = () => openRow(kind, all.find((r) => r.id === tr.dataset.id))));
    };
    draw();
    const q = document.getElementById("q");
    q.oninput = () => { S.q[kind] = q.value; draw(); };
    $m.querySelectorAll("[data-f]").forEach((el) => {
      const ev = el.tagName === "SELECT" ? "onchange" : "oninput";
      el[ev] = () => {
        S.filters[kind][el.dataset.f] = el.value;
        el.classList.toggle("on", el.value !== "");
        if (el.tagName === "SELECT") renderList($m, kind); else draw();
      };
    });
    $m.querySelectorAll("[data-sort]").forEach((th) => (th.onclick = () => {
      const k = th.dataset.sort, s = S.sort[kind];
      S.sort[kind] = !s || s.k !== k ? { k, dir: 1 } : s.dir === 1 ? { k, dir: -1 } : null;
      renderList($m, kind);
    }));
    const cf = document.getElementById("clearF"); if (cf) cf.onclick = () => { S.filters[kind] = {}; renderList($m, kind); };
    const cs = document.getElementById("clearS"); if (cs) cs.onclick = () => { S.sort[kind] = null; renderList($m, kind); };
    const add = document.getElementById("addRow"); if (add) add.onclick = () => openRow(kind, null);
  }

  // ---------- row drawer (add / edit / move) ----------
  function openRow(kind, row) {
    const cols = COLS[kind], d = doc(), all = d[kind];
    const ro = !canEdit();
    const isNew = !row;
    const base = row ? { ...row } : kind === "expenses"
      ? { serial: Math.max(0, ...all.map((r) => num(r.serial))) + 1, year: num(S.year), status: FC, month: new Date().getMonth() + 1 }
      : { status: FC, month: new Date().getMonth() + 1 };
    const field = (c) => {
      const v = base[c.k] ?? "";
      const id = "f_" + c.k;
      const lbl = `<label for="${id}">${esc(c.l)}${c.req ? " *" : ""}</label>`;
      if (ro) return `<div class="field">${lbl}<div class="ro-val">${cellHtml(c, v) || "—"}</div></div>`;
      let inp;
      if (c.opts) inp = `<select class="input" id="${id}">${[...new Set([...c.opts, ...distinct(all, c.k).filter(Boolean)])].map((o) => `<option ${o === String(v) ? "selected" : ""}>${esc(o)}</option>`).join("")}</select>`;
      else if (c.t === "month") inp = `<select class="input" id="${id}"><option value="">—</option>${MONTHS.slice(1).map((mn, i) => `<option value="${i + 1}" ${num(v) === i + 1 ? "selected" : ""}>${i + 1} · ${mn}</option>`).join("")}</select>`;
      else if (c.t === "money" || c.t === "num") inp = `<input class="input" id="${id}" type="number" step="any" dir="ltr" value="${esc(v)}">`;
      else if (c.t === "date") inp = `<input class="input" id="${id}" type="date" value="${esc(String(v).slice(0, 10))}">`;
      else if (c.t === "long") inp = `<textarea class="input" id="${id}" rows="3">${esc(v)}</textarea>`;
      else inp = `<input class="input" id="${id}" value="${esc(v)}" ${c.t === "cat" ? `list="dl_${c.k}" autocomplete="off"` : ""}>
        ${c.t === "cat" ? `<datalist id="dl_${c.k}">${distinct(all, c.k).filter(Boolean).map((o) => `<option value="${esc(o)}">`).join("")}</datalist>` : ""}`;
      return `<div class="field" ${c.t === "long" ? 'style="grid-column:1/-1"' : ""}>${lbl}${inp}${c.hint ? `<div class="audit">${esc(c.hint)}</div>` : ""}</div>`;
    };
    const ov = document.createElement("div");
    ov.className = "overlay";
    ov.innerHTML = `<div class="drawer" role="dialog" aria-label="${isNew ? "הוספה" : "עריכה"}">
      <div class="drawer-head"><h3>${ro ? "צפייה ב" : isNew ? "הוספת " : "עריכת "}${KIND_LABEL[kind]}</h3><button class="btn ghost" data-x>✕</button></div>
      <div class="drawer-body">
        ${!isNew && !ro ? `<div class="note">כדי להעביר את השורה לסעיף אחר – שנה את <b>הקטגוריה</b>${kind === "expenses" ? " ו/או את <b>הסיווג</b>" : " ו/או את <b>הלקוח</b>"} ושמור.</div>` : ""}
        <div class="grid2">${cols.map(field).join("")}</div>
        ${kind === "expenses" && row ? linkedInvoicesHtml(row.id) : ""}
        ${kind === "expenses" && !ro ? `<div class="attach-box">
          <input type="file" id="f_attach" accept="application/pdf,image/*" hidden>
          <button type="button" class="btn sm" id="attBtn">📎 ${row && invList().some((i) => i.expenseId === row.id) ? "צירוף קובץ נוסף" : "צירוף חשבונית / קבלה"} (PDF או תמונה)</button>
          <span id="attName" class="muted"></span>
          <div class="audit" id="attNote">הקובץ יישמר בספריית החשבוניות ויקושר להוצאה. שדות ריקים יושלמו אוטומטית מהמסמך.</div>
        </div>` : ""}
        ${row && row.updatedAt ? `<div class="audit">עודכן ${fmtDate(row.updatedAt)}${row.updatedBy ? " · " + esc(row.updatedBy) : ""}</div>` : ""}
        <div class="err" id="rowErr"></div>
      </div>
      <div class="drawer-foot">${ro ? `<button class="btn" data-x>סגירה</button>` : `
        <button class="btn primary" id="saveRow">${isNew ? "הוספה" : "שמירה"}</button>
        ${!isNew ? `<button class="btn" id="dupRow">שכפול</button>` : ""}
        <button class="btn" data-x>ביטול</button><div class="spacer"></div>
        ${!isNew ? `<button class="btn danger" id="delRow">מחיקה</button>` : ""}`}
      </div></div>`;
    document.body.appendChild(ov); document.body.classList.add("no-scroll");
    const close = () => { ov.remove(); document.body.classList.remove("no-scroll"); };
    ov.onclick = (e) => { if (e.target === ov) close(); };
    ov.querySelectorAll("[data-x]").forEach((b) => (b.onclick = close));
    const onKey = (e) => { if (e.key === "Escape") { close(); document.removeEventListener("keydown", onKey); } };
    document.addEventListener("keydown", onKey);
    if (ro) return;
    const dateEl = ov.querySelector("#f_date"), monthEl = ov.querySelector("#f_month");
    if (dateEl && monthEl) dateEl.onchange = () => { if (dateEl.value) monthEl.value = String(new Date(dateEl.value).getMonth() + 1); };
    const collect = () => {
      const out = {};
      for (const c of cols) {
        let v = ov.querySelector("#f_" + c.k).value.trim();
        if (v === "") out[c.k] = null;
        else if (["money", "num", "month"].includes(c.t) || (c.t === "cat" && ["priority", "impl"].includes(c.k))) out[c.k] = isNum(v) ? num(v) : v;
        else out[c.k] = v;
      }
      return out;
    };
    // attachment (expenses): file is uploaded after the row is saved, then linked to it
    let attach = null, attachParsed = {};
    const attBtn = ov.querySelector("#attBtn");
    if (attBtn) {
      const fin = ov.querySelector("#f_attach");
      attBtn.onclick = () => fin.click();
      fin.onchange = async () => {
        const f0 = fin.files[0]; fin.value = "";
        if (!f0) return;
        const prep = await prepareInvoiceFile(f0);
        if (prep.error) return (ov.querySelector("#attNote").textContent = prep.error);
        attach = prep;
        ov.querySelector("#attName").textContent = f0.name;
        const note = ov.querySelector("#attNote");
        note.textContent = "קורא את המסמך…";
        try {
          const { text } = await extractText(prep.file, (st) => (note.textContent = st));
          attachParsed = parseInvoice(text, f0.name);
          const filled = [];
          const setIf = (id, val, label) => { const el = ov.querySelector(id); if (el && !el.value && val !== null && val !== undefined && val !== "") { el.value = val; filled.push(label); } };
          setIf("#f_amount", attachParsed.amount, "סכום");
          setIf("#f_ref", attachParsed.invoiceNo, "אסמכתא");
          if (attachParsed.date) setIf("#f_month", String(+attachParsed.date.slice(5, 7)), "חודש");
          setIf("#f_detail", attachParsed.supplier, "פירוט");
          note.innerHTML = filled.length ? `הקובץ יצורף בשמירה. הושלמו מהמסמך: <b>${filled.join(", ")}</b> – נא לבדוק.` : "הקובץ יצורף בשמירה.";
        } catch { note.textContent = "הקובץ יצורף בשמירה (לא הצלחתי לקרוא ממנו פרטים)."; }
      };
    }
    const save = async (asNew) => {
      const data = collect();
      const miss = cols.filter((c) => c.req && (data[c.k] === null || data[c.k] === ""));
      if (miss.length) return (ov.querySelector("#rowErr").textContent = "יש למלא: " + miss.map((c) => c.l).join(", "));
      try {
        const r = asNew || isNew
          ? await api(`years/${S.year}/${kind}`, { method: "POST", body: { row: data } })
          : await api(`years/${S.year}/${kind}/${row.id}`, { method: "PUT", body: { row: data } });
        S.docs[S.year] = r.doc;
        if (attach && kind === "expenses") {
          const rowId = asNew || isNew ? r.doc.expenses[r.doc.expenses.length - 1].id : row.id;
          try {
            const r2 = await api(`years/${S.year}/invoices`, { method: "POST", body: {
              file: await blobToB64(attach.file), fileName: attach.origName, mime: attach.file.type,
              meta: { supplier: attachParsed.supplier || data.detail || null, invoiceNo: data.ref ? String(data.ref) : null, date: attachParsed.date || null, amount: isNum(data.amount) ? num(data.amount) : null },
              apply: { expenseId: rowId, linkOnly: true } } });
            S.docs[S.year] = r2.doc;
          } catch (e2) { close(); renderMain(); return toast("ההוצאה נשמרה, אבל העלאת הקובץ נכשלה: " + e2.message, true); }
        }
        close(); renderMain(); toast((asNew ? "השורה שוכפלה" : isNew ? "השורה נוספה" : "השינויים נשמרו") + (attach ? " והקובץ צורף" : ""));
      } catch (er) { ov.querySelector("#rowErr").textContent = er.message; }
    };
    ov.querySelector("#saveRow").onclick = () => save(false);
    const dup = ov.querySelector("#dupRow"); if (dup) dup.onclick = () => save(true);
    const del = ov.querySelector("#delRow"); if (del) del.onclick = () => { close(); confirmDelete(kind, row); };
    setTimeout(() => ov.querySelector(".drawer-body input, .drawer-body select")?.focus(), 50);
  }

  function modal(title, bodyHtml, footHtml) {
    const ov = document.createElement("div");
    ov.className = "overlay center";
    ov.innerHTML = `<div class="modal" role="dialog" aria-label="${esc(title)}"><div class="drawer-head"><h3>${title}</h3><button class="btn ghost" data-x>✕</button></div>
      <div class="drawer-body">${bodyHtml}</div><div class="drawer-foot">${footHtml}</div></div>`;
    document.body.appendChild(ov); document.body.classList.add("no-scroll");
    const close = () => { ov.remove(); document.body.classList.remove("no-scroll"); };
    ov.onclick = (e) => { if (e.target === ov) close(); };
    ov.querySelectorAll("[data-x]").forEach((b) => (b.onclick = close));
    bindEyes(ov);
    return { ov, close, $: (s) => ov.querySelector(s) };
  }
  function confirmDelete(kind, row) {
    if (!row) return;
    const m = modal("מחיקת שורה", `<p>למחוק את ה${KIND_LABEL[kind]} הבאה?</p><div class="preview-box"><b>${esc(row.detail || row.category || "")}</b><br>${shekel(num(row.amount))} · ${esc(row.category || "")} · חודש ${esc(row.month ?? "")}</div><p class="muted">לא ניתן לבטל את הפעולה.</p><div class="err" id="e"></div>`,
      `<button class="btn primary" style="background:var(--danger);border-color:var(--danger)" id="ok">מחיקה</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelector("[data-x]:last-child").onclick = m.close;
    m.$("#ok").onclick = async () => {
      try { const r = await api(`years/${S.year}/${kind}/${row.id}`, { method: "DELETE" }); S.docs[S.year] = r.doc; m.close(); renderMain(); toast("השורה נמחקה"); }
      catch (er) { m.$("#e").textContent = er.message; }
    };
  }

  // ================= PIVOT ENGINE =================
  // keys: [{k,l}] levels; cols: [{l, fn(rows)->num|null}]
  function buildTree(rows, keys) {
    const root = { rows, children: new Map() };
    for (const r of rows) {
      let node = root;
      keys.forEach((key, i) => {
        const v = str(r[key.k]) || "(ריק)";
        if (!node.children.has(v)) node.children.set(v, { label: v, rows: [], children: new Map(), level: i });
        node = node.children.get(v);
        node.rows.push(r);
      });
    }
    return root;
  }
  const monthCols = (valKey) => {
    const cols = [];
    for (let m = 1; m <= 12; m++) cols.push({ l: `${m}<br><span class="muted" style="font-weight:500">${MSHORT[m]}</span>`, x: String(m), fn: (rows) => { const rr = rows.filter((r) => num(r.month) === m); return rr.length ? sum(rr, valKey) : null; } });
    cols.push({ l: "סה״כ", x: "סה\"כ", tot: 1, fn: (rows) => sum(rows, valKey) });
    return cols;
  };
  function pivotHtml(id, rows, keys, cols, opts = {}) {
    const tree = buildTree(rows, keys);
    const col = (S.collapsed[id] = S.collapsed[id] || new Set());
    const fmt = opts.fmt || ((v) => (v === null ? "" : `<span class="${v === 0 ? "z" : ""}">${fmtN(v)}</span>`));
    let html = "";
    const walk = (node, path) => {
      const kids = [...node.children.values()].sort((a, b) => heSort(a.label, b.label));
      for (const n of kids) {
        const p = path + "\u0001" + n.label;
        const hasKids = n.children.size > 0;
        const closed = col.has(p);
        html += `<tr class="l${n.level}"><td class="lbl">${hasKids ? `<button class="tg" data-tg="${esc(p)}">${closed ? "◂" : "▾"}</button>` : `<span class="tg"></span>`}${esc(n.label)}${hasKids ? ` <span class="muted" style="font-weight:500">(${n.rows.length})</span>` : ""}</td>
          ${cols.map((c) => `<td class="${c.tot ? "tot" : ""}">${fmt(c.fn(n.rows), c)}</td>`).join("")}</tr>`;
        if (hasKids && !closed) walk(n, p);
      }
    };
    walk(tree, "");
    if (!rows.length) html = `<tr><td colspan="${cols.length + 1}"><div class="empty">אין נתונים להצגה</div></td></tr>`;
    return `<div class="pv-wrap"><table class="pv" data-pv="${id}"><thead><tr><th class="lbl">${keys.map((k) => esc(k.l)).join(" ← ")}</th>${cols.map((c) => `<th>${c.l}</th>`).join("")}</tr></thead>
      <tbody>${html}</tbody><tfoot><tr class="gt"><td class="lbl">סה״כ כללי</td>${cols.map((c) => `<td>${fmt(c.fn(rows), c)}</td>`).join("")}</tr></tfoot></table></div>`;
  }
  function pivotTools(id, keys) {
    return `<div class="seg-ctl" data-lvl="${id}">${keys.map((k, i) => `<button data-l="${i}">עד ${esc(k.l)}</button>`).join("")}</div>`;
  }
  function bindPivot($m, id, rows, keys, rerender) {
    $m.querySelectorAll(`[data-pv="${id}"] [data-tg]`).forEach((b) => (b.onclick = () => {
      const s = S.collapsed[id]; s.has(b.dataset.tg) ? s.delete(b.dataset.tg) : s.add(b.dataset.tg); rerender();
    }));
    $m.querySelectorAll(`[data-lvl="${id}"] [data-l]`).forEach((b) => (b.onclick = () => {
      const lvl = +b.dataset.l, s = new Set();
      const walk = (node, path, depth) => { for (const n of node.children.values()) { const p = path + "\u0001" + n.label; if (depth >= lvl && n.children.size) s.add(p); walk(n, p, depth + 1); } };
      walk(buildTree(rows, keys), "", 0);
      S.collapsed[id] = s; rerender();
    }));
  }
  function segCtl(name, options, cur) {
    return `<div class="seg-ctl" data-seg="${name}">${options.map(([v, l]) => `<button data-v="${esc(v)}" class="${v === cur ? "on" : ""}">${esc(l)}</button>`).join("")}</div>`;
  }
  function bindSeg($m, name, cb) {
    $m.querySelectorAll(`[data-seg="${name}"] [data-v]`).forEach((b) => (b.onclick = () => cb(b.dataset.v)));
  }
  const INC_KEYS = [{ k: "category", l: "קטגוריה" }, { k: "client", l: "לקוח" }, { k: "level", l: "רמת הסמכה/סיווג" }];

  // ---------- 2/3: income summaries ----------
  function renderIncPivot($m, actualOnly) {
    const id = actualOnly ? "incAct" : "incSum";
    const st = actualOnly ? DONE : S.pv.incStatus;
    const rows = doc().income.filter((r) => !st || r.status === st);
    const draw = () => renderIncPivot($m, actualOnly);
    $m.innerHTML = `${pageHead(actualOnly ? "ריכוז הכנסות בפועל בלבד" : "ריכוז הכנסות", actualOnly ? "כל ההכנסות שסומנו כ״בוצע״, לפי קטגוריה ← לקוח ← רמת הסמכה, לפי חודש" : "סיכום ההכנסות לפי קטגוריה (רמה 1) ← לקוח (רמה 2) ← רמת הסמכה, לפי חודש")}
      <div class="kpis k4">
        <div class="kpi c-inc"><div class="lbl">סה״כ ${actualOnly ? "בוצע" : st ? st : "תחזית + בוצע"}</div><div class="val">${shekel(sum(rows))}</div><div class="sub">${rows.length} שורות</div></div>
        ${[...new Set(rows.map((r) => str(r.category)))].sort(heSort).slice(0, 3).map((c) => `<div class="kpi c-fc"><div class="lbl">${esc(c || "(ריק)")}</div><div class="val">${shekel(sum(rows.filter((r) => str(r.category) === c)))}</div></div>`).join("")}
      </div>
      <div class="pv-tools">${actualOnly ? `<span class="chip-filter">סינון: תחזית/בוצע = בוצע</span>` : segCtl("st", [["", "הכל"], [DONE, "בוצע"], [FC, "תחזית"]], st)}
        <span class="spacer"></span>${pivotTools(id, INC_KEYS)}</div>
      ${pivotHtml(id, rows, INC_KEYS, monthCols("amount"))}`;
    bindPageHead();
    bindSeg($m, "st", (v) => { S.pv.incStatus = v; draw(); });
    bindPivot($m, id, rows, INC_KEYS, draw);
  }

  // ---------- 4/5: expense summaries ----------
  const EXP_KEYS = [{ k: "category", l: "קטגוריה" }, { k: "classification", l: "סיווג" }];
  function expFilter(d = doc()) {
    const f = (d.settings && d.settings.expFilter) || {};
    return { status: f.status || [], priority: f.priority || [], impl: f.impl || [], emergency: f.emergency || [] };
  }
  function expRowsFiltered(d = doc(), actualOnly = false) {
    const f = expFilter(d);
    return d.expenses.filter((r) => EXP_FILTER_KEYS.every(([k]) => !f[k].includes(str(r[k]))) && (!actualOnly || r.status === DONE));
  }
  function filterSummary(d = doc()) {
    const f = expFilter(d);
    const parts = EXP_FILTER_KEYS.filter(([k]) => f[k].length).map(([k, l]) => `${l} ללא ${f[k].map((x) => x || "(ריק)").join(", ")}`);
    return parts.length ? parts.join(" · ") : "ללא סינון – כל ההוצאות";
  }
  function renderExpPivot($m, actualOnly) {
    const id = actualOnly ? "expAct" : "expSum";
    const d = doc(), f = expFilter(d);
    const rows = expRowsFiltered(d, actualOnly);
    const draw = () => renderExpPivot($m, actualOnly);
    const panel = EXP_FILTER_KEYS.map(([k, l]) => {
      const vals = distinct(d.expenses, k);
      return `<div class="fg"><b>${esc(l)}</b><div class="checks">${vals.map((v) => {
        const on = !f[k].includes(v);
        const forced = actualOnly && k === "status" && v !== DONE;
        return `<label class="${on && !forced ? "" : "off"}"><input type="checkbox" data-ef="${k}" value="${esc(v)}" ${on ? "checked" : ""} ${forced ? "disabled" : ""}>${esc(v || "(ריק)")}</label>`;
      }).join("")}</div></div>`;
    }).join("");
    $m.innerHTML = `${pageHead(actualOnly ? "ריכוז הוצאות בפועל בלבד" : "ריכוז הוצאות", actualOnly ? "כל ההוצאות שסומנו כ״בוצע״, לפי קטגוריה ← סיווג, לפי חודש" : "סיכום ההוצאות לפי קטגוריה ← סיווג, לפי חודש")}
      <div class="kpis k4">
        <div class="kpi c-exp"><div class="lbl">סה״כ הוצאות ${actualOnly ? "בפועל" : "בתצוגה"}</div><div class="val">${shekel(sum(rows))}</div><div class="sub">${rows.length} שורות מתוך ${d.expenses.length}</div></div>
        ${!actualOnly ? `<div class="kpi c-done"><div class="lbl">מתוכן בוצע</div><div class="val">${shekel(sum(rows.filter((r) => r.status === DONE)))}</div></div>
        <div class="kpi c-fc"><div class="lbl">מתוכן תחזית</div><div class="val">${shekel(sum(rows.filter((r) => r.status === FC)))}</div></div>` : ""}
        <div class="kpi c-n"><div class="lbl">הוצאות שסוננו החוצה</div><div class="val">${shekel(sum(d.expenses) - sum(expRowsFiltered(d, false)))}</div><div class="sub">לפי המסננים למטה</div></div>
      </div>
      <div class="note">המסננים כאן נשמרים לשנה זו ומשמשים גם את דוח <b>רווח והפסד</b> (כמו מסנני טבלת הפיבוט באקסל). בטל סימון כדי להוציא ערך מהחישוב.${canEdit() ? "" : " (כצופה – השינוי לא יישמר)"}</div>
      <div class="filt-panel">${panel}</div>
      <div class="pv-tools"><span class="chip-filter">${esc(filterSummary(d))}${actualOnly ? " · בוצע בלבד" : ""}</span><span class="spacer"></span>${pivotTools(id, EXP_KEYS)}</div>
      ${pivotHtml(id, rows, EXP_KEYS, monthCols("amount"))}`;
    bindPageHead();
    bindPivot($m, id, rows, EXP_KEYS, draw);
    $m.querySelectorAll("[data-ef]").forEach((cb) => (cb.onchange = async () => {
      const k = cb.dataset.ef, v = cb.value;
      const nf = expFilter(d);
      nf[k] = cb.checked ? nf[k].filter((x) => x !== v) : [...new Set([...nf[k], v])];
      d.settings = { ...(d.settings || {}), expFilter: nf };
      draw();
      if (canEdit()) {
        try { const r = await api(`years/${S.year}/settings`, { method: "PUT", body: { expFilter: nf } }); S.docs[S.year] = r.doc; }
        catch (er) { toast(er.message, true); }
      }
    }));
  }

  // ---------- 7: examinees ----------
  function renderExam($m) {
    const st = S.pv.examStatus;
    const rows = doc().income.filter((r) => !st || r.status === st);
    const draw = () => renderExam($m);
    const done = doc().income.filter((r) => r.status === DONE);
    $m.innerHTML = `${pageHead("כמות נבחנים", "מתעדכן אוטומטית מעמודת ״כמות נבחנים שנגשו״ ברשימת ההכנסות")}
      <div class="kpis k4">
        <div class="kpi c-n"><div class="lbl">סה״כ נבחנים ${st ? "(" + st + ")" : "(תחזית + בוצע)"}</div><div class="val">${fmtN(sum(rows, "examinees"))}</div></div>
        <div class="kpi c-done"><div class="lbl">נבחנים בפועל</div><div class="val">${fmtN(sum(done, "examinees"))}</div></div>
        <div class="kpi c-fc"><div class="lbl">נבחנים בתחזית</div><div class="val">${fmtN(sum(doc().income.filter((r) => r.status === FC), "examinees"))}</div></div>
        <div class="kpi c-inc"><div class="lbl">הכנסה ממוצעת לנבחן (בוצע, מבחנים)</div><div class="val">${(() => { const e = done.filter((r) => num(r.examinees) > 0); const n = sum(e, "examinees"); return n ? shekel(sum(e) / n) : "—"; })()}</div></div>
      </div>
      <div class="pv-tools">${segCtl("st", [["", "הכל"], [DONE, "בוצע"], [FC, "תחזית"]], st)}<span class="spacer"></span>${pivotTools("exam", INC_KEYS)}</div>
      ${pivotHtml("exam", rows, INC_KEYS, monthCols("examinees"))}`;
    bindPageHead();
    bindSeg($m, "st", (v) => { S.pv.examStatus = v; draw(); });
    bindPivot($m, "exam", rows, INC_KEYS, draw);
  }

  // ---------- 8: profit & loss ----------
  const CAT_ORDER = [["מבחנים", "בסיס"], ["פעילות שיווק", "שיווק"], ["אקרדיטציה", "אקרדיטציה"]];
  function plData(d) {
    const inc = d.income, incDone = inc.filter((r) => r.status === DONE);
    const exp = expRowsFiltered(d, false), expDone = expRowsFiltered(d, true);
    const cats = [...new Set(inc.map((r) => str(r.category)))].sort((a, b) => {
      const ia = CAT_ORDER.findIndex((c) => c[0] === a), ib = CAT_ORDER.findIndex((c) => c[0] === b);
      return (ia < 0 ? 99 : ia) - (ib < 0 ? 99 : ib) || heSort(a, b);
    }).map((c) => {
      const fc = sum(inc.filter((r) => str(r.category) === c)), act = sum(incDone.filter((r) => str(r.category) === c));
      return { c, label: (CAT_ORDER.find((x) => x[0] === c) || [0, c])[1], fc, act, gap: fc - act };
    });
    const incFc = sum(inc), incAct = sum(incDone), inc50 = cats.reduce((a, x) => a + x.act + x.gap / 2, 0);
    const expFc = sum(exp), expAct = sum(expDone);
    const open = num(d.settings?.openingBalance);
    return { cats, incFc, incAct, inc50, expFc, expAct, open, profitFc: incFc - expFc, profitAct: incAct - expAct, profit50: inc50 - expFc };
  }
  const sh = (n) => `<span class="${n < 0 ? "neg" : ""}">${shekel(n)}</span>`;
  function renderPL($m) {
    const d = doc(), p = plData(d), s = d.settings || {}, b = s.bank || {};
    const ny = Number(S.year) + 1;
    const bankTotal = num(b.deposits) + num(b.checking) + num(b.eur) * num(b.eurRate) + num(b.usd) * num(b.usdRate);
    const ed = canEdit();
    const nInp = (k, v, ph = "") => ed ? `<input class="input" type="number" step="any" dir="ltr" data-set="${k}" value="${esc(v ?? "")}" placeholder="${ph}">` : `<div class="ro-val">${v ? fmtN(num(v)) : "—"}</div>`;
    $m.innerHTML = `${pageHead("רווח והפסד", "מחושב אוטומטית מהלשוניות: ריכוז הכנסות, ריכוז הכנסות בפועל, ריכוז הוצאות וריכוז הוצאות בפועל")}
      <div class="kpis k4">
        <div class="kpi c-inc"><div class="lbl">הכנסה משוערת ${S.year}</div><div class="val">${shekel(p.incFc)}</div><div class="sub">בפועל עד היום: ${shekel(p.incAct)}</div></div>
        <div class="kpi c-exp"><div class="lbl">הוצאה משוערת ${S.year}</div><div class="val">${shekel(p.expFc)}</div><div class="sub">בפועל עד היום: ${shekel(p.expAct)}</div></div>
        <div class="kpi ${p.profitFc < 0 ? "c-exp" : "c-done"}"><div class="lbl">רווח משוער ${S.year}</div><div class="val ${p.profitFc < 0 ? "neg" : ""}">${shekel(p.profitFc)}</div><div class="sub">נכון להיום: ${sh(p.profitAct)}</div></div>
        <div class="kpi c-n"><div class="lbl">יתרת פתיחה ${ny} (תחזית)</div><div class="val">${shekel(p.open + p.profitFc)}</div><div class="sub">יתרת פתיחה ${S.year}: ${shekel(p.open)}</div></div>
      </div>
      <div class="pl-grid">
        <div class="panel">
          <h2>רווח והפסד ${S.year}</h2><p class="hint">הוצאות לפי מסנני ״ריכוז הוצאות״: ${esc(filterSummary(d))}</p>
          <table class="pl"><thead><tr><th></th><th>תחזית לסוף שנה</th><th>התקדמות נכון להיום</th><th>תחזית 50% משארית יעדים</th></tr></thead><tbody>
            <tr><td>הכנסה משוערת ${S.year}</td><td>${sh(p.incFc)}</td><td>${sh(p.incAct)}</td><td>${sh(p.inc50)}</td></tr>
            <tr><td>הוצאה משוערת ${S.year}</td><td>${sh(p.expFc)}</td><td>${sh(p.expAct)}</td><td>${sh(p.expFc)}</td></tr>
            <tr class="res"><td>רווח משוער ${S.year}</td><td>${sh(p.profitFc)}</td><td>${sh(p.profitAct)}</td><td>${sh(p.profit50)}</td></tr>
            <tr><td>יתרת פתיחה מוערכת ${S.year}</td><td colspan="3">${ed ? `<div class="edit-inline" style="justify-content:center">${nInp("openingBalance", s.openingBalance)}</div>` : shekel(p.open)}</td></tr>
            <tr class="open"><td>יתרת פתיחה ${ny}</td><td>${sh(p.open + p.profitFc)}</td><td>${sh(p.open + p.profitAct)}</td><td>${sh(p.open + p.profit50)}</td></tr>
          </tbody></table>
          <h3 style="margin:22px 0 4px;font-size:16px">הכנסות לפי מקור</h3>
          <p class="hint" style="margin:0">״תחזית 50%״ = מה שבוצע + חצי מהפער שנותר לכל קטגוריה</p>
          <div class="catcards">${p.cats.map((c) => `<div class="catcard"><h4>״${esc(c.label)}״ <span class="muted" style="font-weight:500">${c.label !== c.c ? esc(c.c) : ""}</span></h4>
            <div class="row2"><span>תחזית</span><b>${shekel(c.fc)}</b></div>
            <div class="row2"><span>נכון להיום</span><b>${shekel(c.act)}</b></div>
            <div class="row2"><span>פער נכון להיום</span><b>${sh(c.gap)}</b></div>
            <div class="prog"><i style="width:${c.fc ? Math.min(100, Math.max(0, (c.act / c.fc) * 100)) : 0}%"></i></div></div>`).join("")}</div>
          <h3 style="margin:18px 0 8px;font-size:16px">הכנסות והוצאות לפי חודש</h3>
          ${monthChart(d)}
        </div>
        <div>
          <div class="panel">
            <h2>${ed ? `<input class="input" data-set="balanceTitle" value="${esc(s.balanceTitle || "")}" placeholder="ריכוז יתרות שנתי לסוף ${Number(S.year) - 1} – לא סופי" style="font-weight:700;font-size:16px">` : esc(s.balanceTitle || `ריכוז יתרות שנתי לסוף ${Number(S.year) - 1}`)}</h2>
            <p class="hint">יתרות בנק (₪ / מט״ח) – למעקב בלבד</p>
            <div class="bank-grid">
              <div class="field"><label>פיקדונות (₪)</label>${nInp("bank.deposits", b.deposits)}</div>
              <div class="field"><label>יתרת עו״ש (₪)</label>${nInp("bank.checking", b.checking)}</div>
              <div class="field"><label>יורו (סכום)</label>${nInp("bank.eur", b.eur)}</div>
              <div class="field"><label>שער יורו</label>${nInp("bank.eurRate", b.eurRate, "3.9")}</div>
              <div class="field"><label>דולר (סכום)</label>${nInp("bank.usd", b.usd)}</div>
              <div class="field"><label>שער דולר</label>${nInp("bank.usdRate", b.usdRate, "3.6")}</div>
            </div>
            <div class="preview-box" style="display:flex;justify-content:space-between"><b>סה״כ (בש״ח)</b><b>${shekel(bankTotal)}</b></div>
          </div>
          <div class="panel"><h2>הערות</h2>
            ${ed ? `<textarea class="input" rows="6" data-set="notes">${esc(s.notes || "")}</textarea>` : `<div class="ro-val">${esc(s.notes || "")}</div>`}
          </div>
        </div>
      </div>`;
    bindPageHead();
    bindChartTips($m);
    if (!ed) return;
    $m.querySelectorAll("[data-set]").forEach((el) => (el.onchange = async () => {
      const k = el.dataset.set; const body = {};
      if (k.startsWith("bank.")) body.bank = { ...(s.bank || {}), [k.slice(5)]: num(el.value) };
      else body[k] = el.type === "number" ? num(el.value) : el.value;
      try { const r = await api(`years/${S.year}/settings`, { method: "PUT", body }); S.docs[S.year] = r.doc; toast("נשמר"); renderPL($m); }
      catch (er) { toast(er.message, true); }
    }));
  }
  function monthChart(d) {
    const series = [
      ["הכנסה – תחזית+בוצע", "#8C97E6", (m) => sum(d.income.filter((r) => num(r.month) === m))],
      ["הכנסה – בוצע", "#0F9488", (m) => sum(d.income.filter((r) => num(r.month) === m && r.status === DONE))],
      ["הוצאה (לפי מסננים)", "#F06A2C", (m) => sum(expRowsFiltered(d).filter((r) => num(r.month) === m))],
    ];
    const data = [];
    for (let m = 1; m <= 12; m++) data.push(series.map((s) => s[2](m)));
    const max = Math.max(1, ...data.flat());
    const pow = Math.pow(10, Math.floor(Math.log10(max)));
    const step = [1, 2, 2.5, 5, 10].map((x) => x * pow).find((st) => max / st <= 5) || pow * 10;
    const top = Math.ceil(max / step) * step;
    const W = 900, H = 260, padL = 58, padB = 26, padT = 10;
    const gw = (W - padL - 6) / 12, bw = Math.min(16, (gw * 0.75) / 3);
    const yv = (v) => padT + (H - padT - padB) * (1 - v / top);
    let g = "";
    for (let v = 0; v <= top + 1e-9; v += step) g += `<line x1="${padL}" x2="${W}" y1="${yv(v)}" y2="${yv(v)}" stroke="#E7E9F1"/><text x="${padL - 6}" y="${yv(v) + 4}" text-anchor="end" class="ax">${v >= 1000 ? (v / 1000).toLocaleString("he-IL") + "K" : v}</text>`;
    data.forEach((vals, i) => {
      const m = i + 1, x0 = W - gw * (i + 1) + (gw - bw * 3) / 2; // RTL: January on the right
      vals.forEach((v, j) => {
        const x = x0 + bw * (2 - j), h = Math.max(yv(0) - yv(v), v ? 2 : 0);
        g += `<rect x="${x}" y="${yv(0) - h}" width="${bw - 1.5}" height="${h}" rx="2" fill="${series[j][1]}" data-tip="${MONTHS[m]}|${series[j][0]}|${Math.round(v)}|${series[j][1]}"/>`;
      });
      g += `<text x="${W - gw * i - gw / 2}" y="${H - 8}" text-anchor="middle" class="ax">${MSHORT[m]}</text>`;
    });
    return `<svg viewBox="0 0 ${W} ${H}" class="mchart" role="img" aria-label="הכנסות והוצאות לפי חודש">${g}</svg>
      <div class="legend">${series.map(([l, c]) => `<span style="--c:${c}">${l}</span>`).join("")}</div>`;
  }
  function bindChartTips($m) {
    const tip = document.getElementById("dtip");
    $m.querySelectorAll("[data-tip]").forEach((el) => {
      el.onmousemove = (e) => {
        const [a, b, v, c] = el.dataset.tip.split("|");
        tip.innerHTML = `<div class="tt-l">${esc(a)}</div><span class="sw" style="background:${c}"></span>${esc(b)}: <b>${shekel(+v)}</b>`;
        tip.style.opacity = 1; tip.style.left = e.clientX + 14 + "px"; tip.style.top = e.clientY + 14 + "px";
      };
      el.onmouseleave = () => (tip.style.opacity = 0);
    });
  }

  // ---------- 9: forecast vs actual by sector ----------
  function catOptions() {
    return [["", "כל הקטגוריות"], ...distinct(doc().income, "category").map((c) => [c, c || "(ריק)"])];
  }
  function renderFc($m) {
    const cat = S.pv.fcCat;
    const rows = doc().income.filter((r) => !cat || str(r.category) === cat);
    const keys = [{ k: "sector", l: "מגזר" }];
    const by = (st, k) => (rr) => sum(rr.filter((r) => r.status === st), k);
    const cols = [
      { l: "בוצע<br>סכום ₪", fn: by(DONE, "amount"), m: 1 }, { l: "בוצע<br>נבחנים", fn: by(DONE, "examinees") },
      { l: "תחזית<br>סכום ₪", fn: by(FC, "amount"), m: 1 }, { l: "תחזית<br>נבחנים", fn: by(FC, "examinees") },
      { l: "סה״כ<br>סכום ₪", fn: (rr) => sum(rr), tot: 1, m: 1 }, { l: "סה״כ<br>נבחנים", fn: (rr) => sum(rr, "examinees"), tot: 1 },
    ];
    const sectors = buildTree(rows, keys);
    const list = [...sectors.children.values()].map((n) => ({ l: n.label, done: sum(n.rows.filter((r) => r.status === DONE)), fc: sum(n.rows.filter((r) => r.status === FC)) })).sort((a, b) => b.done + b.fc - (a.done + a.fc));
    const max = Math.max(1, ...list.map((x) => x.done + x.fc));
    $m.innerHTML = `${pageHead("הכנסות – תחזית מול ביצוע", "חישובים בלבד: סכום וכמות נבחנים לפי מגזר, בוצע מול תחזית")}
      <div class="pv-tools">${segCtl("cat", catOptions(), cat)}</div>
      <div class="split2">
        ${pivotHtml("fc", rows, keys, cols)}
        <div class="panel"><h2 style="font-size:16px">התפלגות לפי מגזר</h2>
          <div class="legend"><span style="--c:#0F9488">בוצע</span><span style="--c:#8C97E6">תחזית</span></div>
          <div class="hbars">${list.map((x) => `<div class="hr"><span>${esc(x.l)}</span><div class="tr"><i style="width:${(x.done / max) * 100}%;background:#0F9488"></i><i style="width:${(x.fc / max) * 100}%;background:#8C97E6"></i></div><span class="v">${shekel(x.done + x.fc)}</span></div>`).join("")}</div>
        </div>
      </div>`;
    bindPageHead();
    bindSeg($m, "cat", (v) => { S.pv.fcCat = v; renderFc($m); });
  }

  // ---------- 10/11: income source ----------
  function renderSrc($m, actualOnly) {
    const id = actualOnly ? "srcAct" : "srcAll";
    const ck = actualOnly ? "srcActCat" : "srcAllCat";
    const cat = S.pv[ck];
    const rows = doc().income.filter((r) => (!actualOnly || r.status === DONE) && (!cat || str(r.category) === cat));
    const keys = [{ k: "sector", l: "מגזר" }, { k: "client", l: "לקוח" }];
    const cols = [
      { l: 'סכום בש"ח', fn: (rr) => sum(rr), tot: 1 },
      { l: "כמות נבחנים שנגשו", fn: (rr) => sum(rr, "examinees") },
      { l: "% מהסך", fn: (rr) => (sum(rows) ? (sum(rr) / sum(rows)) * 100 : 0), pct: 1 },
    ];
    const fmt = (v, c) => (c && c.pct ? `${v.toFixed(1)}%` : `<span class="${v === 0 ? "z" : ""}">${fmtN(v)}</span>`);
    const draw = () => renderSrc($m, actualOnly);
    $m.innerHTML = `${pageHead(actualOnly ? "מקור הכנסה – ביצוע עד כה" : "מקור הכנסה – תחזית (כללי)", actualOnly ? "הכנסות שבוצעו בפועל לפי מגזר ← לקוח" : "כל ההכנסות (תחזית + בוצע) לפי מגזר ← לקוח")}
      <div class="kpis k4">
        <div class="kpi c-inc"><div class="lbl">סה״כ הכנסות</div><div class="val">${shekel(sum(rows))}</div></div>
        <div class="kpi c-n"><div class="lbl">כמות נבחנים</div><div class="val">${fmtN(sum(rows, "examinees"))}</div></div>
        <div class="kpi c-fc"><div class="lbl">מגזרים</div><div class="val">${new Set(rows.map((r) => str(r.sector))).size}</div></div>
        <div class="kpi c-fc"><div class="lbl">לקוחות</div><div class="val">${new Set(rows.map((r) => str(r.client))).size}</div></div>
      </div>
      <div class="pv-tools">${actualOnly ? `<span class="chip-filter">תחזית/בוצע = בוצע</span>` : ""}${segCtl("cat", catOptions(), cat)}<span class="spacer"></span>${pivotTools(id, keys)}</div>
      <div style="max-width:900px">${pivotHtml(id, rows, keys, cols, { fmt })}</div>`;
    bindPageHead();
    bindSeg($m, "cat", (v) => { S.pv[ck] = v; draw(); });
    bindPivot($m, id, rows, keys, draw);
  }

  // ================= USERS =================
  async function renderUsers($m) {
    $m.innerHTML = `<div class="loading">טוען…</div>`;
    try { S.users = (await api("users")).users; } catch (e) { $m.innerHTML = `<div class="panel">${esc(e.message)}</div>`; return; }
    $m.innerHTML = `
      <div class="panel"><h2>ניהול משתמשים</h2><p class="hint">הוספת משתמשים עם מייל וסיסמה זמנית. בכניסה הראשונה המשתמש יתבקש לבחור סיסמה אישית.</p>
        <form class="add-user" id="addU">
          <div class="field"><label for="un">שם</label><input class="input" id="un" required></div>
          <div class="field"><label for="ue">מייל</label><input class="input" id="ue" type="email" dir="ltr" required></div>
          <div class="field"><label for="ur">הרשאה</label><select class="input" id="ur"><option value="viewer">צופה</option><option value="editor">עורך</option><option value="admin">מנהל</option></select></div>
          ${pwField("up", "סיסמה זמנית", "new-password", "required minlength=8")}
          <button type="button" class="btn sm" id="gen" style="align-self:end;margin-bottom:6px">יצירת סיסמה</button>
          <button class="btn primary" style="align-self:end">הוספה</button>
        </form>
        <div class="err" id="uerr"></div>
        <div class="roles-help"><div><b>צופה</b>צפייה בכל הלשוניות וייצוא קבצים</div><div><b>עורך</b>הוספה, עריכה ומחיקה של שורות, ייבוא קבצים ויצירת שנים</div><div><b>מנהל</b>כל הרשאות העורך + ניהול משתמשים ומחיקת שנים</div></div>
      </div>
      <div class="panel" style="overflow:auto"><table class="simple"><thead><tr><th>שם</th><th>מייל</th><th>הרשאה</th><th>כניסה אחרונה</th><th>סטטוס</th><th></th></tr></thead><tbody>
        ${S.users.map((u) => `<tr><td><b>${esc(u.name)}</b></td><td dir="ltr" style="text-align:right">${esc(u.email)}</td>
          <td><select class="input" style="width:auto" data-role="${esc(u.email)}">${Object.entries(ROLE_LABEL).map(([k, l]) => `<option value="${k}" ${u.role === k ? "selected" : ""}>${l}</option>`).join("")}</select></td>
          <td>${u.lastLogin ? fmtDate(u.lastLogin) : '<span class="muted">טרם נכנס</span>'}</td>
          <td>${u.mustChange ? '<span class="badge b-fc">ממתין לבחירת סיסמה</span>' : '<span class="badge b-done">פעיל</span>'}</td>
          <td class="nowrap"><button class="icon-btn" data-reset="${esc(u.email)}">איפוס סיסמה</button> ${u.email !== S.user.email ? `<button class="icon-btn del" data-udel="${esc(u.email)}">מחיקה</button>` : ""}</td></tr>`).join("")}
      </tbody></table></div>`;
    bindEyes($m);
    gen.onclick = () => { up.value = genPw(); up.type = "text"; };
    addU.onsubmit = async (e) => {
      e.preventDefault();
      try {
        S.users = (await api("users", { method: "POST", body: { name: un.value, email: ue.value, role: ur.value, password: up.value } })).users;
        const info = `שלום ${un.value},\nנוצר עבורך משתמש במערכת התקציב של ITCB:\n${location.origin}\nמייל: ${ue.value}\nסיסמה זמנית: ${up.value}\nבכניסה הראשונה תתבקש/י לבחור סיסמה אישית.`;
        showInvite(info); renderUsers($m);
      } catch (er) { uerr.textContent = er.message; }
    };
    $m.querySelectorAll("[data-role]").forEach((s) => (s.onchange = async () => {
      try { S.users = (await api("users/" + encodeURIComponent(s.dataset.role), { method: "PUT", body: { role: s.value } })).users; toast("ההרשאה עודכנה"); }
      catch (er) { toast(er.message, true); renderUsers($m); }
    }));
    $m.querySelectorAll("[data-reset]").forEach((b) => (b.onclick = () => {
      const email = b.dataset.reset;
      const m = modal("איפוס סיסמה", `<p>סיסמה זמנית חדשה עבור <b dir="ltr">${esc(email)}</b>. המשתמש יתבקש להחליף אותה בכניסה הבאה.</p>${pwField("rp", "סיסמה זמנית", "new-password")}<div class="err" id="e"></div>`,
        `<button class="btn primary" id="ok">איפוס</button><button class="btn" id="g">יצירת סיסמה</button><button class="btn" data-x>ביטול</button>`);
      m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
      m.$("#g").onclick = () => { m.$("#rp").value = genPw(); m.$("#rp").type = "text"; };
      m.$("#ok").onclick = async () => {
        try {
          await api("users/" + encodeURIComponent(email), { method: "PUT", body: { password: m.$("#rp").value } });
          const pw = m.$("#rp").value; m.close();
          showInvite(`הסיסמה שלך במערכת התקציב של ITCB אופסה.\n${location.origin}\nמייל: ${email}\nסיסמה זמנית: ${pw}`);
          renderUsers($m);
        } catch (er) { m.$("#e").textContent = er.message; }
      };
    }));
    $m.querySelectorAll("[data-udel]").forEach((b) => (b.onclick = () => {
      const email = b.dataset.udel;
      const m = modal("מחיקת משתמש", `<p>למחוק את המשתמש <b dir="ltr">${esc(email)}</b>?</p><div class="err" id="e"></div>`,
        `<button class="btn primary" style="background:var(--danger);border-color:var(--danger)" id="ok">מחיקה</button><button class="btn" data-x>ביטול</button>`);
      m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
      m.$("#ok").onclick = async () => {
        try { await api("users/" + encodeURIComponent(email), { method: "DELETE" }); m.close(); toast("המשתמש נמחק"); renderUsers($m); }
        catch (er) { m.$("#e").textContent = er.message; }
      };
    }));
  }
  function showInvite(text) {
    const m = modal("פרטי כניסה למשלוח", `<p class="muted">העתק ושלח למשתמש (הסיסמה מוצגת רק עכשיו):</p><textarea class="input invite-text" rows="7" readonly>${esc(text)}</textarea>
      <div class="invite-actions"><button class="btn primary" id="cp">העתקה</button><a class="btn wa-btn" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(text)}">שליחה בוואטסאפ</a></div>`,
      `<button class="btn" data-x>סגירה</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    m.$("#cp").onclick = async () => { try { await navigator.clipboard.writeText(text); toast("הועתק"); } catch { m.$("textarea").select(); } };
  }
  function openChangePassword() {
    const m = modal("שינוי סיסמה", `${pwField("c0", "סיסמה נוכחית", "current-password")}${pwField("c1", "סיסמה חדשה (8 תווים לפחות)", "new-password")}${pwField("c2", "אימות סיסמה חדשה", "new-password")}<div class="err" id="e"></div>`,
      `<button class="btn primary" id="ok">שמירה</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    m.$("#ok").onclick = async () => {
      if (m.$("#c1").value !== m.$("#c2").value) return (m.$("#e").textContent = "הסיסמאות אינן תואמות");
      try {
        const r = await api("password", { method: "POST", body: { current: m.$("#c0").value, password: m.$("#c1").value } });
        S.token = r.token; store.set("itcb_budget_token", r.token); m.close(); toast("הסיסמה עודכנה");
      } catch (er) { m.$("#e").textContent = er.message; }
    };
  }

  // ================= YEARS: new / delete =================
  function openNewYear() {
    const next = S.years.length ? String(Math.max(...S.years.map(Number)) + 1) : String(new Date().getFullYear());
    const src = S.years[S.years.length - 1];
    const m = modal("שנה חדשה", `
      <div class="field"><label for="ny">שנה</label><input class="input" id="ny" type="number" dir="ltr" value="${next}" min="2000" max="2100"></div>
      <div class="radio-list">
        <label><input type="radio" name="mode" value="empty" checked><div><b>שנה ריקה</b><small>מתחילים טבלאות הכנסות והוצאות ריקות</small></div></label>
        ${src ? `<label><input type="radio" name="mode" value="copy"><div><b>העתקת השורות מ-<select class="input" id="cf" style="width:auto;display:inline-block;padding:2px 8px">${S.years.map((y) => `<option ${y === src ? "selected" : ""}>${y}</option>`).join("")}</select> כתחזית</b><small>כל ההכנסות וההוצאות יועתקו בסטטוס ״תחזית״, ללא מספרי קבלה/אסמכתא – בסיס נוח לתכנון</small></div></label>` : ""}
        <label><input type="radio" name="mode" value="import"><div><b>ייבוא מקובץ</b><small>קובץ אקסל בפורמט התקציב או קובץ גיבוי (JSON) משנים קודמות</small></div></label>
      </div><div class="err" id="e"></div>`,
      `<button class="btn primary" id="ok">יצירה</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    m.$("#ok").onclick = async () => {
      const year = m.$("#ny").value.trim();
      const mode = m.ov.querySelector("input[name=mode]:checked").value;
      if (!/^\d{4}$/.test(year)) return (m.$("#e").textContent = "שנה לא תקינה");
      if (S.years.includes(year)) return (m.$("#e").textContent = "השנה כבר קיימת");
      if (mode === "import") { m.close(); return openImport(year); }
      try {
        const body = { year };
        if (mode === "copy") { body.copyFrom = m.$("#cf").value; body.copyRows = true; }
        const r = await api("years", { method: "POST", body });
        S.years = r.years; S.docs[year] = r.doc; S.year = year; S.view = "income"; m.close(); render(); toast(`שנת ${year} נוצרה`);
      } catch (er) { m.$("#e").textContent = er.message; }
    };
  }
  function confirmDeleteYear() {
    const y = S.year;
    const m = modal(`מחיקת שנת ${y}`, `<p>כל נתוני שנת <b>${y}</b> (${doc().income.length} הכנסות, ${doc().expenses.length} הוצאות) יימחקו לצמיתות.</p>
      <p>מומלץ לייצא גיבוי קודם. להמשך הקלד את השנה:</p><input class="input" id="cy" dir="ltr"><div class="err" id="e"></div>`,
      `<button class="btn" id="bk">הורדת גיבוי</button><button class="btn primary" style="background:var(--danger);border-color:var(--danger)" id="ok">מחיקה לצמיתות</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    m.$("#bk").onclick = () => downloadJSON({ app: "itcb-budget", version: VERSION, exported: new Date().toISOString(), years: { [y]: stripDoc(doc()) } }, `ITCB_Budget_${y}.json`);
    m.$("#ok").onclick = async () => {
      if (m.$("#cy").value.trim() !== y) return (m.$("#e").textContent = "השנה שהוקלדה אינה תואמת");
      try {
        const r = await api("years/" + y, { method: "DELETE" });
        S.years = r.years; delete S.docs[y]; S.year = S.years[S.years.length - 1] || null;
        if (S.year) await loadYear(S.year);
        m.close(); render(); toast(`שנת ${y} נמחקה`);
      } catch (er) { m.$("#e").textContent = er.message; }
    };
  }

  // ================= EXPORT =================
  const stripDoc = (d) => ({
    year: d.year, settings: d.settings,
    income: d.income.map(({ id, updatedAt, updatedBy, ...r }) => r),
    expenses: d.expenses.map(({ id, updatedAt, updatedBy, ...r }) => r),
  });
  function downloadBlob(blob, name) {
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob); a.download = name; document.body.appendChild(a); a.click();
    setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 500);
  }
  const downloadJSON = (obj, name) => { downloadBlob(new Blob([JSON.stringify(obj, null, 1)], { type: "application/json" }), name); toast("הקובץ הורד"); };
  async function exportAllJSON() {
    const out = {};
    for (const y of S.years) out[y] = stripDoc(await loadYear(y));
    downloadJSON({ app: "itcb-budget", version: VERSION, exported: new Date().toISOString(), years: out }, `ITCB_Budget_all_years.json`);
  }
  function pivotAoa(rows, keys, cols) {
    const aoa = [[...keys.map((k) => k.l), ...cols.map((c) => c.x || c.l.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim())]];
    const walk = (node, prefix) => {
      for (const n of [...node.children.values()].sort((a, b) => heSort(a.label, b.label))) {
        const line = [...prefix, n.label]; while (line.length < keys.length) line.push("");
        aoa.push([...line.map((x, i) => (i === n.level && n.children.size ? x + " – סה״כ" : x)), ...cols.map((c) => { const v = c.fn(n.rows); return v === null ? "" : Math.round(v * 100) / 100; })]);
        walk(n, [...prefix, n.label].map(() => ""));
      }
    };
    walk(buildTree(rows, keys), []);
    aoa.push(["סה״כ כללי", ...keys.slice(1).map(() => ""), ...cols.map((c) => { const v = c.fn(rows); return v === null ? "" : Math.round(v * 100) / 100; })]);
    return aoa;
  }
  function exportXlsx(d) {
    if (!window.XLSX) return toast("רכיב האקסל עדיין נטען, נסה שוב", true);
    XLSX.writeFile(buildWorkbook(d, d.invoices || []), `ITCB_Budget_${d.year}.xlsx`, { compression: true });
    toast("קובץ האקסל הורד");
  }
  function workbookBytes(d, invList) {
    return XLSX.write(buildWorkbook(d, invList), { type: "array", bookType: "xlsx", compression: true });
  }
  function buildWorkbook(d, invList) {
    const wb = XLSX.utils.book_new();
    wb.Workbook = { Views: [{ RTL: true }] };
    const add = (name, aoa, widths) => {
      const ws = XLSX.utils.aoa_to_sheet(aoa, { cellDates: true });
      ws["!cols"] = (widths || aoa[0].map(() => 14)).map((w) => ({ wch: w }));
      if (aoa.length > 1) ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: aoa.length - 1, c: aoa[0].length - 1 } }) };
      XLSX.utils.book_append_sheet(wb, ws, name.slice(0, 31));
    };
    const listAoa = (cols, rows) => [cols.map((c) => c.l), ...rows.map((r) => cols.map((c) => {
      const v = r[c.k];
      if (v === null || v === undefined) return "";
      if (c.t === "date" && v) { const dt = new Date(String(v).slice(0, 10) + "T12:00:00"); return isNaN(dt) ? v : dt; }
      return v;
    }))];
    const inc = d.income, dn = inc.filter((r) => r.status === DONE);
    add("הכנסות", listAoa(INC, inc), [14, 10, 14, 14, 12, 7, 12, 50, 10, 12, 12, 8, 18, 24, 12]);
    add("ריכוז הכנסות", pivotAoa(inc, INC_KEYS, monthCols("amount")));
    add("ריכוז הכנסות בפועל בלבד", pivotAoa(dn, INC_KEYS, monthCols("amount")));
    add("ריכוז הוצאות", pivotAoa(expRowsFiltered(d), EXP_KEYS, monthCols("amount")));
    add("ריכוז הוצאות בפועל בלבד", pivotAoa(expRowsFiltered(d, true), EXP_KEYS, monthCols("amount")));
    add("הוצאות", listAoa(EXP, d.expenses), [7, 6, 14, 10, 60, 12, 16, 6, 12, 7, 14, 7, 9, 24]);
    add("כמות נבחנים", pivotAoa(inc, INC_KEYS, monthCols("examinees")));
    const p = plData(d), ny = Number(d.year) + 1, s = d.settings || {}, b = s.bank || {};
    add("רווח-הפסד", [
      ["", "תחזית לסוף שנה", "התקדמות נכון להיום", "תחזית 50% משארית יעדים"],
      [`הכנסה משוערת ${d.year}`, p.incFc, p.incAct, p.inc50],
      [`הוצאה משוערת ${d.year}`, p.expFc, p.expAct, p.expFc],
      [`רווח משוער ${d.year}`, p.profitFc, p.profitAct, p.profit50],
      [`יתרת פתיחה מוערכת ${d.year}`, p.open, p.open, p.open],
      [`יתרת פתיחה ${ny}`, p.open + p.profitFc, p.open + p.profitAct, p.open + p.profit50],
      [],
      ["מקור", "תחזית", "נכון להיום", "פער"],
      ...p.cats.map((c) => [c.label === c.c ? c.c : `${c.label} (${c.c})`, c.fc, c.act, c.gap]),
      [],
      [s.balanceTitle || "יתרות בנק"], ["פיקדונות", num(b.deposits)], ["יתרת עו״ש", num(b.checking)], ["יורו", num(b.eur), "שער", num(b.eurRate)], ["דולר", num(b.usd), "שער", num(b.usdRate)],
      [], ["הוצאות לפי מסננים:", filterSummary(d)], [], ["הערות"], ...(s.notes || "").split("\n").map((l) => [l]),
    ], [34, 18, 18, 22]);
    const sec = [{ k: "sector", l: "מגזר" }];
    const by = (st, k) => (rr) => sum(rr.filter((r) => r.status === st), k);
    add("הכנסות-תחזית מול ביצוע", pivotAoa(inc, sec, [{ l: "בוצע – סכום", fn: by(DONE, "amount") }, { l: "בוצע – נבחנים", fn: by(DONE, "examinees") }, { l: "תחזית – סכום", fn: by(FC, "amount") }, { l: "תחזית – נבחנים", fn: by(FC, "examinees") }, { l: "סה״כ סכום", fn: (r) => sum(r) }, { l: "סה״כ נבחנים", fn: (r) => sum(r, "examinees") }]));
    const sk = [{ k: "sector", l: "מגזר" }, { k: "client", l: "לקוח" }];
    const sc = [{ l: 'סכום בש"ח', fn: (r) => sum(r) }, { l: "כמות נבחנים שנגשו", fn: (r) => sum(r, "examinees") }];
    add("מקור הכנסה - ביצוע עד כה", pivotAoa(dn, sk, sc), [22, 30, 14, 14]);
    add("מקור הכנסה - כללי", pivotAoa(inc, sk, sc), [22, 30, 14, 14]);
    if (invList && invList.length) add("חשבוניות", invoiceRegisterAoa(d, invList), [12, 26, 14, 12, 40, 18, 22, 12, 30]);
    add("הגדרות", [["מפתח", "ערך"], ["year", d.year], ["settings", JSON.stringify(d.settings || {})]], [12, 80]);
    return wb;
  }

  // ================= IMPORT =================
  const INC_MAP = [["category", ["קטגוריה"]], ["status", ["תחזית/בוצע", "תחזית"]], ["body", ["גוף"]], ["level", ["רמת הסמכה", "סיווג"]], ["receipt", ["קבלה"]], ["month", ["חודש"]], ["amount", ["סכום"]],
    ["detail", ["פירוט"]], ["examinees", ["נבחנים"]], ["training", ["הדרכה"]], ["price", ["מחיר"]], ["sector", ["מגזר"]], ["client", ["לקוח"]], ["examDate", ["תאריך בחינה", "תאריך מבחן"]], ["date", ["תאריך"]]];
  const EXP_MAP = [["serial", ["סידורי"]], ["year", ["שנה"]], ["ref", ["אסמכתא", "חשבונית"]], ["amount", ["סכום"]], ["detail", ["פירוט"]], ["category", ["קטגוריה"]], ["classification", ["סיווג"]],
    ["month", ["חודש"]], ["status", ["תחזית"]], ["priority", ["עדיפות"]], ["manager", ["מנהל"]], ["impl", ["יישום"]], ["emergency", ["חירום"]], ["notes", ["הערות"]]];
  function sheetRows(ws, map, mustHave) {
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: null });
    const hi = aoa.findIndex((r) => r && mustHave.every((h) => r.some((c) => typeof c === "string" && c.includes(h))));
    if (hi < 0) return null;
    const head = aoa[hi].map((h) => (typeof h === "string" ? h.trim() : ""));
    const idx = {}; const used = new Set();
    for (const [k, names] of map) {
      let i = -1;
      for (const n of names) { i = head.findIndex((h, j) => !used.has(j) && h && h.includes(n)); if (i >= 0) break; }
      if (i >= 0) { idx[k] = i; used.add(i); }
    }
    // unnamed trailing column in expenses = notes (as in the original file)
    if (map === EXP_MAP && idx.notes === undefined) {
      const last = Math.max(...Object.values(idx));
      const hasData = aoa.slice(hi + 1).some((r) => r && r[last + 1] !== null && r[last + 1] !== undefined && r[last + 1] !== "");
      if (hasData && !head[last + 1]) idx.notes = last + 1;
    }
    const out = [];
    for (const r of aoa.slice(hi + 1)) {
      if (!r) continue;
      const o = {};
      let any = false;
      for (const [k, i] of Object.entries(idx)) {
        let v = r[i];
        if (v instanceof Date) v = new Date(v.getTime() + 12 * 3600e3).toISOString().slice(0, 10);
        else if (typeof v === "string") v = v.trim();
        if (v === "" || v === undefined) v = null;
        if (v !== null) any = true;
        o[k] = v;
      }
      if (!any) continue;
      if (o.amount === null && !o.detail && !o.category) continue;
      if (typeof o.date === "number") o.date = XLSX.SSF.format("yyyy-mm-dd", o.date);
      out.push(o);
    }
    return out;
  }
  async function parseFile(file) {
    const name = file.name.toLowerCase();
    if (name.endsWith(".json")) {
      const j = JSON.parse(await file.text());
      if (j.years) return Object.entries(j.years).map(([y, d]) => ({ year: y, data: d }));
      if (j.income || j.expenses) return [{ year: String(j.year || ""), data: j }];
      throw new Error("קובץ JSON לא מזוהה");
    }
    if (!window.XLSX) throw new Error("רכיב האקסל עדיין נטען, נסה שוב");
    const wb = XLSX.read(await file.arrayBuffer(), { cellDates: true });
    const pick = (pref, must) => {
      const order = [...wb.SheetNames.filter((n) => n.trim() === pref), ...wb.SheetNames.filter((n) => n.trim() !== pref)];
      for (const n of order) { const r = sheetRows(wb.Sheets[n], pref === "הכנסות" ? INC_MAP : EXP_MAP, must); if (r) return r; }
      return null;
    };
    const income = pick("הכנסות", ["קטגוריה", "סכום", "לקוח"]) || [];
    const expenses = pick("הוצאות", ["פירוט", "סיווג", "עדיפות"]) || [];
    if (!income.length && !expenses.length) throw new Error("לא נמצאו לשוניות ״הכנסות״ / ״הוצאות״ בפורמט המוכר");
    let settings, year = "";
    const st = wb.Sheets["הגדרות"];
    if (st) {
      const aoa = XLSX.utils.sheet_to_json(st, { header: 1 });
      for (const r of aoa) { if (r[0] === "settings") try { settings = JSON.parse(r[1]); } catch {} if (r[0] === "year") year = String(r[1]); }
    }
    if (!settings) { // original workbook: take the opening balance from the P&L sheet
      const pl = wb.SheetNames.find((n) => n.includes("רווח"));
      if (pl) {
        const aoa = XLSX.utils.sheet_to_json(wb.Sheets[pl], { header: 1, raw: true, defval: null });
        for (const r of aoa) {
          const i = (r || []).findIndex((c) => typeof c === "string" && c.includes("פתיחה מוערכת"));
          if (i >= 0 && isNum(r[i + 1])) { settings = { openingBalance: num(r[i + 1]), expFilter: { status: [NOTHAPPEN], priority: [], impl: [], emergency: [] } }; break; }
        }
      }
    }
    if (!year) {
      const cnt = {}; expenses.forEach((r) => { if (isNum(r.year)) cnt[r.year] = (cnt[r.year] || 0) + 1; });
      year = Object.entries(cnt).sort((a, b) => b[1] - a[1])[0]?.[0] || (file.name.match(/20\d\d/) || [""])[0];
    }
    return [{ year: String(year), data: { income, expenses, settings } }];
  }
  function openImport(presetYear) {
    let parsed = null;
    const m = modal("ייבוא קובץ", `
      <p class="hint" style="margin-top:0">אפשר לייבא <b>קובץ אקסל</b> בפורמט התקציב (לשוניות ״הכנסות״ ו״הוצאות״, כמו הקובץ המקורי או קובץ שיוצא מהאתר) או <b>קובץ גיבוי JSON</b> שיוצא מהאתר.</p>
      <label class="file-drop" for="ff"><b>בחירת קובץ</b><br><span class="muted">‎.xlsx · .xls · .json</span><input type="file" id="ff" accept=".xlsx,.xls,.xlsm,.json" hidden></label>
      <div id="pv" style="margin-top:14px"></div><div class="err" id="e"></div>`,
      `<button class="btn primary" id="ok" disabled>ייבוא</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    const pv = m.$("#pv");
    m.$("#ff").onchange = async (e) => {
      const file = e.target.files[0]; if (!file) return;
      m.$("#e").textContent = ""; pv.innerHTML = `<div class="muted">קורא את הקובץ…</div>`;
      try {
        parsed = await parseFile(file);
        if (presetYear && parsed.length === 1) parsed[0].year = presetYear;
        pv.innerHTML = `<div class="preview-box"><b>${esc(file.name)}</b></div>` + parsed.map((p, i) => `
          <div class="preview-box">
            <div class="field" style="margin-bottom:8px"><label>שנת יעד</label><input class="input" data-y="${i}" dir="ltr" value="${esc(p.year)}" style="width:120px"></div>
            ${(p.data.income || []).length} הכנסות (${shekel(sum(p.data.income || []))}) · ${(p.data.expenses || []).length} הוצאות (${shekel(sum(p.data.expenses || []))})
            <div class="radio-list" data-mode="${i}" style="margin:10px 0 0"></div>
          </div>`).join("");
        const upd = () => parsed.forEach((p, i) => {
          const y = m.ov.querySelector(`[data-y="${i}"]`).value.trim();
          p.year = y;
          const box = m.ov.querySelector(`[data-mode="${i}"]`);
          const cur = box.querySelector("input:checked")?.value;
          box.innerHTML = S.years.includes(y)
            ? `<label><input type="radio" name="m${i}" value="replace" ${cur !== "append" ? "checked" : ""}><div><b>החלפת נתוני ${esc(y)}</b><small>הנתונים הקיימים של השנה יוחלפו בנתוני הקובץ</small></div></label>
               <label><input type="radio" name="m${i}" value="append" ${cur === "append" ? "checked" : ""}><div><b>הוספה לשנת ${esc(y)}</b><small>שורות הקובץ יתווספו לשורות הקיימות</small></div></label>`
            : `<div class="muted">${/^\d{4}$/.test(y) ? `תיווצר שנה חדשה: <b>${esc(y)}</b>` : "יש להזין שנה בת 4 ספרות"}</div>`;
        });
        m.ov.querySelectorAll("[data-y]").forEach((inp) => (inp.oninput = upd));
        upd();
        m.$("#ok").disabled = false;
      } catch (er) { pv.innerHTML = ""; m.$("#e").textContent = er.message; }
    };
    m.$("#ok").onclick = async () => {
      if (!parsed) return;
      const btn = m.$("#ok"); btn.disabled = true;
      try {
        for (let i = 0; i < parsed.length; i++) {
          const p = parsed[i];
          if (!/^\d{4}$/.test(p.year)) throw new Error("שנה לא תקינה: " + p.year);
          const mode = m.ov.querySelector(`input[name=m${i}]:checked`)?.value;
          if (S.years.includes(p.year)) {
            const cur = await loadYear(p.year);
            if (mode === "append") {
              let r;
              if ((p.data.income || []).length) r = await api(`years/${p.year}/income`, { method: "POST", body: { rows: p.data.income } });
              if ((p.data.expenses || []).length) r = await api(`years/${p.year}/expenses`, { method: "POST", body: { rows: p.data.expenses } });
              if (r) S.docs[p.year] = r.doc;
            } else {
              const data = { ...p.data, settings: p.data.settings || cur.settings };
              S.docs[p.year] = (await api(`years/${p.year}`, { method: "PUT", body: { data } })).doc;
            }
          } else {
            const data = { ...p.data };
            if (!data.settings) data.settings = { expFilter: { status: [NOTHAPPEN], priority: [], impl: [], emergency: [] } };
            const r = await api("years", { method: "POST", body: { year: p.year, data } });
            S.years = r.years; S.docs[p.year] = r.doc;
          }
          S.year = p.year;
        }
        m.close(); S.view = "income"; render(); toast("הייבוא הושלם");
      } catch (er) { m.$("#e").textContent = er.message; btn.disabled = false; }
    };
  }

  // ================= INVOICES (v1.3) =================
  const INV_MIME = ["application/pdf", "image/jpeg", "image/png", "image/webp", "image/heic", "image/heif", "image/gif"];
  const MAX_UPLOAD = 4.4 * 1024 * 1024;
  const invState = { filter: "all", q: "", month: "", sel: new Set() };
  const invList = () => (doc().invoices || []);
  const expById = (id) => doc().expenses.find((r) => r.id === id);
  const fmtD = (iso) => (iso ? new Date(iso + "T12:00:00").toLocaleDateString("he-IL") : "");

  function linkedInvoicesHtml(rowId) {
    const list = invList().filter((i) => i.expenseId === rowId);
    if (!list.length) return "";
    return `<div class="preview-box"><b>חשבוניות מקושרות:</b> ${list.map((i) => `<button type="button" class="icon-btn" data-inv-view="${i.id}">📎 ${esc(i.supplier || i.fileName)} ${i.invoiceNo ? "· " + esc(i.invoiceNo) : ""}</button>`).join(" ")}</div>`;
  }
  document.addEventListener("click", (e) => {
    const b = e.target.closest("[data-inv-view]");
    if (b) { e.preventDefault(); viewInvoiceFile(b.dataset.invView); }
  });

  // ---------- files ----------
  const blobToB64 = (blob) => new Promise((res, rej) => { const r = new FileReader(); r.onload = () => res(String(r.result).split(",")[1]); r.onerror = rej; r.readAsDataURL(blob); });
  async function compressImage(file) {
    try {
      const bmp = await createImageBitmap(file);
      const scale = Math.min(1, 2200 / Math.max(bmp.width, bmp.height));
      const c = document.createElement("canvas");
      c.width = Math.round(bmp.width * scale); c.height = Math.round(bmp.height * scale);
      c.getContext("2d").drawImage(bmp, 0, 0, c.width, c.height);
      const blob = await new Promise((r) => c.toBlob(r, "image/jpeg", 0.85));
      return blob ? new File([blob], file.name.replace(/\.[^.]+$/, "") + ".jpg", { type: "image/jpeg" }) : file;
    } catch { return file; } // e.g. HEIC on browsers that can't decode it
  }
  async function fetchInvoiceBlob(id, year = S.year) {
    const res = await fetch(`/api/years/${year}/invoices/${id}/file`, { headers: { authorization: "Bearer " + S.token } });
    if (!res.ok) throw new Error("הורדת הקובץ נכשלה");
    return res.blob();
  }
  async function viewInvoiceFile(id) {
    const w = window.open("", "_blank");
    try {
      const blob = await fetchInvoiceBlob(id);
      const url = URL.createObjectURL(blob);
      if (w) w.location = url; else window.location = url;
    } catch (e) { if (w) w.close(); toast(e.message, true); }
  }
  const safeName = (s) => String(s || "").replace(/[\\/:*?"<>|\n\r\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
  function invFileName(i) {
    const ext = i.mime === "application/pdf" ? "pdf" : (i.mime.split("/")[1] || "jpg").replace("jpeg", "jpg");
    return [i.date || "ללא-תאריך", safeName(i.supplier) || "ספק", safeName(i.invoiceNo)].filter(Boolean).join("_") + "." + ext;
  }

  // ---------- lazy loaders (OCR / PDF / ZIP) ----------
  const loadScript = (src) => new Promise((res, rej) => {
    if (document.querySelector(`script[src="${src}"]`)) return res();
    const s = document.createElement("script"); s.src = src; s.onload = res; s.onerror = () => rej(new Error("טעינת רכיב נכשלה: " + src));
    document.head.appendChild(s);
  });
  let _pdfjs = null, _ocr = null;
  async function getPdfjs() {
    if (!_pdfjs) {
      _pdfjs = await import("/vendor/pdfjs/pdf.min.mjs");
      _pdfjs.GlobalWorkerOptions.workerSrc = "/vendor/pdfjs/pdf.worker.min.mjs";
    }
    return _pdfjs;
  }
  async function getOcr() {
    if (!_ocr) {
      await loadScript("/vendor/tesseract/tesseract.min.js");
      _ocr = Tesseract.createWorker("heb+eng", 1, {
        workerPath: "/vendor/tesseract/worker.min.js", corePath: "/vendor/tesseract/", langPath: "/vendor/tesseract/lang", gzip: true, workerBlobURL: false,
      });
    }
    return _ocr;
  }
  async function ocrImage(src) {
    const w = await getOcr();
    const { data } = await w.recognize(src);
    return data.text || "";
  }
  async function extractText(file, onStep) {
    if (file.type === "application/pdf") {
      onStep && onStep("קורא את ה-PDF…");
      const pdfjs = await getPdfjs();
      const pdf = await pdfjs.getDocument({ data: new Uint8Array(await file.arrayBuffer()), cMapUrl: "/vendor/pdfjs/cmaps/", cMapPacked: true }).promise;
      let text = "";
      for (let p = 1; p <= Math.min(pdf.numPages, 2); p++) {
        const page = await pdf.getPage(p);
        const tc = await page.getTextContent();
        const lines = new Map();
        for (const it of tc.items) {
          if (!it.str || !it.str.trim()) continue;
          const y = Math.round(it.transform[5] / 3) * 3;
          if (!lines.has(y)) lines.set(y, []);
          lines.get(y).push({ x: it.transform[4], s: it.str });
        }
        text += [...lines.entries()].sort((a, b) => b[0] - a[0]).map(([, its]) => its.sort((a, b) => b.x - a.x).map((i) => i.s).join(" ")).join("\n") + "\n";
      }
      if (text.replace(/\s/g, "").length >= 40) return { text, method: "טקסט מתוך ה-PDF" };
      onStep && onStep("סורק את המסמך (OCR)…");
      const page = await pdf.getPage(1);
      const vp = page.getViewport({ scale: 2.2 });
      const c = document.createElement("canvas"); c.width = vp.width; c.height = vp.height;
      await page.render({ canvasContext: c.getContext("2d"), viewport: vp }).promise;
      return { text: await ocrImage(c), method: "OCR של PDF סרוק" };
    }
    onStep && onStep("סורק את התמונה (OCR)…");
    return { text: await ocrImage(file), method: "OCR של תמונה" };
  }

  // ---------- field parsing (heuristics; the user always confirms) ----------
  const rev = (s) => [...s].reverse().join("");
  function parseInvoice(raw, fileName) {
    const text = raw.replace(/[״”“]/g, '"').replace(/[׳’]/g, "'").replace(/[ \t]+/g, " ");
    const lines = text.split(/\n/).map((l) => l.trim()).filter(Boolean);
    const has = (l, kws) => kws.some((k) => l.includes(k) || l.includes(rev(k)));
    const NUM = /(?<![\d.,])(\d{1,3}(?:,\d{3})+|\d+)(?:\.(\d{1,2}))?(?![\d])/g;
    const nums = (l) => [...l.replace(/(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{2,4})/g, " ").matchAll(NUM)].map((m) => ({ v: parseFloat(m[1].replace(/,/g, "") + (m[2] ? "." + m[2] : "")), dec: !!m[2] })).filter((n) => n.v > 0 && n.v < 1e7);
    // amount
    let amount = null;
    const tiers = [['לתשלום', 'total due', 'amount due', 'סה"כ לתשלום'], ['כולל מע"מ', 'כולל מעמ', 'total incl', 'grand total'], ['סה"כ', 'סהכ', 'total']];
    for (const kws of tiers) {
      for (let i = 0; i < lines.length && amount === null; i++) {
        if (!has(lines[i].toLowerCase(), kws)) continue;
        let c = nums(lines[i]);
        if (!c.length && lines[i + 1]) c = nums(lines[i + 1]);
        if (c.length) amount = Math.max(...c.map((n) => n.v));
      }
      if (amount !== null) break;
    }
    if (amount === null) {
      const all = lines.flatMap(nums).filter((n) => n.dec);
      if (all.length) amount = Math.max(...all.map((n) => n.v));
    }
    // date
    let date = null;
    const DATE = /(\d{1,2})[\/.\-](\d{1,2})[\/.\-](\d{4}|\d{2})(?!\d)/g;
    const ordered = [...lines.filter((l) => has(l, ["תאריך", "date"])), ...lines];
    outer: for (const l of ordered) {
      for (const m of l.matchAll(DATE)) {
        let [d, mo, y] = [+m[1], +m[2], +m[3]];
        if (y < 100) y += 2000;
        if (d >= 1 && d <= 31 && mo >= 1 && mo <= 12 && y >= 2015 && y <= 2035) { date = `${y}-${String(mo).padStart(2, "0")}-${String(d).padStart(2, "0")}`; break outer; }
      }
    }
    // invoice number
    let invoiceNo = null;
    for (const l of lines) {
      if (!has(l, ["חשבונית", "חשבון עסקה", "קבלה", "invoice", "receipt"])) continue;
      const cand = [...l.replace(DATE, " ").matchAll(/(?<![\d.,])(\d{2,12})(?![\d.,])/g)].map((m) => m[1]).filter((n) => n.length !== 9 || !has(l, ["עוסק", "ח.פ", "ע.מ"]));
      if (cand.length) { invoiceNo = cand[0]; break; }
    }
    // supplier: a company line near the top, else the file name
    let supplier = null;
    const skip = ["חשבונית", "תאריך", "לכבוד", "הבודקים", "ITCB", "עמוד", "מקור", "העתק", "invoice", "page", "טלפון", "פקס", "@", "www"];
    const top = lines.slice(0, 12);
    const company = top.find((l) => has(l, ['בע"מ', "בעמ", "ltd", "inc", "עמותה", "שותפות"]) && !has(l, skip));
    const firstText = top.find((l) => /[א-תA-Za-z]{2,}/.test(l) && l.length <= 40 && !has(l, skip) && !/^\d/.test(l));
    supplier = (company || firstText || "").replace(/\s+/g, " ").trim().slice(0, 60) || null;
    if (!supplier && fileName) supplier = fileName.replace(/\.[^.]+$/, "").replace(/[_-]+/g, " ").slice(0, 60);
    return { supplier, invoiceNo, date, amount };
  }

  // ---------- match suggestions ----------
  function suggestRows(meta, docText = "") {
    const amt = num(meta.amount), mo = meta.date ? +meta.date.slice(5, 7) : null;
    const words = String(meta.supplier || "").toLowerCase().split(/[\s,."'()-]+/).filter((w) => w.length >= 2);
    const textWords = new Set(String(docText).toLowerCase().split(/[^0-9a-zא-ת]+/).filter((w) => w.length >= 3 && !/^\d+$/.test(w)));
    const rows = doc().expenses.filter((r) => r.status !== DONE && r.status !== NOTHAPPEN);
    const scored = rows.map((r) => {
      let sc = 0;
      const ra = num(r.amount);
      if (amt && ra) { const d = Math.abs(ra - amt) / Math.max(amt, 1); sc += d < 0.005 ? 50 : d < 0.05 ? 30 : d < 0.2 ? 10 : 0; }
      if (mo && isNum(r.month)) { const dm = Math.abs(num(r.month) - mo); sc += dm === 0 ? 15 : dm === 1 ? 5 : 0; }
      const hay = `${str(r.detail)} ${str(r.classification)} ${str(r.category)}`.toLowerCase();
      sc += Math.min(40, words.filter((w) => hay.includes(w)).length * 20);
      if (textWords.size) sc += Math.min(25, hay.split(/[^0-9a-zא-ת]+/).filter((w) => w.length >= 3 && textWords.has(w)).length * 8);
      return { r, sc };
    }).sort((a, b) => b.sc - a.sc);
    let top = scored.filter((x) => x.sc > 0).slice(0, 6);
    if (!top.length && amt) top = scored.slice().sort((a, b) => Math.abs(num(a.r.amount) - amt) - Math.abs(num(b.r.amount) - amt)).slice(0, 3);
    return top;
  }

  // ---------- review drawer: new upload (queue) or existing invoice ----------
  async function openInvoiceReview(item, rest = []) {
    // item = { file } for a new upload, or { inv } for an existing invoice
    const isNew = !!item.file;
    const inv = item.inv || {};
    let file = item.file, meta = { supplier: inv.supplier || "", invoiceNo: inv.invoiceNo || "", date: inv.date || "", amount: inv.amount ?? "", detail: inv.detail || "", notes: inv.notes || "" };
    const ov = document.createElement("div");
    ov.className = "overlay";
    ov.innerHTML = `<div class="drawer inv-drawer" role="dialog" aria-label="חשבונית">
      <div class="drawer-head"><h3>${isNew ? "חשבונית חדשה" : "עריכת חשבונית"}${rest.length ? ` <span class="muted" style="font-weight:500">(עוד ${rest.length} בתור)</span>` : ""}</h3><button class="btn ghost" data-x>✕</button></div>
      <div class="drawer-body inv-body">
        <div class="inv-preview" id="invPrev"><div class="muted">טוען תצוגה…</div></div>
        <div class="inv-form">
          <div class="note" id="ocrNote">${isNew ? "מזהה את פרטי החשבונית…" : "אפשר לעדכן פרטים ולשנות שיוך."}</div>
          <div class="grid2">
            <div class="field"><label for="iv_supplier">ספק</label><input class="input" id="iv_supplier"></div>
            <div class="field"><label for="iv_invoiceNo">מס׳ חשבונית</label><input class="input" id="iv_invoiceNo" dir="ltr"></div>
            <div class="field"><label for="iv_date">תאריך חשבונית</label><input class="input" id="iv_date" type="date"></div>
            <div class="field"><label for="iv_amount">סכום כולל (₪)</label><input class="input" id="iv_amount" type="number" step="any" dir="ltr"></div>
            <div class="field" style="grid-column:1/-1"><label for="iv_notes">הערות</label><input class="input" id="iv_notes"></div>
          </div>
          <h4 style="margin:14px 0 6px">שיוך להוצאה בדוח</h4>
          <div class="radio-list" id="matchList"></div>
          <div class="err" id="ivErr"></div>
        </div>
      </div>
      <div class="drawer-foot">
        <button class="btn primary" id="ivSave">${isNew ? "שמירה" : "שמירת שינויים"}</button>
        ${isNew && rest.length ? `<button class="btn" id="ivSkip">דילוג</button>` : ""}
        <button class="btn" data-x>ביטול</button><div class="spacer"></div>
        ${!isNew ? `<button class="btn danger" id="ivDel">מחיקה</button>` : ""}
      </div></div>`;
    document.body.appendChild(ov); document.body.classList.add("no-scroll");
    let previewUrl = null;
    const close = () => { ov.remove(); document.body.classList.remove("no-scroll"); if (previewUrl) URL.revokeObjectURL(previewUrl); };
    ov.querySelectorAll("[data-x]").forEach((b) => (b.onclick = close));
    const $ = (s) => ov.querySelector(s);
    const setFields = () => { for (const k of ["supplier", "invoiceNo", "date", "amount", "notes"]) $("#iv_" + k).value = meta[k] ?? ""; };
    const readFields = () => ({ supplier: $("#iv_supplier").value.trim() || null, invoiceNo: $("#iv_invoiceNo").value.trim() || null, date: $("#iv_date").value || null, amount: $("#iv_amount").value === "" ? null : num($("#iv_amount").value), notes: $("#iv_notes").value.trim() || null });
    setFields();

    // matching list
    const cats = distinct(doc().expenses, "category").filter(Boolean), classes = distinct(doc().expenses, "classification").filter(Boolean);
    const linked = inv.expenseId ? expById(inv.expenseId) : null;
    let userPicked = false;
    const drawMatches = () => {
      const m = readFields();
      const sug = suggestRows(m, ov.dataset.ocrText || "").filter((x) => !linked || x.r.id !== linked.id);
      const cur = userPicked ? $("#matchList").querySelector("input:checked")?.value : null;
      const opt = (val, title, sub, checked) => `<label><input type="radio" name="match" value="${esc(val)}" ${checked ? "checked" : ""}><div><b>${title}</b><small>${sub}</small></div></label>`;
      const rowSub = (r) => `${esc(r.category || "")}${r.classification ? " › " + esc(r.classification) : ""} · חודש ${esc(r.month ?? "—")} · ${shekel(num(r.amount))} · ${esc(r.status || "")}`;
      let html = "";
      if (linked) html += opt("keep", `משויכת כעת: ${esc(linked.detail || "שורת הוצאה")}`, rowSub(linked), !cur || cur === "keep");
      html += sug.map((x, i) => opt("row:" + x.r.id, esc(x.r.detail || "(ללא פירוט)"), rowSub(x.r) + (i === 0 && x.sc >= 50 && (!sug[1] || sug[1].sc < x.sc) ? ' · <b style="color:var(--done)">ההתאמה הטובה ביותר</b>' : ""), cur ? cur === "row:" + x.r.id : !linked && i === 0 && x.sc >= 45)).join("");
      html += opt("new", "שורת הוצאה חדשה", `<span class="new-row-fields"><input class="input" id="nr_cat" list="nr_cats" placeholder="קטגוריה *"><input class="input" id="nr_cls" list="nr_clss" placeholder="סיווג"><input class="input" id="nr_det" placeholder="פירוט"></span>
        <datalist id="nr_cats">${cats.map((c) => `<option value="${esc(c)}">`).join("")}</datalist><datalist id="nr_clss">${classes.map((c) => `<option value="${esc(c)}">`).join("")}</datalist>`, cur === "new" || (!linked && !sug.some((x) => x.sc >= 45) && !cur && false));
      html += opt(linked ? "unlink" : "none", linked ? "ביטול השיוך" : "בלי שיוך כרגע", linked ? "החשבונית תישאר בספרייה; שורת ההוצאה לא משתנה" : "אפשר לשייך מאוחר יותר", cur ? cur === (linked ? "unlink" : "none") : !linked && !sug.some((x) => x.sc >= 45));
      $("#matchList").innerHTML = html;
      const diff = () => {
        const v = $("#matchList").querySelector("input:checked")?.value || "";
        $("#matchList").querySelectorAll(".diff").forEach((d) => d.remove());
        if (!v.startsWith("row:")) return;
        const r = expById(v.slice(4)), f = readFields(), lab = $("#matchList").querySelector("input:checked").closest("label");
        const ch = [`סטטוס: ${esc(r.status || "—")} ← <b>בוצע</b>`];
        if (f.amount !== null && num(r.amount) !== f.amount) ch.push(`סכום: ${shekel(num(r.amount))} ← <b>${shekel(f.amount)}</b>`);
        if (f.invoiceNo) ch.push(`אסמכתא: ${esc(r.ref || "—")} ← <b dir="ltr">${esc(f.invoiceNo)}</b>`);
        if (f.date && +f.date.slice(5, 7) !== num(r.month)) ch.push(`חודש: ${esc(r.month ?? "—")} ← <b>${+f.date.slice(5, 7)}</b>`);
        lab.querySelector("div").insertAdjacentHTML("beforeend", `<small class="diff">יעודכן: ${ch.join(" · ")}</small>`);
      };
      $("#matchList").querySelectorAll("input[name=match]").forEach((r) => (r.onchange = () => { userPicked = true; diff(); }));
      $("#matchList").querySelectorAll(".new-row-fields input").forEach((i) => (i.onfocus = () => { userPicked = true; $("#matchList").querySelector("input[value=new]").checked = true; diff(); }));
      diff();
    };
    drawMatches();
    ["#iv_amount", "#iv_supplier", "#iv_date"].forEach((s) => ($(s).onchange = drawMatches));
    if (!canEdit()) { ov.querySelectorAll("input,button#ivSave,#ivDel").forEach((e) => (e.disabled = true)); $("#ocrNote").textContent = "צפייה בלבד"; }

    // preview + OCR
    (async () => {
      try {
        let blob = file;
        if (!isNew) blob = await fetchInvoiceBlob(inv.id);
        previewUrl = URL.createObjectURL(blob);
        const isPdf = (blob.type || inv.mime) === "application/pdf";
        $("#invPrev").innerHTML = isPdf
          ? `<iframe src="${previewUrl}" title="תצוגת החשבונית"></iframe><a class="btn sm" href="${previewUrl}" target="_blank" rel="noopener">פתיחה בחלון חדש</a>`
          : `<a href="${previewUrl}" target="_blank" rel="noopener"><img src="${previewUrl}" alt="תצוגת החשבונית"></a>`;
        if (!isNew) return;
        const t0 = Date.now();
        const { text, method } = await extractText(file, (s) => ($("#ocrNote").textContent = s));
        const p = parseInvoice(text, item.origName || file.name);
        const found = [];
        for (const k of ["supplier", "invoiceNo", "date", "amount"]) if (p[k] !== null && p[k] !== undefined && p[k] !== "") { meta[k] = p[k]; found.push(k); }
        ov.dataset.ocrText = text;
        setFields(); drawMatches();
        const names = { supplier: "ספק", invoiceNo: "מס׳ חשבונית", date: "תאריך", amount: "סכום" };
        const missing = Object.keys(names).filter((k) => !found.includes(k));
        $("#ocrNote").innerHTML = `זוהו אוטומטית (${esc(method)}, ${Math.round((Date.now() - t0) / 100) / 10} שנ׳): ${found.map((k) => names[k]).join(", ") || "—"}.${missing.length ? ` <b>יש להשלים: ${missing.map((k) => names[k]).join(", ")}.</b>` : ""} נא לבדוק את הפרטים מול המסמך לפני השמירה.`;
        ov.dataset.ocrText = text;
      } catch (e) {
        $("#ocrNote").innerHTML = `לא הצלחתי לזהות את הפרטים אוטומטית (${esc(e.message)}). אפשר למלא ידנית.`;
      }
    })();

    const next = () => { close(); if (rest.length) openInvoiceReview(rest[0], rest.slice(1)); };
    const skip = $("#ivSkip"); if (skip) skip.onclick = next;
    const del = $("#ivDel"); if (del) del.onclick = () => { close(); confirmDeleteInvoice(inv); };
    $("#ivSave").onclick = async () => {
      const f = readFields();
      const choice = $("#matchList").querySelector("input:checked")?.value || "none";
      let apply = null;
      if (choice.startsWith("row:")) apply = { expenseId: choice.slice(4) };
      if (choice === "new") {
        const cat = $("#nr_cat").value.trim();
        if (!cat) return ($("#ivErr").textContent = "לשורה חדשה יש לבחור קטגוריה");
        apply = { newRow: { category: cat, classification: $("#nr_cls").value.trim() || null, detail: $("#nr_det").value.trim() || f.supplier } };
      }
      if (apply && f.amount === null) return ($("#ivErr").textContent = "כדי לשייך להוצאה צריך סכום");
      const btn = $("#ivSave"); btn.disabled = true; btn.textContent = "שומר…";
      try {
        let r;
        if (isNew) {
          r = await api(`years/${S.year}/invoices`, { method: "POST", body: { file: await blobToB64(file), fileName: item.origName || file.name, mime: file.type, meta: f, apply } });
        } else {
          r = await api(`years/${S.year}/invoices/${inv.id}`, { method: "PUT", body: { meta: f, apply, unlink: choice === "unlink" } });
        }
        S.docs[S.year] = r.doc;
        toast(apply ? "החשבונית נשמרה והדוח עודכן" : "החשבונית נשמרה");
        renderMain();
        next();
      } catch (e) { $("#ivErr").textContent = e.message; btn.disabled = false; btn.textContent = isNew ? "שמירה" : "שמירת שינויים"; }
    };
  }

  async function prepareInvoiceFile(f0) {
    let f = f0;
    if (!INV_MIME.includes(f.type)) return { error: `${f0.name}: סוג קובץ לא נתמך – PDF או תמונה בלבד` };
    if (f.type.startsWith("image/") && f.type !== "image/gif") f = await compressImage(f);
    if (f.size > MAX_UPLOAD) return { error: `${f0.name}: הקובץ גדול מ-4.4MB` };
    return { file: f, origName: f0.name };
  }
  async function startUpload(files) {
    const items = [];
    const bad = [];
    for (const f0 of files) {
      const p = await prepareInvoiceFile(f0);
      if (p.error) bad.push(p.error); else items.push(p);
    }
    if (bad.length) toast(bad.join(" · "), true);
    if (items.length) openInvoiceReview(items[0], items.slice(1));
  }

  function confirmDeleteInvoice(inv) {
    const m = modal("מחיקת חשבונית", `<p>למחוק את החשבונית?</p><div class="preview-box"><b>${esc(inv.supplier || inv.fileName)}</b> ${inv.invoiceNo ? "· " + esc(inv.invoiceNo) : ""}<br>${inv.amount != null ? shekel(num(inv.amount)) : ""} · ${fmtD(inv.date)}</div><p class="muted">הקובץ יימחק. שורת ההוצאה המשויכת (אם יש) לא תשתנה.</p><div class="err" id="e"></div>`,
      `<button class="btn primary" style="background:var(--danger);border-color:var(--danger)" id="ok">מחיקה</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    m.$("#ok").onclick = async () => {
      try { const r = await api(`years/${S.year}/invoices/${inv.id}`, { method: "DELETE" }); S.docs[S.year] = r.doc; m.close(); renderMain(); toast("החשבונית נמחקה"); }
      catch (e) { m.$("#e").textContent = e.message; }
    };
  }

  // ---------- the invoices tab ----------
  function filteredInvoices() {
    const q = invState.q.trim().toLowerCase();
    return invList().filter((i) => {
      if (invState.filter === "unlinked" && i.expenseId && expById(i.expenseId)) return false;
      if (invState.filter === "unsent" && i.sentAt) return false;
      if (invState.month && (!i.date || +i.date.slice(5, 7) !== +invState.month)) return false;
      if (q && ![i.supplier, i.invoiceNo, i.notes, i.fileName, expById(i.expenseId)?.detail].some((v) => str(v).toLowerCase().includes(q))) return false;
      return true;
    }).sort((a, b) => str(b.date || b.uploadedAt).localeCompare(str(a.date || a.uploadedAt)));
  }
  function renderInvoices($m) {
    const all = invList();
    const rows = filteredInvoices();
    for (const id of [...invState.sel]) if (!all.some((i) => i.id === id)) invState.sel.delete(id);
    const unlinked = all.filter((i) => !i.expenseId || !expById(i.expenseId));
    const unsent = all.filter((i) => !i.sentAt);
    const months = [...new Set(all.filter((i) => i.date).map((i) => +i.date.slice(5, 7)))].sort((a, b) => a - b);
    $m.innerHTML = `${pageHead("חשבוניות", "ספריית חשבוניות ההוצאה: העלאה, זיהוי אוטומטי, שיוך לדוח ושליחה לרואה החשבון",
      `${canEdit() ? `<button class="btn accent" id="invUp">+ העלאת חשבוניות</button>` : ""}<button class="btn primary" id="invSend" ${all.length ? "" : "disabled"}>שליחה לרואה חשבון</button>`)}
      <input type="file" id="invFile" accept="application/pdf,image/*" multiple hidden>
      <div class="kpis k4">
        <div class="kpi c-n"><div class="lbl">חשבוניות בספרייה</div><div class="val">${all.length}</div><div class="sub">${shekel(sum(all))}</div></div>
        <div class="kpi c-done"><div class="lbl">משויכות לדוח</div><div class="val">${all.length - unlinked.length}</div><div class="sub">${shekel(sum(all) - sum(unlinked))}</div></div>
        <div class="kpi c-exp"><div class="lbl">לא משויכות</div><div class="val">${unlinked.length}</div><div class="sub">${shekel(sum(unlinked))}</div></div>
        <div class="kpi c-fc"><div class="lbl">טרם נשלחו לרו״ח</div><div class="val">${unsent.length}</div><div class="sub">${shekel(sum(unsent))}</div></div>
      </div>
      ${canEdit() ? `<label class="file-drop" id="invDrop" for="invFile"><b>גרירת קבצים לכאן או לחיצה לבחירה</b><br><span class="muted">PDF או צילום של החשבונית · אפשר כמה קבצים יחד · הפרטים מזוהים אוטומטית ותתבקש לאשר</span></label>` : ""}
      <div class="toolbar">
        <input class="input search" id="invQ" type="search" placeholder="חיפוש ספק, מס׳ חשבונית, פירוט…" value="${esc(invState.q)}">
        ${segCtl("invF", [["all", "הכל"], ["unlinked", `לא משויכות (${unlinked.length})`], ["unsent", `לא נשלחו (${unsent.length})`]], invState.filter)}
        <select class="input" id="invMonth" style="width:auto"><option value="">כל החודשים</option>${months.map((mo) => `<option value="${mo}" ${String(mo) === String(invState.month) ? "selected" : ""}>${MONTHS[mo]}</option>`).join("")}</select>
        <span class="spacer"></span>
        <div class="user-chip" style="padding:0;border:0;background:none" id="invExpWrap"><button class="btn" id="invExpBtn" ${all.length ? "" : "disabled"}>ייצוא ▾</button>
          <div class="menu hidden" id="invExpMenu" style="left:0;right:auto">
            <button data-ie="zip">ZIP: חשבוניות + דוח (${invState.sel.size ? "המסומנות" : "בתצוגה"})</button>
            <button data-ie="zipOnly">ZIP: חשבוניות בלבד</button>
            <button data-ie="xlsx">אקסל: רשימת חשבוניות</button>
          </div></div>
      </div>
      <div class="meta-line" style="margin:-4px 0 8px">${invState.sel.size ? `${invState.sel.size} מסומנות · <a href="#" id="invClr">ניקוי סימון</a>` : `${rows.length} חשבוניות בתצוגה`}</div>
      <div class="dt-wrap"><table class="dt"><thead><tr>
        <th style="width:32px"><input type="checkbox" id="invAll" aria-label="סימון הכל" ${rows.length && rows.every((i) => invState.sel.has(i.id)) ? "checked" : ""}></th>
        <th>תאריך</th><th>ספק</th><th>מס׳ חשבונית</th><th>סכום</th><th>שיוך בדוח</th><th>נשלח לרו״ח</th><th class="col-actions"></th></tr></thead>
        <tbody>${rows.length ? rows.map((i) => {
          const ex = i.expenseId ? expById(i.expenseId) : null;
          return `<tr data-id="${i.id}"><td><input type="checkbox" data-sel="${i.id}" ${invState.sel.has(i.id) ? "checked" : ""} aria-label="סימון"></td>
            <td class="numc">${fmtD(i.date) || '<span class="muted">—</span>'}</td>
            <td><b>${esc(i.supplier || "")}</b><div class="audit">${esc(i.fileName)}</div></td>
            <td class="numc">${esc(i.invoiceNo || "")}</td>
            <td class="money">${i.amount != null ? shekel(num(i.amount)) : '<span class="muted">—</span>'}</td>
            <td>${ex ? `<span class="badge b-done">משויכת</span> ${esc(ex.detail || "")}<div class="audit">${esc(ex.category || "")}${ex.classification ? " › " + esc(ex.classification) : ""}</div>` : '<span class="badge b-fc">לא משויכת</span>'}</td>
            <td>${i.sentAt ? fmtDate(i.sentAt) : '<span class="muted">—</span>'}</td>
            <td class="col-actions"><div class="row-actions"><button class="icon-btn" data-iv="${i.id}">צפייה</button>${canEdit() ? `<button class="icon-btn" data-ied="${i.id}">${ex ? "עריכה" : "שיוך"}</button><button class="icon-btn del" data-idel="${i.id}">מחיקה</button>` : ""}</div></td></tr>`;
        }).join("") : `<tr><td colspan="8"><div class="empty">${all.length ? "אין חשבוניות התואמות לסינון" : "עדיין לא הועלו חשבוניות לשנה זו"}</div></td></tr>`}</tbody>
        ${rows.length ? `<tfoot><tr><td></td><td colspan="3">סה״כ (${rows.length})</td><td class="money">${shekel(sum(rows))}</td><td colspan="3"></td></tr></tfoot>` : ""}</table></div>
      <div id="sharesBox"></div>`;
    bindPageHead();
    const rer = () => renderInvoices($m);
    const fileIn = document.getElementById("invFile");
    const up = document.getElementById("invUp"); if (up) up.onclick = () => fileIn.click();
    fileIn.onchange = () => { startUpload([...fileIn.files]); fileIn.value = ""; };
    const drop = document.getElementById("invDrop");
    if (drop) {
      drop.ondragover = (e) => { e.preventDefault(); drop.classList.add("over"); };
      drop.ondragleave = () => drop.classList.remove("over");
      drop.ondrop = (e) => { e.preventDefault(); drop.classList.remove("over"); startUpload([...e.dataTransfer.files]); };
    }
    document.getElementById("invSend").onclick = () => openShareModal();
    const q = document.getElementById("invQ");
    q.oninput = () => { invState.q = q.value; clearTimeout(q._t); q._t = setTimeout(() => { rer(); const n = document.getElementById("invQ"); n.focus(); n.setSelectionRange(n.value.length, n.value.length); }, 250); };
    bindSeg($m, "invF", (v) => { invState.filter = v; rer(); });
    document.getElementById("invMonth").onchange = (e) => { invState.month = e.target.value; rer(); };
    document.getElementById("invAll").onchange = (e) => { rows.forEach((i) => (e.target.checked ? invState.sel.add(i.id) : invState.sel.delete(i.id))); rer(); };
    $m.querySelectorAll("[data-sel]").forEach((c) => (c.onchange = () => { c.checked ? invState.sel.add(c.dataset.sel) : invState.sel.delete(c.dataset.sel); rer(); }));
    const clr = document.getElementById("invClr"); if (clr) clr.onclick = (e) => { e.preventDefault(); invState.sel.clear(); rer(); };
    $m.querySelectorAll("[data-iv]").forEach((b) => (b.onclick = () => viewInvoiceFile(b.dataset.iv)));
    $m.querySelectorAll("[data-ied]").forEach((b) => (b.onclick = () => openInvoiceReview({ inv: all.find((i) => i.id === b.dataset.ied) })));
    $m.querySelectorAll("[data-idel]").forEach((b) => (b.onclick = () => confirmDeleteInvoice(all.find((i) => i.id === b.dataset.idel))));
    $m.querySelectorAll("tbody tr[data-id]").forEach((tr) => (tr.ondblclick = () => openInvoiceReview({ inv: all.find((i) => i.id === tr.dataset.id) })));
    const eb = document.getElementById("invExpBtn"), em = document.getElementById("invExpMenu");
    eb.onclick = (e) => { e.stopPropagation(); em.classList.toggle("hidden"); };
    document.addEventListener("click", () => em.classList.add("hidden"), { once: true });
    const pick = () => (invState.sel.size ? all.filter((i) => invState.sel.has(i.id)) : rows);
    em.querySelectorAll("[data-ie]").forEach((b) => (b.onclick = (e) => {
      e.stopPropagation(); em.classList.add("hidden");
      const list = pick();
      if (b.dataset.ie === "xlsx") return exportInvoiceRegister(list);
      downloadZip(list, b.dataset.ie === "zip");
    }));
    renderShares();
  }

  function invoiceRegisterAoa(d, list) {
    const ex = (id) => d.expenses.find((r) => r.id === id);
    return [["תאריך", "ספק", "מס׳ חשבונית", 'סכום בש"ח', "פירוט ההוצאה בדוח", "קטגוריה", "סיווג", "נשלח לרו״ח", "שם הקובץ"],
      ...list.map((i) => { const r = ex(i.expenseId); return [i.date ? new Date(i.date + "T12:00:00") : "", i.supplier || "", i.invoiceNo || "", i.amount ?? "", r ? r.detail || "" : "לא משויכת", r ? r.category || "" : "", r ? r.classification || "" : "", i.sentAt ? new Date(i.sentAt) : "", invFileName(i)]; }),
      ["סה״כ", "", "", list.reduce((a, i) => a + num(i.amount), 0)]];
  }
  function exportInvoiceRegister(list) {
    if (!window.XLSX) return toast("רכיב האקסל עדיין נטען, נסה שוב", true);
    const wb = XLSX.utils.book_new(); wb.Workbook = { Views: [{ RTL: true }] };
    const ws = XLSX.utils.aoa_to_sheet(invoiceRegisterAoa(doc(), list), { cellDates: true });
    ws["!cols"] = [12, 26, 14, 12, 40, 18, 22, 14, 34].map((w) => ({ wch: w }));
    XLSX.utils.book_append_sheet(wb, ws, "חשבוניות");
    XLSX.writeFile(wb, `ITCB_Invoices_${S.year}.xlsx`);
    toast("הקובץ הורד");
  }
  async function downloadZip(list, withReport) {
    if (!list.length) return toast("אין חשבוניות לייצוא", true);
    toast(`מכין ZIP של ${list.length} חשבוניות…`);
    try {
      await loadScript("/vendor/fflate.min.js");
      const files = {};
      const used = new Set();
      for (const i of list) {
        let n = invFileName(i), k = 2;
        while (used.has(n)) n = invFileName(i).replace(/(\.[^.]+)$/, `_${k++}$1`);
        used.add(n);
        files["חשבוניות/" + n] = new Uint8Array(await (await fetchInvoiceBlob(i.id)).arrayBuffer());
      }
      if (withReport) files[`ITCB_Budget_${S.year}.xlsx`] = new Uint8Array(workbookBytes(doc(), list));
      const zipped = fflate.zipSync(files, { level: 0 });
      downloadBlob(new Blob([zipped], { type: "application/zip" }), `ITCB_Invoices_${S.year}.zip`);
      toast("קובץ ה-ZIP הורד");
    } catch (e) { toast(e.message, true); }
  }

  // ---------- send to accountant ----------
  function openShareModal() {
    const all = invList(), cfg = S.config || {};
    const sel = all.filter((i) => invState.sel.has(i.id)), unsent = all.filter((i) => !i.sentAt);
    const months = [...new Set(all.filter((i) => i.date).map((i) => +i.date.slice(5, 7)))].sort((a, b) => a - b);
    const ed = canEdit();
    const m = modal("שליחה לרואה החשבון", `
      <div class="grid2">
        <div class="field"><label for="acName">שם רואה החשבון</label><input class="input" id="acName" value="${esc(cfg.accountantName || "")}" ${ed ? "" : "disabled"}></div>
        <div class="field"><label for="acMail">מייל</label><input class="input" id="acMail" type="email" dir="ltr" value="${esc(cfg.accountantEmail || "")}" ${ed ? "" : "disabled"}></div>
      </div>
      <b style="font-size:14px">אילו חשבוניות לשלוח?</b>
      <div class="radio-list">
        ${sel.length ? `<label><input type="radio" name="sc" value="sel" checked><div><b>המסומנות (${sel.length})</b><small>${shekel(sum(sel))}</small></div></label>` : ""}
        <label><input type="radio" name="sc" value="unsent" ${!sel.length ? "checked" : ""}><div><b>כל מה שטרם נשלח (${unsent.length})</b><small>${shekel(sum(unsent))}</small></div></label>
        <label><input type="radio" name="sc" value="all"><div><b>כל החשבוניות של ${S.year} (${all.length})</b><small>${shekel(sum(all))}</small></div></label>
        <label><input type="radio" name="sc" value="range"><div><b>לפי חודשים</b><small>מ-<select class="input" id="mFrom" style="width:auto;display:inline-block;padding:2px 6px">${months.map((x) => `<option>${x}</option>`).join("")}</select> עד <select class="input" id="mTo" style="width:auto;display:inline-block;padding:2px 6px">${months.map((x, i) => `<option ${i === months.length - 1 ? "selected" : ""}>${x}</option>`).join("")}</select> (לפי תאריך החשבונית)</small></div></label>
      </div>
      <label class="checks" style="display:flex;margin-bottom:10px"><label><input type="checkbox" id="withRep" checked>לצרף את דוח התקציב המלא (אקסל, כולל לשונית חשבוניות)</label></label>
      <div class="grid2">
        <div class="field"><label for="days">תוקף הקישור</label><select class="input" id="days"><option value="7">שבוע</option><option value="14">שבועיים</option><option value="30" selected>חודש</option><option value="60">חודשיים</option></select></div>
        <div class="field"><label for="note">הודעה לרו״ח (אופציונלי)</label><input class="input" id="note" placeholder="למשל: חשבוניות אוקטובר–נובמבר"></div>
      </div>
      <div class="muted" style="font-size:13px">נוצר קישור הורדה מאובטח (ללא צורך בסיסמה, פג תוקף אוטומטית). המייל נפתח אצלך מוכן לשליחה, והחשבוניות יסומנו כ״נשלחו״.</div>
      <div class="err" id="e"></div>`,
      `<button class="btn primary" id="ok" ${ed ? "" : "disabled"}>יצירת קישור ומייל</button><button class="btn" data-x>ביטול</button>`);
    m.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = m.close));
    if (!ed) m.$("#e").textContent = "רק עורך או מנהל יכולים לשלוח לרואה החשבון. אפשר להשתמש בייצוא ZIP.";
    m.$("#ok").onclick = async () => {
      const sc = m.ov.querySelector("input[name=sc]:checked").value;
      let list = sc === "sel" ? sel : sc === "unsent" ? unsent : all;
      let period = sc === "all" ? `שנת ${S.year}` : "";
      if (sc === "range") {
        const a = +m.$("#mFrom").value, b = +m.$("#mTo").value;
        list = all.filter((i) => i.date && +i.date.slice(5, 7) >= Math.min(a, b) && +i.date.slice(5, 7) <= Math.max(a, b));
        period = `${MONTHS[Math.min(a, b)]}–${MONTHS[Math.max(a, b)]} ${S.year}`;
      }
      const withRep = m.$("#withRep").checked;
      if (!list.length && !withRep) return (m.$("#e").textContent = "אין חשבוניות בטווח שנבחר");
      const email = m.$("#acMail").value.trim(), name = m.$("#acName").value.trim();
      const btn = m.$("#ok"); btn.disabled = true; btn.textContent = "יוצר…";
      try {
        if (email !== (cfg.accountantEmail || "") || name !== (cfg.accountantName || "")) S.config = (await api("config", { method: "PUT", body: { accountantEmail: email, accountantName: name } })).config;
        const report = withRep ? await blobToB64(new Blob([workbookBytes(doc(), list)])) : null;
        const title = `חשבוניות ITCB – ${period || S.year}`;
        const r = await api(`years/${S.year}/shares`, { method: "POST", body: { invoiceIds: list.map((i) => i.id), report, reportName: `ITCB_Budget_${S.year}.xlsx`, title, note: m.$("#note").value.trim(), days: +m.$("#days").value } });
        S.docs[S.year] = r.doc;
        const link = `${location.origin}/share.html?t=${r.token}`;
        const exp = new Date(r.share.expires).toLocaleDateString("he-IL");
        const body = `שלום${name ? " " + name : ""},\n\nמצורף קישור להורדת ${list.length} חשבוניות הוצאה של עמותת הבודקים (ITCB)${period ? " – " + period : ""}, בסך ${Math.round(sum(list)).toLocaleString("he-IL")} ש"ח${withRep ? ", יחד עם דוח התקציב המעודכן (אקסל)" : ""}.\n${m.$("#note").value.trim() ? "\n" + m.$("#note").value.trim() + "\n" : ""}\n${link}\n\nהקישור בתוקף עד ${exp}. אפשר להוריד את כל הקבצים יחד כ-ZIP.\n\nתודה,\n${S.user.name}`;
        const subject = title;
        m.close();
        const res = modal("הקישור מוכן", `<p>נוצר קישור ל-${list.length} חשבוניות${withRep ? " + דוח" : ""}, בתוקף עד ${exp}.</p>
          <div class="field"><label>קישור</label><input class="input" dir="ltr" readonly value="${esc(link)}" id="lnk"></div>
          <div class="field"><label>נוסח ההודעה</label><textarea class="input invite-text" rows="9" readonly>${esc(body)}</textarea></div>
          <div class="invite-actions">
            <a class="btn primary" id="mailBtn" href="mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}">פתיחת מייל מוכן${email ? "" : " (ללא נמען)"}</a>
            <button class="btn" id="cpL">העתקת קישור</button><button class="btn" id="cpB">העתקת הנוסח</button>
            <a class="btn wa-btn" target="_blank" rel="noopener" href="https://wa.me/?text=${encodeURIComponent(body)}">וואטסאפ</a>
          </div>`, `<button class="btn" data-x>סגירה</button>`);
        res.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = res.close));
        const cp = async (t) => { try { await navigator.clipboard.writeText(t); toast("הועתק"); } catch { toast("ההעתקה נכשלה – סמן והעתק ידנית", true); } };
        res.$("#cpL").onclick = () => cp(link); res.$("#cpB").onclick = () => cp(body);
        renderMain();
      } catch (e) { m.$("#e").textContent = e.message; btn.disabled = false; btn.textContent = "יצירת קישור ומייל"; }
    };
  }
  async function renderShares() {
    const box = document.getElementById("sharesBox");
    if (!box) return;
    try {
      const { shares } = await api(`years/${S.year}/shares`);
      if (!shares.length) return;
      box.innerHTML = `<div class="panel" style="margin-top:16px"><h2 style="font-size:16px">קישורים שנשלחו לרואה החשבון</h2>
        <table class="simple"><thead><tr><th>נוצר</th><th>כותרת</th><th>חשבוניות</th><th>בתוקף עד</th><th></th></tr></thead><tbody>
        ${shares.map((s) => { const expd = Date.parse(s.expires) < Date.now(); return `<tr><td>${fmtDate(s.createdAt)}<div class="audit">${esc(s.createdBy || "")}</div></td><td>${esc(s.title)}</td><td>${s.invoiceIds.length}${s.hasReport ? " + דוח" : ""}</td>
          <td>${expd ? '<span class="badge b-no">פג תוקף</span>' : new Date(s.expires).toLocaleDateString("he-IL")}</td>
          <td class="nowrap">${expd ? "" : `<button class="icon-btn" data-cpsh="${s.token}">העתקת קישור</button>`} ${canEdit() ? `<button class="icon-btn del" data-rmsh="${s.token}">${expd ? "הסרה" : "ביטול"}</button>` : ""}</td></tr>`; }).join("")}
        </tbody></table></div>`;
      box.querySelectorAll("[data-cpsh]").forEach((b) => (b.onclick = async () => { try { await navigator.clipboard.writeText(`${location.origin}/share.html?t=${b.dataset.cpsh}`); toast("הקישור הועתק"); } catch { toast("ההעתקה נכשלה", true); } }));
      box.querySelectorAll("[data-rmsh]").forEach((b) => (b.onclick = async () => {
        const cm = modal("ביטול קישור", `<p>לבטל את הקישור? רואה החשבון לא יוכל להוריד ממנו יותר.</p><div class="err" id="e"></div>`, `<button class="btn primary" style="background:var(--danger);border-color:var(--danger)" id="ok">ביטול הקישור</button><button class="btn" data-x>חזרה</button>`);
        cm.ov.querySelectorAll("[data-x]").forEach((x) => (x.onclick = cm.close));
        cm.$("#ok").onclick = async () => { try { await api(`years/${S.year}/shares/${b.dataset.rmsh}`, { method: "DELETE" }); cm.close(); toast("הקישור בוטל"); renderShares(); } catch (e) { cm.$("#e").textContent = e.message; } };
      }));
    } catch {}
  }

  // ---------- start ----------
  window.addEventListener("hashchange", () => {
    if (!S.user) return;
    const h = readHash();
    if (h.view === "users" && isAdmin()) { S.view = "users"; return render(); }
    if (h.year && S.years.includes(h.year) && VIEWS.some((v) => v[0] === h.view) && (h.year !== S.year || h.view !== S.view)) {
      S.year = h.year; S.view = h.view; loadYear(S.year).then(render);
    }
  });
  (async () => {
    if (S.token) return boot();
    try { const st = await api("status"); if (st.needsSetup) return renderSetup(); } catch {}
    renderLogin();
  })();
})();
