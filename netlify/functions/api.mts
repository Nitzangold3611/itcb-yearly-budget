// ITCB yearly budget – API (Netlify Functions + Blobs)
import { getStore } from "@netlify/blobs";
import type { Config, Context } from "@netlify/functions";
import { createHmac, randomBytes, scryptSync, timingSafeEqual } from "node:crypto";

const store = () => getStore({ name: "itcb-budget", consistency: "strong" });
const TOKEN_DAYS = 14;
const KINDS = ["income", "expenses"] as const;
type Kind = (typeof KINDS)[number];
type Role = "admin" | "editor" | "viewer";
interface User { email: string; name: string; role: Role; salt: string; hash: string; mustChange?: boolean; created: string; lastLogin?: string }

const json = (data: unknown, status = 200) => Response.json(data, { status, headers: { "cache-control": "no-store" } });
const fail = (msg: string, status = 400) => json({ error: msg }, status);

// ---------- secrets / tokens ----------
async function secret(): Promise<string> {
  const s = store();
  let v = await s.get("config/secret");
  if (!v) {
    const created = randomBytes(32).toString("hex");
    await s.set("config/secret", created, { onlyIfNew: true });
    v = (await s.get("config/secret")) || created;
  }
  return v;
}
const b64u = (s: string) => Buffer.from(s).toString("base64url");
async function signToken(email: string, ver: string) {
  const body = b64u(JSON.stringify({ e: email, v: ver, x: Date.now() + TOKEN_DAYS * 864e5 }));
  const sig = createHmac("sha256", await secret()).update(body).digest("base64url");
  return `${body}.${sig}`;
}
async function readToken(tok: string | null) {
  if (!tok) return null;
  const [body, sig] = tok.split(".");
  if (!body || !sig) return null;
  const good = createHmac("sha256", await secret()).update(body).digest("base64url");
  if (good.length !== sig.length || !timingSafeEqual(Buffer.from(good), Buffer.from(sig))) return null;
  try {
    const p = JSON.parse(Buffer.from(body, "base64url").toString());
    return p.x > Date.now() ? p : null;
  } catch { return null; }
}
function hashPw(pw: string, salt = randomBytes(16).toString("hex")) {
  return { salt, hash: scryptSync(pw, salt, 64).toString("hex") };
}
function checkPw(pw: string, u: User) {
  const h = scryptSync(pw, u.salt, 64);
  const ref = Buffer.from(u.hash, "hex");
  return ref.length === h.length && timingSafeEqual(ref, h);
}
const pwVer = (u: User) => u.hash.slice(0, 10); // changing password invalidates old tokens
const pubUser = (u: User) => ({ email: u.email, name: u.name, role: u.role, mustChange: !!u.mustChange, created: u.created, lastLogin: u.lastLogin || null });
const normEmail = (e: unknown) => String(e || "").trim().toLowerCase();
const pwOk = (p: unknown) => typeof p === "string" && p.length >= 8;

// ---------- users ----------
async function getUsers(): Promise<User[]> {
  return ((await store().get("users", { type: "json" })) as User[]) || [];
}
async function saveUsers(users: User[]) { await store().setJSON("users", users); }
// setup stays open until the first admin has logged in once (and no data exists yet),
// so a mistyped password at setup can be fixed by running setup again
async function setupOpen() {
  const users = await getUsers();
  if (!users.length) return true;
  return users.length === 1 && !users[0].lastLogin && (await listYears()).length === 0;
}

// ---------- year docs (optimistic concurrency) ----------
const yearKey = (y: string) => `year/${y}`;
const validYear = (y: string) => /^\d{4}$/.test(y);
function emptyDoc(year: string) {
  return { year, income: [], expenses: [], invoices: [], settings: defaultSettings(), updatedAt: new Date().toISOString(), updatedBy: "" };
}
function defaultSettings() {
  return {
    openingBalance: 0,
    bank: { deposits: 0, checking: 0, eur: 0, usd: 0, eurRate: 0, usdRate: 0 },
    balanceTitle: "",
    notes: "יתרות בבנק אינן משקפות הכנסות והוצאות נוספות שטרם בוצעו ביום הבדיקה.\nהון העמותה מומר לאורך השנים לאירו לכן רגיש מאוד לשערים, ייתכנו הפסדים ורווחים שאינם מוצגים בתקציב.",
    expFilter: { status: ["לא התקיים"], priority: [], impl: [], emergency: [] }, // excluded values
  };
}
async function readYear(y: string) {
  const r = await store().getWithMetadata(yearKey(y), { type: "json" });
  return r === null ? null : { doc: r.data as any, etag: r.etag };
}
async function mutateYear(y: string, fn: (doc: any) => void | string, who: string) {
  for (let i = 0; i < 6; i++) {
    const cur = await readYear(y);
    if (!cur) return { error: "השנה לא קיימת", status: 404 };
    const err = fn(cur.doc);
    if (typeof err === "string") return { error: err, status: 400 };
    cur.doc.updatedAt = new Date().toISOString();
    cur.doc.updatedBy = who;
    const { modified } = await store().setJSON(yearKey(y), cur.doc, { onlyIfMatch: cur.etag });
    if (modified) return { doc: cur.doc };
    await new Promise((r) => setTimeout(r, 80 + Math.random() * 150));
  }
  return { error: "התנגשות בשמירה, נסה שוב", status: 409 };
}
async function listYears(): Promise<string[]> {
  const { blobs } = await store().list({ prefix: "year/" });
  return blobs.map((b) => b.key.slice(5)).filter(validYear).sort();
}

// ---------- row sanitizing ----------
const MAX_STR = 4000;
function cleanRow(input: any) {
  const out: Record<string, string | number | null> = {};
  if (!input || typeof input !== "object") return out;
  for (const [k, v] of Object.entries(input)) {
    if (!/^[a-zA-Z][a-zA-Z0-9]{0,30}$/.test(k) || ["id", "updatedAt", "updatedBy"].includes(k)) continue;
    if (v === null || v === undefined || v === "") out[k] = null;
    else if (typeof v === "number" && Number.isFinite(v)) out[k] = v;
    else out[k] = String(v).slice(0, MAX_STR);
  }
  return out;
}
const newId = () => randomBytes(6).toString("hex");
function cleanDoc(year: string, body: any) {
  const doc: any = emptyDoc(year);
  for (const k of KINDS) {
    doc[k] = Array.isArray(body?.[k]) ? body[k].slice(0, 5000).map((r: any) => ({ ...cleanRow(r), id: newId() })) : [];
  }
  if (body?.settings && typeof body.settings === "object") doc.settings = mergeSettings(doc.settings, body.settings);
  return doc;
}
function mergeSettings(base: any, s: any) {
  const out = { ...base };
  if (s.openingBalance !== undefined) out.openingBalance = Number(s.openingBalance) || 0;
  if (s.bank && typeof s.bank === "object") {
    out.bank = { ...base.bank };
    for (const k of Object.keys(base.bank)) if (s.bank[k] !== undefined) out.bank[k] = Number(s.bank[k]) || 0;
  }
  if (typeof s.notes === "string") out.notes = s.notes.slice(0, 5000);
  if (typeof s.balanceTitle === "string") out.balanceTitle = s.balanceTitle.slice(0, 200);
  if (s.expFilter && typeof s.expFilter === "object") {
    const f: any = {};
    for (const k of ["status", "priority", "impl", "emergency"]) f[k] = Array.isArray(s.expFilter[k]) ? s.expFilter[k].map((x: any) => String(x)).slice(0, 50) : [];
    out.expFilter = f;
  }
  return out;
}

// ---------- invoices ----------
const invKey = (y: string, id: string) => `inv/${y}/${id}`;
const shareKey = (t: string) => `share/${t}`;
const shareReportKey = (t: string) => `sharefile/${t}/report`;
const MAX_FILE = 4.4 * 1024 * 1024;
const MIME_OK = ["application/pdf", "image/jpeg", "image/png", "image/webp", "image/heic", "image/heif", "image/gif"];
const DONE = "בוצע";
function cleanInvMeta(b: any) {
  const o: any = {};
  for (const k of ["supplier", "invoiceNo", "detail", "notes"]) if (b?.[k] !== undefined) o[k] = b[k] === null ? null : String(b[k]).slice(0, 500);
  if (b?.date !== undefined) o.date = /^\d{4}-\d{2}-\d{2}$/.test(String(b.date)) ? String(b.date) : null;
  if (b?.amount !== undefined) o.amount = b.amount === null || b.amount === "" ? null : Number.isFinite(Number(b.amount)) ? Number(b.amount) : null;
  return o;
}
// link an invoice to an expense row: existing forecast row, or a new row
function applyInvoice(doc: any, inv: any, apply: any, stamp: any) {
  if (!apply) return;
  const month = inv.date ? Number(inv.date.slice(5, 7)) : null;
  if (apply.expenseId) {
    const row = doc.expenses.find((r: any) => r.id === apply.expenseId);
    if (!row) return "שורת ההוצאה לא נמצאה (אולי נמחקה)";
    row.status = DONE;
    if (inv.amount !== null && inv.amount !== undefined) row.amount = inv.amount;
    if (inv.invoiceNo) row.ref = inv.invoiceNo;
    if (month) row.month = month;
    Object.assign(row, stamp);
    inv.expenseId = row.id;
  } else if (apply.newRow) {
    const r = cleanRow(apply.newRow);
    const serial = Math.max(0, ...doc.expenses.map((x: any) => Number(x.serial) || 0)) + 1;
    const row = { serial, year: Number(doc.year), detail: inv.supplier || null, ...r, status: DONE, amount: inv.amount ?? r.amount ?? null, ref: inv.invoiceNo || r.ref || null, month: month || r.month || null, id: newId(), ...stamp };
    doc.expenses.push(row);
    inv.expenseId = row.id;
  }
}
async function getConfig() {
  return ((await store().get("config/app", { type: "json" })) as any) || { accountantEmail: "", accountantName: "" };
}

// ---------- handler ----------
export default async (req: Request, context: Context) => {
  const url = new URL(req.url);
  const parts = url.pathname.replace(/^\/api\/?/, "").split("/").filter(Boolean).map(decodeURIComponent);
  const m = req.method;
  let body: any = {};
  if (m !== "GET" && m !== "DELETE") body = await req.json().catch(() => ({}));

  try {
    // ----- public -----
    if (parts[0] === "status" && m === "GET") {
      return json({ needsSetup: await setupOpen() });
    }
    if (parts[0] === "setup" && m === "POST") {
      const cur = await store().getWithMetadata("users", { type: "json" });
      if (!(await setupOpen())) return fail("המערכת כבר הוגדרה", 403);
      const email = normEmail(body.email);
      if (!email.includes("@")) return fail("מייל לא תקין");
      if (!pwOk(body.password)) return fail("סיסמה של 8 תווים לפחות");
      const u: User = { email, name: String(body.name || email.split("@")[0]).slice(0, 80), role: "admin", ...hashPw(body.password), mustChange: !!body.mustChange, created: new Date().toISOString() };
      const { modified } = await store().setJSON("users", [u], cur === null ? { onlyIfNew: true } : { onlyIfMatch: cur.etag });
      if (!modified) return fail("המערכת כבר הוגדרה", 403);
      return json({ ok: true });
    }
    if (parts[0] === "login" && m === "POST") {
      const users = await getUsers();
      const email = normEmail(body.email);
      const u = users.find((x) => x.email === email);
      if (!u || typeof body.password !== "string" || !checkPw(body.password, u)) {
        await new Promise((r) => setTimeout(r, 400));
        return fail("מייל או סיסמה שגויים", 401);
      }
      u.lastLogin = new Date().toISOString();
      await saveUsers(users);
      return json({ token: await signToken(u.email, pwVer(u)), user: pubUser(u) });
    }

    // ----- public: accountant share links -----
    if (parts[0] === "share" && m === "GET" && parts[1]) {
      const sh: any = await store().get(shareKey(parts[1]), { type: "json" });
      if (!sh) return fail("הקישור לא קיים או בוטל", 404);
      if (Date.parse(sh.expires) < Date.now()) return fail("תוקף הקישור פג", 410);
      const cur = await readYear(sh.year);
      const invs = (cur?.doc.invoices || []).filter((i: any) => sh.invoiceIds.includes(i.id));
      if (!parts[2]) {
        return json({
          year: sh.year, title: sh.title, note: sh.note || "", createdAt: sh.createdAt, expires: sh.expires, createdBy: sh.createdBy,
          report: sh.hasReport ? { name: sh.reportName } : null,
          files: invs.map((i: any) => ({ id: i.id, fileName: i.fileName, mime: i.mime, size: i.size, supplier: i.supplier, invoiceNo: i.invoiceNo, date: i.date, amount: i.amount })),
        });
      }
      if (parts[2] === "report" && sh.hasReport) {
        const data = await store().get(shareReportKey(parts[1]), { type: "arrayBuffer" });
        if (!data) return fail("הדוח לא נמצא", 404);
        return new Response(data, { headers: { "content-type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "content-disposition": `attachment; filename*=UTF-8''${encodeURIComponent(sh.reportName)}`, "cache-control": "no-store" } });
      }
      if (parts[2] === "file" && parts[3]) {
        const inv = invs.find((i: any) => i.id === parts[3]);
        if (!inv) return fail("הקובץ לא נמצא", 404);
        const data = await store().get(invKey(sh.year, inv.id), { type: "arrayBuffer" });
        if (!data) return fail("הקובץ לא נמצא", 404);
        return new Response(data, { headers: { "content-type": inv.mime, "content-disposition": `inline; filename*=UTF-8''${encodeURIComponent(inv.fileName)}`, "cache-control": "no-store" } });
      }
      return fail("נתיב לא קיים", 404);
    }

    // ----- authenticated -----
    const tok = await readToken((req.headers.get("authorization") || "").replace(/^Bearer\s+/i, "") || null);
    if (!tok) return fail("נדרשת התחברות", 401);
    const users = await getUsers();
    const me = users.find((u) => u.email === tok.e);
    if (!me || pwVer(me) !== tok.v) return fail("נדרשת התחברות", 401);
    const canEdit = me.role === "admin" || me.role === "editor";
    const isAdmin = me.role === "admin";
    const who = me.name || me.email;

    if (parts[0] === "me" && m === "GET") return json({ user: pubUser(me) });

    if (parts[0] === "password" && m === "POST") {
      if (!me.mustChange && (typeof body.current !== "string" || !checkPw(body.current, me))) return fail("הסיסמה הנוכחית שגויה");
      if (!pwOk(body.password)) return fail("סיסמה של 8 תווים לפחות");
      Object.assign(me, hashPw(body.password), { mustChange: false });
      await saveUsers(users);
      return json({ token: await signToken(me.email, pwVer(me)), user: pubUser(me) });
    }

    // users management
    if (parts[0] === "users") {
      if (!isAdmin) return fail("אין הרשאה", 403);
      if (m === "GET") return json({ users: users.map(pubUser) });
      if (m === "POST") {
        const email = normEmail(body.email);
        if (!email.includes("@")) return fail("מייל לא תקין");
        if (users.some((u) => u.email === email)) return fail("משתמש עם מייל זה כבר קיים");
        if (!pwOk(body.password)) return fail("סיסמה זמנית של 8 תווים לפחות");
        const role: Role = ["admin", "editor", "viewer"].includes(body.role) ? body.role : "viewer";
        users.push({ email, name: String(body.name || "").slice(0, 80) || email.split("@")[0], role, ...hashPw(body.password), mustChange: true, created: new Date().toISOString() });
        await saveUsers(users);
        return json({ users: users.map(pubUser) });
      }
      const target = users.find((u) => u.email === normEmail(parts[1]));
      if (!target) return fail("משתמש לא נמצא", 404);
      if (m === "PUT") {
        if (body.role && ["admin", "editor", "viewer"].includes(body.role)) {
          if (target.role === "admin" && body.role !== "admin" && users.filter((u) => u.role === "admin").length === 1) return fail("חייב להישאר לפחות מנהל אחד");
          target.role = body.role;
        }
        if (typeof body.name === "string" && body.name.trim()) target.name = body.name.trim().slice(0, 80);
        if (body.password !== undefined) {
          if (!pwOk(body.password)) return fail("סיסמה זמנית של 8 תווים לפחות");
          Object.assign(target, hashPw(body.password), { mustChange: true });
        }
        await saveUsers(users);
        return json({ users: users.map(pubUser) });
      }
      if (m === "DELETE") {
        if (target.email === me.email) return fail("לא ניתן למחוק את עצמך");
        if (target.role === "admin" && users.filter((u) => u.role === "admin").length === 1) return fail("חייב להישאר לפחות מנהל אחד");
        const rest = users.filter((u) => u !== target);
        await saveUsers(rest);
        return json({ users: rest.map(pubUser) });
      }
    }

    // app config (accountant details)
    if (parts[0] === "config") {
      if (m === "GET") return json({ config: await getConfig() });
      if (m === "PUT") {
        if (!canEdit) return fail("אין הרשאה", 403);
        const c = await getConfig();
        if (typeof body.accountantEmail === "string") c.accountantEmail = body.accountantEmail.trim().slice(0, 200);
        if (typeof body.accountantName === "string") c.accountantName = body.accountantName.trim().slice(0, 100);
        await store().setJSON("config/app", c);
        return json({ config: c });
      }
    }

    // years
    if (parts[0] === "years") {
      const y = parts[1];
      if (!y) {
        if (m === "GET") return json({ years: await listYears() });
        if (m === "POST") {
          if (!canEdit) return fail("אין הרשאה", 403);
          const year = String(body.year || "");
          if (!validYear(year)) return fail("שנה לא תקינה");
          let doc: any;
          if (body.data) doc = cleanDoc(year, body.data);
          else if (body.copyFrom && validYear(String(body.copyFrom))) {
            const src = await readYear(String(body.copyFrom));
            if (!src) return fail("שנת המקור לא קיימת");
            doc = cleanDoc(year, { expenses: body.copyRows ? src.doc.expenses : [], income: body.copyRows ? src.doc.income : [], settings: src.doc.settings });
            if (body.copyRows) {
              for (const r of doc.expenses) { r.status = "תחזית"; if (r.year) r.year = Number(year); r.ref = null; }
              for (const r of doc.income) { r.status = "תחזית"; r.receipt = null; r.date = null; }
            }
          } else doc = emptyDoc(year);
          doc.updatedBy = who;
          const { modified } = await store().setJSON(yearKey(year), doc, { onlyIfNew: !body.replace });
          if (!modified) return fail("השנה כבר קיימת");
          return json({ doc, years: await listYears() });
        }
        return fail("בקשה לא נתמכת", 405);
      }
      if (!validYear(y)) return fail("שנה לא תקינה");

      if (parts.length === 2) {
        if (m === "GET") {
          const cur = await readYear(y);
          return cur ? json({ doc: cur.doc }) : fail("השנה לא קיימת", 404);
        }
        if (m === "PUT") { // replace whole year (import)
          if (!canEdit) return fail("אין הרשאה", 403);
          const doc = cleanDoc(y, body.data || {});
          doc.updatedBy = who;
          const prev = await readYear(y);
          doc.invoices = prev?.doc.invoices || []; // invoice files are kept on import
          await store().setJSON(yearKey(y), doc);
          return json({ doc });
        }
        if (m === "DELETE") {
          if (!isAdmin) return fail("רק מנהל יכול למחוק שנה", 403);
          await store().delete(yearKey(y));
          const { blobs } = await store().list({ prefix: `inv/${y}/` });
          for (const b of blobs) await store().delete(b.key);
          return json({ years: await listYears() });
        }
      }
      if (!canEdit && m !== "GET") return fail("אין הרשאה לעריכה", 403);

      if (parts[2] === "settings" && m === "PUT") {
        const r = await mutateYear(y, (doc) => { doc.settings = mergeSettings(doc.settings || defaultSettings(), body); }, who);
        return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
      }
      if (parts[2] === "invoices") {
        const invId = parts[3];
        const stamp = { updatedAt: new Date().toISOString(), updatedBy: who };
        if (m === "GET" && invId && parts[4] === "file") {
          const cur = await readYear(y);
          const inv = cur?.doc.invoices?.find((i: any) => i.id === invId);
          if (!inv) return fail("החשבונית לא נמצאה", 404);
          const data = await store().get(invKey(y, invId), { type: "arrayBuffer" });
          if (!data) return fail("הקובץ לא נמצא", 404);
          return new Response(data, { headers: { "content-type": inv.mime, "cache-control": "private, no-store" } });
        }
        if (m === "POST" && !invId) {
          const buf = Buffer.from(String(body.file || ""), "base64");
          if (!buf.length) return fail("לא התקבל קובץ");
          if (buf.length > MAX_FILE) return fail("הקובץ גדול מדי (עד 4.4MB)");
          const mime = String(body.mime || "");
          if (!MIME_OK.includes(mime)) return fail("סוג קובץ לא נתמך – PDF או תמונה בלבד");
          const id = newId();
          await store().set(invKey(y, id), new Blob([buf], { type: mime }));
          const inv: any = { id, fileName: String(body.fileName || "invoice").slice(0, 200), mime, size: buf.length, supplier: null, invoiceNo: null, date: null, amount: null, detail: null, notes: null,
            ...cleanInvMeta(body.meta || {}), expenseId: null, uploadedAt: stamp.updatedAt, uploadedBy: who, sentAt: null };
          const r = await mutateYear(y, (doc) => {
            doc.invoices = doc.invoices || [];
            doc.invoices.push(inv);
            return applyInvoice(doc, inv, body.apply, stamp);
          }, who);
          if (r.error) { await store().delete(invKey(y, id)); return fail(r.error, r.status); }
          return json({ doc: r.doc, invoice: inv });
        }
        if (m === "PUT" && invId) {
          const r = await mutateYear(y, (doc) => {
            const inv = (doc.invoices || []).find((i: any) => i.id === invId);
            if (!inv) return "החשבונית לא נמצאה";
            Object.assign(inv, cleanInvMeta(body.meta || {}));
            if (body.unlink) inv.expenseId = null;
            return applyInvoice(doc, inv, body.apply, stamp);
          }, who);
          return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
        }
        if (m === "DELETE" && invId) {
          const r = await mutateYear(y, (doc) => {
            const n = (doc.invoices || []).length;
            doc.invoices = (doc.invoices || []).filter((i: any) => i.id !== invId);
            if (doc.invoices.length === n) return "החשבונית לא נמצאה";
          }, who);
          if (r.error) return fail(r.error, r.status);
          await store().delete(invKey(y, invId));
          return json({ doc: r.doc });
        }
      }

      if (parts[2] === "shares") {
        if (m === "GET") {
          const { blobs } = await store().list({ prefix: "share/" });
          const list = [];
          for (const b of blobs) {
            const sh: any = await store().get(b.key, { type: "json" });
            if (sh && sh.year === y) list.push({ token: b.key.slice(6), ...sh });
          }
          list.sort((a, b) => b.createdAt.localeCompare(a.createdAt));
          return json({ shares: list });
        }
        if (m === "POST") {
          const ids: string[] = Array.isArray(body.invoiceIds) ? body.invoiceIds.map(String) : [];
          const cur = await readYear(y);
          const known = new Set((cur?.doc.invoices || []).map((i: any) => i.id));
          const invoiceIds = ids.filter((i) => known.has(i));
          if (!invoiceIds.length && !body.report) return fail("לא נבחרו חשבוניות");
          const token = randomBytes(24).toString("base64url");
          const days = Math.min(90, Math.max(1, Number(body.days) || 30));
          const hasReport = typeof body.report === "string" && body.report.length > 0;
          if (hasReport) {
            const rb = Buffer.from(body.report, "base64");
            if (rb.length > MAX_FILE) return fail("הדוח גדול מדי");
            await store().set(shareReportKey(token), new Blob([rb]));
          }
          const sh = { year: y, invoiceIds, title: String(body.title || `חשבוניות ${y}`).slice(0, 200), note: String(body.note || "").slice(0, 2000),
            createdAt: new Date().toISOString(), expires: new Date(Date.now() + days * 864e5).toISOString(), createdBy: who, hasReport, reportName: String(body.reportName || `ITCB_Budget_${y}.xlsx`).slice(0, 200) };
          await store().setJSON(shareKey(token), sh);
          const r = await mutateYear(y, (doc) => {
            for (const inv of doc.invoices || []) if (invoiceIds.includes(inv.id)) inv.sentAt = sh.createdAt;
          }, who);
          return json({ token, share: sh, doc: r.doc });
        }
        if (m === "DELETE" && parts[3]) {
          await store().delete(shareKey(parts[3]));
          await store().delete(shareReportKey(parts[3]));
          return json({ ok: true });
        }
      }

      const kind = parts[2] as Kind;
      if (KINDS.includes(kind)) {
        const id = parts[3];
        const stamp = { updatedAt: new Date().toISOString(), updatedBy: who };
        if (m === "POST" && !id) {
          const rows = Array.isArray(body.rows) ? body.rows : [body.row];
          const r = await mutateYear(y, (doc) => { for (const row of rows) doc[kind].push({ ...cleanRow(row), id: newId(), ...stamp }); }, who);
          return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
        }
        if (m === "PUT" && id) {
          const r = await mutateYear(y, (doc) => {
            const i = doc[kind].findIndex((x: any) => x.id === id);
            if (i < 0) return "השורה לא נמצאה (אולי נמחקה)";
            doc[kind][i] = { ...cleanRow(body.row), id, ...stamp };
          }, who);
          return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
        }
        if (m === "DELETE" && id) {
          const r = await mutateYear(y, (doc) => {
            const n = doc[kind].length;
            doc[kind] = doc[kind].filter((x: any) => x.id !== id);
            if (doc[kind].length === n) return "השורה לא נמצאה";
          }, who);
          return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
        }
      }
    }
    return fail("נתיב לא קיים", 404);
  } catch (e: any) {
    console.error(e);
    return fail("שגיאת שרת: " + (e?.message || e), 500);
  }
};

export const config: Config = { path: "/api/*" };
