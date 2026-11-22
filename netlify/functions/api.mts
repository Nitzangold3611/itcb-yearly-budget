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

// ---------- year docs (optimistic concurrency) ----------
const yearKey = (y: string) => `year/${y}`;
const validYear = (y: string) => /^\d{4}$/.test(y);
function emptyDoc(year: string) {
  return { year, income: [], expenses: [], settings: defaultSettings(), updatedAt: new Date().toISOString(), updatedBy: "" };
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
      return json({ needsSetup: (await getUsers()).length === 0 });
    }
    if (parts[0] === "setup" && m === "POST") {
      const users = await getUsers();
      if (users.length) return fail("המערכת כבר הוגדרה", 403);
      const email = normEmail(body.email);
      if (!email.includes("@")) return fail("מייל לא תקין");
      if (!pwOk(body.password)) return fail("סיסמה של 8 תווים לפחות");
      const u: User = { email, name: String(body.name || email.split("@")[0]).slice(0, 80), role: "admin", ...hashPw(body.password), mustChange: !!body.mustChange, created: new Date().toISOString() };
      const { modified } = await store().setJSON("users", [u], { onlyIfNew: true });
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
          await store().setJSON(yearKey(y), doc);
          return json({ doc });
        }
        if (m === "DELETE") {
          if (!isAdmin) return fail("רק מנהל יכול למחוק שנה", 403);
          await store().delete(yearKey(y));
          return json({ years: await listYears() });
        }
      }
      if (!canEdit && m !== "GET") return fail("אין הרשאה לעריכה", 403);

      if (parts[2] === "settings" && m === "PUT") {
        const r = await mutateYear(y, (doc) => { doc.settings = mergeSettings(doc.settings || defaultSettings(), body); }, who);
        return r.error ? fail(r.error, r.status) : json({ doc: r.doc });
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
