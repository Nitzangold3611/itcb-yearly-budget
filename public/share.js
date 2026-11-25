/* ITCB – public download page for the accountant */
(() => {
  const $m = document.getElementById("main");
  const t = new URLSearchParams(location.search).get("t") || "";
  const esc = (s) => String(s ?? "").replace(/[&<>"']/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c]));
  const shekel = (n) => `<bdi dir="ltr">₪${Math.round(Number(n) || 0).toLocaleString("he-IL")}</bdi>`;
  const fmtD = (iso) => (iso ? new Date(iso + "T12:00:00").toLocaleDateString("he-IL") : "");
  const toast = (msg, bad) => { const e = document.getElementById("toast"); e.textContent = msg; e.className = "show" + (bad ? " bad" : ""); clearTimeout(toast.t); toast.t = setTimeout(() => (e.className = ""), 3000); };
  const safe = (s) => String(s || "").replace(/[\\/:*?"<>|\n\r\t]+/g, " ").replace(/\s+/g, " ").trim().slice(0, 60);
  const fname = (f) => [f.date || "ללא-תאריך", safe(f.supplier) || "ספק", safe(f.invoiceNo)].filter(Boolean).join("_") + "." + (f.mime === "application/pdf" ? "pdf" : (f.mime.split("/")[1] || "jpg").replace("jpeg", "jpg"));
  const base = `/api/share/${encodeURIComponent(t)}`;

  async function load() {
    if (!t) return ($m.innerHTML = `<div class="panel"><h2>קישור לא תקין</h2></div>`);
    const res = await fetch(base);
    const d = await res.json().catch(() => ({}));
    if (!res.ok) return ($m.innerHTML = `<div class="panel"><h2>${esc(d.error || "הקישור אינו זמין")}</h2><p class="hint">יש לפנות לשולח לקבלת קישור חדש.</p></div>`);
    const total = d.files.reduce((a, f) => a + (Number(f.amount) || 0), 0);
    $m.innerHTML = `
      <div class="page-head"><div><h1>${esc(d.title)}</h1><p>נשלח על ידי ${esc(d.createdBy || "")} · ${new Date(d.createdAt).toLocaleDateString("he-IL")} · הקישור בתוקף עד ${new Date(d.expires).toLocaleDateString("he-IL")}</p></div>
        <div class="spacer"></div><button class="btn primary" id="zip">הורדת הכל (ZIP)</button></div>
      ${d.note ? `<div class="note">${esc(d.note)}</div>` : ""}
      <div class="kpis k4">
        <div class="kpi c-n"><div class="lbl">חשבוניות</div><div class="val">${d.files.length}</div></div>
        <div class="kpi c-exp"><div class="lbl">סה״כ</div><div class="val">${shekel(total)}</div></div>
        ${d.report ? `<div class="kpi c-done"><div class="lbl">דוח תקציב</div><div class="val" style="font-size:16px"><a href="${base}/report">${esc(d.report.name)}</a></div></div>` : ""}
      </div>
      <div class="dt-wrap"><table class="dt"><thead><tr><th>תאריך</th><th>ספק</th><th>מס׳ חשבונית</th><th>סכום</th><th></th></tr></thead><tbody>
      ${d.files.map((f) => `<tr><td class="numc">${fmtD(f.date)}</td><td>${esc(f.supplier || "")}</td><td class="numc">${esc(f.invoiceNo || "")}</td><td class="money">${f.amount != null ? shekel(f.amount) : ""}</td>
        <td><a class="icon-btn" href="${base}/file/${f.id}" target="_blank" rel="noopener">פתיחה</a></td></tr>`).join("")}
      </tbody><tfoot><tr><td colspan="3">סה״כ (${d.files.length})</td><td class="money">${shekel(total)}</td><td></td></tr></tfoot></table></div>`;
    document.getElementById("zip").onclick = async (e) => {
      const btn = e.target; btn.disabled = true; btn.textContent = "מכין קובץ…";
      try {
        const files = {}, used = new Set();
        for (const f of d.files) {
          let n = fname(f), k = 2;
          while (used.has(n)) n = fname(f).replace(/(\.[^.]+)$/, `_${k++}$1`);
          used.add(n);
          const r = await fetch(`${base}/file/${f.id}`);
          if (!r.ok) throw new Error("הורדת קובץ נכשלה");
          files["חשבוניות/" + n] = new Uint8Array(await r.arrayBuffer());
        }
        if (d.report) { const r = await fetch(`${base}/report`); if (r.ok) files[d.report.name] = new Uint8Array(await r.arrayBuffer()); }
        const blob = new Blob([fflate.zipSync(files, { level: 0 })], { type: "application/zip" });
        const a = document.createElement("a"); a.href = URL.createObjectURL(blob); a.download = `ITCB_Invoices_${d.year}.zip`; document.body.appendChild(a); a.click();
        setTimeout(() => { URL.revokeObjectURL(a.href); a.remove(); }, 1000);
        toast("הקובץ הורד");
      } catch (er) { toast(er.message, true); }
      btn.disabled = false; btn.textContent = "הורדת הכל (ZIP)";
    };
  }
  load();
})();
