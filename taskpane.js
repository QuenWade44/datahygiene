let isRunning = false;

Office.onReady((info) => {
  if (info.host !== Office.HostType.Excel) {
    log("This add-in only runs inside Excel.", "err");
    return;
  }

  document.getElementById("btn-clean")
    .addEventListener("click", runClean);

  scanSheet();
});

/* ---------------- UI Helpers ---------------- */

function setProgress(pct, label) {
  document.getElementById("progress-fill").style.width = pct + "%";
  document.getElementById("progress-pct").textContent = pct + "%";
  document.getElementById("progress-text").textContent = label;
}

function log(msg, type = "info") {
  const el = document.getElementById("log");
  el.innerHTML = `<span class="line ${type}">${msg}</span>`;
}

function setOpState(id, state = "") {
  const el = document.getElementById(id);
  el.className = `op-item ${state}`;
}

function resetOps() {
  ["op-strip","op-email","op-name","op-date","op-address","op-dedup"]
    .forEach(id => setOpState(id));
}

/* ---------------- Sheet Scan ---------------- */

async function scanSheet() {
  try {
    await Excel.run(async ctx => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load("rowCount");
      await ctx.sync();

      const rows = Math.max(0, used.rowCount - 1);
      document.getElementById("stat-rows").textContent = rows;
      document.getElementById("stat-removed").textContent = "—";

      setProgress(0, "Ready");
      log(`Detected ${rows} data rows.`, "info");
    });
  } catch (e) {
    log("Unable to access sheet.", "err");
  }
}

/* ---------------- Main Cleaning ---------------- */

async function runClean() {
  if (isRunning) return;
  isRunning = true;

  const btn = document.getElementById("btn-clean");
  btn.disabled = true;
  btn.classList.add("running");
  btn.textContent = "Cleaning...";

  resetOps();
  setProgress(5, "Starting...");

  try {
    await Excel.run(async ctx => {

      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["values","rowCount","columnCount"]);
      await ctx.sync();

      const values = used.values;
      if (!values || values.length < 2) {
        log("No data rows found.", "err");
        return;
      }

      const headers = values[0].map(h => String(h).toLowerCase().trim());
      let rows = values.slice(1);

      const originalCount = rows.length;

      /* Strip whitespace */
      setOpState("op-strip","active");
      rows = rows.map(r => r.map(c => typeof c === "string" ? c.trim() : c));
      setOpState("op-strip","done");
      setProgress(20,"Stripping whitespace");

      /* Remove empty */
      rows = rows.filter(r => r.some(c => c !== "" && c != null));

      /* Column transforms */
      for (let i=0;i<headers.length;i++) {
        const col = headers[i];

        if (col.includes("email")) {
          setOpState("op-email","active");
          rows = rows.map(r => { r[i] = cleanEmail(r[i]); return r; });
          setOpState("op-email","done");
        }

        if (col.includes("name")) {
          setOpState("op-name","active");
          rows = rows.map(r => { r[i] = capitalizeName(r[i]); return r; });
          setOpState("op-name","done");
        }

        if (col.includes("date")) {
          setOpState("op-date","active");
          rows = rows.map(r => { r[i] = parseDate(r[i]); return r; });
          setOpState("op-date","done");
        }

        if (col.includes("address")) {
          setOpState("op-address","active");
          rows = rows.map(r => { r[i] = normalizeAddress(r[i]); return r; });
          setOpState("op-address","done");
        }

        setProgress(30 + Math.round((i/headers.length)*40),"Transforming columns");
      }

      /* Deduplicate */
      setOpState("op-dedup","active");
      const unique = Array.from(new Set(rows.map(r => JSON.stringify(r))))
        .map(r => JSON.parse(r));
      setOpState("op-dedup","done");

      const removed = originalCount - unique.length;

      /* Write back */
      const newData = [values[0], ...unique];
      sheet.getRangeByIndexes(0,0,newData.length,headers.length).values = newData;
      await ctx.sync();

      document.getElementById("stat-rows").textContent = unique.length;
      document.getElementById("stat-removed").textContent = removed;

      setProgress(100,"Complete");
      log(`Removed ${removed} duplicate rows.`, "ok");
    });

  } catch (e) {
    log("Error: " + e.message, "err");
    setProgress(0,"Failed");
  }

  btn.disabled = false;
  btn.classList.remove("running");
  btn.textContent = "▶ Clean Data";
  isRunning = false;
}

/* ---------------- Transform Helpers ---------------- */

function cleanEmail(email){
  if (!email) return "";
  return String(email).trim().toLowerCase();
}

function capitalizeName(name){
  if (!name) return "";
  return String(name)
    .split(" ")
    .map(w => w.charAt(0).toUpperCase()+w.slice(1).toLowerCase())
    .join(" ");
}

function parseDate(val){
  if (!val) return "";
  const d = new Date(val);
  if (isNaN(d)) return val;
  return d.toISOString().split("T")[0];
}

function normalizeAddress(addr){
  if (!addr) return "";
  return String(addr)
    .split(" ")
    .map(w => w.charAt(0).toUpperCase()+w.slice(1).toLowerCase())
    .join(" ");
}
