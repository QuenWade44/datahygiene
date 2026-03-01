/* ============================================================
   DataHygiene CRM — Office.js Add-in Service
   Mirrors the Python DataCleanerWorker logic on the active sheet
   ============================================================ */

Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    scanSheet(); // populate stats on load
  }
});

// ─── UI HELPERS ─────────────────────────────────────────────

function setProgress(pct, label) {
  document.getElementById("progress-fill").style.width = pct + "%";
  document.getElementById("progress-pct").textContent = pct + "%";
  document.getElementById("progress-text").textContent = label || "";
}

function log(msg, type) {
  const el = document.getElementById("log");
  const line = document.createElement("span");
  line.className = "line " + (type || "");
  line.textContent = msg;
  el.innerHTML = "";
  el.appendChild(line);
}

function setOpState(id, state) {
  // state: '' | 'active' | 'done'
  const el = document.getElementById(id);
  el.className = "op-item " + state;
}

function resetOps() {
  ["op-strip", "op-email", "op-name", "op-date", "op-address", "op-dedup"]
    .forEach(id => setOpState(id, ""));
}

// ─── SCAN SHEET (load stats on open) ────────────────────────

async function scanSheet() {
  try {
    await Excel.run(async (ctx) => {
      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["rowCount", "columnCount"]);
      await ctx.sync();
      const rows = Math.max(0, used.rowCount - 1); // minus header
      document.getElementById("stat-rows").textContent = rows.toLocaleString();
      document.getElementById("stat-removed").textContent = "—";
      log(`Sheet loaded: ${rows} data rows detected.`, "info");
    });
  } catch (e) {
    log("Could not read sheet.", "err");
  }
}

// ─── MAIN ENTRY POINT ───────────────────────────────────────

async function runClean() {
  const btn = document.getElementById("btn-clean");
  btn.disabled = true;
  btn.classList.add("running");
  btn.textContent = "⏳ Cleaning...";
  resetOps();
  setProgress(0, "Starting...");
  document.getElementById("stat-removed").textContent = "—";

  try {
    await Excel.run(async (ctx) => {

      const sheet = ctx.workbook.worksheets.getActiveWorksheet();
      const used = sheet.getUsedRange();
      used.load(["values", "rowCount", "columnCount"]);
      await ctx.sync();

      const values = used.values;
      if (!values || values.length < 2) {
        log("Sheet has no data rows.", "err");
        return;
      }

      const headers = values[0].map(h => (h || "").toString().trim().toLowerCase());
      const originalRowCount = values.length - 1;
      document.getElementById("stat-rows").textContent = originalRowCount.toLocaleString();

      setProgress(10, "Reading sheet...");
      log("Reading active sheet...", "info");

      // ── Step 1: Strip whitespace ──────────────────────────
      setOpState("op-strip", "active");
      let rows = values.slice(1).map(row =>
        row.map(cell => (typeof cell === "string" ? cell.trim() : cell))
      );
      setOpState("op-strip", "done");
      setProgress(20, "Stripping whitespace...");
      await delay(120);

      // ── Step 2: Drop fully-empty rows ─────────────────────
      rows = rows.filter(row => row.some(cell => cell !== "" && cell !== null && cell !== undefined));

      // ── Step 3: Column-level transformations ──────────────
      const totalCols = headers.length;

      for (let ci = 0; ci < totalCols; ci++) {
        const col = headers[ci];

        if (col.includes("email")) {
          setOpState("op-email", "active");
          rows = rows.map(r => { r[ci] = cleanEmail(r[ci]); return r; });
          setOpState("op-email", "done");
        }

        if (col.includes("name")) {
          setOpState("op-name", "active");
          rows = rows.map(r => { r[ci] = capitalizeName(r[ci]); return r; });
          setOpState("op-name", "done");
        }

        if (col.includes("date")) {
          setOpState("op-date", "active");
          rows = rows.map(r => { r[ci] = parseDate(r[ci]); return r; });
          setOpState("op-date", "done");
        }

        if (col.includes("address")) {
          setOpState("op-address", "active");
          rows = rows.map(r => { r[ci] = normalizeAddress(r[ci]); return r; });
          setOpState("op-address", "done");
        }

        const pct = 25 + Math.round(((ci + 1) / totalCols) * 45);
        setProgress(pct, "Transforming columns...");
        await delay(30);
      }

      setProgress(75, "Removing duplicates...");
      await delay(100);

      // ── Step 4: Deduplicate ───────────────────────────────
      setOpState("op-dedup", "active");
      const seen = new Set();
      const deduped = [];
      for (const row of rows) {
        const key = JSON.stringify(row);
        if (!seen.has(key)) {
          seen.add(key);
          deduped.push(row);
        }
      }
      setOpState("op-dedup", "done");

      const removedCount = originalRowCount - deduped.length;
      setProgress(90, "Writing back to sheet...");
      await delay(80);

      // ── Step 5: Write back ────────────────────────────────
      // Clear old data area and write header + cleaned rows
      const newValues = [values[0], ...deduped];
      const writeRange = sheet.getRangeByIndexes(0, 0, newValues.length, headers.length);
      writeRange.values = newValues;

      // Clear any rows below the new data (leftover from original)
      if (originalRowCount > deduped.length) {
        const clearStart = newValues.length;
        const clearRows = originalRowCount - deduped.length;
        const clearRange = sheet.getRangeByIndexes(clearStart, 0, clearRows, headers.length);
        clearRange.clear(Excel.ClearApplyTo.contents);
      }

      await ctx.sync();

      setProgress(100, "Done");
      document.getElementById("stat-rows").textContent = deduped.length.toLocaleString();
      document.getElementById("stat-removed").textContent = removedCount.toLocaleString();
      log(`✓ Cleaned. Removed ${removedCount} row(s). Final: ${deduped.length} rows.`, "ok");
    });

  } catch (e) {
    log("Error: " + e.message, "err");
    setProgress(0, "Failed");
  }

  btn.disabled = false;
  btn.classList.remove("running");
  btn.textContent = "▶ Clean Data";
}

// ─── TRANSFORMATION FUNCTIONS ────────────────────────────────
// Mirrors DataCleanerWorker methods from the Python app

function cleanEmail(email) {
  if (email === null || email === undefined || email === "") return "";
  email = String(email).trim().toLowerCase();
  email = email.replace(/,/g, ".");
  email = email.replace(/atexample/g, "@example");
  if (!email.includes("@") && email.includes("example")) {
    const parts = email.split("example");
    email = parts[0] + "@example" + parts.slice(1).join("example");
  }
  return email;
}

function capitalizeName(name) {
  if (name === null || name === undefined || name === "") return name;
  return String(name)
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
}

function parseDate(val) {
  if (val === null || val === undefined || String(val).trim() === "") return "";
  // Excel stores dates as serial numbers or strings
  const s = String(val).trim();
  // Try to detect Excel serial number (pure integer ~30000-60000)
  if (/^\d{4,5}$/.test(s)) {
    const serial = parseInt(s, 10);
    if (serial > 25569 && serial < 80000) {
      // Convert Excel serial to JS date (Excel epoch: Jan 1 1900, JS epoch: Jan 1 1970)
      const msDate = (serial - 25569) * 86400 * 1000;
      return formatDate(new Date(msDate));
    }
  }
  const d = new Date(s);
  if (!isNaN(d.getTime())) return formatDate(d);
  return val; // return original if unparseable
}

function formatDate(d) {
  const y = d.getUTCFullYear();
  const m = String(d.getUTCMonth() + 1).padStart(2, "0");
  const day = String(d.getUTCDate()).padStart(2, "0");
  return `${y}-${m}-${day}`;
}

function normalizeAddress(address) {
  if (address === null || address === undefined || address === "") return "";
  // Capitalize each word
  let addr = String(address)
    .split(" ")
    .map(w => w.charAt(0).toUpperCase() + w.slice(1).toLowerCase())
    .join(" ");
  // Uppercase 2-letter state codes (word boundaries)
  addr = addr.replace(/\b([A-Za-z]{2})\b/g, (m) => m.toUpperCase());
  return addr;
}

// ─── UTILITY ─────────────────────────────────────────────────

function delay(ms) {
  return new Promise(resolve => setTimeout(resolve, ms));
}
