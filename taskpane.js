// ===========================================================
// ExcelWizPro Taskpane — v12.1.0
// Advanced Smart Mapping + Auto-Refresh
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "12.1.0";

console.log(`🧠 ExcelWizPro Taskpane v${VERSION} loaded`);

// Optional Office error logging
if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

// -----------------------------------------------------------
// Global safety handlers
// -----------------------------------------------------------
window.addEventListener("error", (e) =>
  console.warn("Window error:", e.message || e.error)
);
window.addEventListener("unhandledrejection", (e) =>
  console.warn("Unhandled rejection:", e.reason)
);

// -----------------------------------------------------------
// Helpers
// -----------------------------------------------------------
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
function getEl(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: #${id}`);
  return el;
}
function showToast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2600);
}

function columnIndexToLetter(index) {
  let n = index + 1;
  let letters = "";
  while (n > 0) {
    const rem = (n - 1) % 26;
    letters = String.fromCharCode(65 + rem) + letters;
    n = Math.floor((n - 1) / 26);
  }
  return letters;
}
function normalizeName(name) {
  return String(name || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

// -----------------------------------------------------------
// HTTP safety
// -----------------------------------------------------------
function timeoutSignal(ms) {
  if (typeof AbortController === "undefined") return undefined;
  const ctrl = new AbortController();
  setTimeout(() => ctrl.abort(), ms);
  return ctrl.signal;
}
async function safeFetch(url, { timeout = 15000, ...opts } = {}) {
  if (!navigator.onLine) {
    const err = new Error("offline");
    err.code = "OFFLINE";
    throw err;
  }
  const signal = opts.signal || timeoutSignal(timeout);
  return fetch(url, { ...opts, signal });
}

// -----------------------------------------------------------
// Diagnostics
// -----------------------------------------------------------
function getOfficeDiagnostics() {
  try {
    return {
      host: Office.context?.host || "unknown",
      platform: Office.context?.diagnostics?.platform || "unknown",
      version: Office.context?.diagnostics?.version || "unknown",
      build: Office.context?.diagnostics?.build || "n/a"
    };
  } catch {
    return { host: "unknown", platform: "unknown", version: "unknown" };
  }
}

// ===========================================================
// BOOT SEQUENCE
// ===========================================================
function officeReady() {
  return new Promise((resolve) => {
    if (window.Office && Office.onReady) {
      Office.onReady(resolve);
    } else {
      let tries = 0;
      const timer = setInterval(() => {
        tries++;
        if (window.Office && Office.onReady) {
          clearInterval(timer);
          Office.onReady(resolve);
        }
        if (tries > 40) {
          clearInterval(timer);
          resolve({ host: "unknown" });
        }
      }, 500);
    }
  });
}

async function ensureExcelHost(info) {
  if (!info || info.host !== Office.HostType.Excel) {
    console.warn("⚠️ Not Excel host:", info && info.host);
    showToast("⚠️ Excel host not detected.");
    return false;
  }
  console.log("🟢 Excel host OK");
  return true;
}

async function waitForExcelApi() {
  for (let i = 1; i <= 20; i++) {
    try {
      await Excel.run(async (ctx) => {
        ctx.workbook.properties.load("title");
        await ctx.sync();
      });
      return true;
    } catch {
      await delay(350 + i * 120);
    }
  }
  showToast("⚠️ Excel not ready — try reopening the add-in.");
  return false;
}

// ===========================================================
// BACKEND WARM-UP
// ===========================================================
async function warmUpBackend(max = 5) {
  const status = document.createElement("div");
  Object.assign(status.style, {
    padding: "6px",
    marginBottom: "8px",
    borderRadius: "6px",
    fontSize: "0.9rem",
    textAlign: "center"
  });
  document.querySelector("main.container")?.prepend(status);

  for (let i = 1; i <= max; i++) {
    try {
      const r = await safeFetch(`${API_BASE}/health`, { timeout: 4000 });
      if (r.ok) {
        status.textContent = "✅ Backend ready";
        status.style.background = "#e6ffed";
        status.style.color = "#0c7a0c";
        setTimeout(() => status.remove(), 2000);
        return;
      }
    } catch {}
    status.textContent = `⏳ Waking backend… (${i}/${max})`;
    status.style.background = "#fff4ce";
    status.style.color = "#976f00";
    await delay(1500 + i * 500);
  }

  status.textContent = "❌ Backend unreachable";
  status.style.background = "#fde7e9";
  status.style.color = "#c22";
}

// ===========================================================
// SAFE Excel.run
// ===========================================================
async function safeExcelRun(cb) {
  try {
    return await Excel.run(cb);
  } catch (err) {
    console.warn("Excel.run failed:", err);
    showToast("⚠️ Excel not ready");
    throw err;
  }
}

// ===========================================================
// ADVANCED SMART COLUMN MAPPING (Option B)
// -----------------------------------------------------------
// • Multi-row headers (up to 3 rows)
// • Excel Tables (ListObjects)
// • Named Ranges
// • Pivot Tables (context only)
// • Safe for 1M+ rows
// ===========================================================
let columnMapCache = "";

async function buildColumnMap() {
  return safeExcelRun(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;

    sheets.load("items/name,items/visibility");
    await ctx.sync();

    const lines = [];

    for (const sheet of sheets.items) {
      const vis = sheet.visibility;
      const visText = vis !== "Visible" ? ` (${vis.toLowerCase()})` : "";
      lines.push(`Sheet: ${sheet.name}${visText}`);

      const used = sheet.getUsedRangeOrNullObject(false);
      used.load("rowCount,columnCount,isNullObject");
      await ctx.sync();

      if (used.isNullObject || used.rowCount < 2) continue;

      const headerRows = Math.min(3, used.rowCount);
      const headerRange = sheet.getRangeByIndexes(
        0,
        0,
        headerRows,
        used.columnCount
      );
      headerRange.load("values");
      await ctx.sync();

      const headers = headerRange.values;
      const startRow = headerRows + 1;
      const lastRow = used.rowCount;

      for (let col = 0; col < used.columnCount; col++) {
        const parts = [];
        for (let r = 0; r < headerRows; r++) {
          const v = headers[r][col];
          if (v !== null && v !== "" && v !== undefined) {
            parts.push(String(v).trim());
          }
        }
        if (!parts.length) continue;

        const combined = parts.join(" - ");
        const normalized = normalizeName(combined);
        const colLetter = columnIndexToLetter(col);
        const safe = sheet.name.replace(/'/g, "''");

        lines.push(
          `${normalized} = '${safe}'!${colLetter}${startRow}:${colLetter}${lastRow}`
        );
      }

      // ───────────────────────────────────────── Tables
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      for (const table of tables.items) {
        lines.push(`Table: ${table.name}`);

        const header = table.getHeaderRowRange();
        const body = table.getDataBodyRange();
        header.load("values");
        body.load("address,rowCount,columnCount");
      }
      await ctx.sync();

      for (const table of tables.items) {
        const headerVals = table.getHeaderRowRange().values?.[0] || [];
        headerVals.forEach((h) => {
          if (!h) return;
          const norm = normalizeName(`${table.name}.${h}`);
          const structuredRef = `${table.name}[${h}]`;
          lines.push(`${norm} = ${structuredRef}`);
        });
      }

      // ───────────────────────────────────────── Pivots
      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      await ctx.sync();
      pivots.items.forEach((p) => lines.push(`PivotSource: ${p.name}`));
    }

    // ───────────────────────────────────────── Named Ranges
    const names = wb.names;
    names.load("items/name");
    await ctx.sync();
    const meta = [];

    for (const n of names.items) {
      const r = n.getRange();
      r.load("address");
      meta.push({ name: n.name, range: r });
    }
    await ctx.sync();

    meta.forEach(({ name, range }) => {
      lines.push(`NamedRange: ${name}`);
      const norm = normalizeName(name);
      lines.push(`${norm} = ${range.address}`);
    });

    return lines.join("\n");
  });
}

// -----------------------------------------------------------
// AUTO REFRESH COLUMN MAP (on taskpane visibility)
// -----------------------------------------------------------
async function autoRefreshColumnMap() {
  try {
    console.log("🔄 Auto-refreshing Smart Column Map…");
    columnMapCache = await buildColumnMap();
    console.log("✅ Updated Smart Column Map");
  } catch (err) {
    console.warn("Auto-refresh failed:", err);
    showToast("⚠️ Could not refresh column map");
  }
}

// Auto-refresh when taskpane becomes visible
if (Office.addin?.onVisibilityModeChanged) {
  Office.addin.onVisibilityModeChanged(async (args) => {
    if (args.visibilityMode === "Taskpane") {
      await autoRefreshColumnMap();
    }
  });
} else if (Office.addin?.onVisibilityChanged) {
  Office.addin.onVisibilityChanged(async (visible) => {
    if (visible) await autoRefreshColumnMap();
  });
}

// ===========================================================
// SHEET DROPDOWN
// ===========================================================
async function refreshSheetDropdown(el) {
  return safeExcelRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    el.innerHTML = "";
    sheets.items.forEach((s) => {
      const opt = document.createElement("option");
      opt.value = s.name;
      opt.textContent = s.name;
      el.appendChild(opt);
    });
  });
}

// ===========================================================
// BACKEND FORMULA GENERATION
// ===========================================================
async function generateFormulaFromBackend(payload) {
  const r = await safeFetch(`${API_BASE}/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
    cache: "no-store",
    timeout: 20000
  });

  const data = await r.json();
  return data.formula || "=ERROR(\"No formula returned\")";
}

// ===========================================================
// INSERT FORMULA
// ===========================================================
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();
  const btn = document.createElement("button");
  btn.className = "btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      await safeExcelRun(async (ctx) => {
        ctx.workbook.getSelectedRange().formulas = [[formula]];
        await ctx.sync();
      });
      showToast("✅ Inserted");
    } catch {
      showToast("⚠️ Select a cell first");
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// ===========================================================
// MAIN UI
// ===========================================================
async function initExcelWizPro() {
  const sheetSelect = getEl("sheetSelect");
  const queryInput = getEl("query");
  const output = getEl("output");
  const genBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");

  let lastFormula = "";

  await refreshSheetDropdown(sheetSelect);
  warmUpBackend();

  genBtn.addEventListener("click", async () => {
    const query = queryInput.value.trim();
    if (!query) return showToast("⚠️ Enter a request");

    if (!navigator.onLine) {
      return showToast("📴 Offline");
    }

    output.textContent = "⏳ Generating…";

    // Ensure map is ready
    if (!columnMapCache) await autoRefreshColumnMap();

    const { version } = getOfficeDiagnostics();

    const payload = {
      query,
      columnMap: columnMapCache,
      excelVersion: version,
      mainSheet: sheetSelect.value
    };

    try {
      const formula = await generateFormulaFromBackend(payload);
      lastFormula = formula;
      output.textContent = formula;
      attachInsertButton(output, formula);
    } catch (err) {
      output.textContent = "❌ Error — see console";
      console.error("Generation failed:", err);
    }
  });

  clearBtn.addEventListener("click", () => {
    output.textContent = "";
    queryInput.value = "";
  });

  window.addEventListener("online", () => {
    if (lastFormula) showToast("🌐 Back online — formula restored");
  });

  console.log("🟢 ExcelWizPro UI ready");
}

// ===========================================================
// STARTUP
// ===========================================================
(async function boot() {
  console.log("🧠 Starting ExcelWizPro…");

  const info = await officeReady();
  if (!(await ensureExcelHost(info))) return;
  if (!(await waitForExcelApi())) return;

  await initExcelWizPro();
  showToast("✅ ExcelWizPro ready!");
})();
