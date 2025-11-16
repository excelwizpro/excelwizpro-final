// ===========================================================
// ExcelWizPro Taskpane Script — v12.0.0 (Advanced Smart Mapping)
// - Safe Office/Excel startup
// - Smart column mapping: headers, tables, named ranges, pivots
// - Backend /generate does rule logic + optional AI
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "12.0.0";

console.log(`🧠 ExcelWizPro v${VERSION} taskpane.js loaded`);

// Optional: better Office error logging (guard against undefined Office)
if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

// -----------------------------------------------------------
// Global safety: don't let errors silently kill the WebView
// -----------------------------------------------------------
if (typeof window !== "undefined") {
  window.addEventListener("error", (e) => {
    console.warn("Window error:", e.message || e.error);
  });
  window.addEventListener("unhandledrejection", (e) => {
    console.warn("Unhandled promise rejection:", e.reason);
  });
}

// -----------------------------------------------------------
// Basic helpers
// -----------------------------------------------------------
function delay(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

function getEl(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: #${id}`);
  return el;
}

// Toast UI (uses .toast class from your CSS)
function showToast(msg) {
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 2600);
}

// Convert 0-based column index → Excel column letters (A, B, ... AA, AB, etc.)
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

// Normalize a name for mapping (lowercase, spaces → underscores)
function normalizeName(name) {
  return String(name || "")
    .trim()
    .toLowerCase()
    .replace(/\s+/g, "_");
}

// -----------------------------------------------------------
// Abort + safeFetch for offline / timeouts
// -----------------------------------------------------------
function timeoutSignal(ms) {
  if (typeof AbortController === "undefined") return undefined;
  const ctrl = new AbortController();
  setTimeout(() => ctrl.abort(), ms);
  return ctrl.signal;
}

async function safeFetch(url, { timeout = 12000, ...opts } = {}) {
  if (typeof navigator !== "undefined" && !navigator.onLine) {
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
// SAFE EXCEL BOOT SEQUENCE
// ===========================================================

// Step 1 — wait for Office.js / host
function officeReady() {
  return new Promise((resolve) => {
    if (typeof Office !== "undefined" && Office && Office.onReady) {
      Office.onReady((info) => {
        console.log("📘 Office.onReady:", info);
        resolve(info);
      });
    } else {
      console.log("⏳ Waiting for Office.js injection…");
      let tries = 0;
      const timer = setInterval(() => {
        tries++;
        if (typeof Office !== "undefined" && Office && Office.onReady) {
          clearInterval(timer);
          Office.onReady((info) => {
            console.log("📘 Office.onReady (delayed):", info);
            resolve(info);
          });
        }
        if (tries > 40) {
          clearInterval(timer);
          console.warn("⚠️ Office.js never reported ready — fallback mode.");
          resolve({ host: "unknown" });
        }
      }, 500);
    }
  });
}

// Step 2 — ensure we're really in Excel
async function ensureExcelHost(info) {
  if (!info || info.host !== Office.HostType.Excel) {
    console.warn("⚠️ Not running inside Excel host:", info && info.host);
    showToast("⚠️ Excel host not detected.");
    return false;
  }
  console.log("🟢 Excel host confirmed.");
  return true;
}

// Step 3 — wait for Excel API to be usable
async function waitForExcelApi() {
  for (let i = 1; i <= 15; i++) {
    try {
      console.log(`🔧 Checking Excel API… (${i}/15)`);
      await Excel.run(async (ctx) => {
        ctx.workbook.properties.load("title");
        await ctx.sync();
      });
      console.log("🟢 Excel API ready.");
      return true;
    } catch (e) {
      await delay(500 + i * 100);
    }
  }
  console.error("❌ Excel API did not become ready.");
  showToast("⚠️ Excel still loading — reopen the add-in.");
  return false;
}

// ===========================================================
// BACKEND WARM-UP
// ===========================================================
async function warmUpBackend(max = 5, baseDelay = 2000) {
  try {
    const statusDiv = document.createElement("div");
    Object.assign(statusDiv.style, {
      padding: "6px",
      marginBottom: "8px",
      borderRadius: "6px",
      fontSize: "0.9rem",
      fontWeight: "500",
      textAlign: "center"
    });
    document.querySelector("main.container")?.prepend(statusDiv);

    for (let i = 1; i <= max; i++) {
      try {
        const res = await safeFetch(`${API_BASE}/health`, {
          cache: "no-store",
          timeout: 5000
        });
        if (res.ok) {
          statusDiv.textContent = "✅ Backend awake";
          statusDiv.style.backgroundColor = "#e6ffed";
          statusDiv.style.color = "#0f7b0f";
          setTimeout(() => statusDiv.remove(), 2200);
          console.log("✅ Backend warm-up complete");
          return;
        }
        throw new Error(`HTTP ${res.status}`);
      } catch (err) {
        const offline = err.code === "OFFLINE";
        statusDiv.textContent = offline
          ? "📴 Offline — reconnect to use ExcelWizPro"
          : `⏳ Waking backend… (${i}/${max})`;
        statusDiv.style.backgroundColor = "#fff4ce";
        statusDiv.style.color = "#986f00";
        await delay(baseDelay * (1 + Math.random()));
      }
    }

    statusDiv.textContent = "❌ Cannot reach backend";
    statusDiv.style.backgroundColor = "#fde7e9";
    statusDiv.style.color = "#d13438";
  } catch (e) {
    console.warn("Warm-up error:", e);
  }
}

// ===========================================================
// SAFE Excel.run wrapper
// ===========================================================
async function safeExcelRun(cb) {
  try {
    return await Excel.run(cb);
  } catch (e) {
    console.warn("⚠️ Excel context problem:", e);
    showToast("⚠️ Excel still initializing — try again.");
    throw e;
  }
}

// ===========================================================
// ADVANCED SMART COLUMN MAP LOGIC (Option B)
// - Multi-row headers (up to 3 rows)
// - Tables (ListObjects)
// - Named ranges
// - Pivot table context
// ===========================================================
async function buildColumnMap() {
  return safeExcelRun(async (ctx) => {
    const lines = [];
    const wb = ctx.workbook;
    const sheets = wb.worksheets;

    sheets.load("items/name,items/visibility");
    await ctx.sync();

    // ---------- 1) Sheet-based headers (multi-row) ----------
    for (const sheet of sheets.items) {
      const visibility = sheet.visibility || "Visible";
      const visibilityLabel =
        visibility !== "Visible" ? ` (${visibility.toLowerCase()})` : "";
      lines.push(`Sheet: ${sheet.name}${visibilityLabel}`);

      const used = sheet.getUsedRangeOrNullObject(false); // ignore pure formatting
      used.load("rowCount,columnCount,isNullObject");
      await ctx.sync();

      if (used.isNullObject || used.rowCount < 2 || used.columnCount < 1) {
        // No data rows or effectively empty sheet
        continue;
      }

      const headerDepth = Math.min(3, used.rowCount); // up to 3 header rows
      const headerRange = sheet.getRangeByIndexes(
        0,
        0,
        headerDepth,
        used.columnCount
      );
      headerRange.load("values");
      await ctx.sync();

      const headerValues = headerRange.values || [];
      const dataStartRow = headerDepth + 1; // data starts after header rows
      const lastRow = used.rowCount; // includes header rows

      for (let colIdx = 0; colIdx < used.columnCount; colIdx++) {
        const parts = [];

        for (let r = 0; r < headerDepth; r++) {
          const val = headerValues[r]?.[colIdx];
          if (val !== null && val !== undefined && String(val).trim() !== "") {
            parts.push(String(val).trim());
          }
        }

        if (!parts.length) continue; // no header for this column

        const combinedHeader = parts.join(" - ");
        const normalized = normalizeName(combinedHeader);
        if (!normalized) continue;

        const colLetter = columnIndexToLetter(colIdx);
        const safeSheetName = sheet.name.replace(/'/g, "''");
        const range = `'${safeSheetName}'!${colLetter}${dataStartRow}:${colLetter}${lastRow}`;

        lines.push(`${normalized} = ${range}`);
      }

      // ---------- 2) Tables on this sheet ----------
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      if (tables.items.length) {
        for (const table of tables.items) {
          lines.push(`Table: ${table.name}`);

          const headerRange = table.getHeaderRowRange();
          const dataRange = table.getDataBodyRange();

          headerRange.load("values");
          dataRange.load("rowCount,columnCount,address");
        }

        await ctx.sync();

        for (const table of tables.items) {
          const headerRange = table.getHeaderRowRange();
          const dataRange = table.getDataBodyRange();
          const headers = (headerRange.values && headerRange.values[0]) || [];

          headers.forEach((header, idx) => {
            if (!header) return;
            const headerName = String(header).trim();
            if (!headerName) return;

            const tableNorm = normalizeName(table.name);
            const headerNorm = normalizeName(headerName);
            const combinedName = `${tableNorm}.${headerNorm}`;

            // Use structured references for dynamic ranges
            const structuredRef = `${table.name}[${headerName}]`;

            lines.push(`${combinedName} = ${structuredRef}`);
          });
        }
      }

      // ---------- 3) Pivot tables (context only, for AI) ----------
      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      await ctx.sync();

      if (pivots.items.length) {
        for (const p of pivots.items) {
          lines.push(`PivotSource: ${p.name}`);
          // We don't map pivot fields as ranges (GETPIVOTDATA is more appropriate)
          // but the name + sheet context is very useful for AI prompts.
        }
      }
    }

    // ---------- 4) Workbook Named Ranges ----------
    const names = wb.names;
    names.load("items/name");
    await ctx.sync();

    if (names.items.length) {
      const nameMetadata = [];

      for (const namedItem of names.items) {
        const range = namedItem.getRange();
        range.load("address");
        nameMetadata.push({ namedItem, range });
      }

      await ctx.sync();

      for (const { namedItem, range } of nameMetadata) {
        lines.push(`NamedRange: ${namedItem.name}`);
        const norm = normalizeName(namedItem.name);
        if (!norm) continue;
        // address is already a valid reference like Sheet1!$A$1:$A$1000 or $A$1
        lines.push(`${norm} = ${range.address}`);
      }
    }

    return lines.join("\n");
  });
}

// ===========================================================
// SHEET DROPDOWN POPULATION
// ===========================================================
async function refreshSheetDropdown(selectEl) {
  try {
    await safeExcelRun(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();

      selectEl.innerHTML = "";
      sheets.items.forEach((s) => {
        const opt = document.createElement("option");
        opt.value = s.name;
        opt.textContent = s.name;
        selectEl.appendChild(opt);
      });
    });
  } catch (e) {
    console.warn("Could not refresh sheets:", e);
    showToast("⚠️ Could not read workbook sheets.");
  }
}

// ===========================================================
// BACKEND FORMULA GENERATION
// ===========================================================
async function generateFormulaFromBackend(payload) {
  const res = await safeFetch(`${API_BASE}/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    cache: "no-store",
    timeout: 15000,
    body: JSON.stringify(payload)
  });

  if (!res.ok) throw new Error(`Backend HTTP ${res.status}`);

  const data = await res.json();
  const formula = (data.formula || "").trim();
  return formula || '=ERROR("Empty formula from backend")';
}

// ===========================================================
// INSERT FORMULA BUTTON
// ===========================================================
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();

  const btn = document.createElement("button");
  btn.className = "btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      await safeExcelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.formulas = [[formula]];
        await ctx.sync();
      });
      showToast("✅ Formula inserted");
    } catch (e) {
      console.warn("Insert failed:", e);
      showToast("⚠️ Select a cell and try again.");
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// ===========================================================
// MAIN UI INITIALIZATION
// ===========================================================
async function initExcelWizPro() {
  console.log("🚀 Initializing ExcelWizPro UI…");

  const sheetSelect = getEl("sheetSelect");
  const queryInput = getEl("query");
  const output = getEl("output");
  const generateBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");

  let columnMapCache = "";
  let lastFormula = "";

  await refreshSheetDropdown(sheetSelect);
  warmUpBackend(); // fire-and-forget

  generateBtn.addEventListener("click", async () => {
    try {
      const query = queryInput.value.trim();
      if (!query) {
        showToast("⚠️ Please describe what you want the formula to do.");
        return;
      }

      if (typeof navigator !== "undefined" && !navigator.onLine) {
        showToast("📴 You appear to be offline.");
        return;
      }

      output.textContent = "⏳ Generating formula…";

      // Build smart column map once per taskpane session
      if (!columnMapCache) {
        columnMapCache = await buildColumnMap();
      }

      const { version: excelVersion } = getOfficeDiagnostics();

      const payload = {
        query,
        columnMap: columnMapCache,
        excelVersion,
        mainSheet: sheetSelect.value
      };

      const formula = await generateFormulaFromBackend(payload);
      lastFormula = formula;

      output.textContent = formula;
      attachInsertButton(output, formula);
    } catch (err) {
      console.error("❌ Formula generation failed:", err);
      output.textContent =
        "❌ Could not generate formula. See console for details.";
      showToast("⚠️ Problem contacting the backend.");
    }
  });

  clearBtn.addEventListener("click", () => {
    queryInput.value = "";
    output.textContent = "";
  });

  window.addEventListener("online", () => {
    if (lastFormula) {
      showToast("🌐 Back online — you can re-use your last formula.");
    }
  });

  console.log("🟢 ExcelWizPro UI ready.");
}

// ===========================================================
// MASTER BOOT
// ===========================================================
(async function boot() {
  console.log("🧠 ExcelWizPro boot sequence starting…");

  const info = await officeReady();
  const hostOK = await ensureExcelHost(info);
  if (!hostOK) return;

  const excelReady = await waitForExcelApi();
  if (!excelReady) return;

  console.table(getOfficeDiagnostics());

  await initExcelWizPro();
  showToast("✅ ExcelWizPro ready!");
  console.log("🟢 ExcelWizPro fully initialized.");
})();
