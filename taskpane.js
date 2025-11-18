// ===========================================================
// ExcelWizPro Taskpane — v5.0 Production Stable
// Premium UI + Universal Engine + Web-Safe Insert
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "5.0";

console.log(`🧠 ExcelWizPro Taskpane v${VERSION} loaded`);

// Enable extended Office error logging (optional)
if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

// ===========================================================
// GLOBAL ERROR HANDLERS
// ===========================================================
window.addEventListener("error", (e) =>
  console.warn("Window error:", e.message || e.error)
);
window.addEventListener("unhandledrejection", (e) =>
  console.warn("Unhandled rejection:", e.reason)
);

// ===========================================================
// HELPERS
// ===========================================================
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

// Convert numeric column index → Excel column letters
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

// ===========================================================
// SAFE HTTP REQUESTS
// ===========================================================
function timeoutSignal(ms) {
  if (typeof AbortController === "undefined") return undefined;
  const ctrl = new AbortController();
  const id = setTimeout(() => ctrl.abort(), ms);
  ctrl.signal.addEventListener("abort", () => clearTimeout(id));
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

// ===========================================================
// DIAGNOSTICS UTIL
// ===========================================================
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
// OFFICE BOOTSTRAP
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
    console.warn("❌ Not Excel host:", info && info.host);
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

  const container = document.querySelector("main.container");
  if (container) container.prepend(status);

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
// SAFE Excel.run WRAPPER (stabilized)
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
// SMART COLUMN MAP (cached)
// ===========================================================
let columnMapCache = "";
let lastColumnMapBuild = 0;
const COLUMN_MAP_TTL_MS = 30000;
const MAX_DATA_ROWS_PER_COLUMN = 50000;

async function buildColumnMap() {
  return safeExcelRun(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;

    sheets.load("items/name,items/visibility");
    await ctx.sync();

    const lines = [];
    const globalNameCounts = Object.create(null);

    for (const sheet of sheets.items) {
      const vis = sheet.visibility;
      const visText = vis !== "Visible" ? ` (${vis.toLowerCase()})` : "";
      lines.push(`Sheet: ${sheet.name}${visText}`);

      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,rowIndex,columnIndex,isNullObject");
      await ctx.sync();

      if (used.isNullObject || used.rowCount < 2) continue;

      const headerRows = Math.min(3, used.rowCount);
      const headerRange = sheet.getRangeByIndexes(
        used.rowIndex,
        used.columnIndex,
        headerRows,
        used.columnCount
      );
      headerRange.load("values");
      await ctx.sync();

      const headers = headerRange.values;
      const dataStartRowIndex = used.rowIndex + headerRows;
      const dataLastRowIndex = used.rowIndex + used.rowCount - 1;
      const startRow = dataStartRowIndex + 1;

      const maxLastRow = startRow + MAX_DATA_ROWS_PER_COLUMN - 1;
      const lastRowCandidate = dataLastRowIndex + 1;
      const lastRow = Math.min(lastRowCandidate, maxLastRow);

      if (lastRow < startRow) continue;

      for (let col = 0; col < used.columnCount; col++) {
        const headerTexts = [];
        for (let r = 0; r < headerRows; r++) {
          const v = headers[r][col];
          headerTexts[r] = v ? String(v).trim() : "";
        }

        let primary = "";
        for (let r = headerRows - 1; r >= 0; r--) {
          if (headerTexts[r]) {
            primary = headerTexts[r];
            break;
          }
        }
        if (!primary) continue;

        let combined = primary;
        for (let r = 0; r < headerRows - 1; r++) {
          if (headerTexts[r] && headerTexts[r] !== primary) {
            combined = `${headerTexts[r]} - ${combined}`;
            break;
          }
        }

        let normalized = normalizeName(combined);
        if (globalNameCounts[normalized]) {
          globalNameCounts[normalized]++;
          normalized = `${normalized}__${globalNameCounts[normalized]}`;
        } else {
          globalNameCounts[normalized] = 1;
        }

        const colLetter = columnIndexToLetter(used.columnIndex + col);
        const safeSheetName = sheet.name.replace(/'/g, "''");

        lines.push(
          `${normalized} = '${safeSheetName}'!${colLetter}${startRow}:${colLetter}${lastRow}`
        );
      }
    }

    return lines.join("\n");
  });
}

async function autoRefreshColumnMap(force = false) {
  try {
    const now = Date.now();
    if (!force && columnMapCache && now - lastColumnMapBuild < COLUMN_MAP_TTL_MS) {
      console.log("🔄 Using cached Smart Column Map");
      return;
    }
    console.log("🔄 Building Smart Column Map…");
    columnMapCache = await buildColumnMap();
    lastColumnMapBuild = Date.now();
    console.log("✅ Column Map updated");
  } catch (err) {
    console.warn("Column map failed:", err);
    showToast("⚠️ Could not refresh column map");
  }
}

// Auto-build map when taskpane is shown
if (Office.addin?.onVisibilityModeChanged) {
  Office.addin.onVisibilityModeChanged(async (args) => {
    if (args.visibilityMode === "Taskpane") {
      if (await waitForExcelApi()) await autoRefreshColumnMap();
    }
  });
} else if (Office.addin?.onVisibilityChanged) {
  Office.addin.onVisibilityChanged(async (visible) => {
    if (visible && (await waitForExcelApi())) await autoRefreshColumnMap();
  });
}

// ===========================================================
// SHEET DROPDOWN REFRESH
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
  return data.formula || '=ERROR("No formula returned")';
}

// ===========================================================
// WEB-SAFE INSERT BUTTON (FINAL FIXED VERSION)
// ===========================================================
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();

  const btn = document.createElement("button");
  btn.className = "btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      // Force grid focus (Excel Web)
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, () => {});
      await delay(40);

      await Excel.run(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load(["rowCount", "columnCount"]);
        await ctx.sync();

        if (range.rowCount !== 1 || range.columnCount !== 1) {
          throw { code: "MULTI_CELL_SELECTION" };
        }

        await delay(25); // Web timing stabilization

        range.formulas = [[formula]];
        await ctx.sync();
      });

      showToast("✅ Inserted!");

    } catch (err) {
      console.warn("Insert error:", err);
      if (err?.code === "MULTI_CELL_SELECTION") {
        showToast("⚠️ Select a single cell first");
      } else {
        showToast("⚠️ Could not insert formula");
      }
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// ===========================================================
// MAIN UI INITIALIZATION
// ===========================================================
async function initExcelWizPro() {
  const sheetSelect = getEl("sheetSelect");
  const queryInput = getEl("query");
  const output = getEl("output");
  const genBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");

  let lastFormula = "";

  await refreshSheetDropdown(sheetSelect);
  await autoRefreshColumnMap(true);
  warmUpBackend();

  genBtn.addEventListener("click", async () => {
    const query = queryInput.value.trim();
    if (!query) return showToast("⚠️ Enter a request");

    if (!navigator.onLine) {
      return showToast("📴 Offline");
    }

    output.textContent = "⏳ Generating…";

    await autoRefreshColumnMap(false);
    if (!columnMapCache) await autoRefreshColumnMap(true);

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
// BOOT
// ===========================================================
(async function boot() {
  console.log("🧠 Starting ExcelWizPro…");

  const info = await officeReady();
  if (!(await ensureExcelHost(info))) return;
  if (!(await waitForExcelApi())) return;

  await initExcelWizPro();
  showToast("✅ ExcelWizPro ready!");
})();
