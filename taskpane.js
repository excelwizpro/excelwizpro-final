// ===========================================================
// ExcelWizPro Taskpane — v12.3.0 (Premium UI Compatible)
// - Web-safe insertion
// - Universal backend compatible
// - Smart column map
// - Focus safe for Excel Web
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "12.3.0";

console.log(`🧠 ExcelWizPro Taskpane v${VERSION} loaded`);

if (typeof Office !== "undefined" && Office && Office.config) {
  Office.config = { extendedErrorLogging: true };
}

// ===========================================================
// ERROR HANDLING
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
// HTTP WRAPPER
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
// DIAGNOSTICS
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
// BOOTSTRAP
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
    showToast("⚠️ Excel host not detected.");
    return false;
  }
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
// BACKEND WAKEUP
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
// EXCEL WEB FOCUS FIX
// ===========================================================
function forceGridFocus() {
  return new Promise((resolve) => {
    Office.context.document.getSelectedDataAsync(
      Office.CoercionType.Text,
      () => resolve()
    );
  });
}

// ===========================================================
// SAFE Excel.run (patched for Excel Web)
// ===========================================================
async function safeExcelRun(cb, attempt = 1) {
  try {
    await forceGridFocus();
    await delay(35);

    return await Excel.run(cb);
  } catch (err) {
    console.warn(`Excel.run attempt ${attempt} failed:`, err);

    if (attempt === 1) {
      await delay(80);
      return safeExcelRun(cb, 2);
    }

    showToast("⚠️ Excel not ready");
    throw err;
  }
}

// ===========================================================
// SMART COLUMN MAP BUILDER
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
          normalized += `__${globalNameCounts[normalized]}`;
        } else {
          globalNameCounts[normalized] = 1;
        }

        const colLetter = columnIndexToLetter(used.columnIndex + col);
        const safeSheetName = sheet.name.replace(/'/g, "''");

        lines.push(
          `${normalized} = '${safeSheetName}'!${colLetter}${startRow}:${colLetter}${lastRow}`
        );
      }

      // TABLES
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      const tableMeta = tables.items.map((table) => {
        return {
          table,
          header: table.getHeaderRowRange()
        };
      });

      tableMeta.forEach((m) => m.header.load("values"));
      await ctx.sync();

      for (const { table, header } of tableMeta) {
        lines.push(`Table: ${table.name}`);

        const headerVals = header.values[0] || [];
        headerVals.forEach((h) => {
          if (!h) return;

          let norm = normalizeName(`${table.name}.${h}`);
          if (globalNameCounts[norm]) {
            globalNameCounts[norm]++;
            norm += `__${globalNameCounts[norm]}`;
          } else {
            globalNameCounts[norm] = 1;
          }

          const structured = `${table.name}[${h}]`;
          lines.push(`${norm} = ${structured}`);
        });
      }

      // PIVOTS
      const pivots = sheet.pivotTables;
      pivots.load("items/name");
      await ctx.sync();
      pivots.items.forEach((p) => lines.push(`PivotSource: ${p.name}`));
    }

    // NAMED RANGES
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
      let norm = normalizeName(name);
      if (globalNameCounts[norm]) {
        globalNameCounts[norm]++;
        norm += `__${globalNameCounts[norm]}`;
      } else {
        globalNameCounts[norm] = 1;
      }
      lines.push(`${norm} = ${range.address}`);
    });

    return lines.join("\n");
  });
}

// ===========================================================
// SMART MAP AUTO REFRESH
// ===========================================================
async function autoRefreshColumnMap(force = false) {
  try {
    const now = Date.now();
    if (!force && columnMapCache && now - lastColumnMapBuild < COLUMN_MAP_TTL_MS) {
      return;
    }

    columnMapCache = await buildColumnMap();
    lastColumnMapBuild = Date.now();
  } catch (err) {
    console.warn("Column map refresh failed:", err);
    showToast("⚠️ Could not refresh column map");
  }
}

if (Office.addin?.onVisibilityModeChanged) {
  Office.addin.onVisibilityModeChanged(async (args) => {
    if (args.visibilityMode === "Taskpane") {
      if (await waitForExcelApi()) autoRefreshColumnMap();
    }
  });
} else if (Office.addin?.onVisibilityChanged) {
  Office.addin.onVisibilityChanged(async (visible) => {
    if (visible && (await waitForExcelApi())) autoRefreshColumnMap();
  });
}

// ===========================================================
// SHEET SELECT DROPDOWN
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
// UNIVERSAL BACKEND GENERATION
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
// PATCHED INSERT BUTTON (Excel Web safe)
// ===========================================================
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();

  const btn = document.createElement("button");
  btn.className = "btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      await Office.context.document.getSelectedDataAsync(
        Office.CoercionType.Text,
        () => {}
      );

      await delay(50);

      await safeExcelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("rowCount,columnCount");
        await ctx.sync();

        if (range.rowCount !== 1 || range.columnCount !== 1) {
          const err = new Error("MULTI_CELL_SELECTION");
          err.code = "MULTI_CELL_SELECTION";
          throw err;
        }

        await delay(20); 
        range.formulas = [[formula]];
        await ctx.sync();
      });

      showToast("✅ Inserted");
    } catch (err) {
      if (err.code === "MULTI_CELL_SELECTION") {
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
      output.textContent = "❌ Error (check console)";
      console.error("Generation error:", err);
    }
  });

  clearBtn.addEventListener("click", () => {
    output.textContent = "";
    queryInput.value = "";
  });

  window.addEventListener("online", () => {
    if (lastFormula) showToast("🌐 Back online");
  });
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


