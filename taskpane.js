// ===========================================================
// ExcelWizPro Taskpane Script — v10.0.1 (Stable Boot Edition+)
// Focus: safer timeouts, offline handling, tolerant host detect,
//        no blocking alerts, centralized fetch wrapper
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";   // ✅ PATCHED BACKEND URL
const VERSION = "10.0.1";
console.log(`🧠 ExcelWizPro v${VERSION} taskpane.js loaded`);

Office.config = { extendedErrorLogging: true };

// -----------------------------------------------------------
// Global safety: surface swallowed promise errors (older Office.js)
// -----------------------------------------------------------
window.addEventListener("unhandledrejection", (e) => {
  console.warn("Unhandled promise rejection:", e.reason);
});
window.addEventListener("error", (e) => {
  console.warn("Window error:", e.message || e.error);
});

// -----------------------------------------------------------
// AbortSignal.timeout() fallback for older WebView2
// -----------------------------------------------------------
function timeoutSignal(ms) {
  if (typeof AbortSignal !== "undefined" && typeof AbortSignal.timeout === "function") {
    return AbortSignal.timeout(ms);
  }
  const ctrl = new AbortController();
  setTimeout(() => ctrl.abort(), ms);
  return ctrl.signal;
}

// -----------------------------------------------------------
// Centralized fetch wrapper: offline + timeout + consistent errors
// -----------------------------------------------------------
async function safeFetch(url, { timeout = 8000, ...opts } = {}) {
  if (!navigator.onLine) {
    const err = new Error("offline");
    err.code = "OFFLINE";
    throw err;
  }
  const signal = opts.signal || timeoutSignal(timeout);
  return fetch(url, { ...opts, signal });
}

// ===========================================================
// 🧩 Diagnostics helpers
// ===========================================================
function getOfficeDiagnostics() {
  try {
    return {
      host: Office.context?.host || "unknown",
      platform: Office.context?.diagnostics?.platform || "unknown",
      version: Office.context?.diagnostics?.version || "unknown",
      build: Office.context?.diagnostics?.build || "n/a",
    };
  } catch {
    return { host: "unknown", platform: "unknown", version: "unknown" };
  }
}

// ===========================================================
// 🟢 ExcelWizPro Safe Boot Coordinator
// ===========================================================
(async function bootExcelWizPro() {
  console.log("🚀 Booting ExcelWizPro...");

  await new Promise((resolve) => {
    try { Office.onReady(resolve); } catch { resolve(); }
  });

  for (let i = 1; i <= 10; i++) {
    const host = Office.context?.host;
    if (host && /Excel|Workbook/i.test(String(host))) {
      console.log("✅ Excel host detected.");
      break;
    }
    console.log(`⏳ Waiting for Excel host... (${i}/10)`);
    await delay(1000);
  }

  if (!(Office.context?.host && /Excel|Workbook/i.test(String(Office.context.host)))) {
    console.warn("⚠️ Excel host still not detected — continuing anyway.");
  }

  showToast("🔄 Initializing ExcelWizPro…");

  warmUpBackend().catch((err) => console.warn("⚠️ Backend warm-up failed:", err));

  try {
    await waitForWorkbookReady();
    await ensureExcelIsReady();
  } catch (err) {
    console.warn("⚠️ Excel context slow to respond:", err);
  }

  try {
    console.table(getOfficeDiagnostics());
    await initExcelWizPro();
    showToast("✅ ExcelWizPro ready!");
    console.log("🟢 ExcelWizPro fully initialized.");
  } catch (err) {
    console.error("❌ Excel initialization failed:", err);
    showToast("⚠️ Excel not responding — please reload ExcelWizPro.");
  }
})();

// ===========================================================
// 🩵 Smart Excel Ready Guard — Pings COM Bridge
// ===========================================================
async function waitForWorkbookReady(maxTries = 10, delayMs = 1000) {
  for (let i = 1; i <= maxTries; i++) {
    try {
      await Excel.run(async (ctx) => {
        const app = ctx.workbook.application;
        app.load("calculationMode");
        await ctx.sync();
      });
      console.log(`✅ Excel workbook responded to ping (try ${i})`);
      return true;
    } catch (err) {
      console.warn(`⏳ Excel workbook not ready (try ${i}/${maxTries}) — ${err.code || err.message}`);
    }
    await delay(delayMs * (1 + i / 2));
  }
  showToast("⚠️ Excel not responding — startup timed out.");
  throw new Error("Excel workbook did not respond after retries");
}

// ===========================================================
// 🧱 Robust Excel context guard
// ===========================================================
async function ensureExcelIsReady(maxTries = 12, delayMs = 1000) {
  for (let i = 1; i <= maxTries; i++) {
    try {
      await Excel.run(async (ctx) => {
        const app = ctx.workbook.application;
        app.load("calculationMode");
        await ctx.sync();
        app.calculationMode = app.calculationMode;
        await ctx.sync();
      });
      console.log(`✅ Excel context active (attempt ${i})`);
      return true;
    } catch (err) {
      console.warn(`⏳ Excel not ready (attempt ${i}/${maxTries})`, err.code || err.message);
      await delay(delayMs * (1 + i / 2));
    }
  }
  showToast("⚠️ Excel API context not ready after multiple attempts.");
  throw new Error("Excel context unavailable after retries");
}

// ===========================================================
// Backend warm-up
// ===========================================================
async function warmUpBackend(max = 10, baseDelay = 2500) {
  const statusDiv = document.createElement("div");
  Object.assign(statusDiv.style, {
    padding: "6px",
    marginBottom: "8px",
    borderRadius: "6px",
    fontSize: "0.9rem",
    fontWeight: "500",
    textAlign: "center",
  });
  document.querySelector("main.container")?.prepend(statusDiv);

  for (let i = 1; i <= max; i++) {
    try {
      const res = await safeFetch(`${API_BASE}/health`, {
        cache: "no-store",
        timeout: 3000,
      });
      if (res.ok) {
        statusDiv.textContent = "✅ Backend awake";
        statusDiv.style.backgroundColor = "#e6ffed";
        statusDiv.style.color = "#007a3d";
        setTimeout(() => statusDiv.remove(), 2500);
        console.log("✅ Backend warm-up complete");
        return;
      }
      throw new Error(`HTTP ${res.status}`);
    } catch (err) {
      const offline = err?.code === "OFFLINE";
      statusDiv.textContent = offline
        ? `📴 Offline — reconnect to contact backend`
        : `⏳ Waking backend (attempt ${i}/${max})…`;
      statusDiv.style.backgroundColor = "#fff3cd";
      statusDiv.style.color = "#856404";
      await delay(baseDelay * (1 + Math.random()));
    }
  }

  statusDiv.textContent = "❌ Cannot reach backend";
  statusDiv.style.backgroundColor = "#fdecea";
  statusDiv.style.color = "#d32f2f";
}

// ===========================================================
// ExcelWizPro Core Initialization
// ===========================================================
async function initExcelWizPro() {
  console.log("🚀 Initializing ExcelWizPro environment…");

  if (!Office.context || !Excel?.run) {
    showToast("⚠️ Excel APIs not ready yet. Try reopening ExcelWizPro.");
    return;
  }

  await ensureExcelContext();
  registerUIHandlers();
  console.log("🟢 ExcelWizPro UI active.");
}

// ===========================================================
// Excel context check
// ===========================================================
async function ensureExcelContext(retries = 3) {
  for (let i = 1; i <= retries; i++) {
    try {
      await Excel.run(async (ctx) => {
        const sheet = ctx.workbook.worksheets.getActiveWorksheet();
        sheet.load("name");
        await ctx.sync();
        console.log(`✅ Connected to active sheet: ${sheet.name}`);
      });
      return true;
    } catch (err) {
      console.warn(`Attempt ${i} failed:`, err);
      if (i === retries) throw err;
      await delay(1000 * i);
    }
  }
}

// ===========================================================
// UI + Excel logic
// ===========================================================
function registerUIHandlers() {
  console.log("🔧 Registering UI controls…");

  const sheetSelect = getEl("sheetSelect");
  const queryInput = getEl("query");
  const output = getEl("output");
  const generateBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");

  let columnMapCache = "";
  let lastFormula = "";

  refreshSheetDropdown(sheetSelect);

  generateBtn.addEventListener("click", async () => {
    try {
      if (!navigator.onLine) return showToast("⚠️ Offline — reconnect to generate formulas.");

      const query = queryInput.value.trim();
      if (!query) return showToast("⚠️ Please enter a description.");
      output.textContent = "⏳ Generating formula…";

      if (!columnMapCache) columnMapCache = await buildColumnMap();

      const { version: excelVersion } = getOfficeDiagnostics();
      const payload = {
        query,
        columnMap: columnMapCache,
        excelVersion,
        mainSheet: sheetSelect.value,
      };

      const formula = await generateFormulaWithRetry(payload, 4);

      if (!formula || formula.startsWith("=ERROR"))
        return (output.textContent = `⚠️ ${formula || "No valid formula."}`);

      output.textContent = formula;
      lastFormula = formula;
      attachInsertButton(output, formula);
    } catch (err) {
      console.error("❌ Formula generation failed:", err);
      output.textContent = "❌ Could not generate formula. Check console.";
    }
  });

  clearBtn.addEventListener("click", () => {
    queryInput.value = "";
    output.textContent = "";
  });

  window.addEventListener("online", () => {
    if (lastFormula) showToast("🌐 Back online! You can insert your last formula again.");
  });
}

// ===========================================================
// Backend call with retry
// ===========================================================
async function generateFormulaWithRetry(payload, maxRetries = 3) {
  let lastError = null;
  for (let i = 1; i <= maxRetries; i++) {
    try {
      const response = await safeFetch(`${API_BASE}/generate`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        cache: "no-store",
        timeout: 8000,
        body: JSON.stringify(payload),
      });

      if (response.status === 504) {
        console.warn(`⏳ Backend timeout (attempt ${i})`);
        showToast("☕ Waking backend… retrying.");
        await warmUpBackend(2, 1500);
        await delay(1500 * i);
        continue;
      }

      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      return data.formula?.trim() || "=ERROR('No response')";
    } catch (err) {
      lastError = err;
      const offline = err?.code === "OFFLINE";
      console.warn(`⚠️ Retry ${i}/${maxRetries} failed:`, offline ? "offline" : err.message);
      if (offline) showToast("📴 You’re offline. Reconnect and try again.");
      await delay(1500 * i);
    }
  }
  throw lastError || new Error("All retries failed");
}

// ===========================================================
// Safe Excel wrapper
// ===========================================================
async function safeExcelRun(cb) {
  try {
    return await Excel.run(async (ctx) => await cb(ctx));
  } catch (e) {
    console.warn("⚠️ Excel context lost:", e);
    showToast("Excel still loading — retrying…");
    try {
      await delay(1200);
      await Excel.run(async (ctx) => ctx.workbook.application.load("name"));
      return await Excel.run(cb);
    } catch (err) {
      showToast("⚠️ Excel still initializing — try again.");
      console.error("Excel retry failed:", err);
      throw err;
    }
  }
}

// ===========================================================
// Helpers
// ===========================================================
async function buildColumnMap() {
  return safeExcelRun(async (context) => {
    const sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    const result = [];
    for (const sheet of sheets.items) {
      result.push(`Sheet: ${sheet.name}`);
      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("values,address,isNullObject");
      await context.sync();

      if (used.isNullObject || !used.values?.length) continue;
      const headers = used.values[0] || [];
      headers.forEach((header, i) => {
        if (!header) return;
        const colLetter = String.fromCharCode(65 + i);
        const address = `'${sheet.name}'!${colLetter}2:INDEX('${sheet.name}'!${colLetter}:${colLetter},LOOKUP(2,1/('${sheet.name}'!${colLetter}:${colLetter}<>""),ROW('${sheet.name}'!${colLetter}:${colLetter})))`;
        result.push(`${header.toString().trim().toLowerCase()} = ${address}`);
      });
    }
    return result.join("\n");
  });
}

async function refreshSheetDropdown(select) {
  try {
    await safeExcelRun(async (ctx) => {
      const sheets = ctx.workbook.worksheets;
      sheets.load("items/name");
      await ctx.sync();
      select.innerHTML = "";
      sheets.items.forEach((s) => {
        const opt = document.createElement("option");
        opt.value = s.name;
        opt.textContent = s.name;
        select.appendChild(opt);
      });
    });
  } catch {
    showToast("⚠️ Could not read workbook sheets.");
  }
}

function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();
  const btn = document.createElement("button");
  btn.textContent = "Insert into Excel";
  btn.className = "btn-insert";
  btn.onclick = async () => {
    try {
      await safeExcelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.formulas = [[formula]];
        await ctx.sync();
      });
      showToast("✅ Formula inserted successfully!");
    } catch {
      showToast("⚠️ Could not insert formula. Select a cell first.");
    }
  };
  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

function showToast(msg) {
  ensureToastStyles();
  const toast = document.createElement("div");
  toast.className = "toast";
  toast.textContent = msg;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 2800);
}

let _toastStylesInjected = false;
function ensureToastStyles() {
  if (_toastStylesInjected) return;
  const style = document.createElement("style");
  style.textContent = `
    .toast {
      position: fixed;
      bottom: 20px;
      right: 20px;
      background: #323232;
      color: #fff;
      padding: 8px 14px;
      border-radius: 6px;
      font-size: 13px;
      z-index: 9999;
      opacity: 0.95;
      box-shadow: 0 4px 12px rgba(0,0,0,0.2);
    }
    .btn-insert {
      margin-top: 8px;
      padding: 6px 10px;
      border-radius: 6px;
      border: 1px solid #ccc;
      background: #f7f7f7;
      cursor: pointer;
    }
    .btn-insert:hover { filter: brightness(0.97); }
  `;
  document.head.appendChild(style);
  _toastStylesInjected = true;
}

function getEl(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element: #${id}`);
  return el;
}

function delay(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

