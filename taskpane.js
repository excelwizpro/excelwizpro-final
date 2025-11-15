// ===========================================================
// ExcelWizPro Taskpane Script — v10.0.1 (Stable Production Build)
// Safe for WebView2 + Office Add-in Store deployment
// ===========================================================

/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "10.0.1";

console.log(`🧠 ExcelWizPro v${VERSION} — loading taskpane.js`);

// -----------------------------------------------------------
// SAFETY: prevent WebView2 from crashing on early JS errors
// -----------------------------------------------------------
window.addEventListener("error", (e) => {
  console.warn("JS Error:", e.message);
});
window.addEventListener("unhandledrejection", (e) => {
  console.warn("Promise Rejection:", e.reason);
});

// -----------------------------------------------------------
// Utility: delay
// -----------------------------------------------------------
const delay = (ms) => new Promise((res) => setTimeout(res, ms));

// -----------------------------------------------------------
// SAFETY: Fetch wrapper (timeout + offline protection)
// -----------------------------------------------------------
function safeFetch(url, opts = {}) {
  const timeout = opts.timeout ?? 8000;

  if (!navigator.onLine) {
    const err = new Error("offline");
    err.code = "OFFLINE";
    throw err;
  }

  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), timeout);

  return fetch(url, {
    ...opts,
    signal: ctrl.signal,
  }).finally(() => clearTimeout(timer));
}

// -----------------------------------------------------------
// STARTUP SEQUENCE — PRODUCTION SAFE
// -----------------------------------------------------------
Office.onReady(async (info) => {
  console.log("📘 Office.onReady fired:", info);

  // prevent early Excel context access
  await delay(50);

  if (info.host !== Office.HostType.Excel) {
    console.warn("Not Excel — stopping initialization.");
    return;
  }

  // now safe to run Excel calls
  await safeInit();
});

// -----------------------------------------------------------
// MAIN STARTUP WRAPPER
// -----------------------------------------------------------
async function safeInit() {
  try {
    showToast("🔄 Loading ExcelWizPro…");

    await waitForExcelApi();
    await warmBackend();
    await initializeUI();

    showToast("✅ ExcelWizPro Ready");
    console.log("🟢 ExcelWizPro initialization complete.");
  } catch (err) {
    console.error("❌ Fatal startup error:", err);
    showToast("❌ Excel failed to initialize.");
  }
}

// -----------------------------------------------------------
// STEP 1 — Confirm Excel API is responsive
// -----------------------------------------------------------
async function waitForExcelApi() {
  for (let i = 1; i <= 10; i++) {
    try {
      await Excel.run(async (ctx) => {
        const app = ctx.workbook.application;
        app.load("calculationMode");
        await ctx.sync();
      });

      console.log("✅ Excel is ready on try", i);
      return;
    } catch (err) {
      console.warn(`⏳ Excel API not ready (attempt ${i})`);
      await delay(300 + i * 150);
    }
  }

  throw new Error("Excel failed to respond after 10 attempts");
}

// -----------------------------------------------------------
// STEP 2 — Wake backend
// -----------------------------------------------------------
async function warmBackend() {
  for (let i = 1; i <= 5; i++) {
    try {
      const res = await safeFetch(`${API_BASE}/health`, {
        cache: "no-store",
        timeout: 3000,
      });
      if (res.ok) {
        console.log("✅ Backend awake");
        return;
      }
    } catch (err) {
      console.warn(`Backend warm-up failed (try ${i})`);
    }
    await delay(500 + i * 200);
  }
  console.warn("⚠️ Backend unreachable — continuing offline mode");
}

// -----------------------------------------------------------
// STEP 3 — Build UI + event handlers
// -----------------------------------------------------------
async function initializeUI() {
  console.log("🔧 Initializing UI handlers");

  const sheetSelect = get("sheetSelect");
  const queryInput = get("query");
  const output = get("output");
  const generateBtn = get("generateBtn");
  const clearBtn = get("clearBtn");

  await refreshSheets(sheetSelect);

  let columnMapCache = "";

  // generate
  generateBtn.addEventListener("click", async () => {
    const q = queryInput.value.trim();
    if (!q) return showToast("⚠️ Enter a formula description");

    output.textContent = "⏳ Generating formula…";

    if (!columnMapCache) {
      columnMapCache = await buildColumnMap();
    }

    const payload = {
      query: q,
      columnMap: columnMapCache,
    };

    try {
      const formula = await callGenerator(payload);
      output.textContent = formula;
      attachInsertButton(output, formula);
    } catch (err) {
      output.textContent = "❌ Failed to generate formula";
      console.error(err);
    }
  });

  // clear
  clearBtn.addEventListener("click", () => {
    queryInput.value = "";
    output.textContent = "";
  });
}

// -----------------------------------------------------------
// BUILD COLUMN MAP FROM WORKBOOK
// -----------------------------------------------------------
async function buildColumnMap() {
  return Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    let output = [];

    for (const sheet of sheets.items) {
      output.push(`Sheet: ${sheet.name}`);

      const used = sheet.getUsedRangeOrNullObject(true);
      used.load("values,isNullObject");
      await ctx.sync();

      if (used.isNullObject) continue;

      const headers = used.values[0];
      headers.forEach((h, i) => {
        if (!h) return;
        const col = String.fromCharCode(65 + i);
        const address = `'${sheet.name}'!${col}2:${col}1048576`;
        output.push(`${h.toLowerCase()} = ${address}`);
      });
    }

    return output.join("\n");
  });
}

// -----------------------------------------------------------
// CALL BACKEND
// -----------------------------------------------------------
async function callGenerator(payload) {
  const res = await safeFetch(`${API_BASE}/generate`, {
    method: "POST",
    timeout: 8000,
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  if (!res.ok) {
    return "=ERROR('Backend error')";
  }

  const data = await res.json();
  return data.formula || "=ERROR('No formula')";
}

// -----------------------------------------------------------
// INSERT BUTTON
// -----------------------------------------------------------
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();

  const btn = document.createElement("button");
  btn.textContent = "Insert into Excel";
  btn.className = "btn-insert";

  btn.onclick = async () => {
    try {
      await Excel.run(async (ctx) => {
        ctx.workbook.getSelectedRange().formulas = [[formula]];
      });
      showToast("✅ Inserted!");
    } catch {
      showToast("⚠️ Select a cell first.");
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// -----------------------------------------------------------
// HELPERS
// -----------------------------------------------------------
function get(id) {
  return document.getElementById(id);
}

function showToast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2400);
}

