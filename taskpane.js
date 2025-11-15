// ===========================================================
// ExcelWizPro Taskpane Script — v11.0 Production Edition
// Focus: Reliable Excel host detection, backend warm-up,
// stable formula generation logic, and full UI behavior.
// ===========================================================
/* global Office, Excel, fetch */

const API_BASE = "https://excelwizpro-finalapi.onrender.com";
const VERSION = "11.0.0";

console.log(`🧠 ExcelWizPro v${VERSION} taskpane.js loaded`);

Office.config = { extendedErrorLogging: true };

// ===========================================================
// GLOBAL ERROR SAFETY
// ===========================================================
window.addEventListener("unhandledrejection", (e) => {
  console.warn("Unhandled promise:", e.reason);
});
window.addEventListener("error", (e) => {
  console.warn("Window error:", e.message || e.error);
});

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

let _toastStylesInjected = false;
function ensureToastStyles() {
  if (_toastStylesInjected) return;
  const style = document.createElement("style");
  style.textContent = `
    .ewp-toast {
      position: fixed;
      bottom: 20px;
      right: 20px;
      background: #323232;
      color: #fff;
      padding: 8px 14px;
      border-radius: 6px;
      font-size: 13px;
      z-index: 99999;
      opacity: 0.97;
      max-width: 260px;
    }
    .ewp-btn-insert {
      margin-top: 8px;
      padding: 6px 10px;
      border-radius: 6px;
      border: 1px solid #ccc;
      background: #f7f7f7;
      cursor: pointer;
    }
    .ewp-btn-insert:hover { filter: brightness(0.95); }
  `;
  document.head.appendChild(style);
  _toastStylesInjected = true;
}

function showToast(msg) {
  ensureToastStyles();
  const el = document.createElement("div");
  el.className = "ewp-toast";
  el.textContent = msg;
  document.body.appendChild(el);
  setTimeout(() => el.remove(), 2800);
}

// ===========================================================
// AbortSignal.timeout fallback
// ===========================================================
function timeoutSignal(ms) {
  if (AbortSignal.timeout) return AbortSignal.timeout(ms);
  const ctrl = new AbortController();
  setTimeout(() => ctrl.abort(), ms);
  return ctrl.signal;
}

// ===========================================================
// SAFE FETCH (CORS + TIMEOUT + OFFLINE)
// ===========================================================
async function safeFetch(url, opts = {}) {
  const timeout = opts.timeout || 8000;
  if (!navigator.onLine) {
    const err = new Error("offline");
    err.code = "OFFLINE";
    throw err;
  }
  const signal = opts.signal || timeoutSignal(timeout);
  return fetch(url, { ...opts, signal });
}

// ===========================================================
// OFFICE / EXCEL DIAGNOSTICS
// ===========================================================
function getOfficeDiagnostics() {
  try {
    return {
      host: Office.context?.host,
      platform: Office.context?.diagnostics?.platform,
      version: Office.context?.diagnostics?.version,
      build: Office.context?.diagnostics?.build,
    };
  } catch {
    return { host: "unknown" };
  }
}

// ===========================================================
// BACKEND WARMUP
// ===========================================================
async function warmUpBackend(max = 6, baseDelay = 2000) {
  try {
    const div = document.createElement("div");
    Object.assign(div.style, {
      padding: "6px",
      marginBottom: "10px",
      borderRadius: "6px",
      textAlign: "center",
      fontSize: "0.9rem",
    });
    document.querySelector("main.container")?.prepend(div);

    for (let i = 1; i <= max; i++) {
      try {
        const r = await safeFetch(`${API_BASE}/health`, {
          cache: "no-store",
          timeout: 3000,
        });
        if (r.ok) {
          div.textContent = "✅ Backend awake";
          div.style.background = "#e6ffed";
          div.style.color = "#007a33";
          setTimeout(() => div.remove(), 2500);
          return;
        }
        throw new Error("Bad response");
      } catch (err) {
        const offline = err.code === "OFFLINE";
        div.textContent = offline
          ? "📴 Offline — reconnect"
          : `⏳ Waking backend… (${i}/${max})`;
        div.style.background = "#fff3cd";
        div.style.color = "#8a6d3b";
        await delay(baseDelay * (1 + Math.random()));
      }
    }

    div.textContent = "❌ Cannot reach backend";
    div.style.background = "#fdecea";
    div.style.color = "#b71c1c";
  } catch (e) {
    console.warn("Warmup failed:", e);
  }
}

// ===========================================================
// SAFE Excel.run
// ===========================================================
async function safeExcelRun(cb) {
  try {
    return await Excel.run(cb);
  } catch (err) {
    console.warn("Excel context issue:", err);
    showToast("⚠️ Excel not ready — try again.");
    throw err;
  }
}

// ===========================================================
// COLUMN MAP BUILDER
// ===========================================================
async function buildColumnMap() {
  return safeExcelRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    const output = [];

    for (const s of sheets.items) {
      output.push(`Sheet: ${s.name}`);
      const used = s.getUsedRangeOrNullObject(true);
      used.load("values,isNullObject");
      await ctx.sync();

      if (used.isNullObject || !used.values?.length) continue;

      const headers = used.values[0] || [];

      headers.forEach((header, i) => {
        if (!header) return;
        const col = String.fromCharCode(65 + i);
        const range = `'${s.name}'!${col}2:INDEX('${s.name}'!${col}:${col},LOOKUP(2,1/('${s.name}'!${col}:${col}<>""),ROW('${s.name}'!${col}:${col})))`;
        output.push(`${header.toString().trim().toLowerCase()} = ${range}`);
      });
    }

    return output.join("\n");
  });
}

// ===========================================================
// UI: sheet dropdown
// ===========================================================
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
  } catch (err) {
    console.warn("Sheet dropdown load failed:", err);
    showToast("⚠️ Could not read sheets.");
  }
}

// ===========================================================
// Insert into Excel button
// ===========================================================
function attachInsertButton(container, formula) {
  container.querySelector(".ewp-btn-insert")?.remove();
  const btn = document.createElement("button");
  btn.className = "ewp-btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      await safeExcelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.formulas = [[formula]];
        await ctx.sync();
      });
      showToast("✅ Formula inserted!");
    } catch {
      showToast("⚠️ Select a cell first.");
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// ===========================================================
// BACKEND FORMULA GENERATION (Your Full Logic)
// ===========================================================
async function generateFormula(payload) {
  const res = await safeFetch(`${API_BASE}/generate`, {
    timeout: 8000,
    method: "POST",
    headers: { "Content-Type": "application/json" },
    cache: "no-store",
    body: JSON.stringify(payload),
  });

  if (!res.ok) throw new Error(`Backend ${res.status}`);

  const data = await res.json();
  return (data.formula || "").trim();
}

// ===========================================================
// INIT EXCELWIZPRO (Your Logic)
// ===========================================================
async function initExcelWizPro() {
  console.log("🚀 Initializing ExcelWizPro…");

  const sheetSelect = getEl("sheetSelect");
  const queryEl = getEl("query");
  const outputEl = getEl("output");
  const genBtn = getEl("generateBtn");
  const clearBtn = getEl("clearBtn");

  let columnMapCache = "";
  let lastFormula = "";

  await refreshSheetDropdown(sheetSelect);
  warmUpBackend(); // fire-and-forget

  genBtn.addEventListener("click", async () => {
    try {
      const query = queryEl.value.trim();
      if (!query) return showToast("⚠️ Enter a description.");

      outputEl.textContent = "⏳ Generating...";

      if (!columnMapCache) columnMapCache = await buildColumnMap();

      const { version: excelVersion } = getOfficeDiagnostics();

      const payload = {
        query,
        columnMap: columnMapCache,
        excelVersion,
        mainSheet: sheetSelect.value,
      };

      const formula = await generateFormula(payload);

      outputEl.textContent = formula;
      lastFormula = formula;
      attachInsertButton(outputEl, formula);
    } catch (err) {
      console.error("Generation failed:", err);
      outputEl.textContent = "❌ Could not generate formula.";
      showToast("⚠️ Backend error.");
    }
  });

  clearBtn.addEventListener("click", () => {
    queryEl.value = "";
    outputEl.textContent = "";
  });

  console.log("🟢 ExcelWizPro UI ready.");
}

// ===========================================================
// 🔥 PRODUCTION-SAFE BOOT LOADER (Patch)
// ===========================================================
console.log("🧠 ExcelWizPro starting boot sequence…");

// Step 1 — Office.js ready
function officeReady() {
  return new Promise((resolve) => {
    if (Office?.onReady) {
      Office.onReady((info) => {
        console.log("📘 Office.onReady:", info);
        resolve(info);
      });
    } else {
      let tries = 0;
      const timer = setInterval(() => {
        tries++;
        if (Office?.onReady) {
          clearInterval(timer);
          Office.onReady((info) => resolve(info));
        }
        if (tries > 40) {
          clearInterval(timer);
          resolve({ host: "unknown" });
        }
      }, 500);
    }
  });
}

// Step 2 — ensure Excel host
async function ensureExcelHost(info) {
  if (info.host !== Office.HostType.Excel) {
    console.warn("❌ Not Excel host:", info.host);
    showToast("⚠️ Excel host not detected.");
    return false;
  }
  console.log("🟢 Excel host OK");
  return true;
}

// Step 3 — wait for Excel API
async function waitForExcelApi() {
  for (let i = 1; i <= 15; i++) {
    try {
      await Excel.run(async (ctx) => {
        ctx.workbook.properties.load("title");
        await ctx.sync();
      });
      console.log("🟢 Excel API ready");
      return true;
    } catch {
      await delay(600);
    }
  }
  showToast("⚠️ Excel not ready — reopen the add-in.");
  return false;
}

// Step 4 — Start your logic
(async function boot() {
  const info = await officeReady();
  if (!(await ensureExcelHost(info))) return;
  if (!(await waitForExcelApi())) return;

  console.table(getOfficeDiagnostics());

  await initExcelWizPro();
  showToast("✅ ExcelWizPro ready!");
  console.log("🟢 ExcelWizPro initialized.");
})();
