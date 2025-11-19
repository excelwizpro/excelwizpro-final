/* ===========================================================
   ExcelWizPro — Stable Taskpane Controller (FINAL FIXED)
   =========================================================== */

console.log("🧠 ExcelWizPro Taskpane loaded");

const API_BASE = "https://excelwizpro-finalapi.onrender.com";

// DOM helpers
function $(id) {
  const el = document.getElementById(id);
  if (!el) throw new Error(`Missing element #${id}`);
  return el;
}

function toast(msg) {
  const t = document.createElement("div");
  t.className = "toast";
  t.textContent = msg;
  document.body.appendChild(t);
  setTimeout(() => t.remove(), 2600);
}

// -----------------------------------------------------------
// SAFE Excel.run wrapper
// -----------------------------------------------------------
async function safeExcelRun(cb) {
  try {
    return Excel.run(cb);
  } catch (err) {
    console.error("Excel.run failed:", err);
    toast("⚠️ Excel not ready");
    throw err;
  }
}

// -----------------------------------------------------------
// Column Map Cache
// -----------------------------------------------------------
let columnMapCache = "";
let mapTimestamp = 0;

async function buildColumnMap() {
  return safeExcelRun(async (ctx) => {
    const wb = ctx.workbook;
    const sheets = wb.worksheets;

    sheets.load("items/name");
    await ctx.sync();

    const lines = [];
    const nameCounts = Object.create(null);

    for (const sheet of sheets.items) {
      lines.push(`Sheet: ${sheet.name}`);

      const used = sheet.getUsedRangeOrNullObject();
      used.load("rowCount,columnCount,rowIndex,columnIndex,isNullObject");
      await ctx.sync();

      if (used.isNullObject || used.rowCount < 2) continue;

      const headerRows = Math.min(3, used.rowCount);
      const headers = sheet
        .getRangeByIndexes(
          used.rowIndex,
          used.columnIndex,
          headerRows,
          used.columnCount
        );
      headers.load("values");
      await ctx.sync();

      const hv = headers.values;

      const startRow = used.rowIndex + headerRows + 1;
      const lastRow = used.rowIndex + used.rowCount;

      for (let c = 0; c < used.columnCount; c++) {
        let names = [];
        for (let r = 0; r < headerRows; r++) {
          const t = String(hv[r][c] ?? "").trim();
          names.push(t);
        }
        names = names.reverse().filter((x) => x);
        if (!names.length) continue;

        let label = names[0];
        if (names.length > 1) {
          label = `${names[1]} - ${label}`;
        }

        let norm = label.toLowerCase().replace(/\s+/g, "_");
        if (nameCounts[norm]) {
          nameCounts[norm]++;
          norm += "__" + nameCounts[norm];
        } else {
          nameCounts[norm] = 1;
        }

        const xlCol = indexToLetter(used.columnIndex + c);
        lines.push(`${norm}='${sheet.name}'!${xlCol}${startRow}:${xlCol}${lastRow}`);
      }

      // tables
      const tables = sheet.tables;
      tables.load("items/name");
      await ctx.sync();

      for (const t of tables.items) {
        lines.push(`Table: ${t.name}`);
        const h = t.getHeaderRowRange();
        h.load("values");
        await ctx.sync();

        const cols = h.values[0];
        for (const col of cols) {
          if (!col) continue;
          const key = `${t.name}.${col}`.toLowerCase().replace(/\s+/g, "_");
          lines.push(`${key}=${t.name}[${col}]`);
        }
      }
    }

    return lines.join("\n");
  });
}

function indexToLetter(i) {
  let n = i + 1,
    s = "";
  while (n > 0) {
    const r = (n - 1) % 26;
    s = String.fromCharCode(65 + r) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}

// Refresh map
async function refreshColumnMap(force = false) {
  const now = Date.now();
  if (!force && columnMapCache && now - mapTimestamp < 20000) return;

  try {
    columnMapCache = await buildColumnMap();
    mapTimestamp = now;
  } catch (err) {
    console.error(err);
    toast("⚠️ Column map failed");
  }
}

// -----------------------------------------------------------
// Insert Button
// -----------------------------------------------------------
function attachInsertButton(container, formula) {
  container.querySelector(".btn-insert")?.remove();

  const btn = document.createElement("button");
  btn.className = "btn-insert";
  btn.textContent = "Insert into Excel";

  btn.onclick = async () => {
    try {
      await safeExcelRun(async (ctx) => {
        const range = ctx.workbook.getSelectedRange();
        range.load("rowCount,columnCount");
        await ctx.sync();

        if (range.rowCount !== 1 || range.columnCount !== 1) {
          return toast("⚠️ Select a single cell");
        }

        // Fix Excel API escaping issues
        const safeFormula = formula.replace(/"/g, '""');

        range.formulas = [[safeFormula]];
        await ctx.sync();
      });

      toast("✅ Inserted");
    } catch (err) {
      console.error(err);
      toast("⚠️ Could not insert");
    }
  };

  container.appendChild(document.createElement("br"));
  container.appendChild(btn);
}

// -----------------------------------------------------------
// Backend Call
// -----------------------------------------------------------
async function generateFormula(payload) {
  const r = await fetch(`${API_BASE}/generate`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload),
  });

  const data = await r.json();
  return data.formula || '=ERROR("No formula returned")';
}

// -----------------------------------------------------------
// UI Init
// -----------------------------------------------------------
async function initUI() {
  const sheetSelect = $("sheetSelect");
  const queryInput = $("query");
  const output = $("output");
  const generateBtn = $("generateBtn");
  const clearBtn = $("clearBtn");

  // Fill dropdown
  await safeExcelRun(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    sheetSelect.innerHTML = "";
    sheets.items.forEach((s) => {
      const op = document.createElement("option");
      op.value = s.name;
      op.textContent = s.name;
      sheetSelect.appendChild(op);
    });
  });

  // Build column map immediately
  await refreshColumnMap(true);

  generateBtn.onclick = async () => {
    const q = queryInput.value.trim();
    if (!q) return toast("Enter a request");

    output.textContent = "⏳ Generating…";

    await refreshColumnMap();

    const payload = {
      query: q,
      columnMap: columnMapCache,
      excelVersion: "web/desktop",
      mainSheet: sheetSelect.value,
    };

    try {
      const formula = await generateFormula(payload);
      output.textContent = formula;
      attachInsertButton(output, formula);
    } catch (err) {
      console.error(err);
      output.textContent = "❌ Error";
    }
  };

  clearBtn.onclick = () => {
    queryInput.value = "";
    output.textContent = "";
  };
}

// Start when Excel ready
Office.onReady(() => {
  initUI().then(() => toast("✅ ExcelWizPro Ready"));
});
