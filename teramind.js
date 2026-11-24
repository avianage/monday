// Teramind weekly connectivity dashboard: UI, data loading and Excel export.
// Calculations live in teramind-core.js (window.TMCore).
// Uses helpers from app.js at call time: $, el, cssVar, showError.

const TMUI = (() => {
  const C = window.TMCore;
  const SETTINGS_KEY = "tm_settings";
  const API_DATA_HEADERS = ["Name", "Section", "Talent ID", "Teramind Status", "Teramind plan", "Teramind - Exclude", "Last Working Date in USAA Account", "Release Category", "Valid"];

  const tm = {
    source: null, // { kind: "api" | "xlsx", label, headers, records, ...xlsx extras }
    map: {},
    valid: [],
    dropped: null,
    result: null,
    charts: [],
    settings: loadSettings(),
  };

  function loadSettings() {
    try {
      const s = JSON.parse(localStorage.getItem(SETTINGS_KEY) || "{}");
      return { startSerial: s.startSerial ?? C.DEFAULT_START_SERIAL, remarks: Array.isArray(s.remarks) ? s.remarks : [] };
    } catch {
      return { startSerial: C.DEFAULT_START_SERIAL, remarks: [] };
    }
  }
  function saveSettings() {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(tm.settings));
  }

  // ---------- Loading ----------

  function loadFromBoard(board, items) {
    const headers = [{ id: "__name", title: "Name", type: "name" }].concat(
      board.columns.filter((c) => c.type !== "name").map((c) => ({ id: c.id, title: c.title, type: c.type }))
    );
    const records = items.map((it, i) => {
      const cells = { __name: it.name };
      // Mirror / connect-boards columns return text = null; their value is in display_value.
      for (const cv of it.column_values) cells[cv.id] = cv.text || cv.display_value || "";
      return { rowRef: i, section: it.group?.title || "", cells };
    });
    tm.source = {
      kind: "api", label: `monday board "${board.name}"`, boardName: board.name, headers, records,
      mapKey: `tm_map:board:${board.id}`,
    };
    tm.map = applySavedMap(C.autoMap(headers));
    // Live boards have no header/section/summary rows, which is what the Talent ID filter
    // removes from the export. If the board has no Talent ID column, skip the filter.
    if (tm.map.talentId == null) tm.map.talentId = C.NONE;
    render();
  }

  // ---------- Remembered column mapping ----------

  function applySavedMap(auto) {
    let saved = {};
    try { saved = JSON.parse(localStorage.getItem(tm.source.mapKey) || "{}"); } catch { /* ignore */ }
    const ids = new Set(tm.source.headers.map((h) => String(h.id)));
    const map = { ...auto };
    for (const [k, v] of Object.entries(saved)) {
      if (v === C.NONE && tm.source.kind === "api" && k === "talentId") map[k] = v;
      else if (v !== null && ids.has(String(v))) map[k] = v; // ignore columns that no longer exist
    }
    return map;
  }
  function saveMap() {
    localStorage.setItem(tm.source.mapKey, JSON.stringify(tm.map));
  }

  function cellText(cell) {
    const v = cell.value;
    if (v === null || v === undefined) return "";
    if (v instanceof Date) return v.toISOString().slice(0, 10);
    if (typeof v === "object") {
      if (v.richText) return v.richText.map((t) => t.text).join("");
      if (v.text !== undefined) return String(v.text); // hyperlink
      if (v.result !== undefined) return v.result instanceof Date ? v.result.toISOString().slice(0, 10) : String(v.result ?? "");
      if (v.error) return "";
    }
    return String(v);
  }

  async function loadFromFile(file) {
    if (typeof ExcelJS === "undefined") throw new Error("The Excel library didn't load. Check your internet connection and reload the page.");
    const buffer = await file.arrayBuffer();
    const wb = new ExcelJS.Workbook();
    try {
      await wb.xlsx.load(buffer.slice(0));
    } catch (e) {
      throw new Error(`Couldn't read "${file.name}" as an .xlsx file. Re-export it from monday.com as Excel. (${e.message})`);
    }
    const ws = wb.worksheets[0];
    if (!ws) throw new Error("The workbook has no sheets.");

    // Read every row as { rowNumber, cells: { col: text } }.
    const rows = [];
    ws.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      const cells = {};
      row.eachCell({ includeEmpty: false }, (cell, col) => {
        const t = cellText(cell).trim();
        if (t) cells[String(col)] = t;
      });
      rows.push({ rowNumber, cells });
    });

    const isHeader = (r) => Object.values(r.cells).some((t) => t.toLowerCase() === "talent id");
    const headerRow = rows.find(isHeader);
    if (!headerRow) {
      throw new Error(`INSUFFICIENT_DATA: no header row containing "Talent ID" was found in the first sheet ("${ws.name}"). Make sure this is the PMO onboarding export.`);
    }
    const headers = Object.entries(headerRow.cells).map(([col, title]) => ({
      id: col, title, letter: ws.getColumn(Number(col)).letter,
    }));

    const records = [];
    let section = "";
    for (const r of rows) {
      if (r === headerRow) continue;
      const filled = Object.values(r.cells);
      if (filled.length === 1 && !isHeader(r)) section = filled[0]; // a section title row, e.g. "New Joinee"
      if (r.rowNumber < headerRow.rowNumber) continue;
      records.push({ rowRef: r.rowNumber, section, cells: r.cells });
    }

    tm.source = {
      kind: "xlsx",
      label: `file "${file.name}" (sheet "${ws.name}")`,
      fileName: file.name,
      fileBuffer: buffer,
      sheetName: ws.name,
      headerRowNumber: headerRow.rowNumber,
      lastRow: Math.max(ws.rowCount, headerRow.rowNumber + 1),
      maxCol: ws.columnCount,
      headers,
      records,
      mapKey: `tm_map:xlsx:${headers.map((h) => h.letter + "=" + h.title).join("|")}`,
    };
    tm.map = applySavedMap(C.autoMap(headers));
    render();
  }

  // ---------- Rendering ----------

  function render() {
    if (!tm.source) return;
    renderMapping();
    const box = $("tm-insufficient");
    const stop = (lines) => {
      box.replaceChildren(...lines.map((l) => el("div", {}, l)));
      box.classList.remove("hidden");
      $("tm-content").classList.add("hidden");
      $("tm-mapping-card").open = true;
      $("tm-source").textContent = `Source: ${tm.source.label}`;
      tm.result = null;
    };

    const missingKeys = C.FIELDS.filter((f) => f.required && tm.map[f.key] == null);
    if (missingKeys.length) {
      const lines = [`INSUFFICIENT_DATA: couldn't find these columns in the ${tm.source.label}: ${missingKeys.map((f) => f.label).join(", ")}.`];
      for (const f of missingKeys) {
        const cands = C.candidatesFor(f.key, tm.source.headers);
        lines.push(cands.length
          ? `• ${f.label}: did you mean ${cands.slice(0, 6).map((h) => `"${h.title}"${h.letter ? ` (${h.letter})` : ""}${h.type ? ` [${h.type}]` : ""}`).join(", ")}?`
          : `• ${f.label}: no similar column found on this ${tm.source.kind === "api" ? "board" : "sheet"}.`);
      }
      lines.push(tm.source.kind === "api" && missingKeys.some((f) => f.key === "talentId")
        ? `Pick the right column in "Column mapping" above (it's remembered for this board). If the board truly has no Talent ID column, choose "(none: use every item)". On a live board there are no header/summary rows to filter out.`
        : `Pick the right column in "Column mapping" above (it's remembered), or check that this is the right board/export.`);
      return stop(lines);
    }

    const { valid, dropped } = C.cleanRows(tm.source.records, tm.map);
    const talentOff = tm.map.talentId === C.NONE;
    if (!talentOff && tm.source.records.length && dropped.noTalentId === tm.source.records.length) {
      const h = tm.source.headers.find((x) => String(x.id) === String(tm.map.talentId));
      return stop([
        `INSUFFICIENT_DATA: "${h ? h.title : "?"}" is mapped as Talent ID but it's empty on all ${tm.source.records.length} rows.`,
        `It may be a formula or lookup column whose value the API doesn't return, or the wrong column. Pick another column in "Column mapping" above.`,
      ]);
    }

    box.classList.add("hidden");
    $("tm-content").classList.remove("hidden");
    tm.valid = valid;
    tm.dropped = dropped;
    tm.result = C.computeTeramind(valid, { startSerial: tm.settings.startSerial, talentFilterOff: talentOff });

    $("tm-source").textContent = talentOff
      ? `Source: ${tm.source.label} · ${valid.length} resource rows used · Talent ID filter not used (no Talent ID column mapped); only ${dropped.copy} "(copy)" rows dropped`
      : `Source: ${tm.source.label} · ${valid.length} resource rows used · dropped ${dropped.noTalentId} without Talent ID ` +
        `(section titles, summaries, blanks), ${dropped.header} repeated header rows, ${dropped.copy} "(copy)" rows`;
    const d = C.serialToDate(tm.settings.startSerial);
    $("tm-start").value = d.toISOString().slice(0, 10);

    renderWeekly();
    renderKPIs();
    renderStatus();
    renderChecks();
    renderCharts();
    renderIssues();
    renderProfile();
    renderFormulas();
  }

  function renderMapping() {
    const box = $("tm-mapping");
    box.replaceChildren();
    for (const f of C.FIELDS) {
      const wrap = el("label", { class: "map-row" });
      wrap.append(el("span", {}, f.label));
      const sel = el("select");
      sel.append(el("option", { value: "" }, "(not found: pick a column)"));
      if (f.key === "talentId" && tm.source.kind === "api") {
        const none = el("option", { value: C.NONE }, "(none: use every item, no Talent ID filter)");
        if (tm.map.talentId === C.NONE) none.selected = true;
        sel.append(none);
      }
      for (const h of tm.source.headers) {
        const opt = el("option", { value: h.id }, h.letter ? `${h.letter}: ${h.title}` : h.title);
        if (String(tm.map[f.key]) === String(h.id)) opt.selected = true;
        sel.append(opt);
      }
      if (tm.map[f.key] == null) wrap.classList.add("unmapped");
      sel.addEventListener("change", () => {
        tm.map[f.key] = sel.value === "" ? null : sel.value;
        saveMap();
        render();
      });
      wrap.append(sel);
      box.append(wrap);
    }
  }

  function td(text, cls) {
    const c = el("td", cls ? { class: cls } : {}, text === undefined ? "" : String(text));
    return c;
  }

  function renderWeekly() {
    const r = tm.result;
    const table = $("tm-weekly");
    table.replaceChildren();
    const head = el("tr");
    ["Week #", "Plan Installations", "Actual", "Cumulative", "Week Ending", "Rolloffs", "Exemptions", "Remarks"].forEach((h) => head.append(el("th", {}, h)));
    const thead = el("thead");
    thead.append(head);
    const tbody = el("tbody");

    r.weekly.forEach((w, i) => {
      const tr = el("tr");
      tr.append(td(w.week), td(w.plan, "num"), td(w.actual, "num"), td(w.cumulative, "num"), td(C.formatSerial(w.weekEndingSerial)), td(w.rolloffs, "num"), td(w.exemptions, "num"));
      const rc = el("td");
      const input = el("input", { type: "text", class: "remark", placeholder: "Add remark…" });
      input.value = tm.settings.remarks[i] || "";
      input.addEventListener("input", () => { tm.settings.remarks[i] = input.value; saveSettings(); });
      rc.append(input);
      tr.append(rc);
      tbody.append(tr);
    });

    const g = r.grand;
    const gt = el("tr", { class: "total" });
    gt.append(td("Grand Total"), td(g.plan, "num"), td(g.actual, "num"), td(g.cumulative, "num"), td(""), td(g.rolloffs, "num"), td(g.exemptions, "num"), td(""));
    const ex = el("tr", { class: "exempt" });
    ex.append(td("Exempt (Plan = NA)"), td(r.exempt.plan, "num"), td(r.exempt.actual, "num"), td("—"), td(""), td(r.exempt.rolloffs, "num"), td(r.exempt.exemptions, "num"), td(""));
    tbody.append(gt, ex);
    table.append(thead, tbody);
  }

  function renderKPIs() {
    const k = tm.result.key;
    const box = $("tm-kpis");
    box.replaceChildren();
    const tiles = [
      ["Total planned", k.planned, "Week0–Week9"],
      ["Installed", k.installed, 'status exactly "Installed"'],
      ["Remaining", k.remaining, "planned − installed"],
      ["% installed", `${(k.pctInstalled * 100).toFixed(1)}%`, `${k.installed} of ${k.planned}`],
      ["Rolloffs", k.rolloffs, "release category or last working date set"],
      ["Exemptions", k.exemptions, 'Teramind - Exclude = "Exclude"'],
    ];
    for (const [label, value, sub] of tiles) {
      const t = el("div", { class: "kpi" });
      t.append(el("div", { class: "label" }, label), el("div", { class: "value" }, String(value)), el("div", { class: "sub" }, sub));
      box.append(t);
    }
  }

  function renderStatus() {
    const r = tm.result;
    const table = $("tm-status");
    table.replaceChildren();
    const head = el("tr");
    head.append(el("th", {}, "Teramind Status"), el("th", { class: "num" }, "Count"));
    const tbody = el("tbody");
    for (const s of r.status) {
      const tr = el("tr");
      tr.append(td(s.label + (s.unknown ? " (unrecognised)" : "")), td(s.count, "num"));
      tbody.append(tr);
    }
    const tot = el("tr", { class: "total" });
    tot.append(td("Total"), td(r.statusTotal, "num"));
    const vr = el("tr");
    const ok = r.statusTotal === r.key.validRows;
    vr.append(td(`Valid resource rows (${ok ? "reconciles" : "DOES NOT reconcile"})`), td(r.key.validRows, "num"));
    tbody.append(tot, vr);
    const thead = el("thead");
    thead.append(head);
    table.append(thead, tbody);
  }

  function renderChecks() {
    const box = $("tm-checks");
    box.replaceChildren();
    for (const c of tm.result.checks) {
      const li = el("li", { class: c.pass ? "pass" : "fail" });
      li.append(el("span", { class: "badge" }, c.pass ? "✓ Pass" : "✗ Fail"), document.createTextNode(" " + c.name));
      if (!c.pass && c.detail) li.append(el("div", { class: "small muted" }, c.detail));
      box.append(li);
    }
  }

  function renderIssues() {
    const box = $("tm-issues");
    box.replaceChildren();
    const issues = tm.result.issues;
    if (!issues.length) {
      box.append(el("p", { class: "muted" }, "No issues found."));
      return;
    }
    for (const iss of issues) {
      const d = el("details", { class: "issue" });
      d.append(el("summary", {}, iss.title));
      const ul = el("ul");
      iss.items.slice(0, 100).forEach((t) => ul.append(el("li", {}, t)));
      if (iss.items.length > 100) ul.append(el("li", { class: "muted" }, `…and ${iss.items.length - 100} more`));
      d.append(ul);
      box.append(d);
    }
  }

  function renderProfile() {
    const box = $("tm-profile");
    box.replaceChildren();
    for (const [field, counts] of Object.entries(tm.result.profile)) {
      const card = el("div", { class: "profile-col" });
      card.append(el("h4", {}, field));
      const ul = el("ul");
      [...counts.entries()].sort((a, b) => b[1] - a[1]).forEach(([v, c]) => ul.append(el("li", {}, `${v}: ${c}`)));
      card.append(ul);
      box.append(card);
    }
  }

  function renderCharts() {
    tm.charts.forEach((c) => c.destroy());
    tm.charts = [];
    const r = tm.result;
    const labels = r.weekly.map((w) => w.week);
    const ink = cssVar("--text-secondary");
    const grid = cssVar("--grid");
    const s1 = cssVar("--series-1");
    const s2 = cssVar("--series-2");
    const scales = {
      x: { grid: { display: false }, ticks: { color: ink } },
      y: { beginAtZero: true, grid: { color: grid }, ticks: { color: ink, precision: 0 } },
    };
    const legend = { display: true, position: "top", align: "start", labels: { color: ink, boxWidth: 12, boxHeight: 12 } };

    tm.charts.push(new Chart($("tm-chart-plan"), {
      type: "bar",
      data: {
        labels,
        datasets: [
          { label: "Plan", data: r.weekly.map((w) => w.plan), backgroundColor: s1, borderRadius: 4, borderSkipped: "start", maxBarThickness: 22 },
          { label: "Actual", data: r.weekly.map((w) => w.actual), backgroundColor: s2, borderRadius: 4, borderSkipped: "start", maxBarThickness: 22 },
        ],
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false }, plugins: { legend }, scales },
    }));

    tm.charts.push(new Chart($("tm-chart-cum"), {
      type: "line",
      data: {
        labels,
        datasets: [{ label: "Cumulative installed", data: r.weekly.map((w) => w.cumulative), borderColor: s1, backgroundColor: s1, borderWidth: 2, pointRadius: 4, pointHoverRadius: 6, tension: 0 }],
      },
      options: { responsive: true, maintainAspectRatio: false, interaction: { mode: "index", intersect: false }, plugins: { legend: { display: false } }, scales },
    }));
  }

  // ---------- Formulas (shared by the web view and the Excel export) ----------

  function quoteSheet(name) {
    return `'${String(name).replace(/'/g, "''")}'!`;
  }
  function escCriterion(s) {
    return `"${String(s).replace(/[*?~]/g, "~$&").replace(/"/g, '""')}"`;
  }

  // Ranges for every field the formulas need, plus the helper "Valid" column.
  function buildRefs() {
    const src = tm.source;
    if (src.kind === "xlsx") {
      const letter = (key) => src.headers.find((h) => String(h.id) === String(tm.map[key])).letter;
      const first = src.headerRowNumber + 1;
      const last = src.lastRow;
      const sheet = quoteSheet(src.sheetName);
      const validCol = src.maxCol + 1;
      const validLetter = colLetter(validCol);
      const rng = (L) => `${sheet}$${L}$${first}:$${L}$${last}`;
      return {
        sheet, first, last, validCol, validLetter,
        S: rng(letter("status")), P: rng(letter("plan")), E: rng(letter("exclude")),
        LWD: rng(letter("lastWorkingDate")), RC: rng(letter("releaseCategory")), V: rng(validLetter),
        letters: { status: letter("status"), plan: letter("plan"), exclude: letter("exclude"), lastWorkingDate: letter("lastWorkingDate"), releaseCategory: letter("releaseCategory"), valid: validLetter },
      };
    }
    const first = 2;
    const last = Math.max(src.records.length + 1, 2);
    const rng = (L) => `Data!$${L}$${first}:$${L}$${last}`;
    return {
      sheet: "Data!", first, last,
      S: rng("D"), P: rng("E"), E: rng("F"), LWD: rng("G"), RC: rng("H"), V: rng("I"),
      letters: { status: "D", plan: "E", exclude: "F", lastWorkingDate: "G", releaseCategory: "H", valid: "I" },
    };
  }

  function colLetter(n) {
    let s = "";
    while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  function weekFormulas(R, weekRef) {
    return {
      plan: `COUNTIFS(${R.P},${weekRef},${R.V},1)`,
      actual: `COUNTIFS(${R.P},${weekRef},${R.S},"Installed",${R.V},1)`,
      rolloffs: `SUMPRODUCT((${R.P}=${weekRef})*(${R.V}=1)*(((${R.RC}<>"")+(${R.LWD}<>""))>0))`,
      exemptions: `COUNTIFS(${R.P},${weekRef},${R.E},"Exclude",${R.V},1)`,
    };
  }

  function renderFormulas() {
    const R = buildRefs();
    const f = weekFormulas(R, "$A5");
    const na = weekFormulas(R, '"NA"');
    const L = R.letters;
    const lines = [
      `Columns used: Teramind Status = ${L.status}, Teramind plan = ${L.plan}, Teramind - Exclude = ${L.exclude}, Last Working Date = ${L.lastWorkingDate}, Release Category = ${L.releaseCategory}, Valid helper = ${L.valid}`,
      `(Valid = 1 for real resource rows: Talent ID filled, not a header, name without "(copy)". The export copy gets this helper column so formulas skip section titles, headers and copies.)`,
      "",
      "Dashboard sheet, Week0 row (row 5; Week # in A5). Copy down for Week1–Week9:",
      `Plan        ${f.plan}`,
      `Actual      ${f.actual}`,
      `Cumulative  C5   (then D5+C6, D6+C7, …)`,
      `Week Ending ${tm.settings.startSerial}   (then E5+7, …; number format dd-mmm-yy)`,
      `Rolloffs    ${f.rolloffs}`,
      `Exemptions  ${f.exemptions}`,
      "",
      "Exempt (Plan = NA) row:",
      `Plan        ${na.plan}`,
      `Actual      ${na.actual}`,
      `Rolloffs    ${na.rolloffs}`,
      `Exemptions  ${na.exemptions}`,
      "",
      "Grand Total: SUM of each column over Week0–Week9; Cumulative = last week's Cumulative.",
    ];
    $("tm-formulas").textContent = lines.join("\n");
  }

  // ---------- Excel export ----------

  async function downloadExcel() {
    if (!tm.result) throw new Error("Build the Teramind dashboard first.");
    if (typeof ExcelJS === "undefined") throw new Error("The Excel library didn't load. Check your internet connection and reload the page.");
    const R = buildRefs();
    const validRefs = new Set(tm.valid.map((v) => v.rowRef));
    let wb;
    let fileName;

    if (tm.source.kind === "xlsx") {
      // Work on a COPY: reload the original bytes into a new workbook object.
      wb = new ExcelJS.Workbook();
      await wb.xlsx.load(tm.source.fileBuffer.slice(0));
      const ws = wb.getWorksheet(tm.source.sheetName) || wb.worksheets[0];
      ws.getCell(tm.source.headerRowNumber, R.validCol).value = "Valid (dashboard)";
      for (const ref of validRefs) ws.getCell(ref, R.validCol).value = 1;
      fileName = tm.source.fileName.replace(/\.xlsx$/i, "") + "_Dashboard.xlsx";
    } else {
      wb = new ExcelJS.Workbook();
      const data = wb.addWorksheet("Data");
      data.addRow(API_DATA_HEADERS);
      data.getRow(1).font = { bold: true };
      const get = (rec, key) => (tm.map[key] == null ? "" : rec.cells[tm.map[key]] ?? "");
      for (const rec of tm.source.records) {
        data.addRow([
          get(rec, "name"), rec.section, get(rec, "talentId"), get(rec, "status"), get(rec, "plan"),
          get(rec, "exclude"), get(rec, "lastWorkingDate"), get(rec, "releaseCategory"),
          validRefs.has(rec.rowRef) ? 1 : 0,
        ].map((v) => (typeof v === "string" ? v.trim() : v)));
      }
      data.columns.forEach((c, i) => (c.width = [32, 22, 14, 26, 14, 16, 18, 18, 8][i]));
      const safe = (tm.source.boardName || "monday").replace(/[\\/:*?"<>|]+/g, "_");
      fileName = `${safe}_Teramind_Dashboard.xlsx`;
    }

    writeDashboardSheet(wb, R);
    const buf = await wb.xlsx.writeBuffer();
    const blob = new Blob([buf], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = fileName;
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }

  function writeDashboardSheet(wb, R) {
    let name = "Dashboard";
    for (let i = 2; wb.getWorksheet(name); i++) name = `Dashboard (${i})`;
    const ws = wb.addWorksheet(name);
    const r = tm.result;
    const bold = { bold: true };
    const set = (addr, formula, result, numFmt) => {
      const c = ws.getCell(addr);
      c.value = formula === null ? result : { formula, result };
      if (numFmt) c.numFmt = numFmt;
      return c;
    };

    ws.getCell("A1").value = "Teramind Connectivity Dashboard";
    ws.getCell("A1").font = { bold: true, size: 14 };
    ws.getCell("A2").value = `Source: ${tm.source.label} · generated ${new Date().toISOString().slice(0, 10)}`;

    const header = ["Week #", "Plan Installations", "Actual", "Cumulative", "Week Ending", "Rolloffs", "Exemptions", "Remarks"];
    ws.getRow(4).values = header;
    ws.getRow(4).font = bold;

    r.weekly.forEach((w, i) => {
      const row = 5 + i;
      const f = weekFormulas(R, `$A${row}`);
      ws.getCell(`A${row}`).value = w.week;
      set(`B${row}`, f.plan, w.plan);
      set(`C${row}`, f.actual, w.actual);
      set(`D${row}`, i === 0 ? `C${row}` : `D${row - 1}+C${row}`, w.cumulative);
      set(`E${row}`, i === 0 ? null : `E${row - 1}+7`, w.weekEndingSerial, "dd-mmm-yy");
      set(`F${row}`, f.rolloffs, w.rolloffs);
      set(`G${row}`, f.exemptions, w.exemptions);
      ws.getCell(`H${row}`).value = tm.settings.remarks[i] || "";
    });

    const g = r.grand;
    ws.getCell("A15").value = "Grand Total";
    set("B15", "SUM(B5:B14)", g.plan);
    set("C15", "SUM(C5:C14)", g.actual);
    set("D15", "D14", g.cumulative);
    set("F15", "SUM(F5:F14)", g.rolloffs);
    set("G15", "SUM(G5:G14)", g.exemptions);
    ws.getRow(15).font = bold;

    const na = weekFormulas(R, '"NA"');
    ws.getCell("A16").value = "Exempt (Plan = NA)";
    set("B16", na.plan, r.exempt.plan);
    set("C16", na.actual, r.exempt.actual);
    set("F16", na.rolloffs, r.exempt.rolloffs);
    set("G16", na.exemptions, r.exempt.exemptions);

    // Key metrics
    const k = r.key;
    ws.getCell("A18").value = "Key Metrics";
    ws.getCell("A18").font = bold;
    const km = [
      ["Total planned (Week0–Week9)", "B15", k.planned],
      ["Installed", "C15", k.installed],
      ["Remaining", "B19-B20", k.remaining],
      ["% installed", "IF(B19=0,0,B20/B19)", k.pctInstalled, "0.0%"],
      ["Rolloffs (all resource rows)", `SUMPRODUCT((${R.V}=1)*(((${R.RC}<>"")+(${R.LWD}<>""))>0))`, k.rolloffs],
      ["Exemptions (Exclude flag, all resource rows)", `COUNTIFS(${R.E},"Exclude",${R.V},1)`, k.exemptions],
      ["Valid resource rows", `COUNTIF(${R.V},1)`, k.validRows],
    ];
    km.forEach(([label, formula, val, fmt], i) => {
      ws.getCell(`A${19 + i}`).value = label;
      set(`B${19 + i}`, formula, val, fmt);
    });

    // Status breakdown
    const sRow = 19 + km.length + 1;
    ws.getCell(`A${sRow}`).value = "Teramind Status";
    ws.getCell(`B${sRow}`).value = "Count";
    ws.getRow(sRow).font = bold;
    r.status.forEach((s, i) => {
      const row = sRow + 1 + i;
      ws.getCell(`A${row}`).value = s.label;
      const crit = s.blank ? '""' : escCriterion(s.label);
      set(`B${row}`, `COUNTIFS(${R.S},${crit},${R.V},1)`, s.count);
    });
    const totRow = sRow + 1 + r.status.length;
    ws.getCell(`A${totRow}`).value = "Total";
    set(`B${totRow}`, `SUM(B${sRow + 1}:B${totRow - 1})`, r.statusTotal);
    ws.getRow(totRow).font = bold;
    ws.getCell(`A${totRow + 1}`).value = "Reconciles with valid resource rows?";
    set(`B${totRow + 1}`, `IF(B${totRow}=B${19 + km.length - 1},"Yes","No")`, r.statusTotal === k.validRows ? "Yes" : "No");

    ws.getCell(`D${sRow}`).value = "Charts: see the web dashboard (Plan vs Actual clustered columns; Cumulative line).";
    ws.getCell(`D${sRow}`).font = { italic: true, color: { argb: "FF52514E" } };

    ws.columns.forEach((c, i) => (c.width = [42, 18, 10, 12, 13, 10, 12, 40][i] || 12));
    ws.views = [{ state: "frozen", ySplit: 4 }];
    // Open the workbook on the Dashboard sheet, and have Excel recalculate every formula on open
    // (cached results of 0 aren't stored by ExcelJS, so cells would otherwise look empty).
    wb.views = [{ activeTab: wb.worksheets.indexOf(ws) }];
    wb.calcProperties = Object.assign({}, wb.calcProperties, { fullCalcOnLoad: true });
  }

  // ---------- Compare with user-quoted numbers ----------

  function compare() {
    const out = $("tm-compare-result");
    out.replaceChildren();
    if (!tm.result) return;
    const res = C.compareExpected($("tm-compare-input").value, tm.result.weekly);
    if (!res.compared) {
      out.append(el("p", { class: "muted" }, "Nothing to compare. Use one line per week, e.g. \"Week3: Plan 24, Actual 19, Rolloffs 1, Exemptions 0\"."));
      return;
    }
    if (!res.diffs.length) out.append(el("p", { class: "pass" }, `✓ All ${res.compared} numbers you quoted match the data.`));
    else {
      out.append(el("p", { class: "fail" }, `✗ ${res.diffs.length} of ${res.compared} numbers differ (the table shows what the data says):`));
      const ul = el("ul");
      res.diffs.forEach((d) => ul.append(el("li", {}, d)));
      out.append(ul);
    }
    if (res.unparsed.length) out.append(el("p", { class: "small muted" }, `Ignored lines: ${res.unparsed.join(" | ")}`));
  }

  function setStartDate(value) {
    if (!value) return;
    tm.settings.startSerial = C.dateToSerial(new Date(value + "T00:00:00Z"));
    saveSettings();
    render();
  }

  function reset() {
    tm.charts.forEach((c) => c.destroy());
    tm.charts = [];
    tm.source = null;
    tm.result = null;
  }

  return { loadFromBoard, loadFromFile, render, downloadExcel, compare, setStartDate, reset, hasData: () => !!tm.result, _tm: tm };
})();
