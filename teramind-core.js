// Teramind weekly connectivity dashboard: pure calculation logic.
// No DOM access, so it runs in the browser (window.TMCore) and in Node (require) for tests.
// Every number comes from the input rows. Nothing is estimated or invented.

(function (root) {
  const WEEKS = Array.from({ length: 10 }, (_, i) => `Week${i}`);
  const DEFAULT_START_SERIAL = 46290; // 25-Sep-26 as an Excel serial date
  const KNOWN_STATUSES = [
    "Installed",
    "Installed and not active",
    "Mail requested for Activation",
    "Pending with issues",
    "Exemption",
  ];

  // Header titles are compared in a compact form (lowercase, letters and digits only),
  // so "Talent ID#", "TalentID", "Talent_ID" and "Teramind-Exclude" all match.
  // `exact` is tried first across all headers; `match` is the looser fallback.
  const compact = (s) => String(s ?? "").toLowerCase().replace(/[^a-z0-9]+/g, "");
  const tokens = (s) => String(s ?? "").toLowerCase().split(/[^a-z0-9]+/).filter(Boolean);

  const FIELDS = [
    { key: "name", label: "Name", required: true, exact: ["name"], match: (c) => c === "item" || c === "itemname" },
    {
      key: "talentId", label: "Talent ID", required: true, exact: ["talentid"], hint: ["talent", "id"],
      match: (c, w) => c.includes("talentid") || (w.includes("talent") && w.includes("id")),
    },
    { key: "status", label: "Teramind Status", required: true, exact: ["teramindstatus"], hint: ["teramind", "status"], match: (c) => c.startsWith("teramindstatus") },
    { key: "plan", label: "Teramind plan", required: true, exact: ["teramindplan"], hint: ["teramind", "plan"], match: (c) => c.startsWith("teramindplan") },
    { key: "exclude", label: "Teramind - Exclude", required: true, exact: ["teramindexclude"], hint: ["exclude"], match: (c) => c.includes("teramind") && c.includes("exclude") },
    { key: "lastWorkingDate", label: "Last Working Date in USAA Account", required: true, exact: ["lastworkingdateinusaaaccount"], hint: ["working", "date", "lwd"], match: (c) => c.startsWith("lastworkingdate") },
    { key: "releaseCategory", label: "Release Category", required: true, exact: ["releasecategory"], hint: ["release"], match: (c) => c.startsWith("releasecategory") },
  ];

  // Special mapping value: the board has no Talent ID column, so use every item.
  // Offered only for live monday boards, where there are no header/summary rows to filter out.
  const NONE = "__none__";

  const norm = (s) => String(s ?? "").toLowerCase().replace(/\s+/g, " ").trim();
  const clean = (s) => String(s ?? "").replace(/\s+/g, " ").trim();

  // headers: [{ id, title }] -> { fieldKey: headerId | null }
  function autoMap(headers) {
    const map = {};
    for (const f of FIELDS) {
      const hit =
        headers.find((h) => f.exact.includes(compact(h.title))) ||
        headers.find((h) => f.match(compact(h.title), tokens(h.title)));
      map[f.key] = hit ? hit.id : null;
    }
    return map;
  }

  function missingFields(map) {
    return FIELDS.filter((f) => f.required && (map[f.key] === null || map[f.key] === undefined)).map((f) => f.label);
  }

  // Headers that look related to a field, to suggest when it couldn't be matched.
  function candidatesFor(fieldKey, headers) {
    const f = FIELDS.find((x) => x.key === fieldKey);
    if (!f || !f.hint) return [];
    return headers.filter((h) => tokens(h.title).some((w) => f.hint.includes(w)) || f.hint.some((k) => compact(h.title).includes(k) && k.length > 3));
  }

  // records: [{ rowRef, section, cells: { headerId: text } }]
  // Returns the valid rows plus counts of what was dropped and why.
  function cleanRows(records, map) {
    const valid = [];
    const dropped = { header: 0, noTalentId: 0, copy: 0 };
    const talentFilter = map.talentId !== NONE;
    for (const rec of records) {
      const get = (key) => clean(map[key] == null || map[key] === NONE ? "" : rec.cells[map[key]]);
      const talentId = get("talentId");
      const name = get("name");
      if (talentFilter && !talentId) { dropped.noTalentId++; continue; }
      if (talentFilter && norm(talentId) === "talent id") { dropped.header++; continue; }
      if (name.toLowerCase().includes("(copy)")) { dropped.copy++; continue; }
      valid.push({
        rowRef: rec.rowRef,
        section: rec.section || "",
        name,
        talentId,
        status: get("status"),
        plan: get("plan"),
        exclude: get("exclude"),
        lastWorkingDate: get("lastWorkingDate"),
        releaseCategory: get("releaseCategory"),
      });
    }
    return { valid, dropped };
  }

  // Case-insensitive equality, matching how Excel COUNTIFS compares text.
  const eq = (a, b) => String(a).toLowerCase() === String(b).toLowerCase();
  const isInstalled = (r) => eq(r.status, "Installed"); // "Installed and not active" is NOT counted
  const isExcluded = (r) => eq(r.exclude, "Exclude");
  const isRolloff = (r) => r.releaseCategory !== "" || r.lastWorkingDate !== "";
  const planKey = (r) => {
    const w = WEEKS.find((wk) => eq(r.plan, wk));
    if (w) return w;
    if (eq(r.plan, "NA")) return "NA";
    return r.plan === "" ? "(blank)" : "OTHER";
  };

  function metricsFor(rows) {
    return {
      plan: rows.length,
      actual: rows.filter(isInstalled).length,
      rolloffs: rows.filter(isRolloff).length,
      exemptions: rows.filter(isExcluded).length,
    };
  }

  function serialToDate(serial) {
    return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
  }
  function dateToSerial(d) {
    return Math.round((Date.UTC(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate()) - Date.UTC(1899, 11, 30)) / 86400000);
  }
  const MONTHS = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
  function formatSerial(serial) {
    const d = serialToDate(serial);
    return `${String(d.getUTCDate()).padStart(2, "0")}-${MONTHS[d.getUTCMonth()]}-${String(d.getUTCFullYear()).slice(-2)}`;
  }

  function countBy(values) {
    const m = new Map();
    for (const v of values) m.set(v, (m.get(v) || 0) + 1);
    return m;
  }

  function computeTeramind(rows, opts = {}) {
    const startSerial = opts.startSerial ?? DEFAULT_START_SERIAL;
    const todaySerial = opts.todaySerial ?? dateToSerial(new Date());

    // 1) Weekly table: one filter per week.
    let cumulative = 0;
    const weekly = WEEKS.map((week, i) => {
      const m = metricsFor(rows.filter((r) => eq(r.plan, week)));
      cumulative += m.actual;
      return { week, ...m, cumulative, weekEndingSerial: startSerial + 7 * i };
    });
    const grand = weekly.reduce(
      (t, w) => ({ plan: t.plan + w.plan, actual: t.actual + w.actual, rolloffs: t.rolloffs + w.rolloffs, exemptions: t.exemptions + w.exemptions }),
      { plan: 0, actual: 0, rolloffs: 0, exemptions: 0 }
    );
    grand.cumulative = cumulative;
    const exempt = metricsFor(rows.filter((r) => eq(r.plan, "NA")));

    // 2) Key metrics.
    const key = {
      validRows: rows.length,
      planned: grand.plan,
      installed: grand.actual,
      remaining: grand.plan - grand.actual,
      pctInstalled: grand.plan ? grand.actual / grand.plan : 0,
      rolloffs: rows.filter(isRolloff).length,
      exemptions: rows.filter(isExcluded).length,
    };

    // 3) Status breakdown: known statuses first, then anything else, then blank.
    const statusCounts = countBy(rows.map((r) => r.status));
    const status = [];
    const seen = new Set();
    for (const s of KNOWN_STATUSES) {
      const hit = [...statusCounts.keys()].find((k) => eq(k, s));
      status.push({ label: s, count: hit !== undefined ? statusCounts.get(hit) : 0 });
      if (hit !== undefined) seen.add(hit);
    }
    for (const [k, c] of statusCounts) {
      if (k !== "" && !seen.has(k)) status.push({ label: k, count: c, unknown: true });
    }
    status.push({ label: "Not started (blank)", count: statusCounts.get("") || 0, blank: true });
    const statusTotal = status.reduce((s, x) => s + x.count, 0);

    // 4) Profile of distinct values.
    const profile = {
      "Teramind Status": countBy(rows.map((r) => r.status || "(blank)")),
      "Teramind plan": countBy(rows.map((r) => r.plan || "(blank)")),
      "Teramind - Exclude": countBy(rows.map((r) => r.exclude || "(blank)")),
      "Release Category": countBy(rows.map((r) => r.releaseCategory || "(blank)")),
      "Last Working Date": countBy(rows.map((r) => (r.lastWorkingDate ? "filled" : "(blank)"))),
    };

    const result = { weekly, grand, exempt, key, status, statusTotal, profile, startSerial, todaySerial, talentFilterOff: !!opts.talentFilterOff };
    result.checks = reconcile(rows, result);
    result.issues = findIssues(rows, result);
    return result;
  }

  // Separate single-pass recount. It doesn't reuse the per-week filters, so a bug
  // in one method shows up as a mismatch instead of being hidden.
  function reconcile(rows, res) {
    const buckets = {};
    const bump = (k) => (buckets[k] ??= { plan: 0, actual: 0, rolloffs: 0, exemptions: 0 });
    for (const r of rows) {
      const b = bump(planKey(r));
      b.plan++;
      if (isInstalled(r)) b.actual++;
      if (isRolloff(r)) b.rolloffs++;
      if (isExcluded(r)) b.exemptions++;
    }
    const checks = [];
    const metrics = ["plan", "actual", "rolloffs", "exemptions"];
    const diffs = [];
    for (const w of res.weekly) {
      const b = buckets[w.week] || { plan: 0, actual: 0, rolloffs: 0, exemptions: 0 };
      for (const m of metrics) if (b[m] !== w[m]) diffs.push(`${w.week} ${m}: table ${w[m]} vs recount ${b[m]}`);
    }
    const na = buckets.NA || { plan: 0, actual: 0, rolloffs: 0, exemptions: 0 };
    for (const m of metrics) if (na[m] !== res.exempt[m]) diffs.push(`NA ${m}: table ${res.exempt[m]} vs recount ${na[m]}`);
    checks.push({ name: "Weekly Plan/Actual/Rolloffs/Exemptions match an independent row-by-row recount", pass: diffs.length === 0, detail: diffs.join("; ") });

    checks.push({
      name: "Status breakdown total equals valid rows",
      pass: res.statusTotal === rows.length,
      detail: `${res.statusTotal} vs ${rows.length}`,
    });

    const other = (buckets.OTHER?.plan || 0) + (buckets["(blank)"]?.plan || 0);
    const sum = res.grand.plan + res.exempt.plan + other;
    checks.push({
      name: "Week0–9 planned + NA + blank/other plan equals valid rows",
      pass: sum === rows.length,
      detail: `${res.grand.plan} + ${res.exempt.plan} + ${other} = ${sum} vs ${rows.length}`,
    });

    const lastCum = res.weekly[res.weekly.length - 1].cumulative;
    checks.push({
      name: "Final Cumulative equals Grand Total Actual",
      pass: lastCum === res.grand.actual,
      detail: `${lastCum} vs ${res.grand.actual}`,
    });
    return checks;
  }

  function findIssues(rows, res) {
    const issues = [];
    const names = (list) => list.map((r) => r.name);
    const add = (title, list, extra) => { if (list.length) issues.push({ title: `${title} (${list.length})`, items: extra ? list.map(extra) : names(list) }); };
    const withWeek = (r) => `${r.name} (${r.plan || "no plan"})`;

    add("Installed but not active, so not counted as Actual", rows.filter((r) => eq(r.status, "Installed and not active")), withWeek);
    add("Planned for a week but status is blank", rows.filter((r) => WEEKS.some((w) => eq(r.plan, w)) && r.status === ""), withWeek);
    add("Pending with issues", rows.filter((r) => eq(r.status, "Pending with issues")), withWeek);
    add("Mail requested for activation", rows.filter((r) => eq(r.status, "Mail requested for Activation")), withWeek);

    const unknown = rows.filter((r) => planKey(r) === "OTHER");
    if (unknown.length) {
      const vals = [...countBy(unknown.map((r) => r.plan))].map(([v, c]) => `"${v}": ${c}`);
      issues.push({ title: `Unrecognised Teramind plan values, not in any week (${unknown.length})`, items: vals });
    }
    const noPlan = rows.filter((r) => r.plan === "");
    add("No Teramind plan set", noPlan);

    const futureWeeks = res.weekly.filter((w) => w.weekEndingSerial > res.todaySerial).map((w) => w.week);
    add(
      "Rolling off but still planned for an upcoming week and not installed",
      rows.filter((r) => isRolloff(r) && futureWeeks.some((w) => eq(r.plan, w)) && !isInstalled(r)),
      withWeek
    );

    const unknownStatus = res.status.filter((s) => s.unknown);
    if (unknownStatus.length) issues.push({ title: "Unrecognised Teramind Status values", items: unknownStatus.map((s) => `"${s.label}": ${s.count}`) });

    for (const c of res.checks) if (!c.pass) issues.push({ title: `Reconciliation failed: ${c.name}`, items: [c.detail] });
    if (res.talentFilterOff) {
      issues.push({
        title: "Note: Talent ID filter not applied",
        items: ["Every board item except \"(copy)\" rows is counted as a resource. To filter by an ID column (e.g. Eagle ID), map it as Talent ID in Column mapping."],
      });
    }
    return issues;
  }

  // Compare against numbers the user quotes. One line per week, e.g.
  // "Week3: Plan 24, Actual 19, Rolloffs 1, Exemptions 0" or "Week3 24 19 1 0".
  function compareExpected(text, weekly) {
    const out = { compared: 0, diffs: [], unparsed: [] };
    for (const line of String(text).split(/\r?\n/)) {
      if (!line.trim()) continue;
      const m = line.match(/week\s*(\d+)/i);
      if (!m) { out.unparsed.push(line.trim()); continue; }
      const w = weekly.find((x) => x.week === `Week${Number(m[1])}`);
      if (!w) { out.unparsed.push(line.trim()); continue; }
      const nums = line.slice(m.index + m[0].length).match(/-?\d+/g) || [];
      const fields = ["plan", "actual", "rolloffs", "exemptions"];
      // Labelled form ("Plan 24") wins over positional order.
      const labelled = {};
      for (const f of fields) {
        const lm = line.match(new RegExp(f.replace(/s$/, "") + "s?\\D{0,3}(-?\\d+)", "i"));
        if (lm) labelled[f] = Number(lm[1]);
      }
      const vals = Object.keys(labelled).length ? labelled : Object.fromEntries(nums.slice(0, 4).map((n, i) => [fields[i], Number(n)]));
      for (const [f, v] of Object.entries(vals)) {
        out.compared++;
        if (w[f] !== v) out.diffs.push(`${w.week} ${f}: data says ${w[f]}, you quoted ${v}`);
      }
    }
    return out;
  }

  const api = {
    WEEKS, FIELDS, KNOWN_STATUSES, DEFAULT_START_SERIAL, NONE,
    autoMap, missingFields, candidatesFor, cleanRows, computeTeramind, compareExpected,
    serialToDate, dateToSerial, formatSerial,
  };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.TMCore = api;
})(typeof window !== "undefined" ? window : globalThis);
