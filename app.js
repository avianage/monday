// monday.com Board Dashboard — v1
// Runs entirely in the browser: calls https://api.monday.com/v2 directly with the
// user's own token. Nothing is sent anywhere else.

const API_URL = "https://api.monday.com/v2";
const TOKEN_KEY = "monday_dashboard_token";
const PAGE_SIZE = 500;
const DONE_PATTERN = /\b(done|complete|completed|closed|finished|resolved)\b/i;

const $ = (id) => document.getElementById(id);
const state = { token: "", demo: false, board: null, items: [], charts: [] };

// ---------- API ----------

async function mondayQuery(query, variables = {}) {
  let res;
  try {
    res = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", Authorization: state.token },
      body: JSON.stringify({ query, variables }),
    });
  } catch (e) {
    throw new Error("Could not reach monday.com. Check your internet connection.");
  }
  let body = null;
  try { body = await res.json(); } catch { /* non-JSON body */ }

  if (res.status === 401 || res.status === 403) throw new Error("Invalid or expired API token.");
  if (res.status === 429) throw new Error("monday.com rate limit hit. Wait a minute and try again.");
  if (!res.ok) throw new Error(`monday.com returned HTTP ${res.status}.`);
  if (body?.errors?.length) throw new Error(body.errors.map((e) => e.message).join("; "));
  if (body?.error_message) throw new Error(body.error_message);
  return body.data;
}

async function fetchMe() {
  const data = await mondayQuery(`query { me { name account { name } } }`);
  return data.me;
}

async function fetchBoards() {
  const data = await mondayQuery(
    `query { boards(limit: 100, order_by: used_at, state: active) { id name items_count type } }`
  );
  return data.boards.filter((b) => b.type !== "sub_items_board" && b.type !== "document");
}

// Mirror (lookup) and connect-boards columns return text = null, so also ask for display_value.
const ITEM_FIELDS = `id name group { title } column_values { id type text ... on MirrorValue { display_value } ... on BoardRelationValue { display_value } }`;

async function fetchBoardWithItems(boardId) {
  const data = await mondayQuery(
    `query ($ids: [ID!]) {
      boards(ids: $ids) {
        id name
        columns { id title type }
        groups { id title }
        items_page(limit: ${PAGE_SIZE}) { cursor items { ${ITEM_FIELDS} } }
      }
    }`,
    { ids: [String(boardId)] }
  );
  const board = data.boards[0];
  if (!board) throw new Error("Board not found or you don't have access to it.");

  const items = [...board.items_page.items];
  let cursor = board.items_page.cursor;
  while (cursor) {
    setLoading(`Loading items… (${items.length} so far)`);
    const page = await mondayQuery(
      `query ($cursor: String!) {
        next_items_page(limit: ${PAGE_SIZE}, cursor: $cursor) { cursor items { ${ITEM_FIELDS} } }
      }`,
      { cursor }
    );
    items.push(...page.next_items_page.items);
    cursor = page.next_items_page.cursor;
  }
  return { board, items };
}

// ---------- Data processing ----------

function countBy(values) {
  const map = new Map();
  for (const v of values) map.set(v, (map.get(v) || 0) + 1);
  return [...map.entries()].sort((a, b) => b[1] - a[1]);
}

function valueOf(item, colId) {
  const cv = item.column_values.find((c) => c.id === colId);
  const t = cv ? cv.text || cv.display_value : "";
  return t ? String(t).trim() : "";
}

function summarize(board, items) {
  const cols = board.columns;
  const statusCol = cols.find((c) => c.type === "status");
  const peopleCol = cols.find((c) => c.type === "people");
  const dateCol = cols.find((c) => c.type === "date");

  const summary = { total: items.length, groups: board.groups.length, statusCol, peopleCol, dateCol };

  summary.byGroup = countBy(items.map((i) => i.group?.title || "(no group)"));

  if (statusCol) {
    summary.byStatus = countBy(items.map((i) => valueOf(i, statusCol.id) || "(blank)"));
    summary.done = items.filter((i) => DONE_PATTERN.test(valueOf(i, statusCol.id))).length;
  }

  if (peopleCol) {
    const names = [];
    for (const i of items) {
      const v = valueOf(i, peopleCol.id);
      if (!v) names.push("Unassigned");
      else v.split(",").forEach((n) => names.push(n.trim()));
    }
    summary.byPerson = countBy(names);
  }

  if (dateCol) {
    const today = new Date().toISOString().slice(0, 10);
    const months = new Map();
    let overdue = 0;
    for (const i of items) {
      const d = valueOf(i, dateCol.id).slice(0, 10);
      if (!/^\d{4}-\d{2}-\d{2}$/.test(d)) continue;
      const m = d.slice(0, 7);
      months.set(m, (months.get(m) || 0) + 1);
      const isDone = statusCol && DONE_PATTERN.test(valueOf(i, statusCol.id));
      if (d < today && !isDone) overdue++;
    }
    summary.overdue = overdue;
    summary.byMonth = fillMonths(months);
  }
  return summary;
}

// Return [["2026-01", n], ...] with empty months filled in (max 36 months, most recent kept).
function fillMonths(months) {
  if (!months.size) return [];
  const keys = [...months.keys()].sort();
  const out = [];
  let [y, m] = keys[0].split("-").map(Number);
  const [ey, em] = keys[keys.length - 1].split("-").map(Number);
  while (y < ey || (y === ey && m <= em)) {
    const k = `${y}-${String(m).padStart(2, "0")}`;
    out.push([k, months.get(k) || 0]);
    m++; if (m > 12) { m = 1; y++; }
  }
  return out.slice(-36);
}

// ---------- Rendering ----------

function cssVar(name) {
  return getComputedStyle(document.documentElement).getPropertyValue(name).trim();
}

function el(tag, attrs = {}, text) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs)) e.setAttribute(k, v);
  if (text !== undefined) e.textContent = text;
  return e;
}

function renderKPIs(s) {
  const box = $("kpis");
  box.replaceChildren();
  const tiles = [
    ["Total items", s.total.toLocaleString(), `${s.groups} group${s.groups === 1 ? "" : "s"}`],
  ];
  if (s.statusCol) {
    const pct = s.total ? Math.round((s.done / s.total) * 100) : 0;
    tiles.push(["Completed", `${pct}%`, `${s.done} of ${s.total} marked done in "${s.statusCol.title}"`]);
    tiles.push(["Open", (s.total - s.done).toLocaleString(), "not marked done"]);
  }
  if (s.dateCol) {
    tiles.push(["Overdue", s.overdue.toLocaleString(), `past "${s.dateCol.title}" and not done`]);
  }
  if (s.peopleCol) {
    const people = s.byPerson.filter(([n]) => n !== "Unassigned").length;
    tiles.push(["People", people.toLocaleString(), `assigned in "${s.peopleCol.title}"`]);
  }
  for (const [label, value, sub] of tiles) {
    const t = el("div", { class: "kpi" });
    t.append(el("div", { class: "label" }, label), el("div", { class: "value" }, value), el("div", { class: "sub" }, sub));
    box.append(t);
  }
}

// Long category lists fold into "Other" so charts stay readable.
function topN(entries, n = 12) {
  if (entries.length <= n) return entries;
  const head = entries.slice(0, n - 1);
  const rest = entries.slice(n - 1).reduce((sum, [, c]) => sum + c, 0);
  return [...head, ["Other", rest]];
}

function addChart(title, type, entries, opts = {}) {
  const card = el("div", { class: "card" });
  card.append(el("h3", {}, title));
  const box = el("div", { class: "chart-box" });
  const canvas = el("canvas", { role: "img", "aria-label": `${title} chart` });
  box.append(canvas);
  card.append(box);
  $("charts").append(card);

  const color = cssVar("--series-1");
  const ink = cssVar("--text-secondary");
  const grid = cssVar("--grid");
  const horizontal = opts.horizontal;

  const chart = new Chart(canvas, {
    type,
    data: {
      labels: entries.map(([k]) => k),
      datasets: [{
        label: "Items",
        data: entries.map(([, v]) => v),
        backgroundColor: color,
        borderColor: color,
        borderWidth: type === "line" ? 2 : 0,
        borderRadius: type === "bar" ? 4 : 0,
        borderSkipped: "start",
        maxBarThickness: 28,
        pointRadius: type === "line" ? 4 : 0,
        pointHoverRadius: 6,
        tension: 0,
        fill: false,
      }],
    },
    options: {
      indexAxis: horizontal ? "y" : "x",
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        legend: { display: false },
        tooltip: { callbacks: { label: (ctx) => ` ${ctx.parsed[horizontal ? "x" : "y"]} items` } },
      },
      scales: {
        x: {
          grid: { display: !!horizontal, color: grid },
          ticks: { color: ink, precision: 0 },
          beginAtZero: true,
        },
        y: {
          grid: { display: !horizontal, color: grid },
          ticks: { color: ink, precision: 0 },
          beginAtZero: true,
        },
      },
    },
  });
  state.charts.push(chart);
}

function renderCharts(s) {
  state.charts.forEach((c) => c.destroy());
  state.charts = [];
  $("charts").replaceChildren();

  if (s.byStatus) addChart(`Items by ${s.statusCol.title}`, "bar", topN(s.byStatus), { horizontal: true });
  addChart("Items by group", "bar", topN(s.byGroup), { horizontal: true });
  if (s.byPerson) addChart(`Items by ${s.peopleCol.title}`, "bar", topN(s.byPerson), { horizontal: true });
  if (s.byMonth?.length) addChart(`Items by ${s.dateCol.title} (month)`, "line", s.byMonth);
}

function tableColumns(board) {
  const skip = new Set(["subtasks", "button", "doc", "file", "board_relation", "mirror", "dependency"]);
  return board.columns.filter((c) => c.type !== "name" && !skip.has(c.type)).slice(0, 10);
}

function renderTable(filter = "") {
  const table = $("items-table");
  table.replaceChildren();
  const cols = tableColumns(state.board);

  const thead = el("thead");
  const hr = el("tr");
  ["Item", "Group", ...cols.map((c) => c.title)].forEach((h) => hr.append(el("th", {}, h)));
  thead.append(hr);

  const q = filter.trim().toLowerCase();
  const tbody = el("tbody");
  let shown = 0;
  for (const item of state.items) {
    const cells = [item.name, item.group?.title || "", ...cols.map((c) => valueOf(item, c.id))];
    if (q && !cells.some((v) => v.toLowerCase().includes(q))) continue;
    const tr = el("tr");
    cells.forEach((v) => tr.append(el("td", {}, v)));
    tbody.append(tr);
    if (++shown >= 1000) break; // keep the DOM light on very large boards
  }
  if (!shown) {
    const tr = el("tr");
    tr.append(el("td", { colspan: String(cols.length + 2) }, "No matching items."));
    tbody.append(tr);
  }
  table.append(thead, tbody);
}

function renderDashboard() {
  const s = summarize(state.board, state.items);
  $("board-title").textContent = state.board.name;
  renderKPIs(s);
  renderCharts(s);
  $("search").value = "";
  renderTable();
  TMUI.loadFromBoard(state.board, state.items);
  $("tab-overview").classList.remove("hidden");
  $("views").classList.remove("hidden");
}

// ---------- Tabs ----------

function showTab(id) {
  document.querySelectorAll(".tab").forEach((b) => b.classList.toggle("active", b.dataset.tab === id));
  document.querySelectorAll(".tab-panel").forEach((p) => p.classList.toggle("hidden", p.id !== id));
  // Charts drawn while hidden have no size yet, so resize them once visible.
  [...state.charts, ...TMUI._tm.charts].forEach((c) => c.resize());
}

// ---------- UI helpers ----------

function showError(msg) {
  $("error").textContent = msg;
  $("error").classList.toggle("hidden", !msg);
}
function setLoading(msg) {
  $("loading").textContent = msg || "";
  $("loading").classList.toggle("hidden", !msg);
}
async function run(task) {
  showError("");
  document.querySelectorAll("button").forEach((b) => (b.disabled = true));
  try {
    await task();
  } catch (e) {
    showError(e.message || String(e));
  } finally {
    setLoading("");
    document.querySelectorAll("button").forEach((b) => (b.disabled = false));
  }
}

// ---------- Actions ----------

async function connect() {
  const token = $("token").value.trim();
  if (!token) return showError("Please paste your monday.com API token first.");
  state.token = token;
  state.demo = false;
  await run(async () => {
    setLoading("Connecting…");
    const me = await fetchMe();
    $("whoami").textContent = `Connected as ${me.name} · ${me.account.name}`;
    if ($("remember").checked) localStorage.setItem(TOKEN_KEY, token);
    else localStorage.removeItem(TOKEN_KEY);

    setLoading("Loading boards…");
    const boards = await fetchBoards();
    const sel = $("board-select");
    sel.replaceChildren();
    if (!boards.length) throw new Error("No boards found for this account.");
    for (const b of boards) sel.append(el("option", { value: b.id }, `${b.name} (${b.items_count ?? "?"} items)`));
    $("board-card").classList.remove("hidden");
  });
}

async function loadBoard() {
  if (state.demo) return loadDemo();
  const id = $("board-select").value;
  if (!id) return;
  await run(async () => {
    setLoading("Loading board…");
    const { board, items } = await fetchBoardWithItems(id);
    state.board = board;
    state.items = items;
    renderDashboard();
  });
}

async function uploadFile(file) {
  if (!file) return;
  await run(async () => {
    setLoading(`Reading ${file.name}…`);
    state.demo = false;
    $("views").classList.remove("hidden");
    // The generic overview needs monday column types, which an export doesn't carry.
    $("tab-overview").classList.add("hidden");
    showTab("tm-view");
    await TMUI.loadFromFile(file);
  });
  $("xlsx-file").value = ""; // allow re-uploading the same file
}

function forget() {
  localStorage.removeItem(TOKEN_KEY);
  state.token = "";
  state.demo = false;
  $("token").value = "";
  $("remember").checked = false;
  $("whoami").textContent = "";
  $("board-card").classList.add("hidden");
  $("views").classList.add("hidden");
  TMUI.reset();
  showError("");
}

// ---------- Demo data (lets you preview without a token) ----------

function loadDemo() {
  state.demo = true;
  const groups = ["New Joinee", "QA India Resources Assigned", "IND_TransformationSquad", "Asset Engineering team"];
  const managers = ["Priya Sharma", "Alex Chen", "Sam Rivera", "Jordan Lee"];
  const board = {
    id: "demo", name: "Demo: TCOE PMO Onboarding",
    columns: [
      { id: "name", title: "Name", type: "name" },
      { id: "talent", title: "Talent ID", type: "text" },
      { id: "tstatus", title: "Teramind Status", type: "status" },
      { id: "tplan", title: "Teramind plan", type: "dropdown" },
      { id: "texclude", title: "Teramind - Exclude", type: "status" },
      { id: "manager", title: "Manager", type: "people" },
      { id: "lwd", title: "Last Working Date in USAA Account", type: "date" },
      { id: "release", title: "Release Category", type: "status" },
    ],
    groups: groups.map((g, i) => ({ id: "g" + i, title: g })),
  };
  let seed = 11;
  const rnd = (n) => { seed = (seed * 9301 + 49297) % 233280; return Math.floor((seed / 233280) * n); };
  const pick = (weights) => { let x = rnd(100); for (const [v, w] of weights) { if ((x -= w) < 0) return v; } return weights[0][0]; };
  const items = [];
  const mk = (name, talent, status, plan, exclude, lwd, release) => ({
    id: String(items.length + 1), name, group: { title: groups[rnd(groups.length)] },
    column_values: [
      { id: "talent", type: "text", text: talent },
      { id: "tstatus", type: "status", text: status },
      { id: "tplan", type: "dropdown", text: plan },
      { id: "texclude", type: "status", text: exclude },
      { id: "manager", type: "people", text: managers[rnd(managers.length)] },
      { id: "lwd", type: "date", text: lwd },
      { id: "release", type: "status", text: release },
    ],
  });
  for (let i = 1; i <= 150; i++) {
    const na = rnd(100) < 6;
    const week = na ? "NA" : `Week${rnd(10)}`;
    const w = na ? -1 : Number(week.slice(4));
    let status;
    if (na) status = "Exemption";
    else if (w <= 7) status = pick([["Installed", 78], ["Installed and not active", 8], ["Pending with issues", 6], ["Mail requested for Activation", 4], ["", 4]]);
    else if (w === 8) status = pick([["Installed", 45], ["Mail requested for Activation", 25], ["Pending with issues", 10], ["", 20]]);
    else status = pick([["", 60], ["Mail requested for Activation", 30], ["Installed", 10]]);
    const exclude = na || rnd(100) < 3 ? "Exclude" : "";
    const roll = rnd(100) < 7;
    items.push(mk(`Resource ${String(i).padStart(3, "0")}`, String(100200 + i), status, week, exclude,
      roll ? `2026-12-${String(1 + rnd(28)).padStart(2, "0")}` : "", roll && rnd(2) ? "Project end" : ""));
  }
  // Rows the dashboard must ignore.
  items.push(mk("Resource 004 (copy)", "100204", "Installed", "Week1", "", "", ""));
  items.push(mk("Resource 019 (copy)", "100219", "Installed", "Week3", "", "", ""));
  items.push(mk("Placeholder: hiring", "", "", "Week5", "", "", ""));

  state.board = board;
  state.items = items;
  $("whoami").textContent = "Demo mode (no token)";
  $("board-card").classList.add("hidden");
  showError("");
  renderDashboard();
}

// ---------- Wire up ----------

$("btn-connect").addEventListener("click", connect);
$("token").addEventListener("keydown", (e) => e.key === "Enter" && connect());
$("btn-forget").addEventListener("click", forget);
$("btn-load").addEventListener("click", loadBoard);
$("btn-refresh").addEventListener("click", loadBoard);
$("btn-demo").addEventListener("click", (e) => { e.preventDefault(); loadDemo(); });
$("search").addEventListener("input", (e) => renderTable(e.target.value));
$("xlsx-file").addEventListener("change", (e) => uploadFile(e.target.files[0]));
document.querySelectorAll(".tab").forEach((b) => b.addEventListener("click", () => showTab(b.dataset.tab)));
$("tm-download").addEventListener("click", () => run(() => TMUI.downloadExcel()));
$("tm-compare").addEventListener("click", () => TMUI.compare());
$("tm-start").addEventListener("change", (e) => TMUI.setStartDate(e.target.value));

const saved = localStorage.getItem(TOKEN_KEY);
if (saved) {
  $("token").value = saved;
  $("remember").checked = true;
  connect();
}
