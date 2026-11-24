# monday.com Dashboards: Teramind Connectivity (v1)

A static web page with no backend and no build step. It has two tabs:

- **Teramind weekly** (default): the weekly Teramind connectivity dashboard built from the PMO onboarding board.
- **Board overview**: a generic dashboard for any monday.com board.

## Load data
- **Option A, live:** paste your monday API token (**Avatar → Developers → My access tokens**), click **Connect**, pick the board, then click **Build dashboard**.
- **Option B, export:** click **Upload export (.xlsx)** and choose the board export (e.g. `TCOE_PMO_Onboarding_*.xlsx`). The file is read in the browser and never uploaded.
- **Demo:** click **try with demo data**.

The token and the file stay in your browser. The only network call is to `https://api.monday.com/v2`.

## Teramind rules
| Rule | How it's applied |
|---|---|
| Resource rows | Talent ID filled and not the header text "Talent ID"; name doesn't contain "(copy)". This drops section titles, repeated headers and date-range summary rows. |
| Columns | Matched by header name: Teramind Status (BF), Teramind plan (BG), Teramind - Exclude (BH), Last Working Date in USAA Account (AT), Release Category (AU). Check or override them in **Column mapping**. |
| Plan | plan = the week (Week0–Week9) |
| Actual | plan = the week **and** status exactly "Installed" ("Installed and not active" is **not** counted) |
| Cumulative | running total of Actual |
| Rolloffs | plan = the week **and** (Release Category or Last Working Date filled) |
| Exemptions | plan = the week **and** Exclude = "Exclude". They stay in Plan. |
| Exempt row | the same metrics for plan = "NA" |
| Week Ending | Week0 date (default 25-Sep-26), +7 days per week. Editable. |
| Remarks | typed into the table and saved in this browser |

The page also shows:
- Key Metrics
- a Status breakdown that reconciles to the resource-row count
- Plan vs Actual and Cumulative charts
- a **Reconciliation** panel: an independent recount, with pass/fail shown, never assumed
- **Compare with your numbers**: paste the figures you were given and see every difference
- an **Issues** list: installed but not active, planned with no status, pending, unknown plan values, and rolloffs still planned
- the Excel formulas

If a required column is missing, the page says `INSUFFICIENT_DATA` and names the column.

## Download Excel
**Download Excel** creates a new file and never changes the original:
- **From an upload:** a copy of your export with a `Dashboard` sheet added, plus a helper column `Valid (dashboard)` (1 = resource row). The formulas use the export's own columns.
- **From the API:** a `Data` sheet plus a `Dashboard` sheet.

The Dashboard sheet uses live formulas:
- `COUNTIFS` for Plan, Actual and Exemptions
- `SUMPRODUCT((plan=week)*(valid=1)*(((AU<>"")+(AT<>""))>0))` for Rolloffs
- dates as Excel serials formatted `dd-mmm-yy`

Excel recalculates everything when the file opens.

The Excel file has no charts, because the library can't create native Excel charts. The charts are on the web page.

## Host it (free)
- **Netlify Drop:** go to https://app.netlify.com/drop and drag this folder in.
- **GitHub Pages:** push to a repo, then go to **Settings → Pages → Deploy from branch**.

## Files
- `index.html`: layout
- `style.css`: styles
- `app.js`: monday API, generic overview, tabs, demo data
- `teramind-core.js`: Teramind calculations (pure, testable in Node)
- `teramind.js`: Teramind UI, .xlsx reading and Excel export
