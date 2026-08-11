# Daily Transport Service Report (GitHub Actions version)

Manual-trigger version of the daily transport report: click a button in the
Actions tab, it pulls the latest data from Google Sheets, builds the report,
converts it to PDF, and commits the result back to this repo so you can
download it.

This is a companion to a Claude Cowork scheduled task that runs the same
report daily from a live Google Drive connection. This repo exists because
that connection can't be used from GitHub Actions -- see "Why CSV, not the
Sheets API" below.

## Setup (one-time)

### 1. Add the two CSV URLs as repository secrets

This repo is **public**. The two Google Sheets behind this report contain
internal SCM data, so the published CSV URLs themselves must **not** be
committed anywhere in this repo (anyone with the URL can pull live data,
indefinitely, with no login) -- they go in encrypted repo secrets instead.

In this repo: **Settings → Secrets and variables → Actions → New repository
secret**, add two secrets:

| Name | Value |
|---|---|
| `CSV_URL_DM` | Published CSV URL for the Delivery Monitoring sheet's tab |
| `CSV_URL_VM` | Published CSV URL for the Daily VM sheet's "Today" tab |

To get/refresh these URLs in Google Sheets: open the sheet → make sure the
**correct tab is active** (this matters -- see warning below) → **File →
Share → Publish to web** → in the first dropdown pick the specific tab (not
"Entire Document") → format **Comma-separated values (.csv)** → **Publish**
→ copy the link. It should contain `&gid=...&single=true&output=csv`.

> **Picking "Entire Document" instead of the specific tab is a common
> mistake** -- it silently publishes whichever tab happens to be sheet #1 in
> the workbook (for the Daily VM file, that's an unrelated lookup tab called
> `m`, not `Today`), and the CSV comes back empty or wrong. If a run fails
> with "Could not locate section marker(s)" for the VM sheet, this is the
> first thing to check.

### 2. Run it

**Actions tab → "Daily Transport Service Report" → Run workflow.** Manual
only for now (see "Why manual-only" below). Takes 1-2 minutes. When it
finishes, `Daily_Transport_Service_Report.pdf` and `.html` in the repo root
will be updated (new commit) and also attached to the run as a downloadable
artifact.

## Why CSV, not the Sheets API

The Cowork version of this pipeline reads Google Sheets through a live
Drive OAuth session. GitHub Actions runners have no such session and can't
easily get one: a Google Cloud service account would normally fill that
role, but service account creation is blocked by IT policy in this
environment (as of 2026-08-11). "Publish to web" as CSV was chosen instead
because it needs **no Google Cloud project, no credentials, and no IT
exception** -- it's a plain HTTP GET.

**Trade-off, and it's a real one:** a published CSV URL is a bearer link --
anyone who has it can read the live sheet data, in real time, with no
authentication, for as long as the sheet stays published. That's *more*
exposure than just having yesterday's PDF snapshot sitting in a repo. This
was a deliberate, discussed decision (not a default), and it's the reason
the URLs live in secrets rather than the workflow file even though the repo
itself is public -- keeping the *access mechanism* out of public view is
the only mitigation available here. If that trade-off stops being
acceptable, revisit whether an IT-approved service account or OAuth client
has become available.

## Why manual-only (no daily cron)

Deliberate choice for the initial rollout: the CSV-based parser is new and
hasn't been through the weeks of real-world edge cases the original
Cowork/openpyxl parser has (see `scripts/parse_vm.py`'s docstring for the
list of ways the source sheets have changed shape without notice in the
past). Running it unattended on a schedule before it's proven reliable
means a broken run fails silently overnight. Once it's been triggered
manually enough times to build confidence, add a `schedule:` trigger to
`.github/workflows/daily-report.yml`, e.g.:

```yaml
on:
  workflow_dispatch: {}
  schedule:
    - cron: '0 3 * * *'   # 10:00 WIB
```

## Known limitations vs. the Cowork-generated report

- **Isu Pending Harian is shorter here.** The Cowork parser also merges in
  a dated historical "archive" table that lives on a different tab of the
  Delivery Monitoring spreadsheet (not part of what got published to web
  for this pipeline). In practice that archive has only ever contributed
  stale data (weeks behind), so the loss is expected to be minor -- but if
  you know there's an internal-reason pending order today and it's missing
  from the report, this is why.
- **Sheet structure can change without notice.** Both source sheets have
  changed column layouts, header wording, and table shapes multiple times
  already (see the docstrings in `scripts/parse_dm.py` and
  `scripts/parse_vm.py`). `scripts/build_report.py` does basic sanity
  checks and prints warnings to the Actions log if something looks
  structurally wrong (e.g. a section came back empty, or total-row counts
  are off) -- check the log if the report looks thin.
- If a run fails outright, the previous good report stays in the repo
  untouched (the commit step only runs after everything else succeeds).

## Repo layout

```
templates/report_template.html   HTML/CSS/JS shell -- renders pre-parsed
                                  dm/vm JSON objects. No parsing logic here
                                  (unlike the Cowork version's template,
                                  which parses raw markdown/JSON client-side)
scripts/parse_dm.py              Delivery Monitoring CSV -> dm JSON
scripts/parse_vm.py              Daily VM "Today" CSV -> vm JSON
scripts/build_report.py          Fetches both CSVs, runs both parsers,
                                  injects into the template -> final HTML
scripts/render_pdf.js            Puppeteer: HTML -> portrait + landscape PDF
scripts/merge_pdf.py             Merges the two PDFs into one file
.github/workflows/daily-report.yml   The button + pipeline glue
```

## Running it locally (for debugging a failed run)

```bash
pip install -r requirements.txt
npm install
export CSV_URL_DM="<published DM csv url>"
export CSV_URL_VM="<published VM csv url>"
python scripts/build_report.py Daily_Transport_Service_Report.html
node scripts/render_pdf.js Daily_Transport_Service_Report.html .
python scripts/merge_pdf.py page1_portrait.pdf page2_landscape.pdf Daily_Transport_Service_Report.pdf
```

`scripts/parse_dm.py` and `scripts/parse_vm.py` can also be run standalone
against a saved CSV file to debug just the parsing step without touching
the network or Puppeteer:

```bash
python scripts/parse_dm.py path/to/dm.csv > dm.json
python scripts/parse_vm.py path/to/vm.csv > vm.json
```
