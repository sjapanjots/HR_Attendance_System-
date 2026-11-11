# HR Attendance System — monthly attendance marking

A Streamlit app that marks the attendance code for **every employee and every day of the month in one pass**,
using only the punch-in and punch-out times. It writes the codes into the sheet's existing day-wise
classification columns and changes nothing else in the workbook.

It is built for the month-wide layout used in the attendance workbooks (one row per employee, then one block of
columns per day), for example:

| … | SHIFT HOURS | … | In Time 26 | Out Time 26 | OT Hr 26 | Duration Hr 26 | **26 Fr** | In Time 27 | … | **27 Sa** | … |
|---|---|---|---|---|---|---|---|---|---|---|---|
| … | 12 HRS | … | 07:38 | 19:51 | | 12.13 | **P** | 19:48 | … | **P** | … |

A sheet with only one day (one in, out, worked-hours and classification column) works the same way.

## Run

```bash
python -m pip install -r requirements.txt
streamlit run app.py
```

Tests (optional):

```bash
python -m pip install -r requirements-dev.txt
python -m pytest -q tests
```

To also run the end-to-end test against a real month workbook, set `ATTENDANCE_SAMPLE=/path/to/workbook.xlsx`.

## Interface

A minimal, single-column layout: a progress stepper (Upload → Worksheet → Map columns → Thresholds → Preview → Export),
one card per step, colour-coded attendance chips and summary tiles. The look comes from the theme in
[`.streamlit/config.toml`](.streamlit/config.toml) (neutral palette, Inter font) and the stylesheet in
[`assets/style.css`](assets/style.css); change colours or spacing there without touching the app logic.

## Workflow

1. **Upload** an `.xlsx` workbook. The upload is never modified.
2. **Select the worksheet** and header row. A read-only preview is shown.
3. **Map the columns** explicitly — nothing is guessed:
   - shift length column;
   - for **day 1**: in-punch, out-punch, worked hours (or "not in this sheet") and the attendance classification column;
   - the number of days, and **day 2's in-punch** column. The step between the two in-punch columns is the width of one
     day block, and the day-1 layout is repeated for every day.

   The full day-by-day mapping is shown in a table. If any generated block's headers do not match day 1
   (e.g. `In Time 27` vs `In Time 26`, or the four columns of a block do not share a day number), processing stops.
   You must tick that you have checked the mapping.
   If the sheet has no classification column, you can explicitly opt in to adding one (single-day sheets only).
4. **Review the thresholds** and tick that you have reviewed them.
5. **Preview**: every employee-day is classified. You see how many cells will be updated, how many already hold the
   right code, and how many are left unchanged for review, with Excel row, cell, punches, worked time, old code,
   new code and the reason.
6. **Approve and create** a separate output workbook. It is verified before the download button appears.

## How a day is classified

Worked time = out-punch − in-punch. When the out-punch is earlier than the in-punch it is the next morning
(night shift, e.g. 19:48 → 08:20 = 12:32).

The punch times alone cannot show a punch pair that spans **24 hours or more** (e.g. in 07:34, out 08:03 the
next day). When the worked-hours cell is mapped and equals the punch time plus whole days (±5 minutes), those days
are added (07:34 → 08:03 with `24.29` = 24:29). Apart from that, the worked-hours cell is not used: the code comes from
the punches. Choose the format the column uses: `12.13` in these workbooks means 12 h 13 min (hours.minutes).

| Situation | Result |
|---|---|
| Exactly one of in-punch / out-punch present | `MP` (punch cells are not changed) |
| No punches | cell left unchanged (keeps `A`, `WO`, `A/CO`, …), listed for review |
| A punch is not a readable time, shift length missing or not 9/10/11/12 h | cell left unchanged, listed for review |
| Proposed code already in the cell (ignoring letter case, e.g. `Hf`) | cell left unchanged |
| Otherwise | code from the table below |

### Thresholds (worked time, h:mm)

| Shift | A | HF | SP | P | PP |
|---|---|---|---|---|---|
| 9 hours | < 5:00 | 5:00 to < 8:00 | 8:00 to < 8:45 | 8:45 to 17:00 inclusive | > 17:00 |
| 10 hours | < 5:00 | 5:00 to < 9:00 | 9:00 to < 9:45 | 9:45 to 18:00 inclusive | > 18:00 |
| 11 hours | < 5:00 | 5:00 to < 10:00 | 10:00 to < 10:45 | 10:45 to 20:00 inclusive | > 20:00 |
| 12 hours | < 6:00 | 6:00 to < 10:00 | 10:00 to < 11:45 | 11:45 to 22:00 inclusive | > 22:00 |

**15-minute grace period:** P starts 15 minutes before the full shift length, so 11:45 worked on a 12-hour shift is
`P`, and 11:44 is `SP`. All thresholds live in one place, [`attendance/rules.py`](attendance/rules.py).

## What is (and is not) changed in the workbook

- Only the mapped classification cells whose code actually changes are written. Their style is kept; when the code
  already exists in the workbook's shared strings it is reused, so new cells look exactly like the existing ones.
- Every other cell, column, row, worksheet, formula, format, filter, conditional format, sheet order and workbook
  setting is left exactly as it was. The edit is made by splicing bytes into the selected worksheet's XML at the target
  cells only. All other parts of the `.xlsx` package are copied byte-for-byte.
- No data cleaning: values are never trimmed, renamed, re-dated, de-duplicated, filled or sorted. No sheets, reports
  or extra columns are added (except the explicit single-day "add a classification column" opt-in).
- Before the download is offered, the output is verified: same package members in the same order, every member other
  than the edited worksheet byte-identical, the edited worksheet byte-identical once the target cells are removed,
  the XML well-formed, and each target cell holds its approved code. If any check fails, no file is offered.
- The tool refuses (and explains why) to write into a cell that holds a formula, sits inside a merged range (other than
  its top-left cell) or an array formula, or holds rich-value metadata. It also refuses to add a column to a sheet with
  merged cells or an Excel table, and refuses worksheets that use prefixed or Strict Open XML.
- Formulas elsewhere that depend on the classification cells keep their cached results until Excel recalculates.

## Why Python + Streamlit (and not Node.js/TSX)

- **The work is data processing, not a website.** About 30,000 employee-days have to be calculated and reviewed per month.
  Polars handles that in well under a second and is part of the requirements. Node.js has no comparable dataframe library.
- **Workbook fidelity is the hard part.** The common JS libraries (ExcelJS, SheetJS) load the workbook into objects and
  write it out again, which drops or rewrites features they do not model. They cannot prove that only the target cells
  changed. The byte-level patch and verification used here would have to be hand-written in Node too.
- **One HR user, one file, one screen.** Streamlit gives the upload → map → review → approve → download flow in one Python
  file with no separate frontend, API or build step. A Node.js/TSX version needs a React frontend, an API server and a
  build pipeline for the same result.
- Node.js/TSX becomes the better choice if this grows into a multi-user product: logins and roles, a database of
  employees and punches, live device integration, or approval workflows. The rules and workbook-patching logic in
  `attendance/` are small and can be ported, or kept as a Python service behind such a frontend.

## Project layout

```
app.py                  Streamlit UI (workflow steps 1–6)
assets/style.css        styling for the clean, minimal look
.streamlit/config.toml  theme (colours, font) and hidden Streamlit toolbar
attendance/rules.py     thresholds, grace period and the Polars classification expression
attendance/parsing.py   read-only parsing of punch times, worked hours and shift length
attendance/layout.py    day-block mapping (repeat day 1 across the month) and header checks
attendance/pipeline.py  one record per employee-day, classified with Polars
attendance/workbook.py  reading the sheet, byte-level cell patching and verification
tests/                  unit tests and an optional end-to-end test with a real workbook
```
