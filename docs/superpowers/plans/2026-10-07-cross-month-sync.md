# Cross-Month KPI Sync Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Make the daily KPI sync retry the previous month together with the current month so a missed or failed month-boundary run does not leave the prior month stale.

**Architecture:** Add a small pure-Python month-window helper that returns the previous and current `YYYYMM` keys in chronological order. Update the three KPI scripts to accept an explicit month, process the two-month window, and only write a month after its SAP export and local processing complete; a failure exits nonzero so the next scheduled run retries it. Keep the existing Google Sheet columns and DOMO connector unchanged.

**Tech Stack:** Python 3.11, pytest, SAP GUI COM automation, openpyxl, gspread, PowerShell orchestration.

**Spec:** User-approved request in the conversation: update the 00/01/02 KPI jobs so cross-month absences or failures are automatically recovered on the next run.

## Global Constraints

- Daily processing covers exactly the previous calendar month and the current calendar month.
- Month-specific SAP queries must use an explicit `YYYYMM`; they must not derive the query month from the current date inside the query function.
- A failed month must not overwrite its existing Google Sheet row and must produce a nonzero process exit code.
- Existing `MasterData` column ownership remains unchanged: 00 writes A:D, 01 writes E:G, and 02 writes H:J.
- Do not run SAP GUI or Excel during tests; tests must cover pure month logic only.

## Review Focus

- January boundary: previous month must be December of the prior year.
- A failure in the previous month must leave the month eligible for retry on the next run.
- Current-month no-data behavior must not be confused with a successful historical-month refresh.
- Existing output filenames and Google Sheet row matching must remain keyed by the explicit month.
- The three KPI scripts must use the same chronological month window.

---

### Task 1: Add and test the shared month window

**Files:**
- Create: `month_utils.py`
- Test: `tests/test_month_utils.py`

**Interfaces:**
- Produces `months_to_sync(today: datetime.date, lookback: int = 1) -> list[str]`.

- [ ] **Step 1: Write the failing tests** for normal month rollover, January rollover, and invalid lookback.
- [ ] **Step 2: Run `pytest tests/test_month_utils.py -q` and verify the tests fail because `month_utils` is missing.**
- [ ] **Step 3: Implement `months_to_sync` with chronological previous-to-current output.**
- [ ] **Step 4: Run the focused tests and verify they pass.**

### Task 2: Convert KPI scripts to previous-plus-current processing

**Files:**
- Modify: `00 - Maintenance_Plan_Adherence_Sync/main.py`
- Modify: `01 - Maintenance_Effectiveness_Sync/main.py`
- Modify: `02 - Critical_A_ &_H_equipment_with_Maintenance_Plan_Sync/main.py`

**Interfaces:**
- Consumes `months_to_sync` from `month_utils.py`.
- Each SAP extraction function receives an explicit `year_month` and derives its date range from it.

- [ ] **Step 1: Update 00** to process the two-month window and return/raise clearly around month-level completion.
- [ ] **Step 2: Update 01** to process the two-month window, return a distinct no-data result, and treat historical-month no-data as a failure while preserving current-month no-data behavior.
- [ ] **Step 3: Update 02** to process the same two-month window and keep its A-class equipment detail sheet aligned with each explicit month.
- [ ] **Step 4: Run Python syntax compilation for all changed scripts.**

### Task 3: Document operation and run the full verification suite

**Files:**
- Modify: `README.md`

- [ ] **Step 1: Document the previous-plus-current recovery behavior and the one-time backfill requirement for already stale months such as 202608.**
- [ ] **Step 2: Run the complete available pytest suite.**
- [ ] **Step 3: Run `py_compile` on all changed Python files and inspect the diff/status.**

