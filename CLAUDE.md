# SASE Attendance Analysis

Python script that cleans the SASE (UConn chapter) Microsoft Forms attendance export and runs analyses on it. Owner: Agronil Das (SASE e-board).

## Files

- `sase_analysis.py`: everything lives here. `load_clean(path)` returns a `CleanData` with:
  - `attendance`: one row per (person, meeting), duplicates removed, `is_new`, `is_eboard`
  - `people`: one row per unique person: profile (from their most complete form), `meetings_attended`, `meeting_dates`, `returned`, `incentive_eligible`, `is_eboard`, `heard_list`
  - `meetings`: one row per meeting: `attendees`, `new`, `returning`, `eboard`, `non_eboard`, plus any `MEETING_INFO` fields
  - `raw`: original rows plus the cleaned columns
- Run it with `python sase_analysis.py <file>.xlsx [--out output]`. It writes 2 PNG charts and `sase_report.xlsx` to `output/`.
- Requires pandas, openpyxl, matplotlib, and optionally statsmodels (the regressions are skipped without it).

## User preferences

- **No extra CSV/config files.** All cleaning must be automatic inside the script. Settings (`EBOARD`, `MEETING_INFO`, `INCENTIVE_THRESHOLD`) are Python constants at the top of the script.
- Cleaned data should be reusable for any future task, so new analyses go in new functions that take `CleanData`.

## Data quirks (Fall 2026 export) and how they're handled

- The **"Date of Meeting" typed field is unreliable**: 5 rows had the wrong date. The meeting is instead taken from the date of `Start time`. Any date not in `MEETINGS` gets snapped back to the nearest earlier real meeting (or forward to the first one if it predates all meetings).
- Real meetings: 9/17, 9/24, 10/1, 10/8, with 22 / 17 / 15 / 19 unique attendees. The `MEETINGS` constant (date -> meeting name) at the top of the script holds the real dates; submissions on any other date snap back to the nearest earlier listed date, and the names replace dates in charts/tables (`meetings["label"]`). If `MEETINGS` is empty, every distinct submission date is its own meeting, labelled by date.
- **Identity resolution**: union-find on normalized email OR normalized name. This merges people with two emails or typo'd emails (Andy Vo, Laine Elliott, Fahim Rifat pbl25005/pbk25005). The 76 raw rows become 45 people.
- Same-meeting double submissions are dropped (Rishabh Raj, Bianca Abrau).
- Returning members skip the profile questions, so profile/sizes come from each person's most complete submission.
- "Have you filled out this form yet?" is inconsistent, so new vs. returning is derived from the data instead.
- `How did you hear about SASE?` is a `;`-separated multi-select.

## Open items / next steps

- User still needs to fill in the constants at the top of the script: `EBOARD` (names or emails), `MEETINGS` (date -> name), `MEETING_TYPES` (number -> description), and `MEETING_INFO` (date -> `{"IS_FOOD": 0/1, "MEETING_TYPE": n}`). The cleaning report warns about `EBOARD` entries or `MEETINGS`/`MEETING_INFO` dates that don't match the data.
- `meeting_model` shows average non-e-board attendance per `IS_FOOD` / `MEETING_TYPE` value with few meetings, and only fits a Poisson regression once meetings >= 3 x predictors + 3 (`MEETING_TYPE` is treated as categorical).
- `return_model`: exploratory logistic regression on "came back after first meeting" (first-year, friend/fair/Instagram channel). Still untested: statsmodels isn't installed here.
- Possible additions: arrival-time analysis, fuzzy name matching for nicknames, multi-semester support.
