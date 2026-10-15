# SASE Attendance Analysis

A Python script that cleans the Microsoft Forms attendance export for SASE (UConn chapter) and produces charts, a summary, and an Excel report: attendance per meeting, new vs. returning members, retention, recruitment channels, incentive eligibility, merch sizes, and what drives attendance (food, meeting type).

Everything lives in one file, `sase_analysis.py`. All settings are Python constants at the top of that file; there are no config or CSV files to maintain.

## Requirements

- Python 3.9 or newer
- `pandas`, `openpyxl`, `matplotlib`
- `statsmodels` (optional): only needed for the regression models. Without it, those sections print a "not installed" message and everything else still runs.

## Setup

From the project folder:

```powershell
python -m venv venv
venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

On macOS/Linux, activate with `source venv/bin/activate` instead.

## Running

1. **Get your data.** Export the form responses from Microsoft Forms to Excel (`.xlsx`) or CSV and put the file in the project folder. The export is not included in this repo because it contains members' personal information.
2. **Point the script at it.** Open `sase_analysis.py` and set `DATA_FILE` near the top:

   ```python
   DATA_FILE = "SASE Fall_2026 Member_Attendance.xlsx"
   ```

   The dataset is deliberately not a command-line argument. Change this line whenever you switch to a new file.
3. **Fill in the settings** (see below).
4. **Run it:**

   ```powershell
   python sase_analysis.py
   ```

   Add `--out <folder>` to write the results somewhere other than `output/`.

## Settings to edit (top of `sase_analysis.py`)

| Constant | What to put in it |
| --- | --- |
| `DATA_FILE` | File name of the form export (`.xlsx` or `.csv`). |
| `EBOARD` | The names of all e-board members, as `"First Last"` (case-insensitive), matching what they typed on the form. Emails also work. They are excluded from the "non-e-board" counts, retention, channels and demographics. |
| `MEETINGS` | Real meeting dates and names: `"YYYY-MM-DD": "Meeting name"`. The names replace dates on every chart and table. |
| `MEETING_TYPES` | What each meeting-type number means, e.g. `1: "Social"`, `2: "Professional"`, `3: "Company Visit"`. |
| `MEETING_INFO` | Per-meeting details, keyed by the same dates as `MEETINGS`: `{"IS_FOOD": 0 or 1, "MEETING_TYPE": n}`. `IS_FOOD` is 1 if food was there and 0 if not. `MEETING_TYPE` is a number from `MEETING_TYPES`. Any extra numeric field you add also becomes a predictor in the meeting model. |
| `INCENTIVE_THRESHOLD` | Number of meetings needed to qualify for incentives (default 3). |
| `COL` | Column headers of the form export. Only change these if your form's question text differs. |

To add a new meeting, add one line to `MEETINGS` and one to `MEETING_INFO` with the same date.

## What the script does to the data

- **Meeting dates:** the typed "Date of Meeting" answer is unreliable, so the meeting is taken from the date of the `Start time` timestamp. A submission on a date not listed in `MEETINGS` (for example a late submission) is snapped back to the nearest earlier listed meeting, or forward to the first one if it predates all of them. If `MEETINGS` is empty, every distinct submission date is treated as its own meeting.
- **Identifying people:** rows that share a normalized email *or* a normalized name are merged into one person. This catches people with two emails and typo'd emails.
- **Duplicates:** if the same person submits twice for one meeting, only one counts.
- **New vs. returning:** derived from the data (a person is new at their first meeting), because the form's "Have you filled out this form yet?" answer is inconsistent.
- **Profiles:** returning members skip the profile questions, so each person's profile and sizes come from their most complete submission.
- **"How did you hear about SASE?"** is split on `;` because it is a multi-select.

## Output

Everything is written to `output/` (or `--out`). The terminal only shows a one-line "saved" message.

| File | Contents |
| --- | --- |
| `attendance_per_meeting.png` | Stacked bars of returning vs. first-time attendees per meeting, with counts. |
| `attendance_without_eboard.png` | Attendance with and without e-board, per meeting. |
| `attendance_by_food.png` | Average attendance at meetings with food vs. without. |
| `attendance_by_meeting_type.png` | Average attendance per meeting type. |
| `sase_summary.txt` | The cleaning report, meeting counts, incentive-eligible members, retention, recruitment channels, model results and event suggestions. |
| `member_attendance.csv` | One row per member: name, email, meetings attended and whether they qualify for incentives, ordered by meetings attended (most first, ties alphabetical). |
| `sase_report.xlsx` | Sheets: Attendees, Meetings, Attendance (clean), Retention, Channels, Merch sizes, and one demographics sheet per question. |

The two food/meeting-type charts are skipped, with a note in `sase_summary.txt`, until `MEETING_INFO` has values.

The cleaning report at the top of `sase_summary.txt` also warns about `EBOARD` names that don't match anyone in the data and `MEETINGS` / `MEETING_INFO` dates with no submissions. Check it after each run for typos.

## About the models

- **Return model** (exploratory): a logistic regression of whether a first-time attendee came back, using first-year status and whether they heard about SASE through a friend, the involvement fair, or Instagram. E-board members are excluded. With a small sample the confidence intervals are very wide, so treat it as a hint rather than a result.
- **Meeting model:** with few meetings it prints average non-e-board attendance per `IS_FOOD` / `MEETING_TYPE` value. It only fits a Poisson regression once there are at least 3 × (number of predictors) + 3 meetings.

## Using the cleaned data in your own analysis

```python
from sase_analysis import load_clean, DATA_FILE

data = load_clean(DATA_FILE)
data.attendance   # one row per (person, meeting), duplicates removed
data.people       # one row per unique person: profile + attendance stats
data.meetings     # one row per meeting: attendees, new, returning, eboard, non_eboard, label
data.raw          # original rows plus the cleaned columns
```

New analyses are best added as functions in `sase_analysis.py` that take a `CleanData`.

## Privacy

The attendance export and `output/` contain names, emails and clothing sizes. Both are listed in `.gitignore`, so keep them out of version control.
