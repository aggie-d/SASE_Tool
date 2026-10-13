"""
SASE attendance analysis.

Usage
-----
    python sase_analysis.py SASE_Fall_2026_Member_Attendance.xlsx [--out output]

Or reuse the cleaned data anywhere:

    from sase_analysis import load_clean
    data = load_clean("SASE_Fall_2026_Member_Attendance.xlsx")
    data.attendance   # one row per (person, meeting) -- duplicates removed
    data.people       # one row per unique person, with profile + attendance stats
    data.meetings     # one row per meeting, with counts
    data.raw          # original rows + the cleaned columns added
"""
from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path

import pandas as pd

# ---------------------------------------------------------------------------
# Optional settings -- edit here, no extra files needed.
# ---------------------------------------------------------------------------

# E-board members: emails or "first last" names (case-insensitive).
EBOARD: set[str] = {
    "Aiden Lau",
    "Sahasra Darisi",
    "Agronil Das",
    "Aggie Das",
    "Jake Saravanan",
    "Danny Nguyen",
    "Nithya Rajeshkanna",
    "Laine Elliott",
    "Jason Huang",
    "Shen Wan",
    "Aaron Wong",
    "Nicole Wang",
    "Aashna Gaikwad",
}

# The real meetings: date (YYYY-MM-DD) -> meeting name. Names replace the dates
# on every chart and table. A submission on any other date is snapped back to
# the nearest earlier listed date (or the first one, if it predates them all).
# If left empty, every distinct submission date is treated as its own meeting.
MEETINGS: dict[str, str] = {
    "2026-09-17": "Popsicle Bridge Building",
    "2026-09-24": "Career Fair Simulation",
    "2026-10-01": "ATLA Movie Night",
    "2026-10-08": "Medtronic Company Visit",
    # "2026-10-15": "ASML Company Visit",
    # "2026-10-22": "Pumpkin Painting",
    # "2026-10-29": "Synchrony Company Visit",
    # "2026-11-05": "Henkel Company Visit",
    # "2026-11-12": "Engineer Your Career",
    # "2026-12-03": "Christmas Movie Night",
}

# What each MEETING_TYPE number means (you define these).
MEETING_TYPES: dict[int, str] = {
    1: "Social",
    2: "Professional",
    3: "Company Visit",
}

# Per-meeting details for the "what brings people in" analysis.
# Key = meeting date (YYYY-MM-DD), matching a date in MEETINGS.
#   IS_FOOD:      1 if food was present, 0 if not
#   MEETING_TYPE: a number defined in MEETING_TYPES above
# Any extra numeric field you add also becomes a predictor.
MEETING_INFO: dict[str, dict] = {
    "2026-09-17": {"IS_FOOD": 1, "MEETING_TYPE": 1},
}

INCENTIVE_THRESHOLD = 3  # meetings needed to qualify for incentives

# Column names in the Microsoft Forms export
COL = {
    "time": "Start time",
    "typed_date": "Date of Meeting",
    "email": "Email Address",
    "first": "First Name",
    "last": "Last Name",
    "standing": "Class Standing",
    "grad": "Graduation Year",
    "stem": "Are you a STEM major?",
    "major": "What is your major(s)?",
    "minor": "What is your minor(s)?",
    "heritage": "Asian Heritage",
    "generation": "Immigration Generation",
    "gender": "Gender Self Identification",
    "pronouns": "Preferred Pronouns",
    "heard": "How did you hear about SASE?",
    "suggestions": "Suggestions for SASE events or GB Meetings?",
    "summer": "What did you do over the summer?",
    "hoodie": "For incentive purposes, what is your hoodie size?",
    "shirt": "For incentive purposes, what is your T-Shirt size?",
}
PROFILE_COLS = ["standing", "grad", "stem", "major", "minor", "heritage", "generation",
                "gender", "pronouns", "heard", "summer", "hoodie", "shirt"]


# ---------------------------------------------------------------------------
# Cleaning
# ---------------------------------------------------------------------------

@dataclass
class CleanData:
    raw: pd.DataFrame
    attendance: pd.DataFrame
    people: pd.DataFrame
    meetings: pd.DataFrame


def _norm_text(s) -> str | None:
    if pd.isna(s):
        return None
    s = re.sub(r"\s+", " ", str(s)).strip()
    return s or None


def _norm_email(s) -> str | None:
    s = _norm_text(s)
    return s.lower() if s else None


def _norm_name(first, last) -> str | None:
    parts = [p for p in (_norm_text(first), _norm_text(last)) if p]
    return " ".join(parts).lower() if parts else None


def _assign_person_ids(df: pd.DataFrame) -> pd.Series:
    """Union-find: rows sharing an email OR a normalized name are one person.
    Catches typo'd emails (pbl25005 vs pbk25005) and people using two emails."""
    parent = list(range(len(df)))

    def find(i):
        while parent[i] != i:
            parent[i] = parent[parent[i]]
            i = parent[i]
        return i

    def union(a, b):
        ra, rb = find(a), find(b)
        if ra != rb:
            parent[rb] = ra

    for key in ("email_norm", "name_norm"):
        first_seen: dict[str, int] = {}
        for i, v in enumerate(df[key]):
            if v is None:
                continue
            if v in first_seen:
                union(first_seen[v], i)
            else:
                first_seen[v] = i

    roots = [find(i) for i in range(len(df))]
    codes = {r: n for n, r in enumerate(dict.fromkeys(roots), start=1)}
    return pd.Series([f"P{codes[r]:03d}" for r in roots], index=df.index)


def _assign_meetings(times: pd.Series) -> pd.Series:
    """Meeting = calendar date of the submission timestamp (the typed-in date is
    unreliable). Any other date is snapped to the nearest past real meeting."""
    dates = pd.to_datetime(times).dt.normalize()
    if not MEETINGS:
        return dates
    real = sorted(pd.to_datetime(list(MEETINGS)))

    def snap(d):
        # Most recent meeting on or before d; a submission that predates every
        # meeting falls forward to the first one.
        past = [m for m in real if m <= d]
        return past[-1] if past else real[0]

    return dates.map(snap)


def _split_multiselect(s) -> list[str]:
    if pd.isna(s):
        return []
    return [p.strip() for p in str(s).split(";") if p.strip()]


def _is_eboard(row) -> bool:
    keys = {row["email_norm"], row["name_norm"]}
    configured = {re.sub(r"\s+", " ", k).strip().lower() for k in EBOARD}
    return bool(keys & configured)


def load_clean(path: str | Path) -> CleanData:
    raw = pd.read_excel(path)
    raw.columns = [str(c).strip() for c in raw.columns]

    df = raw.copy()
    for c in ("first", "last"):
        df[COL[c]] = df[COL[c]].map(_norm_text)
    df["email_norm"] = df[COL["email"]].map(_norm_email)
    df["name_norm"] = [_norm_name(f, l) for f, l in zip(df[COL["first"]], df[COL["last"]])]
    df = df[df["email_norm"].notna() | df["name_norm"].notna()].reset_index(drop=True)

    df["submitted_at"] = pd.to_datetime(df[COL["time"]])
    df["meeting"] = _assign_meetings(df["submitted_at"])
    df["typed_date_mismatch"] = (
        pd.to_datetime(df[COL["typed_date"]], errors="coerce").dt.normalize() != df["meeting"]
    )
    df["person_id"] = _assign_person_ids(df)
    df["n_profile_fields"] = df[[COL[c] for c in PROFILE_COLS]].notna().sum(axis=1)

    # --- people: one row per person, profile from their most complete form ---
    best = (df.sort_values(["n_profile_fields", "submitted_at"], ascending=[False, True])
              .drop_duplicates("person_id"))
    people = best[["person_id", "email_norm", COL["first"], COL["last"]]
                  + [COL[c] for c in PROFILE_COLS]].copy()
    people.columns = ["person_id", "email", "first_name", "last_name"] + PROFILE_COLS
    people["name"] = (people["first_name"].fillna("") + " " + people["last_name"].fillna("")).str.strip()
    people["all_emails"] = people["person_id"].map(
        df.groupby("person_id")["email_norm"].agg(lambda s: ", ".join(sorted(set(s.dropna())))))
    eboard_ids = set(df.loc[df.apply(_is_eboard, axis=1), "person_id"])
    people["is_eboard"] = people["person_id"].isin(eboard_ids)
    people["heard_list"] = people["heard"].map(_split_multiselect)

    # --- attendance: one row per (person, meeting), duplicates removed ---
    attendance = (df.sort_values("submitted_at")
                    .drop_duplicates(["person_id", "meeting"])
                    [["person_id", "meeting", "submitted_at"]]
                    .reset_index(drop=True))
    first_meeting = attendance.groupby("person_id")["meeting"].min()
    attendance["is_new"] = attendance["meeting"].eq(attendance["person_id"].map(first_meeting))
    attendance = attendance.merge(people[["person_id", "name", "is_eboard"]], on="person_id")

    stats = attendance.groupby("person_id").agg(
        meetings_attended=("meeting", "nunique"),
        first_meeting=("meeting", "min"),
        last_meeting=("meeting", "max"),
        meeting_dates=("meeting", lambda s: ", ".join(d.strftime("%m/%d") for d in sorted(s))),
    )
    people = people.merge(stats, on="person_id")
    people["returned"] = people["meetings_attended"] > 1
    people["incentive_eligible"] = people["meetings_attended"] >= INCENTIVE_THRESHOLD
    people = people.sort_values(["meetings_attended", "name"], ascending=[False, True]).reset_index(drop=True)

    # --- meetings ---
    g = attendance.groupby("meeting")
    meetings = pd.DataFrame({
        "attendees": g["person_id"].nunique(),
        "new": g["is_new"].sum(),
        "eboard": g["is_eboard"].sum(),
        "raw_submissions": df.groupby("meeting").size(),
    }).astype(int)
    meetings["returning"] = meetings["attendees"] - meetings["new"]
    meetings["non_eboard"] = meetings["attendees"] - meetings["eboard"]
    info = pd.DataFrame.from_dict(MEETING_INFO, orient="index")
    if not info.empty:
        info.index = pd.to_datetime(info.index)
        info.columns = [str(c).upper() for c in info.columns]
        meetings = meetings.join(info)
    meetings = meetings.reset_index().rename(columns={"index": "meeting"})
    if "MEETING_TYPE" in meetings:
        meetings["MEETING_TYPE_NAME"] = meetings["MEETING_TYPE"].map(
            lambda t: MEETING_TYPES.get(int(t), str(int(t))) if pd.notna(t) else None)
    names = {pd.Timestamp(d): n for d, n in MEETINGS.items()}
    meetings["label"] = [names.get(d) or d.strftime("%b %d") for d in meetings["meeting"]]

    return CleanData(raw=df, attendance=attendance, people=people, meetings=meetings)


def cleaning_report(data: CleanData) -> str:
    df = data.raw
    lines = [
        f"Raw submissions:           {len(df)}",
        f"Unique emails (raw):       {df['email_norm'].nunique()}",
        f"Unique people (merged):    {data.people.shape[0]}",
        f"Duplicate same-meeting submissions removed: {len(df) - len(data.attendance)}",
        f"Typed meeting date was wrong: {int(df['typed_date_mismatch'].sum())} rows",
    ]
    found = set(data.meetings["meeting"].dt.strftime("%Y-%m-%d"))
    for name, keys in (("MEETINGS", MEETINGS), ("MEETING_INFO", MEETING_INFO)):
        missing = sorted(set(keys) - found)
        if missing:
            lines.append(f"WARNING: {name} dates with no submissions: {', '.join(missing)}")
    unmatched = {k for k in EBOARD if k.strip().lower() not in
                 set(df["email_norm"].dropna()) | set(df["name_norm"].dropna())}
    if unmatched:
        lines.append(f"WARNING: EBOARD entries not found in the data: {', '.join(sorted(unmatched))}")
    multi = data.people[data.people["all_emails"].str.contains(",")]
    if len(multi):
        lines.append("People merged across multiple emails:")
        lines += [f"  - {r.name}: {r.all_emails}" for r in multi.itertuples()]
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# Analyses (each takes CleanData; add your own the same way)
# ---------------------------------------------------------------------------

def plot_attendance(data: CleanData, out: Path):
    import matplotlib.pyplot as plt
    m = data.meetings
    x = range(len(m))

    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.bar(x, m["returning"], color="#2a6f97", label="Returning")
    ax.bar(x, m["new"], bottom=m["returning"], color="#a9d6e5", label="First-time")
    for i, total in zip(x, m["attendees"]):
        ax.text(i, total + 0.4, str(total), ha="center", fontsize=10)
    ax.set_xticks(list(x), m["label"])
    ax.set_ylabel("Unique attendees")
    ax.set_title("Attendance per meeting")
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(out / "attendance_per_meeting.png", dpi=150)
    plt.close(fig)

    fig, ax = plt.subplots(figsize=(8, 4.5))
    ax.plot(m["label"], m["attendees"], marker="o", color="#adb5bd", label="All attendees")
    ax.plot(m["label"], m["non_eboard"], marker="o", color="#2a6f97", label="Excluding e-board")
    for lbl, v in zip(m["label"], m["non_eboard"]):
        ax.annotate(str(v), (lbl, v), textcoords="offset points", xytext=(0, 7), ha="center")
    ax.set_ylim(bottom=0)
    ax.set_ylabel("Unique attendees")
    ax.set_title("Attendance with and without e-board")
    ax.spines[["top", "right"]].set_visible(False)
    ax.legend(frameon=False)
    fig.tight_layout()
    fig.savefig(out / "attendance_without_eboard.png", dpi=150)
    plt.close(fig)


def retention(data: CleanData) -> pd.DataFrame:
    """Of the first-timers at each meeting, how many came back at least once."""
    p = data.people[~data.people["is_eboard"]]
    r = p.groupby("first_meeting").agg(first_timers=("person_id", "size"),
                                       came_back=("returned", "sum"))
    r["return_rate"] = (r["came_back"] / r["first_timers"]).round(2)
    r = r.reset_index()
    r.insert(0, "meeting", r["first_meeting"].map(dict(zip(data.meetings["meeting"], data.meetings["label"]))))
    return r.drop(columns="first_meeting")


def recruitment_channels(data: CleanData) -> pd.DataFrame:
    p = data.people[~data.people["is_eboard"]].explode("heard_list").dropna(subset=["heard_list"])
    return (p.groupby("heard_list")
              .agg(people=("person_id", "nunique"),
                   avg_meetings=("meetings_attended", "mean"),
                   return_rate=("returned", "mean"))
              .round(2).sort_values("people", ascending=False).reset_index()
              .rename(columns={"heard_list": "channel"}))


def demographics(data: CleanData) -> dict[str, pd.Series]:
    p = data.people[~data.people["is_eboard"]]
    return {c: p[c].fillna("(blank)").value_counts() for c in
            ["standing", "stem", "heritage", "generation", "gender"]}


def merch_sizes(data: CleanData, eligible_only=False) -> pd.DataFrame:
    p = data.people[data.people["incentive_eligible"]] if eligible_only else data.people
    order = ["XS", "S", "M", "L", "XL", "XXL", "2XL", "3XL"]
    out = pd.DataFrame({
        "hoodie": p["hoodie"].str.upper().str.strip().value_counts(),
        "shirt": p["shirt"].str.upper().str.strip().value_counts(),
    }).fillna(0).astype(int)
    return out.reindex([s for s in order if s in out.index] + [s for s in out.index if s not in order])


def return_model(data: CleanData):
    """Member-level logistic regression: what predicts coming back?
    Exploratory only -- small sample."""
    try:
        import statsmodels.formula.api as smf
    except ImportError:
        return "statsmodels not installed -- run: pip install statsmodels"
    p = data.people[~data.people["is_eboard"] & (data.people["first_meeting"] < data.meetings["meeting"].max())].copy()
    if len(p) < 20:
        return None
    p["y"] = p["returned"].astype(int)
    p["first_year"] = (p["standing"].fillna("").str.lower().str.startswith("first")).astype(int)
    p["via_friend"] = p["heard_list"].map(lambda l: any("friend" in x.lower() for x in l)).astype(int)
    p["via_fair"] = p["heard_list"].map(lambda l: any("fair" in x.lower() for x in l)).astype(int)
    p["via_instagram"] = p["heard_list"].map(lambda l: any("instagram" in x.lower() for x in l)).astype(int)
    try:
        return smf.logit("y ~ first_year + via_friend + via_fair + via_instagram", p).fit(disp=False)
    except Exception as e:  # perfect separation etc. with tiny samples
        return f"Model could not be fit: {e}"


def meeting_model(data: CleanData):
    """Meeting-level Poisson regression on whatever is in MEETING_INFO.
    Only runs once there are enough meetings to say anything."""
    m = data.meetings
    predictors = [c for c in m.columns if c not in
                  {"meeting", "label", "attendees", "new", "eboard", "raw_submissions", "returning",
                   "non_eboard", "MEETING_TYPE_NAME"}]
    if not predictors:
        return "No meeting details yet: fill in MEETING_INFO at the top of the script."
    m = m.dropna(subset=predictors)
    if m.empty:
        return "MEETING_INFO dates don't match any meeting -- check the dates in MEETING_INFO."
    # MEETING_TYPE is a category code, not a quantity
    categorical = {c for c in predictors if c == "MEETING_TYPE" or m[c].dtype == object}
    if len(m) < 3 * len(predictors) + 3:
        out = []
        for c in predictors:
            avg = m.groupby(c)["non_eboard"].mean().round(1)
            if c == "MEETING_TYPE":
                avg.index = [MEETING_TYPES.get(int(i), str(int(i))) for i in avg.index]
            out.append(f"Avg non-e-board attendance by {c}:\n{avg.to_string()}")
        return (f"Only {len(m)} meetings for {len(predictors)} predictor(s) -- too few for a "
                f"regression. Showing averages instead:\n\n" + "\n\n".join(out))
    try:
        import statsmodels.formula.api as smf
    except ImportError:
        return "statsmodels not installed -- run: pip install statsmodels"
    rhs = " + ".join(f"C({c})" if c in categorical else c for c in predictors)
    return smf.poisson(f"non_eboard ~ {rhs}", m).fit(disp=False)


def suggestions(data: CleanData) -> list[str]:
    return data.raw[COL["suggestions"]].dropna().map(_norm_text).dropna().unique().tolist()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    ap = argparse.ArgumentParser(description="SASE attendance analysis")
    ap.add_argument("xlsx")
    ap.add_argument("--out", default="output")
    args = ap.parse_args()
    out = Path(args.out)
    out.mkdir(parents=True, exist_ok=True)

    data = load_clean(args.xlsx)
    print("=== Cleaning ===\n" + cleaning_report(data))

    print("\n=== Meetings ===")
    print(data.meetings[["label", "attendees", "new", "returning", "eboard", "non_eboard"]].to_string(index=False))

    plot_attendance(data, out)

    print(f"\n=== Incentives (>= {INCENTIVE_THRESHOLD} meetings) ===")
    print(data.people.loc[data.people["incentive_eligible"], ["name", "meetings_attended", "meeting_dates"]]
          .to_string(index=False))

    print("\n=== Retention by first meeting (non-e-board) ===")
    print(retention(data).to_string(index=False))

    print("\n=== Recruitment channels ===")
    print(recruitment_channels(data).to_string(index=False))

    print("\n=== Return model (exploratory) ===")
    rm = return_model(data)
    print(rm.summary2().tables[1].round(3) if hasattr(rm, "summary2") else rm)

    print("\n=== Meeting model ===")
    mm = meeting_model(data)
    print(mm.summary() if hasattr(mm, "summary") else mm)

    print("\n=== Suggestions ===")
    for s in suggestions(data):
        print(" -", s)

    attendees_cols = ["name", "email", "all_emails", "meetings_attended", "meeting_dates",
                      "incentive_eligible", "is_eboard", "standing", "hoodie", "shirt"]
    with pd.ExcelWriter(out / "sase_report.xlsx") as xw:
        data.people[attendees_cols].to_excel(xw, sheet_name="Attendees", index=False)
        data.meetings.to_excel(xw, sheet_name="Meetings", index=False)
        data.attendance.to_excel(xw, sheet_name="Attendance (clean)", index=False)
        retention(data).to_excel(xw, sheet_name="Retention", index=False)
        recruitment_channels(data).to_excel(xw, sheet_name="Channels", index=False)
        merch_sizes(data).to_excel(xw, sheet_name="Merch sizes")
        for name, s in demographics(data).items():
            s.to_frame("count").to_excel(xw, sheet_name=f"Demo - {name}"[:31])
    print(f"\nSaved charts and sase_report.xlsx to {out.resolve()}")


if __name__ == "__main__":
    main()
