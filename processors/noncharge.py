"""
processors/noncharge.py

Loads the TAS Team Time report (the new OpenAir export format) and builds the
Non-Charge Time view that replaces the old Utilization tab.

Report columns: Project - Name, Date, Employee, Time (Hours), Task, Notes, Description
The file has one title row above the header and three footer rows at the bottom.
"""

import pandas as pd

# Projects that represent non-charge time. Filtering on project matches the
# task-prefix filter exactly (NONCHARGE HRS *, NCH *, TTG *, Available Time)
# and is more robust to new task names being added.
NONCHARGE_PROJECTS = (
    "GTM - NONCHG",
    "GTM - TRAINING",
    "GTM - PTO",
    "GTM - HOLIDAYS",
)

AVAILABLE_TIME_TASK = "Available Time"

# Tasks that legitimately never carry a note, so we don't prompt for one.
NO_NOTE_EXPECTED = {
    "NONCHARGE HRS HOLIDAY",
    "NONCHARGE HRS FLOATING HOLIDAY",
    "NONCHARGE HRS PTO",
    "NCH MATERNITY/PATERNITY/ADOPT",
    "NCH BEREAVEMENT",
    "NCH JURY DUTY",
    "NCH PPL",
}


def load_time_report(file) -> pd.DataFrame:
    """Read the time report into a clean frame. Accepts a path or file-like object."""
    df = pd.read_csv(file, skiprows=1, encoding="utf-8-sig")
    df.columns = [str(c).strip() for c in df.columns]

    # Drop the grand-total row and the "Generated on" / "Filter set applied" footers.
    df = df[df["Employee"].notna() & (df["Employee"].astype(str).str.strip() != "")]

    df["Date"] = pd.to_datetime(df["Date"], format="%m/%d/%Y", errors="coerce")
    df = df[df["Date"].notna()]

    df["Time (Hours)"] = pd.to_numeric(df["Time (Hours)"], errors="coerce").fillna(0.0)
    for col in ("Project - Name", "Employee", "Task", "Notes", "Description"):
        df[col] = df[col].fillna("").astype(str).str.strip()

    df["Period"] = df["Date"].apply(half_month_period)
    df["Is Non-Charge"] = df["Project - Name"].isin(NONCHARGE_PROJECTS)

    return df.reset_index(drop=True)


def half_month_period(d: pd.Timestamp) -> str:
    """Bucket a date into a half-month period, e.g. 'Jan 1-15' / 'Jan 16-31'."""
    month = d.strftime("%b")
    if d.day <= 15:
        return f"{month} 1-15"
    last = d.days_in_month
    return f"{month} 16-{last}"


def build_noncharge(df: pd.DataFrame, periods=None, people=None) -> pd.DataFrame:
    """
    Detail rows for the Non-Charge Time tab / export.

    Available Time sorts to the top for each person, since that's the line
    that actually needs an answer.
    """
    nc = df[df["Is Non-Charge"]].copy()

    if periods:
        nc = nc[nc["Period"].isin(periods)]
    if people:
        nc = nc[nc["Employee"].isin(people)]

    nc["Available Time"] = nc["Task"] == AVAILABLE_TIME_TASK
    nc["Needs Response"] = nc["Available Time"] | (
        (nc["Notes"] == "")
        & (nc["Description"] == "")
        & (~nc["Task"].isin(NO_NOTE_EXPECTED))
    )
    nc["Response"] = ""

    nc = nc.sort_values(
        ["Employee", "Available Time", "Date"], ascending=[True, False, True]
    )

    return nc[[
        "Employee", "Date", "Period", "Project - Name", "Task",
        "Time (Hours)", "Notes", "Description", "Needs Response", "Response",
    ]].reset_index(drop=True)


def noncharge_summary(df: pd.DataFrame, periods=None) -> pd.DataFrame:
    """Hours by person by task, with an Available Time column pulled out."""
    nc = df[df["Is Non-Charge"]].copy()
    if periods:
        nc = nc[nc["Period"].isin(periods)]

    summary = nc.pivot_table(
        index="Employee", columns="Task", values="Time (Hours)",
        aggfunc="sum", fill_value=0.0,
    )
    summary["Total Non-Charge"] = summary.sum(axis=1)
    return summary.sort_values("Total Non-Charge", ascending=False)


def actual_hours(df: pd.DataFrame) -> pd.DataFrame:
    """
    Chargeable actuals by person / project / period — the replacement feed for
    variance.py, which previously read this from the OpenAir export.
    """
    ch = df[~df["Is Non-Charge"]]
    return (
        ch.groupby(["Employee", "Project - Name", "Period"], as_index=False)["Time (Hours)"]
        .sum()
        .rename(columns={"Time (Hours)": "Actual Hours"})
    )
