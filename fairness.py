"""30-day fairness tracking and end-of-day reconciliation."""

from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd

HISTORY_PATH = Path(__file__).parent / "rotation_history.csv"

HISTORY_COLUMNS = [
    "Date",
    "Side",
    "Employee",
    "Employee ID",
    "Shift",
    "Position",
    "Area",
    "Family",
    "Source",
    "Reconciled By",
]


# Families the user asked to rotate fairly
FAIR_FAMILIES = {
    "DIVEST": ["divest"],
    "MAIN_FRONT": ["front entrance"],
    "SOUTH_FRONT": ["checkpoint - front entrance", "checkpoint – front entrance"],
    "ADA": ["ada entrance"],
    "PRECHECK_OVERFLOW": ["pre-check overflow", "precheck overflow"],
    "NORTH_CORRIDOR": ["north corridor"],
}


def _norm(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def classify_family(position_name: str) -> str:
    name = _norm(position_name).lower().replace("—", "-").replace("–", "-")
    if "divest" in name:
        return "DIVEST"
    if "ada" in name:
        return "ADA"
    if "pre-check overflow" in name or "precheck overflow" in name:
        return "PRECHECK_OVERFLOW"
    if "north corridor" in name:
        return "NORTH_CORRIDOR"
    if "front entrance" in name and "checkpoint" in name:
        return "SOUTH_FRONT"
    if "front entrance" in name:
        return "MAIN_FRONT"
    return ""


def is_fair_tracked(position_name: str) -> bool:
    return classify_family(position_name) != ""


def load_history(path: Path | None = None) -> pd.DataFrame:
    path = path or HISTORY_PATH
    if path.exists():
        df = pd.read_csv(path)
        for col in HISTORY_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        return df[HISTORY_COLUMNS]
    return pd.DataFrame(columns=HISTORY_COLUMNS)


def save_history(df: pd.DataFrame, path: Path | None = None) -> Path:
    path = path or HISTORY_PATH
    out = df.copy()
    for col in HISTORY_COLUMNS:
        if col not in out.columns:
            out[col] = ""
    out[HISTORY_COLUMNS].to_csv(path, index=False)
    return path


def _as_date(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    if hasattr(value, "date") and not isinstance(value, str):
        try:
            return value.date() if hasattr(value, "date") and not isinstance(value, datetime) else value
        except Exception:
            pass
    try:
        return pd.to_datetime(value).date()
    except Exception:
        return None


def last_30(history: pd.DataFrame, schedule_date) -> pd.DataFrame:
    if history is None or history.empty:
        return history if history is not None else pd.DataFrame(columns=HISTORY_COLUMNS)
    end = _as_date(schedule_date)
    if end is None:
        return history
    start = end - timedelta(days=29)
    dates = pd.to_datetime(history["Date"], errors="coerce").dt.date
    mask = dates.notna() & (dates >= start) & (dates <= end)
    return history.loc[mask].copy()


def burden_table(history: pd.DataFrame, schedule_date) -> pd.DataFrame:
    window = last_30(history, schedule_date)
    if window is None or window.empty:
        return pd.DataFrame(columns=["Employee", "Family", "Days", "Consecutive"])
    rows = []
    for emp, grp in window.groupby(window["Employee"].map(_norm)):
        if not emp:
            continue
        for family, fgrp in grp.groupby(grp["Family"].map(_norm)):
            if not family:
                continue
            days = sorted({d for d in pd.to_datetime(fgrp["Date"], errors="coerce").dt.date if pd.notna(d)})
            consec = 0
            end = _as_date(schedule_date)
            if end is not None:
                cursor = end
                dayset = set(days)
                while cursor in dayset:
                    consec += 1
                    cursor = cursor - timedelta(days=1)
            rows.append({
                "Employee": emp,
                "Family": family,
                "Days": len(days),
                "Assignments": len(fgrp),
                "Consecutive": consec,
            })
    return pd.DataFrame(rows)


def score_for(employee_name: str, position_name: str, history: pd.DataFrame, schedule_date) -> tuple:
    """Lower tuple = higher priority for the next hard post."""
    family = classify_family(position_name)
    name = _norm(employee_name)
    table = burden_table(history, schedule_date)
    family_days = 0
    family_consec = 0
    total_days = 0
    if table is not None and not table.empty:
        mine = table[table["Employee"].map(_norm).str.lower() == name.lower()]
        total_days = int(mine["Days"].sum()) if not mine.empty else 0
        if family:
            fam = mine[mine["Family"] == family]
            if not fam.empty:
                family_days = int(fam["Days"].sum())
                family_consec = int(fam["Consecutive"].max())
    return (family_days, family_consec, total_days, name.lower())


def sort_pool_by_fairness(pool: pd.DataFrame, position_name: str, history: pd.DataFrame, schedule_date) -> pd.DataFrame:
    if pool is None or pool.empty or not is_fair_tracked(position_name):
        return pool
    scored = pool.copy()
    scored["_fair"] = scored["Employee Name"].map(
        lambda n: score_for(n, position_name, history, schedule_date)
    )
    scored = scored.sort_values("_fair")
    return scored.drop(columns=["_fair"])


def rows_from_board(board: pd.DataFrame, schedule_date, side: str, source: str, reconciled_by: str) -> list:
    rows = []
    date_text = schedule_date.isoformat() if hasattr(schedule_date, "isoformat") else str(schedule_date)
    if board is None or board.empty:
        return rows
    for _, rec in board.iterrows():
        name = _norm(rec.get("Assigned Employee"))
        pname = _norm(rec.get("Position Name"))
        if name in {"", "UNFILLED", "NOT NEEDED", "NOT STAFFED", "NOT STAFFED PM"}:
            continue
        if _norm(rec.get("Area")) == "Rover / Relief Pool":
            continue
        family = classify_family(pname)
        if not family:
            continue
        rows.append({
            "Date": date_text,
            "Side": side,
            "Employee": name,
            "Employee ID": _norm(rec.get("Employee ID")),
            "Shift": _norm(rec.get("Shift")),
            "Position": pname,
            "Area": _norm(rec.get("Area")),
            "Family": family,
            "Source": source,
            "Reconciled By": reconciled_by,
        })
    return rows


def apply_reconciliation(history: pd.DataFrame, am_board: pd.DataFrame, pm_board: pd.DataFrame,
                         schedule_date, reconciled_by: str, sides: str = "BOTH") -> pd.DataFrame:
    """Replace that day's history rows for the chosen side(s) with actual board names."""
    date_text = schedule_date.isoformat() if hasattr(schedule_date, "isoformat") else str(schedule_date)
    hist = history.copy() if history is not None else pd.DataFrame(columns=HISTORY_COLUMNS)
    sides = _norm(sides).upper()
    keep_mask = hist["Date"].astype(str) != date_text
    if sides == "AM":
        keep_mask = (hist["Date"].astype(str) != date_text) | (hist["Side"].astype(str) != "AM")
    elif sides == "PM":
        keep_mask = (hist["Date"].astype(str) != date_text) | (hist["Side"].astype(str) != "PM")
    kept = hist.loc[keep_mask].copy() if not hist.empty else hist

    new_rows = []
    if sides in {"AM", "BOTH"}:
        new_rows.extend(rows_from_board(am_board, schedule_date, "AM", "reconciled", reconciled_by))
    if sides in {"PM", "BOTH"}:
        new_rows.extend(rows_from_board(pm_board, schedule_date, "PM", "reconciled", reconciled_by))
    added = pd.DataFrame(new_rows, columns=HISTORY_COLUMNS)
    out = pd.concat([kept, added], ignore_index=True)
    return out
