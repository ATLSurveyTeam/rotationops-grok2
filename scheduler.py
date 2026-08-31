"""
ATL Employee Scheduler 3.1 — AM engine (first Streamlit version)

Rules encoded from ATL_Employee_Scheduler_3.1:
- Fill all 43 Client Required slots.
- HARD 3:45 OPEN posts must open with 3:45 AM staff.
- START 5:45 posts start with 5:45 AM staff.
- Leads are excluded from posts and from break relief.
- Mentors + trainees stay together later; first version keeps them off split posts.
- Survey seats fill only when surveys_needed is True.
- Break relief follows Break Relief Guide (Primary → Secondary → Tertiary).
- Relief person must have a different lunch window than the person they cover.
"""

from collections import defaultdict
import pandas as pd


LUNCH_BY_SHIFT = {
    "3:45 AM": ["8:00 AM", "9:00 AM"],
    "5:45 AM": ["9:00 AM", "10:00 AM"],
    "12:15 PM": ["4:00 PM", "5:00 PM"],
    "2:00 PM": ["6:00 PM", "7:00 PM"],
}


def _norm(val) -> str:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return ""
    return str(val).strip()


def _yes(val) -> bool:
    return _norm(val).upper() in {"Y", "YES", "TRUE", "1"}


def pick_lunch(shift: str, counts: dict) -> str:
    options = LUNCH_BY_SHIFT.get(_norm(shift), ["12:00 PM"])
    best = min(options, key=lambda t: counts.get(t, 0))
    counts[best] = counts.get(best, 0) + 1
    return best


def load_workbook_tables(path: str) -> dict:
    positions = pd.read_excel(path, sheet_name="Position Library")
    positions.columns = [str(c).strip() for c in positions.columns]
    positions = positions.dropna(subset=["Position"]).copy()

    relief = pd.read_excel(path, sheet_name="Break Relief Guide")
    relief.columns = [str(c).strip() for c in relief.columns]

    employees = pd.read_excel(path, sheet_name="Employee Master")
    employees.columns = [str(c).strip() for c in employees.columns]
    employees = employees[employees["Active"].map(_yes)].copy()

    return {
        "positions": positions,
        "relief": relief,
        "employees": employees,
    }


def _needs_divest(position_name: str) -> bool:
    return "divest" in position_name.lower()


def _needs_info(position_name: str) -> bool:
    return "information desk" in position_name.lower()


def _needs_main_inside(position_name: str) -> bool:
    return "inside center" in position_name.lower()


def _eligible_for_position(emp: pd.Series, position_name: str) -> bool:
    if _yes(emp.get("Lead")):
        return False
    if _needs_divest(position_name) and not _yes(emp.get("Divest")):
        return False
    if _needs_info(position_name) and not _yes(emp.get("Information Desk")):
        return False
    if _needs_main_inside(position_name) and not _yes(emp.get("Main Inside")):
        return False
    return True


def build_am_schedule(
    employees: pd.DataFrame,
    positions: pd.DataFrame,
    relief_guide: pd.DataFrame,
    surveys_needed: bool = True,
) -> pd.DataFrame:
    emp = employees.copy()
    pos = positions.copy()
    guide = relief_guide.copy()

    emp.columns = [str(c).strip() for c in emp.columns]
    pos.columns = [str(c).strip() for c in pos.columns]
    guide.columns = [str(c).strip() for c in guide.columns]

    # Leads never take posts
    staff = emp[~emp["Lead"].map(_yes)].copy()

    am_345 = staff[staff["Shift"].map(_norm) == "3:45 AM"].copy()
    am_545 = staff[staff["Shift"].map(_norm) == "5:45 AM"].copy()

    used = set()
    lunch_counts = defaultdict(int)
    assignment = {}  # position name -> dict

    def take_eligible(pool: pd.DataFrame, position_name: str):
        for idx, row in pool.iterrows():
            eid = row["Employee ID"]
            if eid in used:
                continue
            if not _eligible_for_position(row, position_name):
                continue
            used.add(eid)
            return row
        return None

    # Sort: HARD 3:45 first, then START 5:45, then optional, surveys last
    def start_rank(rule: str) -> int:
        r = _norm(rule).upper()
        if "HARD 3:45" in r:
            return 0
        if "START 5:45" in r:
            return 1
        if "SURVEY" in r:
            return 3
        return 2

    pos["_rank"] = pos["Start of Day Rule"].map(start_rank)
    pos = pos.sort_values(["_rank", "Display Order"])

    for _, p in pos.iterrows():
        pname = _norm(p["Position"])
        area = _norm(p.get("Area"))
        rule = _norm(p.get("Start of Day Rule"))
        tier = _norm(p.get("Legacy Tier"))
        display = p.get("Display Order")

        is_survey = "survey" in area.lower() or "surveyor" in pname.lower()
        if is_survey and not surveys_needed:
            assignment[pname] = {
                "Pos #": display,
                "Area": area,
                "Position Name": pname,
                "Tier": tier or "Survey",
                "Start Rule": rule,
                "Assigned Employee": "NOT NEEDED",
                "Shift": "",
                "Assigned Lunch": "",
                "Status": "SURVEYS OFF",
                "Notes": "Survey flag is N",
            }
            continue

        person = None
        shift = ""
        notes = ""

        if "HARD 3:45" in rule.upper():
            person = take_eligible(am_345, pname)
            shift = "3:45 AM"
            notes = "HARD 3:45 OPEN"
            if person is None:
                person = take_eligible(am_545, pname)
                shift = "5:45 AM"
                notes = "HARD 3:45 OPEN — fallback 5:45"
        elif "START 5:45" in rule.upper():
            person = take_eligible(am_545, pname)
            shift = "5:45 AM"
            notes = "START 5:45"
            if person is None:
                person = take_eligible(am_345, pname)
                shift = "3:45 AM"
                notes = "START 5:45 — fallback 3:45 still on duty"
        elif is_survey:
            # Prefer survey-qualified 5:45 then 3:45
            survey_pool = pd.concat([am_545, am_345], ignore_index=True)
            survey_pool = survey_pool[survey_pool["Survey"].map(_yes)]
            person = take_eligible(survey_pool, pname)
            if person is not None:
                shift = _norm(person["Shift"])
            notes = "Survey seat"
        else:
            person = take_eligible(am_545, pname)
            if person is None:
                person = take_eligible(am_345, pname)
            if person is not None:
                shift = _norm(person["Shift"])
            notes = "Optional / zone fill"

        if person is None:
            assignment[pname] = {
                "Pos #": display,
                "Area": area,
                "Position Name": pname,
                "Tier": tier,
                "Start Rule": rule,
                "Assigned Employee": "UNFILLED",
                "Shift": "",
                "Assigned Lunch": "",
                "Status": "UNFILLED",
                "Notes": notes or "No eligible AM staff left",
            }
            continue

        lunch = pick_lunch(shift, lunch_counts)
        assignment[pname] = {
            "Pos #": display,
            "Area": area,
            "Position Name": pname,
            "Tier": tier,
            "Start Rule": rule,
            "Assigned Employee": _norm(person["Employee Name"]),
            "Employee ID": _norm(person["Employee ID"]),
            "Shift": shift,
            "Assigned Lunch": lunch,
            "Status": "ASSIGNED",
            "Notes": notes,
        }

    # Relief lookup by position name
    relief_map = {}
    for _, row in guide.iterrows():
        relief_map[_norm(row.get("Position Needing Relief"))] = [
            _norm(row.get("Primary Relief")),
            _norm(row.get("Secondary Relief")),
            _norm(row.get("Tertiary Relief")),
        ]

    rows = []
    for _, p in pos.sort_values("Display Order").iterrows():
        pname = _norm(p["Position"])
        rec = dict(assignment.get(pname, {}))
        if not rec:
            continue

        candidates = [c for c in relief_map.get(pname, []) if c]
        relief_name = ""
        relief_from = ""
        relief_lunch = ""
        status = rec.get("Status", "")

        if rec.get("Assigned Employee") in {"UNFILLED", "NOT NEEDED"}:
            rec["Break Relief Agent"] = ""
            rec["Relief From Position"] = ""
            rec["Relief Lunch"] = ""
            rows.append(rec)
            continue

        if not candidates or candidates[0].lower().startswith("no relief"):
            rec["Break Relief Agent"] = "No Relief Needed"
            rec["Relief From Position"] = ""
            rec["Relief Lunch"] = ""
            rec["Status"] = "OK"
            rows.append(rec)
            continue

        primary_lunch = rec.get("Assigned Lunch", "")
        found = False
        for cand_pos in candidates:
            if cand_pos.lower() == "rover/relief":
                rec["Break Relief Agent"] = "Rover / Relief"
                rec["Relief From Position"] = "Rover / Relief"
                rec["Relief Lunch"] = ""
                rec["Status"] = "OK — Rover"
                found = True
                break
            other = assignment.get(cand_pos)
            if not other:
                continue
            other_name = other.get("Assigned Employee", "")
            if other_name in {"", "UNFILLED", "NOT NEEDED"}:
                continue
            other_lunch = other.get("Assigned Lunch", "")
            if other_lunch and other_lunch == primary_lunch:
                continue
            relief_name = other_name
            relief_from = cand_pos
            relief_lunch = other_lunch
            found = True
            break

        rec["Break Relief Agent"] = relief_name
        rec["Relief From Position"] = relief_from
        rec["Relief Lunch"] = relief_lunch
        if found and relief_name:
            rec["Status"] = "OK"
        elif found and rec.get("Status") == "OK — Rover":
            pass
        else:
            rec["Status"] = "NO RELIEF" if not relief_name else rec.get("Status")
            if not found:
                rec["Status"] = "NO RELIEF / SAME LUNCH"
                rec["Notes"] = (rec.get("Notes", "") + " | Relief chain blocked").strip(" |")

        rows.append(rec)

    cols = [
        "Pos #", "Area", "Position Name", "Tier", "Start Rule",
        "Assigned Employee", "Shift", "Assigned Lunch",
        "Break Relief Agent", "Relief From Position", "Relief Lunch",
        "Status", "Notes",
    ]
    out = pd.DataFrame(rows)
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    return out[cols]


def build_schedule(employees: pd.DataFrame, positions: pd.DataFrame, relief_guide: pd.DataFrame | None = None, surveys_needed: bool = True) -> pd.DataFrame:
    """Back-compatible wrapper used by app.py."""
    if relief_guide is None:
        relief_guide = pd.DataFrame(columns=["Position Needing Relief", "Primary Relief", "Secondary Relief", "Tertiary Relief"])
    return build_am_schedule(employees, positions, relief_guide, surveys_needed=surveys_needed)

def _handoff_shift(pm_handoff: str, am_shift: str) -> str:
    text = _norm(pm_handoff).lower()
    if "not staffed" in text:
        return "NOT STAFFED"
    if "survey" in text:
        return "SURVEY"
    if "12:15" in text:
        return "12:15 PM"
    if "2:00" in text:
        return "2:00 PM"
    # Covered by either — follow AM person still on post
    if "3:45" in _norm(am_shift):
        return "12:15 PM"
    return "2:00 PM"


def build_pm_schedule(
    employees: pd.DataFrame,
    positions: pd.DataFrame,
    relief_guide: pd.DataFrame,
    am_board: pd.DataFrame,
    surveys_needed: bool = True,
) -> pd.DataFrame:
    emp = employees.copy()
    pos = positions.copy()
    staff = emp[~emp["Lead"].map(_yes)].copy()

    pm_1215 = staff[staff["Shift"].map(_norm) == "12:15 PM"].copy()
    pm_200 = staff[staff["Shift"].map(_norm) == "2:00 PM"].copy()

    used = set()
    lunch_counts = defaultdict(int)
    am_by_name = {
        _norm(r["Position Name"]): r for _, r in am_board.iterrows()
    }

    def take_eligible(pool: pd.DataFrame, position_name: str):
        for idx, row in pool.iterrows():
            eid = row["Employee ID"]
            if eid in used:
                continue
            if not _eligible_for_position(row, position_name):
                continue
            used.add(eid)
            return row
        return None

    assignment = {}
    pos = pos.sort_values("Display Order")

    for _, p in pos.iterrows():
        pname = _norm(p["Position"])
        area = _norm(p.get("Area"))
        tier = _norm(p.get("Legacy Tier"))
        display = p.get("Display Order")
        handoff = _norm(p.get("PM Handoff"))
        am_row = am_by_name.get(pname, {})
        am_shift = _norm(am_row.get("Shift", ""))
        am_name = _norm(am_row.get("Assigned Employee", ""))

        is_survey = "survey" in area.lower() or "surveyor" in pname.lower()
        wanted = _handoff_shift(handoff, am_shift)

        if wanted == "NOT STAFFED":
            assignment[pname] = {
                "Pos #": display, "Area": area, "Position Name": pname, "Tier": tier,
                "PM Handoff Rule": handoff, "Assigned Employee": "NOT STAFFED",
                "Shift": "", "Assigned Lunch": "", "Status": "NOT STAFFED PM",
                "Notes": f"AM was {am_name}" if am_name else "Not staffed in PM",
            }
            continue

        if is_survey and not surveys_needed:
            assignment[pname] = {
                "Pos #": display, "Area": area, "Position Name": pname, "Tier": tier or "Survey",
                "PM Handoff Rule": handoff, "Assigned Employee": "NOT NEEDED",
                "Shift": "", "Assigned Lunch": "", "Status": "SURVEYS OFF",
                "Notes": "PM survey flag is N",
            }
            continue

        person = None
        shift = wanted if wanted in {"12:15 PM", "2:00 PM"} else "12:15 PM"

        if is_survey:
            survey_pool = pd.concat([pm_1215, pm_200], ignore_index=True)
            survey_pool = survey_pool[survey_pool["Survey"].map(_yes)]
            person = take_eligible(survey_pool, pname)
            if person is not None:
                shift = _norm(person["Shift"])
        elif shift == "12:15 PM":
            person = take_eligible(pm_1215, pname)
            if person is None:
                person = take_eligible(pm_200, pname)
                if person is not None:
                    shift = "2:00 PM"
        else:
            person = take_eligible(pm_200, pname)
            if person is None:
                person = take_eligible(pm_1215, pname)
                if person is not None:
                    shift = "12:15 PM"

        if person is None:
            assignment[pname] = {
                "Pos #": display, "Area": area, "Position Name": pname, "Tier": tier,
                "PM Handoff Rule": handoff, "Assigned Employee": "UNFILLED",
                "Shift": "", "Assigned Lunch": "", "Status": "UNFILLED",
                "Notes": f"Handoff from AM {am_name} ({am_shift})" if am_name else handoff,
            }
            continue

        lunch = pick_lunch(shift, lunch_counts)
        assignment[pname] = {
            "Pos #": display, "Area": area, "Position Name": pname, "Tier": tier,
            "PM Handoff Rule": handoff,
            "Assigned Employee": _norm(person["Employee Name"]),
            "Employee ID": _norm(person["Employee ID"]),
            "Shift": shift, "Assigned Lunch": lunch, "Status": "ASSIGNED",
            "Notes": f"Handoff from {am_name} ({am_shift})" if am_name else handoff,
        }

    relief_map = {}
    for _, row in relief_guide.iterrows():
        relief_map[_norm(row.get("Position Needing Relief"))] = [
            _norm(row.get("Primary Relief")),
            _norm(row.get("Secondary Relief")),
            _norm(row.get("Tertiary Relief")),
        ]

    rows = []
    for _, p in pos.iterrows():
        pname = _norm(p["Position"])
        rec = dict(assignment.get(pname, {}))
        if not rec:
            continue
        if rec.get("Assigned Employee") in {"UNFILLED", "NOT NEEDED", "NOT STAFFED"}:
            rec["Break Relief Agent"] = ""
            rec["Relief From Position"] = ""
            rec["Relief Lunch"] = ""
            rows.append(rec)
            continue

        candidates = [c for c in relief_map.get(pname, []) if c]
        if not candidates or candidates[0].lower().startswith("no relief"):
            rec["Break Relief Agent"] = "No Relief Needed"
            rec["Relief From Position"] = ""
            rec["Relief Lunch"] = ""
            rec["Status"] = "OK"
            rows.append(rec)
            continue

        primary_lunch = rec.get("Assigned Lunch", "")
        found = False
        for cand_pos in candidates:
            if cand_pos.lower() == "rover/relief":
                rec["Break Relief Agent"] = "Rover / Relief"
                rec["Relief From Position"] = "Rover / Relief"
                rec["Relief Lunch"] = ""
                rec["Status"] = "OK — Rover"
                found = True
                break
            other = assignment.get(cand_pos)
            if not other:
                continue
            other_name = other.get("Assigned Employee", "")
            if other_name in {"", "UNFILLED", "NOT NEEDED", "NOT STAFFED"}:
                continue
            other_lunch = other.get("Assigned Lunch", "")
            if other_lunch and other_lunch == primary_lunch:
                continue
            rec["Break Relief Agent"] = other_name
            rec["Relief From Position"] = cand_pos
            rec["Relief Lunch"] = other_lunch
            rec["Status"] = "OK"
            found = True
            break
        if not found:
            rec["Break Relief Agent"] = ""
            rec["Relief From Position"] = ""
            rec["Relief Lunch"] = ""
            rec["Status"] = "NO RELIEF / SAME LUNCH"
        rows.append(rec)

    cols = [
        "Pos #", "Area", "Position Name", "Tier", "PM Handoff Rule",
        "Assigned Employee", "Shift", "Assigned Lunch",
        "Break Relief Agent", "Relief From Position", "Relief Lunch",
        "Status", "Notes",
    ]
    out = pd.DataFrame(rows)
    for c in cols:
        if c not in out.columns:
            out[c] = ""
    return out[cols]


def build_full_day(employees, positions, relief_guide, am_surveys=True, pm_surveys=True):
    am = build_am_schedule(employees, positions, relief_guide, surveys_needed=am_surveys)
    pm = build_pm_schedule(employees, positions, relief_guide, am, surveys_needed=pm_surveys)
    return am, pm
