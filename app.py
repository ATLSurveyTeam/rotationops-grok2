import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date
from io import BytesIO
from scheduler import (
    load_workbook_tables,
    build_full_day,
    build_official_sheet,
    fill_official_docx,
    apply_lead_sheet_overrides,
    duplicate_name_report,
)
from fairness import (
    load_history,
    save_history,
    apply_reconciliation,
    burden_table,
)
from employee_admin import (
    SHIFTS,
    YN,
    next_employee_id,
    save_employee_master,
    upsert_employee,
)

st.set_page_config(
    page_title="ATL Employee Scheduler 3.1",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("ATL Employee Scheduler 3.1")
st.caption("43-position client list  •  Lead sheet includes rovers  •  30-day fairness + end-of-day reconciliation")

WORKBOOK_PATH = Path(__file__).parent / "ATL_Employee_Scheduler_3.1.xlsx"

page = st.sidebar.radio("Page", ["Daily Schedule", "Employee Admin"])
st.sidebar.divider()
st.sidebar.header("Schedule Controls")
selected_date = st.sidebar.date_input("Select Schedule Date", value=date.today())
leads = st.sidebar.text_input("Leads", value="")
am_surveys = st.sidebar.selectbox("AM Surveys Needed?", ["Y", "N"]) == "Y"
pm_surveys = st.sidebar.selectbox("PM Surveys Needed?", ["Y", "N"]) == "Y"
run_button = st.sidebar.button("Run Schedule", type="primary")
st.sidebar.divider()
st.sidebar.caption("Type over names on the Official sheet, then click Apply Overrides.")
st.sidebar.caption("Reconcile actual posts at end of day to keep the 30-day rotation fair.")
st.sidebar.caption("Use Employee Admin to hire or deactivate people.")


@st.cache_data
def load_tables():
    if not WORKBOOK_PATH.exists():
        return None
    return load_workbook_tables(str(WORKBOOK_PATH))


data = load_tables()
if data is None:
    st.error("3.1 workbook not found.")
    st.stop()

positions = data["positions"]
employees = data["employees"]
relief = data["relief"]

if page == "Employee Admin":
    st.header("Employee Admin")
    st.caption("Add a new hire or set Active to N when someone leaves. Do not delete rows.")
    roster = pd.read_excel(WORKBOOK_PATH, sheet_name="Employee Master")
    roster.columns = [str(c).strip() for c in roster.columns]
    show_inactive = st.checkbox("Show inactive employees", value=False)
    view = roster.copy()
    if "Active" in view.columns and not show_inactive:
        view = view[view["Active"].astype(str).str.upper().isin(["Y", "YES", "TRUE", "1"])]
    display_cols = [c for c in [
        "Employee ID", "Employee Name", "Shift", "Days Off", "Lead", "Mentor",
        "Active", "LOA", "Divest", "Information Desk", "Main Inside", "Survey",
    ] if c in view.columns]
    st.dataframe(view[display_cols], use_container_width=True, hide_index=True)

    st.subheader("Add or update one employee")
    mode = st.radio("Action", ["Add new hire", "Update existing"], horizontal=True)
    existing_ids = roster["Employee ID"].astype(str).tolist() if "Employee ID" in roster.columns else []
    existing_names = roster["Employee Name"].astype(str).tolist() if "Employee Name" in roster.columns else []

    if mode == "Update existing":
        pick = st.selectbox("Select employee", sorted(existing_names))
        current = roster[roster["Employee Name"].astype(str) == pick].iloc[0]
        default_id = str(current.get("Employee ID", ""))
        default_name = str(current.get("Employee Name", ""))
        default_shift = str(current.get("Shift", "5:45 AM"))
        default_off = str(current.get("Days Off", ""))
        default_lead = "Y" if str(current.get("Lead", "N")).upper() in {"Y", "YES"} else "N"
        default_mentor = "Y" if str(current.get("Mentor", "N")).upper() in {"Y", "YES"} else "N"
        default_active = "Y" if str(current.get("Active", "Y")).upper() in {"Y", "YES"} else "N"
        default_divest = "Y" if str(current.get("Divest", "N")).upper() in {"Y", "YES"} else "N"
        default_info = "Y" if str(current.get("Information Desk", "N")).upper() in {"Y", "YES"} else "N"
        default_inside = "Y" if str(current.get("Main Inside", "N")).upper() in {"Y", "YES"} else "N"
        default_survey = "Y" if str(current.get("Survey", "N")).upper() in {"Y", "YES"} else "N"
        default_loa = "Y" if str(current.get("LOA", "N")).upper() in {"Y", "YES"} else "N"
    else:
        default_id = next_employee_id(roster)
        default_name = ""
        default_shift = "5:45 AM"
        default_off = ""
        default_lead = "N"
        default_mentor = "N"
        default_active = "Y"
        default_divest = "N"
        default_info = "N"
        default_inside = "N"
        default_survey = "N"
        default_loa = "N"

    c1, c2, c3 = st.columns(3)
    with c1:
        emp_id = st.text_input("Employee ID", value=default_id)
        emp_name = st.text_input("Employee Name", value=default_name)
        emp_shift = st.selectbox("Shift", SHIFTS, index=SHIFTS.index(default_shift) if default_shift in SHIFTS else 1)
    with c2:
        emp_off = st.text_input("Days Off (example: Sat, Sun)", value="" if default_off == "nan" else default_off)
        emp_active = st.selectbox("Active", YN, index=0 if default_active == "Y" else 1)
        emp_lead = st.selectbox("Lead", YN, index=0 if default_lead == "Y" else 1)
    with c3:
        emp_mentor = st.selectbox("Mentor", YN, index=0 if default_mentor == "Y" else 1)
        emp_divest = st.selectbox("Divest qualified", YN, index=0 if default_divest == "Y" else 1)
        emp_info = st.selectbox("Information Desk", YN, index=0 if default_info == "Y" else 1)
        emp_inside = st.selectbox("Main Inside", YN, index=0 if default_inside == "Y" else 1)
        emp_survey = st.selectbox("Survey", YN, index=0 if default_survey == "Y" else 1)
        emp_loa = st.selectbox("LOA", YN, index=0 if default_loa == "Y" else 1)

    if st.button("Save employee to workbook", type="primary"):
        if not emp_name.strip():
            st.warning("Name is required.")
        else:
            fields = {
                "Employee ID": emp_id.strip(),
                "Employee Name": emp_name.strip(),
                "Shift": emp_shift,
                "Days Off": emp_off.strip(),
                "Lead": emp_lead,
                "Mentor": emp_mentor,
                "Active": emp_active,
                "Information Desk": emp_info,
                "Main Inside": emp_inside,
                "Survey": emp_survey,
                "Divest": emp_divest,
                "LOA": emp_loa,
            }
            updated = upsert_employee(roster, fields)
            save_employee_master(WORKBOOK_PATH, updated)
            load_tables.clear()
            st.success(f"Saved {emp_name.strip()} ({emp_id.strip()}). Active = {emp_active}.")
            st.rerun()
    st.stop()

st.header("1. Selected Date")
st.write(f"**{selected_date.strftime('%A, %B %d, %Y')}**")

if run_button:
    history = load_history()
    am_board, pm_board = build_full_day(
        employees, positions, relief,
        am_surveys=am_surveys,
        pm_surveys=pm_surveys,
        schedule_date=selected_date,
        history=history,
    )
    official = build_official_sheet(am_board, pm_board, selected_date.isoformat(), leads=leads)
    st.session_state["am_board"] = am_board
    st.session_state["pm_board"] = pm_board
    st.session_state["official"] = official.drop(columns=["_header"], errors="ignore")

if "official" not in st.session_state:
    st.write("Set the date and click **Run Schedule** to build the first draft.")
    st.stop()

am_board = st.session_state["am_board"]
pm_board = st.session_state["pm_board"]
official = st.session_state["official"]

st.header("2. Lead sheet")
st.caption("Click a name cell and start typing. Matching employee names will appear. Clear a name for a call-out, then click Apply Overrides.")

name_options = [""]
if "Employee Name" in employees.columns:
    name_options += sorted({
        str(n).strip()
        for n in employees["Employee Name"].dropna().tolist()
        if str(n).strip()
    })
for extra in ("UNFILLED", "NOT NEEDED", "NOT STAFFED"):
    if extra not in name_options:
        name_options.append(extra)
# Keep any name already on the sheet so the editor does not blank it
for col in ("Name (AM)", "Name (PM)"):
    if col in official.columns:
        for n in official[col].dropna().astype(str):
            n = n.strip()
            if n and n not in name_options:
                name_options.append(n)

edited = st.data_editor(
    official,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key="lead_sheet_editor",
    column_config={
        "Name (AM)": st.column_config.SelectboxColumn(
            "Name (AM)",
            options=name_options,
            required=False,
        ),
        "Name (PM)": st.column_config.SelectboxColumn(
            "Name (PM)",
            options=name_options,
            required=False,
        ),
    },
)

apply = st.button("Apply Overrides", type="primary")
if apply:
    am_board, pm_board = apply_lead_sheet_overrides(edited, am_board, pm_board, employees)
    official = build_official_sheet(am_board, pm_board, selected_date.isoformat(), leads=leads)
    official = official.drop(columns=["_header"], errors="ignore")
    st.session_state["am_board"] = am_board
    st.session_state["pm_board"] = pm_board
    st.session_state["official"] = official
    st.success("Overrides applied. Rovers were used to fill cleared required posts when possible.")
    st.rerun()

ok_status = {"OK", "OK — Rover", "SURVEYS OFF", "NOT STAFFED PM", "LOCKED OVERRIDE", "FILLED FROM ROVER"}
am_flags = am_board[~am_board["Status"].astype(str).isin(ok_status)]
pm_flags = pm_board[~pm_board["Status"].astype(str).isin(ok_status)]

tab_am, tab_pm, tab_fair = st.tabs(["AM Coverage Board", "PM Coverage Board", "30-Day Fairness"])
with tab_am:
    st.dataframe(am_board, use_container_width=True, hide_index=True)
    if len(am_flags) > 0:
        st.subheader("AM rows that need review")
        st.dataframe(am_flags, use_container_width=True, hide_index=True)
with tab_pm:
    st.dataframe(pm_board, use_container_width=True, hide_index=True)
    if len(pm_flags) > 0:
        st.subheader("PM rows that need review")
        st.dataframe(pm_flags, use_container_width=True, hide_index=True)
with tab_fair:
    hist_now = load_history()
    burden = burden_table(hist_now, selected_date)
    st.caption("Tracked families: Divest (all), Main Front 1–2, South Front 1–2, ADA, Precheck Overflow, North Corridor.")
    if burden is None or burden.empty:
        st.info("No reconciled history yet. After the first end-of-day reconciliation, this table will fill in.")
    else:
        st.dataframe(burden.sort_values(["Family", "Days", "Consecutive"], ascending=[True, False, False]),
                     use_container_width=True, hide_index=True)

output = BytesIO()
with pd.ExcelWriter(output, engine="openpyxl") as writer:
    official.to_excel(writer, sheet_name="Official Assignment Sheet", index=False)
    am_board.to_excel(writer, sheet_name="AM Coverage Board", index=False)
    pm_board.to_excel(writer, sheet_name="PM Coverage Board", index=False)
dups = duplicate_name_report(am_board, pm_board)
if dups is not None and not dups.empty:
    st.error("Duplicate names found on the assignment sheet. Fix these before printing.")
    st.dataframe(dups, use_container_width=True, hide_index=True)
else:
    st.success("No duplicated names on the AM/PM assignment boards.")

template = Path(__file__).parent / "CXR_Daily_Assignment_Template.docx"
lead_docx = fill_official_docx(str(template), am_board, pm_board, selected_date, leads=leads, include_rovers=True)
recon_docx = fill_official_docx(str(template), am_board, pm_board, selected_date, leads=leads, include_rovers=False)

st.download_button(
    "Download Lead Word sheet (includes rovers at bottom)",
    data=lead_docx,
    file_name=f"CXR_Lead_Sheet_{selected_date.isoformat()}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
st.download_button(
    "Download Reconciled Word sheet (no rover / relief)",
    data=recon_docx,
    file_name=f"CXR_Reconciled_Sheet_{selected_date.isoformat()}.docx",
    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
)
st.download_button(
    "Download Excel copy (includes lead sheet + rovers)",
    data=output.getvalue(),
    file_name=f"CXR_Daily_Assignment_{selected_date.isoformat()}.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.header("3. End-of-day reconciliation")
st.caption(
    "AM supervisor reconciles AM actual posts. "
    "PM supervisor reconciles PM, and can also reconcile AM if it was missed. "
    "This updates the 30-day rotation and should be done from the names actually worked."
)
recon_who = st.text_input("Supervisor name", value="")
recon_side = st.selectbox("Reconcile which side?", ["BOTH", "AM", "PM"])
recon_btn = st.button("Save Reconciliation", type="primary")
if recon_btn:
    if not recon_who.strip():
        st.warning("Enter the supervisor name before saving.")
    else:
        hist = load_history()
        updated = apply_reconciliation(
            hist, am_board, pm_board, selected_date,
            reconciled_by=recon_who.strip(),
            sides=recon_side,
        )
        save_history(updated)
        st.success(
            f"Saved {recon_side} reconciliation for {selected_date.isoformat()} "
            f"by {recon_who.strip()}. Next Run Schedule will use this 30-day history."
        )
        st.dataframe(updated[updated["Date"].astype(str) == selected_date.isoformat()],
                     use_container_width=True, hide_index=True)
