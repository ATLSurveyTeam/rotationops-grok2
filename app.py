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
)
from fairness import (
    load_history,
    save_history,
    apply_reconciliation,
    burden_table,
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

st.sidebar.header("Schedule Controls")
selected_date = st.sidebar.date_input("Select Schedule Date", value=date.today())
leads = st.sidebar.text_input("Leads", value="")
am_surveys = st.sidebar.selectbox("AM Surveys Needed?", ["Y", "N"]) == "Y"
pm_surveys = st.sidebar.selectbox("PM Surveys Needed?", ["Y", "N"]) == "Y"
run_button = st.sidebar.button("Run Schedule", type="primary")
st.sidebar.divider()
st.sidebar.caption("Type over names on the Official sheet, then click Apply Overrides.")
st.sidebar.caption("Reconcile actual posts at end of day to keep the 30-day rotation fair.")
st.sidebar.caption("Overnight comes later.")


@st.cache_data
def load_tables():
    if not WORKBOOK_PATH.exists():
        return None
    return load_workbook_tables(str(WORKBOOK_PATH))


data = load_tables()
st.header("1. Selected Date")
st.write(f"**{selected_date.strftime('%A, %B %d, %Y')}**")

if data is None:
    st.error("3.1 workbook not found.")
    st.stop()

positions = data["positions"]
employees = data["employees"]
relief = data["relief"]

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
st.caption("Type a new name, or clear a name for a call-out. Then click Apply Overrides.")

edited = st.data_editor(
    official,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    key="lead_sheet_editor",
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
template = Path(__file__).parent / "CXR_Daily_Assignment_Template.docx"
docx_bytes = fill_official_docx(str(template), am_board, pm_board, selected_date, leads=leads)

st.download_button(
    "Download Word Assignment Sheet (43 client positions only)",
    data=docx_bytes,
    file_name=f"CXR_Daily_Assignment_{selected_date.isoformat()}.docx",
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
