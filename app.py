import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date
from io import BytesIO
from scheduler import load_workbook_tables, build_full_day, build_official_sheet

st.set_page_config(
    page_title="ATL Employee Scheduler 3.1",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("ATL Employee Scheduler 3.1")
st.caption("43-position plan  •  AM + PM  •  Official Lead assignment sheet")

WORKBOOK_PATH = Path(__file__).parent / "ATL_Employee_Scheduler_3.1.xlsx"

st.sidebar.header("Schedule Controls")
selected_date = st.sidebar.date_input("Select Schedule Date", value=date.today())
leads = st.sidebar.text_input("Leads", value="")
am_surveys = st.sidebar.selectbox("AM Surveys Needed?", ["Y", "N"]) == "Y"
pm_surveys = st.sidebar.selectbox("PM Surveys Needed?", ["Y", "N"]) == "Y"
run_button = st.sidebar.button("Run Schedule", type="primary")
st.sidebar.divider()
st.sidebar.caption("Using ATL Employee Scheduler 3.1 — 43 client-required slots.")
st.sidebar.caption("Overnight will be added later.")


@st.cache_data
def load_tables():
    if not WORKBOOK_PATH.exists():
        return None
    return load_workbook_tables(str(WORKBOOK_PATH))


data = load_tables()

st.header("1. Selected Date")
st.write(f"**{selected_date.strftime('%A, %B %d, %Y')}**")

if data is None:
    st.error("3.1 workbook not found. Place ATL_Employee_Scheduler_3.1.xlsx next to app.py.")
    st.stop()

positions = data["positions"]
employees = data["employees"]
relief = data["relief"]

c1, c2, c3 = st.columns(3)
with c1:
    st.metric("Active employees", len(employees))
with c2:
    st.metric("Client-required positions", len(positions))
with c3:
    hard = positions["Start of Day Rule"].astype(str).str.contains("HARD 3:45", case=False, na=False)
    st.metric("HARD 3:45 OPEN posts", int(hard.sum()))

st.divider()
st.header("2. Schedule Generation")

if run_button:
    am_board, pm_board = build_full_day(
        employees, positions, relief,
        am_surveys=am_surveys,
        pm_surveys=pm_surveys,
    )
    official = build_official_sheet(am_board, pm_board, selected_date.isoformat(), leads=leads)
    show_official = official.drop(columns=["_header"], errors="ignore")

    ok_status = {"OK", "OK — Rover", "SURVEYS OFF", "NOT STAFFED PM"}
    am_flags = am_board[~am_board["Status"].astype(str).isin(ok_status)]
    pm_flags = pm_board[~pm_board["Status"].astype(str).isin(ok_status)]

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("AM posts", len(am_board))
    m2.metric("PM posts", len(pm_board))
    m3.metric("AM needs review", int(len(am_flags)))
    m4.metric("PM needs review", int(len(pm_flags)))

    tab_off, tab_am, tab_pm = st.tabs([
        "Official Assignment Sheet",
        "AM Coverage Board",
        "PM Coverage Board",
    ])
    with tab_off:
        st.caption("This matches the CXR Daily Assignment Sheet the leads already use.")
        st.dataframe(show_official, use_container_width=True, hide_index=True)
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

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        show_official.to_excel(writer, sheet_name="Official Assignment Sheet", index=False)
        am_board.to_excel(writer, sheet_name="AM Coverage Board", index=False)
        pm_board.to_excel(writer, sheet_name="PM Coverage Board", index=False)
    st.download_button(
        "Download Official Assignment Sheet",
        data=output.getvalue(),
        file_name=f"CXR_Daily_Assignment_{selected_date.isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.write("Set the date, AM surveys, and PM surveys. Then click **Run Schedule**.")
    st.write("The first tab will look like the official Lead assignment sheet.")
