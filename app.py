import streamlit as st
import pandas as pd
from pathlib import Path
from datetime import date
from io import BytesIO
from scheduler import load_workbook_tables, build_am_schedule

st.set_page_config(
    page_title="ATL Employee Scheduler 3.1",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("ATL Employee Scheduler 3.1")
st.caption("43-position plan  •  AM engine  •  3:45 AM – 11:00 PM coverage")

WORKBOOK_PATH = Path(__file__).parent / "ATL_Employee_Scheduler_3.1.xlsx"

st.sidebar.header("Schedule Controls")
selected_date = st.sidebar.date_input("Select Schedule Date", value=date.today())
surveys_needed = st.sidebar.selectbox("AM Surveys Needed?", ["Y", "N"]) == "Y"
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
    board = build_am_schedule(
        employees=employees,
        positions=positions,
        relief_guide=relief,
        surveys_needed=surveys_needed,
    )

    unfilled = int((board["Status"] == "UNFILLED").sum())
    flags = board[~board["Status"].astype(str).isin(["OK", "OK — Rover", "SURVEYS OFF"])]

    m1, m2, m3 = st.columns(3)
    m1.metric("Posts generated", len(board))
    m2.metric("UNFILLED", unfilled)
    m3.metric("Needs review", int(len(flags)))

    st.subheader("Full Coverage Board")
    st.dataframe(board, use_container_width=True, hide_index=True)

    if len(flags) > 0:
        st.subheader("Rows that need review")
        st.dataframe(flags, use_container_width=True, hide_index=True)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        board.to_excel(writer, sheet_name="AM Coverage Board", index=False)
        for area, group in board.groupby("Area", sort=False):
            safe = str(area)[:28].replace("/", "-")
            group.to_excel(writer, sheet_name=safe or "Area", index=False)
    st.download_button(
        "Download AM Coverage Board",
        data=output.getvalue(),
        file_name=f"AM_Coverage_{selected_date.isoformat()}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
else:
    st.write("Select the date, set the survey flag, then click **Run Schedule**.")
    st.write("The board will list all 43 positions with assigned names and named break-relief positions.")
