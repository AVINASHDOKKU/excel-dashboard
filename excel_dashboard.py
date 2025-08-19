import streamlit as st
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.image(logo_url, width=200)
st.markdown("### COE Student Analyzer\nPowered by TEK4DAY")

# Launch control
if 'launch' not in st.session_state:
    st.session_state.launch = False
if not st.session_state.launch:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.launch = True
    st.stop()

# 📅 Course Date Calculator on Home Page
st.subheader("📅 Course Date Calculator")

# --- Section 1: Calculate End Date ---
st.markdown("### 1. Calculate Course End Date")
start_date = st.date_input("Start Date", value=datetime.today())
duration_weeks = st.number_input("Duration (weeks)", min_value=1, value=1)
if st.button("Calculate End Date"):
    end_date = start_date + timedelta(weeks=duration_weeks)
    st.success(f"📅 End Date: {end_date.strftime('%Y-%m-%d')}")

# --- Divider ---
st.markdown("---")

# --- Section 2: Calculate Weeks Between Dates ---
st.markdown("### 2. Calculate Weeks Between Two Dates")
proposed_start = st.date_input("Proposed Start Date", value=datetime.today(), key="start")
proposed_end = st.date_input("Proposed End Date", value=datetime.today(), key="end")
if st.button("Calculate Weeks Between"):
    if proposed_end >= proposed_start:
        delta_days = (proposed_end - proposed_start).days
        weeks_between = delta_days // 7
        st.success(f"📆 Total Weeks: {weeks_between} weeks")
    else:
        st.error("End date must be after start date.")
