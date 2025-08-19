import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.image(logo_url, width=200)

# Launch control
if 'mode' not in st.session_state:
    st.session_state.mode = None

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.mode = "analyzer"
with col2:
    if st.button("📅 Course Date Calculator"):
        st.session_state.mode = "calculator"

# 📅 Course Date Calculator
if st.session_state.mode == "calculator":
    st.subheader("📅 Course Date Calculator")

    st.markdown("### 1. Calculate Course End Date")
    start_date = st.date_input("Start Date", value=datetime.today(), format="DD/MM/YYYY")
    duration_weeks = st.number_input("Duration (weeks)", min_value=1, value=1)
    if st.button("Calculate End Date"):
        end_date = start_date + timedelta(weeks=duration_weeks)
        st.success(f"📅 End Date: {end_date.strftime('%d/%m/%Y')}")

    st.markdown("---")

    st.markdown("### 2. Calculate Weeks Between Two Dates")
    proposed_start = st.date_input("Proposed Start Date", value=datetime.today(), key="start", format="DD/MM/YYYY")
    proposed_end = st.date_input("Proposed End Date", value=datetime.today(), key="end", format="DD/MM/YYYY")
    if st.button("Calculate Weeks Between"):
        if proposed_end >= proposed_start:
            delta_days = (proposed_end - proposed_start).days
            weeks_between = delta_days // 7
            st.success(f"📆 Total Weeks: {weeks_between} weeks")
        else:
            st.error("End date must be after start date.")

# Analyzer functionality
elif st.session_state.mode == "analyzer":
    st.subheader("📊 COE Student Analyzer")

    uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])

    if uploaded_file:
        try:
            df = pd.read_excel(uploaded_file, engine="openpyxl")

            # Normalize column names
            df.columns = df.columns.str.strip().str.upper()

            # Rename expected columns
            rename_map = {
                "COE CODE": "COE CODE",
                "COE STATUS": "COE STATUS",
                "FIRST NAME": "FIRST NAME",
                "SECOND NAME": "SECOND NAME",
                "FAMILY NAME": "FAMILY NAME",
                "COURSE CODE": "COURSE CODE",
                "COURSE NAME": "COURSE NAME",
                "DURATION IN WEEKS": "DURATION IN WEEKS",
                "PROPOSED START DATE": "Proposed Start Date",
                "PROPOSED END DATE": "Proposed End Date",
                "PROVIDER STUDENT ID": "Provider Student ID",
                "VISA END DATE": "Visa End Date",
                "VISA NON GRANT STATUS": "Visa Non Grant Status",
                "AGENT": "AGENT"
            }
            df = df.rename(columns={col: rename_map[col] for col in df.columns if col in rename_map})

            # Convert date columns
            df["Proposed Start Date"] = pd.to_datetime(df.get("Proposed Start Date"), errors="coerce")
            df["Proposed End Date"] = pd.to_datetime(df.get("Proposed End Date"), errors="coerce")
            if "Visa End Date" in df.columns:
                df["Visa End Date"] = pd.to_datetime(df.get("Visa End Date"), errors="coerce")

            # Drop rows with missing dates
            df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)

            # Tabs
            tab1, tab2, tab3 = st.tabs(["📅 Start Date Filter", "🛂 Visa Expiry", "📈 Weekly Starts"])

            with tab1:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status", statuses, default=statuses)
                min_date = df["Proposed Start Date"].min()
                max_date = df["Proposed Start Date"].max()
                start_date, end_date = st.date_input("📆 Select Start Date Range", [min_date, max_date])
                filtered_df = df[
                    (df["COE STATUS"].isin(selected_statuses)) &
                    (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
                    (df["Proposed Start Date"] <= pd.to_datetime(end_date))
                ]
                st.write(f"🔍 {len(filtered_df)} students found")
                st.dataframe(filtered_df)

            with tab2:
                if "Visa End Date" in df.columns:
                    visa_days = st.slider("📆 Visa expiring in next X days", 7, 180, 30)
                    today = pd.to_datetime(datetime.today().date())
                    future_limit = today + pd.to_timedelta(visa_days, unit="d")
                    df_visa_expiring = df[
                        (df["Visa End Date"] >= today) & (df["Visa End Date"] <= future_limit)
                    ]
                    st.write(f"🛂 {len(df_visa_expiring)} students with visa expiring in {visa_days} days")
                    st.dataframe(df_visa_expiring)
                else:
                    st.info("ℹ️ 'Visa End Date' column not found.")

            with tab3:
                df["Start Week"] = df["Proposed Start Date"].dt.isocalendar().week
                weekly_counts = df.groupby("Start Week")["Provider Student ID"].count().reset_index(name="Number of Starts")
                st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("📤 Upload an Excel file to begin")
