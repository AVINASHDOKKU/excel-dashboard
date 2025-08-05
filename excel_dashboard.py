import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="COE Student Analyzer", layout="wide")
pd.set_option("display.max_columns", None)
pd.set_option("styler.render.max_elements", 200000)

st.title("📊 COE Student Analyzer")

uploaded_file = st.sidebar.file_uploader("Upload Excel File", type=[".xlsx"])

if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Standardize column names
    df.columns = df.columns.str.upper()

    required_columns = [
        "COE STATUS", "FIRST NAME", "SECOND NAME", "FAMILY NAME",
        "STUDENT ID", "PROVIDER STUDENT ID", "COURSE NAME",
        "DURATION IN WEEKS", "PROPOSED START DATE", "PROPOSED END DATE"
    ]

    if not all(col in df.columns for col in required_columns):
        st.error(f"❌ Error: Missing one or more required columns: {required_columns}")
        st.stop()

    # Clean and parse dates
    df["PROPOSED START DATE"] = pd.to_datetime(df["PROPOSED START DATE"], errors="coerce")
    df["PROPOSED END DATE"] = pd.to_datetime(df["PROPOSED END DATE"], errors="coerce")

    df["FULL NAME"] = (
        df["FIRST NAME"].fillna("").astype(str).str.strip() + " " +
        df["SECOND NAME"].fillna("").astype(str).str.strip() + " " +
        df["FAMILY NAME"].fillna("").astype(str).str.strip()
    ).str.upper().str.replace("  ", " ").str.strip()

    df["PROVIDER STUDENT ID"] = df["PROVIDER STUDENT ID"].astype(str)

    st.sidebar.markdown("### Date Range Filter (Proposed Start Date)")
    min_date = df["PROPOSED START DATE"].min()
    max_date = df["PROPOSED START DATE"].max()
    start_date, end_date = st.sidebar.date_input(
        "Select Proposed Start Date Range (dd/mm/yyyy)",
        value=(min_date, max_date),
        format="DD/MM/YYYY"
    )

    st.sidebar.markdown("### COE Status Selection")
    available_statuses = df["COE STATUS"].dropna().unique().tolist()
    selected_statuses = st.sidebar.multiselect("Select COE Statuses", available_statuses, default=available_statuses)

    # Filter dataset
    filtered_df = df[
        (df["PROPOSED START DATE"] >= pd.to_datetime(start_date)) &
        (df["PROPOSED START DATE"] <= pd.to_datetime(end_date)) &
        (df["COE STATUS"].isin(selected_statuses))
    ].copy()

    st.write(f"Selected: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")
    st.write(f"\n\U0001F50D {len(filtered_df)} records found")

    # Highlight duplicates within the filtered data by PROVIDER STUDENT ID
    def highlight_duplicates(df):
        dup_mask = df["PROVIDER STUDENT ID"].duplicated(keep=False)
        styles = pd.DataFrame("" , index=df.index, columns=df.columns)
        styles.loc[dup_mask, :] = "background-color: gold"
        return styles

    st.dataframe(filtered_df.style.apply(highlight_duplicates, axis=None), use_container_width=True)

    # Duration mismatch validator
    def course_duration_validator(df):
        df["DURATION IN WEEKS"] = pd.to_numeric(df["DURATION IN WEEKS"], errors="coerce")
        df_clean = df.dropna(subset=["DURATION IN WEEKS", "PROPOSED START DATE", "PROPOSED END DATE"]).copy()
        df_clean["ACTUAL WEEKS"] = ((df_clean["PROPOSED END DATE"] - df_clean["PROPOSED START DATE"]).dt.days // 7).astype("Int64")
        df_clean["DURATION MISMATCH"] = df_clean["DURATION IN WEEKS"].astype("Int64") != df_clean["ACTUAL WEEKS"]
        return df_clean[df_clean["DURATION MISMATCH"]]

    with st.expander("⚠️ Duration Mismatch Checker"):
        mismatch_df = course_duration_validator(filtered_df)
        st.write(f"Mismatched Durations: {len(mismatch_df)}")
        st.dataframe(mismatch_df, use_container_width=True)

    # COE Expiry Tracking
    with st.expander("📆 COE Expiry Tracker"):
        today = pd.Timestamp.today()
        for days in [14, 30, 60]:
            expiring = filtered_df[
                (df["PROPOSED END DATE"] >= today) &
                (df["PROPOSED END DATE"] <= today + pd.Timedelta(days=days))
            ]
            st.write(f"Expiring in {days} days: {len(expiring)}")
            st.dataframe(expiring, use_container_width=True)

    # Visa Expiry Tracker
    visa_col = [col for col in df.columns if "VISA EXPIRY" in col.upper()]
    if visa_col:
        with st.expander("🛂 Visa Expiry Tracker"):
            df[visa_col[0]] = pd.to_datetime(df[visa_col[0]], errors="coerce")
            for days in [14, 30, 60]:
                expiring = filtered_df[
                    (df[visa_col[0]] >= today) &
                    (df[visa_col[0]] <= today + pd.Timedelta(days=days))
                ]
                st.write(f"Visa Expiring in {days} days: {len(expiring)}")
                st.dataframe(expiring, use_container_width=True)

    # Weekly Start Count
    with st.expander("📅 Weekly Start Count"):
        filtered_df["START WEEK"] = filtered_df["PROPOSED START DATE"].dt.to_period("W").astype(str)
        weekly_count = filtered_df.groupby("START WEEK").size().reset_index(name="Student Count")
        st.dataframe(weekly_count, use_container_width=True)

    # Contact sheet
    with st.expander("📥 Download Contact Sheet"):
        contact_df = filtered_df[["FULL NAME", "STUDENT ID", "COE STATUS", "PROPOSED START DATE"]]
        csv = contact_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Contact Sheet as CSV", data=csv, file_name="contact_sheet.csv", mime="text/csv")

    # Agent-wise summary
    agent_cols = [col for col in df.columns if "AGENT" in col.upper()]
    if agent_cols:
        with st.expander("👨‍💼 Agent Summary"):
            agent_df = filtered_df.groupby(agent_cols[0])["STUDENT ID"].count().reset_index(name="Total Students")
            st.dataframe(agent_df, use_container_width=True)
else:
    st.info("⬅️ Please upload a file to begin.")
