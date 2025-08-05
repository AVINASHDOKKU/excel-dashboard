import streamlit as st
import pandas as pd
from datetime import datetime
from io import BytesIO

st.set_page_config(page_title="COE Student Analyzer", layout="wide")
pd.set_option("display.max_columns", None)

st.title("📊 COE Student Analyzer")

uploaded_file = st.sidebar.file_uploader("📁 Upload Excel File", type=[".xlsx"])

# Required columns
required_columns = [
    "COE STATUS", "FIRST NAME", "SECOND NAME", "FAMILY NAME",
    "STUDENT ID", "PROVIDER STUDENT ID", "COURSE NAME",
    "DURATION IN WEEKS", "PROPOSED START DATE", "PROPOSED END DATE"
]

# Main logic if file uploaded
if uploaded_file:
    df = pd.read_excel(uploaded_file)

    # Show detected columns for debugging
    st.sidebar.markdown("### 🔍 Detected Columns:")
    st.sidebar.code("\n".join(df.columns.astype(str)), language="text")

    # Standardize column names
    df.columns = df.columns.str.upper().str.strip()

    # Check if all required columns are present
    if not all(col in df.columns for col in required_columns):
        missing = [col for col in required_columns if col not in df.columns]
        st.error(f"❌ Error: Missing one or more required columns:\n{missing}")
        st.stop()

    # Parse dates safely
    df["PROPOSED START DATE"] = pd.to_datetime(df["PROPOSED START DATE"], errors="coerce")
    df["PROPOSED END DATE"] = pd.to_datetime(df["PROPOSED END DATE"], errors="coerce")

    # Generate full name
    df["FULL NAME"] = (
        df["FIRST NAME"].fillna("") + " " +
        df["SECOND NAME"].fillna("") + " " +
        df["FAMILY NAME"].fillna("")
    ).str.upper().str.replace("  ", " ").str.strip()

    df["PROVIDER STUDENT ID"] = df["PROVIDER STUDENT ID"].astype(str)

    # Sidebar filters
    st.sidebar.markdown("### 📅 Date Range Filter")
    min_date = df["PROPOSED START DATE"].min()
    max_date = df["PROPOSED START DATE"].max()

    start_date, end_date = st.sidebar.date_input(
        "Filter by Proposed Start Date",
        value=(min_date, max_date),
        format="DD/MM/YYYY"
    )

    st.sidebar.markdown("### 🏷️ COE Status Filter")
    status_options = df["COE STATUS"].dropna().unique().tolist()
    selected_status = st.sidebar.multiselect("Select COE Status", status_options, default=status_options)

    # Filtered data
    filtered_df = df[
        (df["PROPOSED START DATE"] >= pd.to_datetime(start_date)) &
        (df["PROPOSED START DATE"] <= pd.to_datetime(end_date)) &
        (df["COE STATUS"].isin(selected_status))
    ].copy()

    st.success(f"✅ {len(filtered_df)} Records Found from {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")

    # Highlight duplicate PROVIDER STUDENT ID
    def highlight_duplicates(df):
        dup = df["PROVIDER STUDENT ID"].duplicated(keep=False)
        style = pd.DataFrame("", index=df.index, columns=df.columns)
        style.loc[dup, :] = "background-color: gold"
        return style

    st.dataframe(filtered_df.style.apply(highlight_duplicates, axis=None), use_container_width=True)

    # Duration mismatch checker
    def check_duration_mismatch(df):
        df["DURATION IN WEEKS"] = pd.to_numeric(df["DURATION IN WEEKS"], errors="coerce")
        df = df.dropna(subset=["PROPOSED START DATE", "PROPOSED END DATE", "DURATION IN WEEKS"]).copy()
        df["ACTUAL WEEKS"] = ((df["PROPOSED END DATE"] - df["PROPOSED START DATE"]).dt.days // 7).astype("Int64")
        df["DURATION MISMATCH"] = df["ACTUAL WEEKS"] != df["DURATION IN WEEKS"].astype("Int64")
        return df[df["DURATION MISMATCH"]]

    with st.expander("⚠️ Duration Mismatch Checker"):
        mismatch_df = check_duration_mismatch(filtered_df)
        st.write(f"⚠️ Mismatched: {len(mismatch_df)}")
        st.dataframe(mismatch_df, use_container_width=True)

    # COE expiry tracker
    with st.expander("📆 COE Expiry Tracker"):
        today = pd.Timestamp.today()
        for days in [14, 30, 60]:
            soon_expiring = filtered_df[
                (filtered_df["PROPOSED END DATE"] >= today) &
                (filtered_df["PROPOSED END DATE"] <= today + pd.Timedelta(days=days))
            ]
            st.write(f"⏳ Expiring in {days} days: {len(soon_expiring)}")
            st.dataframe(soon_expiring, use_container_width=True)

    # Visa expiry (if available)
    visa_cols = [col for col in df.columns if "VISA EXPIRY" in col]
    if visa_cols:
        with st.expander("🛂 Visa Expiry Tracker"):
            df[visa_cols[0]] = pd.to_datetime(df[visa_cols[0]], errors="coerce")
            for days in [14, 30, 60]:
                visa_exp = filtered_df[
                    (df[visa_cols[0]] >= today) &
                    (df[visa_cols[0]] <= today + pd.Timedelta(days=days))
                ]
                st.write(f"🛑 Visa expiring in {days} days: {len(visa_exp)}")
                st.dataframe(visa_exp, use_container_width=True)

    # Weekly starts
    with st.expander("📅 Weekly Start Count"):
        filtered_df["START WEEK"] = filtered_df["PROPOSED START DATE"].dt.to_period("W").astype(str)
        weekly_count = filtered_df.groupby("START WEEK").size().reset_index(name="Student Count")
        st.dataframe(weekly_count, use_container_width=True)

    # Contact Sheet
    with st.expander("📥 Download Contact Sheet"):
        contact_df = filtered_df[["FULL NAME", "STUDENT ID", "COE STATUS", "PROPOSED START DATE"]]
        csv = contact_df.to_csv(index=False).encode('utf-8')
        st.download_button("Download Contact Sheet", data=csv, file_name="contact_sheet.csv", mime="text/csv")

    # Agent Summary
    agent_cols = [col for col in df.columns if "AGENT" in col.upper()]
    if agent_cols:
        with st.expander("👤 Agent Summary"):
            agent_summary = filtered_df.groupby(agent_cols[0])["STUDENT ID"].count().reset_index(name="Total Students")
            st.dataframe(agent_summary, use_container_width=True)

else:
    st.info("⬅️ Please upload a valid Excel file to get started.")
