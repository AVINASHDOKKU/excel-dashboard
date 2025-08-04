import streamlit as st
import pandas as pd
from datetime import datetime

st.set_page_config(page_title="Excel Data Analyzer", layout="wide")
st.title("📊 Excel Data Analyzer - Export for CoE and Student Details")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])
if uploaded_file:
    df = pd.read_excel(uploaded_file, sheet_name=0)

    # --- CLEANING ---
    df.columns = df.columns.str.strip()
    for col in df.select_dtypes(include='object').columns:
        df[col] = df[col].astype(str).str.strip()

    date_columns = [
        "Proposed Start Date", "Proposed End Date", "Actual Start Date", "Actual End Date",
        "COE Created Date", "Visa Effective Date", "Visa End Date"
    ]
    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')

    st.success("✅ File uploaded and cleaned successfully.")
    st.subheader("📌 Sheet Preview (Full Data with Pagination)")

    # Show ALL rows with pagination and dynamic sizing
    st.dataframe(
        df,
        use_container_width=True,
        height=min(800, len(df) * 35 + 50)
    )

    # --- ANALYSIS SECTION ---
    analysis_type = st.selectbox("Choose Analysis Type", ["Active Students by Qualification"])

    if analysis_type == "Active Students by Qualification":
        st.subheader("🎓 Active Students by Qualification")
        ref_date = st.date_input("Select Date", datetime.today())
        min_weeks = st.number_input("Show students with duration less than (weeks)", min_value=1, value=12)

        status_active = ["APPROVED", "STUDYING", "VISA GRANTED"]
        status_all = status_active + ["CANCELLED", "FINISHED"]
        exclude_status = ["SAVED"]

        def filter_by_date(df_sub, mode):
            if mode == "Today":
                status_filter = status_active
                date = pd.to_datetime(ref_date)
            elif mode == "Past":
                status_filter = status_all
                date = pd.to_datetime(ref_date)
            else:
                status_filter = status_active
                date = pd.to_datetime(ref_date)

            df_filtered = df_sub[
                df_sub["COE Status"].isin(status_filter) &
                ~df_sub["COE Status"].isin(exclude_status) &
                df_sub["Proposed Start Date"].notna() &
                df_sub["Proposed End Date"].notna()
            ]

            if mode == "Past":
                return df_filtered[
                    (df_filtered["Proposed Start Date"] <= date) &
                    (df_filtered["Proposed End Date"] >= date)
                ]
            elif mode == "Future":
                return df_filtered[
                    (df_filtered["Proposed Start Date"] > date)
                ]
            else:
                return df_filtered[
                    (df_filtered["Proposed Start Date"] <= date) &
                    (df_filtered["Proposed End Date"] >= date)
                ]

        tabs = st.tabs(["📅 Active Today", "🕒 Past", "📆 Future"])
        modes = ["Today", "Past", "Future"]

        for tab, mode in zip(tabs, modes):
            with tab:
                filtered = filter_by_date(df, mode)
                filtered = filtered.drop_duplicates(subset=["COE Code"])
                if "Duration In Weeks" in filtered.columns:
                    filtered = filtered[pd.to_numeric(filtered["Duration In Weeks"], errors='coerce') < min_weeks]

                st.write(f"### {mode} Active Students: {len(filtered)} records found")
                st.dataframe(filtered, use_container_width=True)
                csv = filtered.to_csv(index=False).encode("utf-8")
                st.download_button("Download CSV", csv, file_name=f"active_students_{mode.lower()}.csv", mime="text/csv")
