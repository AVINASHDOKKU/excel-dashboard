import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Excel Data Analyzer", layout="wide")

st.title("📊 Excel Data Analyzer - Export for CoE and Student Details")
st.markdown("Upload your Excel file")

uploaded_file = st.file_uploader("Drag and drop file here", type=["xlsx"])

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, skiprows=0)
    except Exception as e:
        st.error(f"❌ Failed to read Excel file: {e}")
        st.stop()

    st.success("✅ File uploaded and cleaned successfully.")
    
    # Remove fully empty columns
    df_raw.dropna(axis=1, how='all', inplace=True)

    # Convert date columns
    date_cols = ["Proposed Start Date", "Proposed End Date"]
    for col in date_cols:
        if col in df_raw.columns:
            df_raw[col] = pd.to_datetime(df_raw[col], errors='coerce')

    # Status selection
    available_statuses = df_raw["COE Status"].dropna().unique().tolist()
    selected_statuses = st.multiselect(
        "Select COE Status(es) to include in analysis",
        options=available_statuses,
        default=available_statuses
    )

    # Date reference
    ref_date = st.date_input(
        "Reference date for analysis",
        datetime.date.today()
    )

    st.markdown("## 📌 Sheet Preview")
    st.dataframe(df_raw.head(50), use_container_width=True)

    st.markdown("---")

    # Filter function
    def filter_by_date(df_sub, mode):
        date = pd.to_datetime(ref_date).normalize()
        df_filtered = df_sub[
            df_sub["COE Status"].isin(selected_statuses) &
            df_sub["Proposed Start Date"].notna()
        ]

        if mode == "Past":
            return df_filtered[df_filtered["Proposed End Date"] < date]

        elif mode == "Future":
            return df_filtered[df_filtered["Proposed Start Date"] > date]

        else:  # Today: show all started on/before today
            df_today = df_filtered[df_filtered["Proposed Start Date"] <= date]

            # Mark duplicates based on Provider Student ID
            if "Provider Student ID" in df_today.columns:
                df_today = df_today.copy()
                df_today["Is Duplicate"] = df_today.duplicated(subset="Provider Student ID", keep=False)
            return df_today

    # Tabs
    tab_past, tab_today, tab_future = st.tabs(["⏪ Past Students", "🟢 Today Active Students", "⏩ Future Students"])

    with tab_past:
        past_df = filter_by_date(df_raw, "Past")
        st.write(f"Past Students: {len(past_df)} records found")
        st.dataframe(past_df, use_container_width=True)

    with tab_today:
        today_df = filter_by_date(df_raw, "Today")
        st.write(f"Today Active Students (all students started on or before today): {len(today_df)} records found")

        # Highlight duplicates if present
        if "Is Duplicate" in today_df.columns:
            st.dataframe(today_df.style.apply(
                lambda row: ['background-color: yellow' if row['Is Duplicate'] else '' for _ in row],
                axis=1
            ), use_container_width=True)
        else:
            st.dataframe(today_df, use_container_width=True)

    with tab_future:
        future_df = filter_by_date(df_raw, "Future")
        st.write(f"Future Students: {len(future_df)} records found")
        st.dataframe(future_df, use_container_width=True)

else:
    st.info("👆 Upload an Excel file above to begin analysis.")
