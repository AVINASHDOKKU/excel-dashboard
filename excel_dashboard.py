import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="COE Student Analyzer", layout="wide")

st.title("📘 COE Student Analyzer")
uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# Define expected columns
expected_columns = {
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
    "PROVIDER STUDENT ID": "Provider Student ID"
}

# --- Utility Functions ---

def normalize_columns(df):
    df.columns = df.columns.str.strip().str.upper()
    return df

def rename_columns(df):
    return df.rename(columns={k: expected_columns[k] for k in df.columns if k in expected_columns})

def preprocess_data(df):
    df = normalize_columns(df)
    df = rename_columns(df)
    df = df[list(expected_columns.values())]
    df["Proposed Start Date"] = pd.to_datetime(df["Proposed Start Date"], errors="coerce")
    df["Proposed End Date"] = pd.to_datetime(df["Proposed End Date"], errors="coerce")
    df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
    return df

def detect_duplicates_by_id(filtered_df):
    # Detect duplicates only within the filtered DataFrame
    dup_key = ["Provider Student ID"]
    filtered_df["Is Duplicate"] = filtered_df.duplicated(subset=dup_key, keep=False)
    return filtered_df

def style_duplicates(df):
    def highlight_row(row):
        if row.get("Is Duplicate", False):
            return ['background-color: khaki'] * len(row)
        return [''] * len(row)

    return df.style.apply(highlight_row, axis=1).format({
        "Proposed Start Date": lambda x: x.strftime('%d-%m-%Y') if pd.notnull(x) else "",
        "Proposed End Date": lambda x: x.strftime('%d-%m-%Y') if pd.notnull(x) else ""
    })

def filter_by_date(df, mode, today, selected_statuses):
    df = df[df["COE STATUS"].isin(selected_statuses)]
    if mode == "Past":
        return df[df["Proposed End Date"] < today]
    elif mode == "Today":
        return df[(df["Proposed Start Date"] <= today) & (df["Proposed End Date"] >= today)]
    elif mode == "Future":
        return df[df["Proposed Start Date"] > today]
    return df

# --- Main App Logic ---

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = preprocess_data(df_raw)

        all_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status to include", all_statuses, default=all_statuses)

        today = pd.to_datetime(datetime.date.today())

        tab1, tab2, tab3 = st.tabs(["⏪ Past Students", "🟢 Today Active Students", "⏩ Future Students"])

        with tab1:
            df_past = filter_by_date(df, "Past", today, selected_statuses)
            df_past = detect_duplicates_by_id(df_past)
            st.write(f"Past Students: {len(df_past)} found")
            st.dataframe(style_duplicates(df_past), use_container_width=True)

        with tab2:
            df_today = filter_by_date(df, "Today", today, selected_statuses)
            df_today = detect_duplicates_by_id(df_today)
            st.write(f"Today Active Students: {len(df_today)} found")
            st.dataframe(style_duplicates(df_today), use_container_width=True)

        with tab3:
            df_future = filter_by_date(df, "Future", today, selected_statuses)
            df_future = detect_duplicates_by_id(df_future)
            st.write(f"Future Students: {len(df_future)} found")
            st.dataframe(style_duplicates(df_future), use_container_width=True)

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("👆 Upload an Excel file to begin analysis.")
