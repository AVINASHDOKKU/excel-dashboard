import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Show logo only once
if 'logo_shown' not in st.session_state:
    logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
    st.image(logo_url, width=200)
    st.session_state.logo_shown = True

# Launch control
if 'mode' not in st.session_state:
    st.session_state.mode = None

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🚀 Launch Analyzer", key="launch_analyzer"):
        st.session_state.mode = "analyzer"
with col2:
    if st.button("📅 Course Date Calculator", key="launch_calculator"):
        st.session_state.mode = "calculator"
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Show logo only once
if 'logo_shown' not in st.session_state:
    logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
    st.image(logo_url, width=200)
    st.session_state.logo_shown = True

# Launch control
if 'mode' not in st.session_state:
    st.session_state.mode = None

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🚀 Launch Analyzer", key="launch_analyzer"):
        st.session_state.mode = "analyzer"
with col2:
    if st.button("📅 Course Date Calculator", key="launch_calculator"):
        st.session_state.mode = "calculator"
# Analyzer functionality
elif st.session_state.mode == "analyzer":
    st.subheader("📊 COE Student Analyzer")

    uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"], key="file_uploader")

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
        "PROVIDER STUDENT ID": "Provider Student ID",
        "VISA END DATE": "Visa End Date",
        "VISA NON GRANT STATUS": "Visa Non Grant Status",
        "AGENT": "AGENT"
    }

    def normalize_columns(df):
        df.columns = df.columns.str.strip().str.upper()
        return df

    def rename_columns(df):
        return df.rename(columns={col: expected_columns[col] for col in df.columns if col in expected_columns})

    def preprocess_data(df):
        df = normalize_columns(df)
        df = rename_columns(df)
        df["Proposed Start Date"] = pd.to_datetime(df.get("Proposed Start Date"), errors="coerce")
        df["Proposed End Date"] = pd.to_datetime(df.get("Proposed End Date"), errors="coerce")
        if "Visa End Date" in df.columns:
            df["Visa End Date"] = pd.to_datetime(df.get("Visa End Date"), errors="coerce")
        df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
        return df

    def detect_duplicates_by_id(df):
        df["Is Duplicate"] = df.duplicated(subset=["Provider Student ID"], keep=False)
        return df

    def style_dates_and_duplicates(df):
        def highlight_row(row):
            return ['background-color: khaki'] * len(row) if row.get("Is Duplicate", False) else [''] * len(row)

        return df.style.apply(highlight_row, axis=1).format({
            "Proposed Start Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
            "Proposed End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
            "Visa End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
        })

    def visa_expiry_tracker(df, days=30):
        today = pd.to_datetime(datetime.today().date())
        future_limit = today + pd.to_timedelta(days, unit="d")
        return df[(df["Visa End Date"] >= today) & (df["Visa End Date"] <= future_limit)]

    def coe_expiry_tracker(df, within_days=30):
        today = pd.to_datetime(datetime.today().date())
        future_limit = today + pd.to_timedelta(within_days, unit="d")
        return df[(df["Proposed End Date"] >= today) & (df["Proposed End Date"] <= future_limit)]

    def course_duration_validator(df):
        df["Actual Weeks"] = (df["Proposed End Date"] - df["Proposed Start Date"]).dt.days // 7
        df["Duration Mismatch"] = df["Actual Weeks"] != df["DURATION IN WEEKS"]
        return df[df["Duration Mismatch"]]

    def weekly_start_count(df):
        df["Start Week"] = df["Proposed Start Date"].dt.isocalendar().week
        return df.groupby("Start Week")["Provider Student ID"].count().reset_index(name="Number of Starts")

    def agent_summary(df):
        if "AGENT" in df.columns:
            return df.groupby("AGENT").agg(
                Total_Students=("Provider Student ID", "count"),
                Active_Students=("COE STATUS", lambda x: (x == "Active").sum())
            ).reset_index()
        return pd.DataFrame()
# Analyzer functionality
elif st.session_state.mode == "analyzer":
    st.subheader("📊 COE Student Analyzer")

    uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"], key="file_uploader")

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
        "PROVIDER STUDENT ID": "Provider Student ID",
        "VISA END DATE": "Visa End Date",
        "VISA NON GRANT STATUS": "Visa Non Grant Status",
        "AGENT": "AGENT"
    }

    def normalize_columns(df):
        df.columns = df.columns.str.strip().str.upper()
        return df

    def rename_columns(df):
        return df.rename(columns={col: expected_columns[col] for col in df.columns if col in expected_columns})

    def preprocess_data(df):
        df = normalize_columns(df)
        df = rename_columns(df)
        df["Proposed Start Date"] = pd.to_datetime(df.get("Proposed Start Date"), errors="coerce")
        df["Proposed End Date"] = pd.to_datetime(df.get("Proposed End Date"), errors="coerce")
        if "Visa End Date" in df.columns:
            df["Visa End Date"] = pd.to_datetime(df.get("Visa End Date"), errors="coerce")
        df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
        return df

    def detect_duplicates_by_id(df):
        df["Is Duplicate"] = df.duplicated(subset=["Provider Student ID"], keep=False)
        return df

    def style_dates_and_duplicates(df):
        def highlight_row(row):
            return ['background-color: khaki'] * len(row) if row.get("Is Duplicate", False) else [''] * len(row)

        return df.style.apply(highlight_row, axis=1).format({
            "Proposed Start Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
            "Proposed End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
            "Visa End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
        })

    def visa_expiry_tracker(df, days=30):
        today = pd.to_datetime(datetime.today().date())
        future_limit = today + pd.to_timedelta(days, unit="d")
        return df[(df["Visa End Date"] >= today) & (df["Visa End Date"] <= future_limit)]

    def coe_expiry_tracker(df, within_days=30):
        today = pd.to_datetime(datetime.today().date())
        future_limit = today + pd.to_timedelta(within_days, unit="d")
        return df[(df["Proposed End Date"] >= today) & (df["Proposed End Date"] <= future_limit)]

    def course_duration_validator(df):
        df["Actual Weeks"] = (df["Proposed End Date"] - df["Proposed Start Date"]).dt.days // 7
        df["Duration Mismatch"] = df["Actual Weeks"] != df["DURATION IN WEEKS"]
        return df[df["Duration Mismatch"]]

    def weekly_start_count(df):
        df["Start Week"] = df["Proposed Start Date"].dt.isocalendar().week
        return df.groupby("Start Week")["Provider Student ID"].count().reset_index(name="Number of Starts")

    def agent_summary(df):
        if "AGENT" in df.columns:
            return df.groupby("AGENT").agg(
                Total_Students=("Provider Student ID", "count"),
                Active_Students=("COE STATUS", lambda x: (x == "Active").sum())
            ).reset_index()
        return pd.DataFrame()
except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("📤 Upload an Excel file to begin")
