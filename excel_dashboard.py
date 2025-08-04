import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

st.set_page_config(page_title="COE Student Analyzer", layout="wide")

st.title("📘 COE Student Analyzer")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

# Expected columns mapping
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
    "AGENT NAME": "AGENT NAME",  # Optional column if available
    "VISA EXPIRY DATE": "Visa Expiry Date"  # Optional
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
    # Keep only available expected columns
    cols_to_keep = [col for col in expected_columns.values() if col in df.columns]
    df = df[cols_to_keep]
    if "Proposed Start Date" in df:
        df["Proposed Start Date"] = pd.to_datetime(df["Proposed Start Date"], errors="coerce")
    if "Proposed End Date" in df:
        df["Proposed End Date"] = pd.to_datetime(df["Proposed End Date"], errors="coerce")
    if "Visa Expiry Date" in df:
        df["Visa Expiry Date"] = pd.to_datetime(df["Visa Expiry Date"], errors="coerce")
    return df

def detect_duplicates_by_id(filtered_df):
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
        "Proposed End Date": lambda x: x.strftime('%d-%m-%Y') if pd.notnull(x) else "",
        "Visa Expiry Date": lambda x: x.strftime('%d-%m-%Y') if pd.notnull(x) else ""
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

def course_duration_validator(df):
    df = df.copy()
    df["Actual Weeks"] = ((df["Proposed End Date"] - df["Proposed Start Date"]).dt.days / 7).round(1)
    df["Duration Mismatch"] = abs(df["Actual Weeks"] - df["DURATION IN WEEKS"]) > 1
    return df

def weekly_start_count(df):
    weekly = df.copy()
    weekly["Week Start"] = weekly["Proposed Start Date"].dt.to_period("W").apply(lambda r: r.start_time)
    count = weekly.groupby("Week Start").size().reset_index(name="Students Starting")
    return count

def monthly_start_count(df):
    monthly = df.copy()
    monthly["Month"] = monthly["Proposed Start Date"].dt.to_period("M").astype(str)
    count = monthly.groupby("Month").size().reset_index(name="Students Starting")
    return count

def agent_summary(df):
    if "AGENT NAME" not in df.columns:
        return None
    return df.groupby("AGENT NAME").agg(
        Students=("Provider Student ID", "count")
    ).reset_index()

def visa_expiry_filter(df, days_threshold, today):
    if "Visa Expiry Date" not in df.columns:
        return pd.DataFrame()
    future_date = today + pd.Timedelta(days=days_threshold)
    return df[(df["Visa Expiry Date"].notna()) & (df["Visa Expiry Date"] <= future_date)]

def coe_expiry_filter(df, days_threshold, today):
    future_date = today + pd.Timedelta(days=days_threshold)
    return df[df["Proposed End Date"] <= future_date]

def generate_contact_csv(df):
    contact_cols = ["FIRST NAME", "FAMILY NAME", "Provider Student ID"]
    available_cols = [c for c in contact_cols if c in df.columns]
    return df[available_cols].to_csv(index=False).encode("utf-8")


# --- Main App Logic ---

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = preprocess_data(df_raw)

        all_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status to include", all_statuses, default=all_statuses)

        today = pd.to_datetime(datetime.date.today())

        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "⏪ Past Students",
            "🟢 Today Active Students",
            "⏩ Future Students",
            "⚠️ Visa Expiry Tracker",
            "📅 COE Expiry Tracker",
            "✅ Course Duration Validator",
            "📈 Weekly/Monthly Starts"
        ])

        # Past Students
        with tab1:
            df_past = filter_by_date(df, "Past", today, selected_statuses)
            df_past = detect_duplicates_by_id(df_past)
            st.write(f"Past Students: {len(df_past)} found")
            st.dataframe(style_duplicates(df_past), use_container_width=True)

        # Today Active
        with tab2:
            df_today = filter_by_date(df, "Today", today, selected_statuses)
            df_today = detect_duplicates_by_id(df_today)
            st.write(f"Today Active Students: {len(df_today)} found")
            st.dataframe(style_duplicates(df_today), use_container_width=True)

        # Future Students
        with tab3:
            df_future = filter_by_date(df, "Future", today, selected_statuses)
            df_future = detect_duplicates_by_id(df_future)
            st.write(f"Future Students: {len(df_future)} found")
            st.dataframe(style_duplicates(df_future), use_container_width=True)

        # Visa Expiry
        with tab4:
            if "Visa Expiry Date" not in df.columns:
                st.warning("Visa Expiry Date column not found.")
            else:
                days = st.slider("Show visas expiring in next X days:", 7, 180, 30)
                visa_expiring = visa_expiry_filter(df, days, today)
                st.write(f"Students with Visa expiring in next {days} days: {len(visa_expiring)}")
                st.dataframe(visa_expiring, use_container_width=True)

        # COE Expiry
        with tab5:
            days = st.slider("Show COEs expiring in next X days:", 7, 180, 30)
            coe_expiring = coe_expiry_filter(df, days, today)
            st.write(f"Students with COE expiring in next {days} days: {len(coe_expiring)}")
            st.dataframe(coe_expiring, use_container_width=True)

        # Duration Validator
        with tab6:
            df_validated = course_duration_validator(df)
            st.write("Flagging inconsistencies between actual and recorded duration (>1 week difference):")
            st.dataframe(df_validated[df_validated["Duration Mismatch"]], use_container_width=True)

        # Weekly/Monthly Starts
        with tab7:
            st.subheader("Weekly Starts")
            weekly = weekly_start_count(df)
            st.dataframe(weekly, use_container_width=True)

            st.subheader("Monthly Starts")
            monthly = monthly_start_count(df)
            st.dataframe(monthly, use_container_width=True)

        # Agent Summary (if available)
        if "AGENT NAME" in df.columns:
            st.subheader("Agent-wise Summary")
            summary = agent_summary(df)
            st.dataframe(summary, use_container_width=True)

        # Download Contact Sheet
        st.subheader("🎯 Download Contact Sheet")
        filtered_df = df[df["COE STATUS"].isin(selected_statuses)]
        csv_data = generate_contact_csv(filtered_df)
        st.download_button("Download Contacts CSV", data=csv_data, file_name="contacts.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("👆 Upload an Excel file to begin analysis.")
