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
    "PROVIDER STUDENT ID": "Provider Student ID",
    "VISA EXPIRY DATE": "Visa Expiry Date",
    "AGENT": "AGENT"
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
    df["Proposed Start Date"] = pd.to_datetime(df.get("Proposed Start Date"), errors="coerce")
    df["Proposed End Date"] = pd.to_datetime(df.get("Proposed End Date"), errors="coerce")
    if "Visa Expiry Date" in df.columns:
        df["Visa Expiry Date"] = pd.to_datetime(df.get("Visa Expiry Date"), errors="coerce")
    df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
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

# --- New Suggested Modules ---

def visa_expiry_tracker(df, days=30):
    if "Visa Expiry Date" not in df.columns:
        return pd.DataFrame()
    future_limit = pd.to_datetime(datetime.date.today()) + pd.to_timedelta(days, unit="d")
    return df[(df["Visa Expiry Date"] >= pd.to_datetime(datetime.date.today())) & (df["Visa Expiry Date"] <= future_limit)]

def coe_expiry_tracker(df, within_days=30):
    future_limit = pd.to_datetime(datetime.date.today()) + pd.to_timedelta(within_days, unit="d")
    return df[(df["Proposed End Date"] >= pd.to_datetime(datetime.date.today())) & (df["Proposed End Date"] <= future_limit)]

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
    else:
        return pd.DataFrame()

# --- Main App Logic ---

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = preprocess_data(df_raw)

        all_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status to include", all_statuses, default=all_statuses)

        today = pd.to_datetime(datetime.date.today())

        tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
            "⏪ Past Students", "🟢 Today Active", "⏩ Future Students",
            "🛂 Visa Expiry Tracker", "📝 COE Expiry Tracker", "⏰ Course Duration Validator",
            "📅 Weekly Starts", "👤 Agent Summary"
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

        # Visa Expiry Tracker
        with tab4:
            visa_days = st.slider("Visa expiring in next X days", 7, 180, 30)
            df_visa = visa_expiry_tracker(df, visa_days)
            st.write(f"{len(df_visa)} students with visa expiring in {visa_days} days")
            st.dataframe(df_visa, use_container_width=True)

        # COE Expiry Tracker
        with tab5:
            coe_days = st.slider("COE expiring in next X days", 7, 180, 30)
            df_coe = coe_expiry_tracker(df, coe_days)
            st.write(f"{len(df_coe)} students with COE expiring in {coe_days} days")
            st.dataframe(df_coe, use_container_width=True)

        # Course Duration Validator
        with tab6:
            df_mismatch = course_duration_validator(df)
            st.write(f"{len(df_mismatch)} students with duration mismatch")
            st.dataframe(df_mismatch, use_container_width=True)

        # Weekly Starts
        with tab7:
            weekly_counts = weekly_start_count(df)
            st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

        # Agent Summary
        with tab8:
            df_agent = agent_summary(df)
            if not df_agent.empty:
                st.dataframe(df_agent, use_container_width=True)
            else:
                st.info("No agent column found in the uploaded file.")

        # Student Contact Sheet Generator
        st.sidebar.subheader("📥 Download Contact Sheet")
        contact_df = df[["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]].drop_duplicates()
        csv = contact_df.to_csv(index=False).encode('utf-8')
        st.sidebar.download_button("Download Contact Sheet CSV", csv, file_name="contact_sheet.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("👆 Upload an Excel file to begin analysis.")
