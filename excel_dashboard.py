import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="COE Student Analyzer", layout="wide")
st.title("📘 COE Student Analyzer")

if 'launch' not in st.session_state:
    st.session_state.launch = False

if not st.session_state.launch:
    if st.button("🔍 Launch COE Analyzer"):
        st.session_state.launch = True
    st.stop()

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx"])

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
    filtered_df["Is Duplicate"] = filtered_df.duplicated(subset=["Provider Student ID"], keep=False)
    return filtered_df

def format_dates(df, columns):
    for col in columns:
        if col in df.columns:
            df[col] = df[col].dt.strftime('%d/%m/%Y')
    return df

def visa_expiry_tracker(df, days=30):
    if "Visa Expiry Date" not in df.columns:
        return pd.DataFrame()
    today = pd.to_datetime(datetime.date.today())
    future_limit = today + pd.Timedelta(days=days)
    return df[(df["Visa Expiry Date"] >= today) & (df["Visa Expiry Date"] <= future_limit)]

def coe_expiry_tracker(df, days=30):
    today = pd.to_datetime(datetime.date.today())
    future_limit = today + pd.Timedelta(days=days)
    return df[(df["Proposed End Date"] >= today) & (df["Proposed End Date"] <= future_limit)]

def course_duration_validator(df):
    df["Actual Weeks"] = (df["Proposed End Date"] - df["Proposed Start Date"]).dt.days // 7
    df["Duration Mismatch"] = df["Actual Weeks"] != pd.to_numeric(df["DURATION IN WEEKS"], errors='coerce')
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

def apply_duplicate_color(df):
    # Color duplicates for display
    df = df.copy()
    df["Highlight"] = df["Is Duplicate"].apply(lambda x: "background-color: yellow" if x else "")
    return df

if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = preprocess_data(df_raw)

        all_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status to include", all_statuses, default=all_statuses)

        min_date = df["Proposed Start Date"].min().date()
        max_date = df["Proposed Start Date"].max().date()
        start_date, end_date = st.date_input("Select Proposed Start Date Range", [min_date, max_date], format="DD/MM/YYYY")

        filtered_df = df[
            (df["COE STATUS"].isin(selected_statuses)) &
            (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
            (df["Proposed Start Date"] <= pd.to_datetime(end_date))
        ]

        filtered_df = detect_duplicates_by_id(filtered_df)
        formatted_filtered_df = format_dates(filtered_df.copy(), ["Proposed Start Date", "Proposed End Date", "Visa Expiry Date"])

        visa_days = st.slider("Visa expiring in next X days", 7, 180, 30)
        df_visa = visa_expiry_tracker(df, visa_days)
        formatted_visa_df = format_dates(df_visa.copy(), ["Visa Expiry Date"])

        coe_days = st.slider("COE expiring in next X days", 7, 180, 30)
        df_coe = coe_expiry_tracker(df, coe_days)
        formatted_coe_df = format_dates(df_coe.copy(), ["Proposed End Date"])

        df_mismatch = course_duration_validator(df)
        formatted_mismatch_df = format_dates(df_mismatch.copy(), ["Proposed Start Date", "Proposed End Date"])

        weekly_counts = weekly_start_count(df)
        df_agent = agent_summary(df)

        tabs = st.tabs(["🎯 Filtered Students", "🛂 Visa Expiry", "📝 COE Expiry", "⚠️ Duration Mismatch", "📅 Weekly Starts", "👤 Agent Summary", "📥 Contact Sheet"])

        with tabs[0]:
            st.subheader("🎯 Filtered Students")
            st.write(f"{len(filtered_df)} students found")
            # Add color on duplicate rows
            styled_df = formatted_filtered_df.style.apply(
                lambda x: ['background-color: yellow' if val else '' for val in x["Is Duplicate"]], axis=1
            )
            st.dataframe(styled_df, use_container_width=True)

        with tabs[1]:
            st.subheader("🛂 Visa Expiry Tracker")
            st.write(f"{len(formatted_visa_df)} students with visa expiring in {visa_days} days")
            st.dataframe(formatted_visa_df, use_container_width=True)

        with tabs[2]:
            st.subheader("📝 COE Expiry Tracker")
            st.write(f"{len(formatted_coe_df)} students with COE expiring in {coe_days} days")
            st.dataframe(formatted_coe_df, use_container_width=True)

        with tabs[3]:
            st.subheader("⚠️ Duration Mismatch")
            st.write(f"{len(formatted_mismatch_df)} students with mismatched durations")
            st.dataframe(formatted_mismatch_df, use_container_width=True)

        with tabs[4]:
            st.subheader("📅 Weekly Start Count")
            st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

        with tabs[5]:
            st.subheader("👤 Agent Summary")
            if not df_agent.empty:
                st.dataframe(df_agent, use_container_width=True)
            else:
                st.info("No agent data available.")

        with tabs[6]:
            st.subheader("📥 Download Contact Sheet")
            contact_df = df[["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]].drop_duplicates()
            csv = contact_df.to_csv(index=False).encode('utf-8')
            st.download_button("Download CSV", csv, file_name="contact_sheet.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("👆 Upload an Excel file to begin.")
