import streamlit as st
import pandas as pd
import datetime

# --- Branding ---
st.set_page_config(page_title="COE Student Analyzer", layout="wide")
st.markdown("""
<div style='text-align: center'>
    <img src='TEK4DAY.PNG' width='120'>
    <h2>COE Student Analyzer</h2>
    <p><em>Powered by TEK4DAY</em></p>
</div>
""", unsafe_allow_html=True)

# --- Launch Button ---
if 'launch' not in st.session_state:
    st.session_state.launch = False

if not st.session_state.launch:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.launch = True
    st.stop()

# --- File Upload ---
uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])

# --- Expected Columns ---
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
    return df.rename(columns={col: expected_columns[col] for col in df.columns if col in expected_columns})

def preprocess_data(df):
    df = normalize_columns(df)
    df = rename_columns(df)
    df["Proposed Start Date"] = pd.to_datetime(df.get("Proposed Start Date"), errors="coerce", dayfirst=True)
    df["Proposed End Date"] = pd.to_datetime(df.get("Proposed End Date"), errors="coerce", dayfirst=True)
    if "Visa Expiry Date" in df.columns:
        df["Visa Expiry Date"] = pd.to_datetime(df.get("Visa Expiry Date"), errors="coerce", dayfirst=True)
    df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
    return df

def detect_duplicates_by_id(filtered_df):
    dup_key = ["Provider Student ID"]
    filtered_df["Is Duplicate"] = filtered_df.duplicated(subset=dup_key, keep=False)
    return filtered_df

def style_dates_and_duplicates(df):
    max_cells = 250000
    if df.size > max_cells:
        st.warning("⚠️ Too many cells to style. Displaying without formatting.")
        return df
    def highlight_row(row):
        if row.get("Is Duplicate", False):
            return ['background-color: khaki'] * len(row)
        return [''] * len(row)
    return df.style.apply(highlight_row, axis=1).format({
        "Proposed Start Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
        "Proposed End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
        "Visa Expiry Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
    })

# --- Analysis Modules ---
def visa_expiry_tracker(df, days=30):
    if "Visa Expiry Date" not in df.columns:
        return pd.DataFrame()
    today = pd.to_datetime(datetime.date.today())
    future_limit = today + pd.to_timedelta(days, unit="d")
    return df[(df["Visa Expiry Date"] >= today) & (df["Visa Expiry Date"] <= future_limit)]

def coe_expiry_tracker(df, within_days=30):
    today = pd.to_datetime(datetime.date.today())
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
    else:
        return pd.DataFrame()

# --- Main App Logic ---
if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file)
        df = preprocess_data(df_raw)

        all_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("🎯 Select COE Status to include", all_statuses, default=all_statuses)

        today = pd.to_datetime(datetime.date.today())

        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📅 Start Date Filter", "🛂 Visa Expiry", "📄 COE Expiry",
            "⏳ Duration Check", "📈 Weekly Starts", "🤝 Agent Summary", "📥 Download Contacts"
        ])

        with tab1:
            min_date = df["Proposed Start Date"].min()
            max_date = df["Proposed Start Date"].max()
            start_date, end_date = st.date_input("📆 Select Proposed Start Date Range", [min_date, max_date], format="DD/MM/YYYY")
            filtered_df = df[
                (df["COE STATUS"].isin(selected_statuses)) &
                (df["Proposed Start Date"] >= pd.to_datetime(start_date, dayfirst=True)) &
                (df["Proposed Start Date"] <= pd.to_datetime(end_date, dayfirst=True))
            ]
            filtered_df = detect_duplicates_by_id(filtered_df)
            st.write(f"🔎 {len(filtered_df)} students found in selected date range")
            styled = style_dates_and_duplicates(filtered_df)
            st.dataframe(styled, use_container_width=True)

        with tab2:
            visa_days = st.slider("📆 Visa expiring in next X days", 7, 180, 30)
            df_visa = visa_expiry_tracker(df, visa_days)
            st.write(f"🛂 {len(df_visa)} students with visa expiring in {visa_days} days")
            st.dataframe(df_visa.style.format({
                "Visa Expiry Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
            }), use_container_width=True)

        with tab3:
            coe_days = st.slider("📆 COE expiring in next X days", 7, 180, 30)
            df_coe = coe_expiry_tracker(df, coe_days)
            st.write(f"📄 {len(df_coe)} students with COE expiring in {coe_days} days")
            st.dataframe(df_coe.style.format({
                "Proposed End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
            }), use_container_width=True)

        with tab4:
            df_mismatch = course_duration_validator(df)
            st.write(f"⏳ {len(df_mismatch)} students with duration mismatch")
            st.dataframe(df_mismatch.style.format({
                "Proposed Start Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else "",
                "Proposed End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
            }), use_container_width=True)

        with tab5:
            weekly_counts = weekly_start_count(df)
            st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

        with tab6:
            df_agent = agent_summary(df)
            if not df_agent.empty:
                st.dataframe(df_agent, use_container_width=True)
            else:
                st.info("ℹ️ No agent column found in the uploaded file.")

        with tab7:
            contact_df = df[["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]].drop_duplicates()
            if "Is Duplicate" in df.columns:
                contact_df = pd.merge(contact_df
