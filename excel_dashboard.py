import streamlit as st
import pandas as pd
import datetime

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.markdown("""
### COE Student Analyzer  
Powered by TEK4DAY  
""", unsafe_allow_html=True)

# Feature Tabs
feature_tab = st.tabs(["📊 Analyzer", "📅 Course Date Calculator"])

# -------------------- 📊 ANALYZER --------------------
with feature_tab[0]:
    uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])

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
            "Visa End Date": lambda x: x.strftime('%d/%m/%Y') if pd.notnull(x) else ""
        })

    def visa_expiry_tracker(df, days=30):
        if "Visa End Date" not in df.columns:
            return pd.DataFrame()
        today = pd.to_datetime(datetime.date.today())
        future_limit = today + pd.to_timedelta(days, unit="d")
        return df[(df["Visa End Date"] >= today) & (df["Visa End Date"] <= future_limit)]

    def coe_expiry_tracker(df, within_days=30):
        future_limit = pd.to_datetime(datetime.date.today()) + pd.to_timedelta(within_days, unit="d")
        return df[(df["Proposed End Date"] >= pd.to_datetime(datetime.date.today())) &
                  (df["Proposed End Date"] <= future_limit)]

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

    if uploaded_file:
        try:
            df_raw = pd.read_excel(uploaded_file)
            df = preprocess_data(df_raw)

            tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
                "📅 Start Date Filter", "🛂 Visa Expiry", "📄 COE Expiry",
                "⏳ Duration Check", "📈 Weekly Starts", "🤝 Agent Summary", "📥 Download Contacts"
            ])

            with tab1:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab1")
                min_date = df["Proposed Start Date"].min()
                max_date = df["Proposed Start Date"].max()
                start_date, end_date = st.date_input("📆 Select Proposed Start Date Range", [min_date, max_date])
                filtered_df = df[
                    (df["COE STATUS"].isin(selected_statuses)) &
                    (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
                    (df["Proposed Start Date"] <= pd.to_datetime(end_date))
                ]
                filtered_df = detect_duplicates_by_id(filtered_df)
                st.write(f"🔍 {len(filtered_df)} students found in selected date range")
                styled = style_dates_and_duplicates(filtered_df)
                st.dataframe(styled, use_container_width=True)

            with tab2:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab2")
                df_filtered = df[df["COE STATUS"].isin(selected_statuses)]
                st.subheader("📅 Visa Expiring in Next X Days")
                visa_days = st.slider("📆 Visa expiring in next X days", 7, 180, 30)
                df_visa_expiring = visa_expiry_tracker(df_filtered, visa_days)
                st.write(f"🛂 {len(df_visa_expiring)} students with visa expiring in {visa_days} days")
                st.dataframe(df_visa_expiring, use_container_width=True)

                st.subheader("❌ Visa Refused Students")
                if "Visa Non Grant Status" in df_filtered.columns:
                    df_visa_refused = df_filtered[df_filtered["Visa Non Grant Status"].astype(str).str.lower() == "refused"]
                    st.write(f"❌ {len(df_visa_refused)} students with refused visa status")
                    st.dataframe(df_visa_refused, use_container_width=True)
                else:
                    st.info("ℹ️ 'Visa Non Grant Status' column not found in the uploaded file.")

            with tab3:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab3")
                df_coe = coe_expiry_tracker(df[df["COE STATUS"].isin(selected_statuses)])
                st.write(f"📄 {len(df_coe)} students with COE expiring in next 30 days")
                st.dataframe(df_coe, use_container_width=True)

            with tab4:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab4")
                df_mismatch = course_duration_validator(df[df["COE STATUS"].isin(selected_statuses)])
                st.write(f"⏳ {len(df_mismatch)} students with duration mismatch")
                st.dataframe(df_mismatch, use_container_width=True)

            with tab5:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab5")
                weekly_counts = weekly_start_count(df[df["COE STATUS"].isin(selected_statuses)])
                st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

            with tab6:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab6")
                df_agent = agent_summary(df[df["COE STATUS"].isin(selected_statuses)])
                if not df_agent.empty:
                    st.dataframe(df_agent, use_container_width=True)
                else:
                    st.info("ℹ️ No agent column found in the uploaded file.")

            with tab7:
                statuses = df["COE STATUS"].dropna().unique().tolist()
                selected_statuses = st.multiselect("🎯 Select COE Status to include", statuses, default=statuses, key="status_tab7")
                contact_df = df[df["COE STATUS"].isin(selected_statuses)][["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]].drop_duplicates()
                if "Is Duplicate" in df.columns:
                    contact_df = pd.merge(contact_df, df[["Provider Student ID", "Is Duplicate"]], on="Provider Student ID", how="left")
                    contact_df["Duplicate Flag"] = contact_df["Is Duplicate"].apply(lambda x: "Yes" if x else "No")
                    contact_df.drop(columns=["Is Duplicate"], inplace=True)
                csv = contact_df.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Download Contact Sheet CSV", csv, file_name="contact_sheet.csv", mime="text/csv")

        except Exception as e:
            st.error(f"❌ Error: {e}")
    else:
        st.info("📤 Upload an Excel file to begin")

# -------------------- 📅 COURSE DATE CALCULATOR --------------------
with feature_tab[1]:
    st.subheader("📅 Course Date Calculator")

    # Forward Calculator
    st.markdown("### 📘 Forward Calculator")
    start_date = st.date_input("Enter Course Start Date (dd/mm/yyyy)", key="forward_start")
    duration_weeks = st.number_input("Enter Duration (in weeks)", min_value=1, step=1, key="forward_duration")

    if start_date and duration_weeks:
        finish_date = start_date + datetime.timedelta(weeks=duration_weeks)
        st.success(f"🎯 Course Finish Date: {finish_date.strftime('%d/%m/%Y')}")

    # Reverse Calculator
    st.markdown("### 🔁 Reverse Calculator")
    proposed_start = st.date_input("Enter Proposed Start Date (dd/mm/yyyy)", key="reverse_start")
    proposed_end = st.date_input("Enter Proposed End Date (dd/mm/yyyy)", key="reverse_end")

    if proposed_start and proposed_end and proposed_end > proposed_start:
        total_weeks = (proposed_end - proposed_start).days // 7
        st.success(f"📆 Total Number of Weeks: {total_weeks}")
