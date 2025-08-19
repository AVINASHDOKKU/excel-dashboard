import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.image(logo_url, width=200)

# Launch control
if 'mode' not in st.session_state:
    st.session_state.mode = None

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.mode = "analyzer"
with col2:
    if st.button("📅 Course Date Calculator"):
        st.session_state.mode = "calculator"

# 📅 Course Date Calculator
if st.session_state.mode == "calculator":
    st.subheader("📅 Course Date Calculator")

    st.markdown("### 1. Calculate Course End Date")
    start_date = st.date_input("Start Date", value=datetime.today(), format="DD/MM/YYYY")
    duration_weeks = st.number_input("Duration (weeks)", min_value=1, value=1)
    if st.button("Calculate End Date"):
        end_date = start_date + timedelta(weeks=duration_weeks)
        st.success(f"📅 End Date: {end_date.strftime('%d/%m/%Y')}")

    st.markdown("---")

    st.markdown("### 2. Calculate Weeks Between Two Dates")
    proposed_start = st.date_input("Proposed Start Date", value=datetime.today(), key="start", format="DD/MM/YYYY")
    proposed_end = st.date_input("Proposed End Date", value=datetime.today(), key="end", format="DD/MM/YYYY")
    if st.button("Calculate Weeks Between"):
        if proposed_end >= proposed_start:
            delta_days = (proposed_end - proposed_start).days
            weeks_between = delta_days // 7
            st.success(f"📆 Total Weeks: {weeks_between} weeks")
        else:
            st.error("End date must be after start date.")
import streamlit as st
import pandas as pd
from datetime import datetime, timedelta

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.image(logo_url, width=200)

# Launch control
if 'mode' not in st.session_state:
    st.session_state.mode = None

col1, col2 = st.columns([1, 1])
with col1:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.mode = "analyzer"
with col2:
    if st.button("📅 Course Date Calculator"):
        st.session_state.mode = "calculator"

# 📅 Course Date Calculator
if st.session_state.mode == "calculator":
    st.subheader("📅 Course Date Calculator")

    st.markdown("### 1. Calculate Course End Date")
    start_date = st.date_input("Start Date", value=datetime.today(), format="DD/MM/YYYY")
    duration_weeks = st.number_input("Duration (weeks)", min_value=1, value=1)
    if st.button("Calculate End Date"):
        end_date = start_date + timedelta(weeks=duration_weeks)
        st.success(f"📅 End Date: {end_date.strftime('%d/%m/%Y')}")

    st.markdown("---")

    st.markdown("### 2. Calculate Weeks Between Two Dates")
    proposed_start = st.date_input("Proposed Start Date", value=datetime.today(), key="start", format="DD/MM/YYYY")
    proposed_end = st.date_input("Proposed End Date", value=datetime.today(), key="end", format="DD/MM/YYYY")
    if st.button("Calculate Weeks Between"):
        if proposed_end >= proposed_start:
            delta_days = (proposed_end - proposed_start).days
            weeks_between = delta_days // 7
            st.success(f"📆 Total Weeks: {weeks_between} weeks")
        else:
            st.error("End date must be after start date.")
# Analyzer Tabs
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
                contact_df = df[df["COE STATUS"].isin(selected_statuses)][
                    ["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]
                ].drop_duplicates()
                if "Is Duplicate" in df.columns:
                    contact_df = pd.merge(contact_df, df[["Provider Student ID", "Is Duplicate"]], on="Provider Student ID", how="left")
                    contact_df["Duplicate Flag"] = contact_df["Is Duplicate"].apply(lambda x: "Yes" if x else "No")
                    contact_df.drop(columns=["Is Duplicate"], inplace=True)
                csv = contact_df.to_csv(index=False).encode('utf-8')
                st.download_button("📥 Download Contact Sheet CSV", csv, file_name="contact_sheet.csv", mime="text/csv")
