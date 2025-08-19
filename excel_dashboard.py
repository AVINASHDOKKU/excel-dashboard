import streamlit as st
import pandas as pd
import datetime

# Page config
st.set_page_config(page_title="COE Student Analyzer", layout="wide")

# Logo and header
logo_url = "https://github.com/AVINASHDOKKU/excel-dashboard/blob/main/TEK4DAY.png?raw=true"
st.image(logo_url, width=300)  # Increased logo size
st.markdown("""
### COE Student Analyzer  
Powered by TEK4DAY  
""", unsafe_allow_html=True)

# Feature Tabs
feature_tab = st.tabs(["📊 Analyzer", "📅 Course Date Calculator"])
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
