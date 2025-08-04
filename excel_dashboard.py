import streamlit as st
import pandas as pd
import datetime

st.set_page_config(page_title="Excel Analyser", layout="wide")

st.title("📊 Excel Data Analyser")
st.markdown("Upload your Excel file and choose the type of analysis to perform.")

ANALYSER_OPTIONS = [
    "Export for CoE and Student Details"
]

REQUIRED_COLUMNS = [
    "Provider Code", "COE Code", "COE Status", "COE Type", "Principal CoE", "Immigration Post",
    "Provider Student ID", "Courtesy Title", "First Name", "Second Name", "Family Name", "Gender",
    "Date Of Birth", "Country Of Birth", "Nationality", "Country Of Passport", "Passport Number",
    "Email Address", "Mobile", "Phone", "Student Address Line 1", "Student Address Line 2",
    "Student Address Line 3", "Student Address Line 4", "Student Address Locality",
    "Student Address State", "Student Address Country", "Student Address Post Code",
    "ProviderArrangedHealthCover (OSHC)", "OSHC Start Date", "OSHC End Date", "OSHC Provider",
    "English Test Exemption Reason", "English Test Type", "English Test Score", "English Test Date",
    "Other Form Of Testing", "Other Form Of Testing Comments", "Provider Confirmed Study Commencement",
    "Course Code", "Course Name", "Course Sector", "Is Dual Qualification", "Broad Field Of Education 1",
    "Narrow Field Of Education 1", "Detailed Field Of Education 1", "Broad Field Of Education 2",
    "Narrow Field Of Education 2", "Detailed Field Of Education 2", "Course Level", "VET National Code",
    "Foundation Studies", "Work Component", "Work Component Hrs/ Wk", "Work Component Weeks",
    "Work Component Total Hours", "Course Language", "Duration In Weeks", "Current Total Course Fee",
    "Proposed Start Date", "Proposed End Date", "Actual Start Date", "Actual End Date", "Total Course Fee",
    "Prepaid Course Fee", "Visa Granted", "Visa Grant Status", "Visa Grant Number", "Visa Non  Grant Status",
    "Visa Non Grant Action Date", "Visa Effective Date", "Visa End Date", "COE Created Date", "Created By",
    "COE Last Updated", "Enrolment Comments", "Student Comments", "Location Name", "Active Location"
]

ACTIVE_STATUSES_TODAY_FUTURE = ["APPROVED", "STUDYING", "VISA GRANTED"]
EXCLUDE_STATUSES_TODAY_FUTURE = ["FINISHED", "CANCELLED", "SAVED"]
ACTIVE_STATUSES_PAST = ["APPROVED", "STUDYING", "VISA GRANTED", "CANCELLED", "FINISHED"]
EXCLUDE_STATUSES_PAST = ["SAVED"]

def load_excel(file):
    try:
        df = pd.read_excel(file)
        return df
    except Exception as e:
        st.error(f"Error reading file: {e}")
        return None

def validate_columns(df):
    missing = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing:
        st.error(f"Missing required columns: {missing}")
        return False
    return True

def filter_active_students(df, ref_date, duration_weeks, mode):
    ref_date = pd.to_datetime(ref_date)
    duration_days = duration_weeks * 7
    date_from = ref_date
    date_to = ref_date + pd.Timedelta(days=duration_days)

    if mode == "today":
        status_incl = ACTIVE_STATUSES_TODAY_FUTURE
        status_excl = EXCLUDE_STATUSES_TODAY_FUTURE
    elif mode == "past":
        status_incl = ACTIVE_STATUSES_PAST
        status_excl = EXCLUDE_STATUSES_PAST
    elif mode == "future":
        status_incl = ACTIVE_STATUSES_TODAY_FUTURE
        status_excl = EXCLUDE_STATUSES_TODAY_FUTURE
    else:
        st.error("Invalid mode selected")
        return pd.DataFrame()

    filtered = df[df["COE Status"].isin(status_incl) & ~df["COE Status"].isin(status_excl)]

    cond = (
        (filtered["Proposed Start Date"] <= date_to) &
        (filtered["Proposed End Date"] >= date_from)
    )
    filtered = filtered[cond]

    return filtered

uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

if uploaded_file:
    df = load_excel(uploaded_file)
    if df is not None and validate_columns(df):
        analyser = st.selectbox("Select Analysis Type", ANALYSER_OPTIONS)

        if analyser == "Export for CoE and Student Details":
            st.subheader("Active Students Analysis")

            col1, col2 = st.columns(2)
            with col1:
                input_date = st.date_input("Select Reference Date", datetime.date.today())
            with col2:
                input_weeks = st.number_input("Duration in Weeks", min_value=1, value=4)

            tab1, tab2, tab3 = st.tabs(["Today", "Past", "Future"])

            with tab1:
                result_today = filter_active_students(df, input_date, input_weeks, "today")
                st.write("### Active Students - Today")
                st.dataframe(result_today)
            with tab2:
                past_date = st.date_input("Select Past Date", value=datetime.date.today() - datetime.timedelta(days=90))
                result_past = filter_active_students(df, past_date, input_weeks, "past")
                st.write("### Active Students - Past")
                st.dataframe(result_past)
            with tab3:
                future_date = st.date_input("Select Future Date", value=datetime.date.today() + datetime.timedelta(days=30))
                result_future = filter_active_students(df, future_date, input_weeks, "future")
                st.write("### Active Students - Future")
                st.dataframe(result_future)

