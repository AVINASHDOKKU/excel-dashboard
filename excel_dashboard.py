import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

# --- Constants ---

EXPECTED_COLUMNS = [
    "Provider Code", "COE Code", "COE Status", "COE Type", "Principal CoE",
    "Immigration Post", "Provider Student ID", "Courtesy Title", "First Name",
    "Second Name", "Family Name", "Gender", "Date Of Birth", "Country Of Birth",
    "Nationality", "Country Of Passport", "Passport Number", "Email Address",
    "Mobile", "Phone", "Student Address Line 1", "Student Address Line 2",
    "Student Address Line 3", "Student Address Line 4", "Student Address Locality",
    "Student Address State", "Student Address Country", "Student Address Post Code",
    "ProviderArrangedHealthCover (OSHC)", "OSHC Start Date", "OSHC End Date",
    "OSHC Provider", "English Test Exemption Reason", "English Test Type",
    "English Test Score", "English Test Date", "Other Form Of Testing",
    "Other Form Of Testing Comments", "Provider Confirmed Study Commencement",
    "Course Code", "Course Name", "Course Sector", "Is Dual Qualification",
    "Broad Field Of Education 1", "Narrow Field Of Education 1",
    "Detailed Field Of Education 1", "Broad Field Of Education 2",
    "Narrow Field Of Education 2", "Detailed Field Of Education 2", "Course Level",
    "VET National Code", "Foundation Studies", "Work Component",
    "Work Component Hrs/ Wk", "Work Component Weeks", "Work Component Total Hours",
    "Course Language", "Duration In Weeks", "Current Total Course Fee",
    "Proposed Start Date", "Proposed End Date", "Actual Start Date",
    "Actual End Date", "Total Course Fee", "Prepaid Course Fee", "Visa Granted",
    "Visa Grant Status", "Visa Grant Number", "Visa Non  Grant Status",
    "Visa Non Grant Action Date", "Visa Effective Date", "Visa End Date",
    "COE Created Date", "Created By", "COE Last Updated", "Enrolment Comments",
    "Student Comments", "Location Name", "Active Location"
]

# Relevant fields for analysis
ANALYSIS_FIELDS = [
    "COE Code", "COE Status", "First Name", "Second Name", "Family Name",
    "Course Code", "Course Name", "Duration In Weeks", "Proposed Start Date", "Proposed End Date"
]

ACTIVE_STATUSES_TODAY = {"APPROVED", "STUDYING", "VISA GRANTED"}
EXCLUDED_STATUSES_TODAY = {"FINISHED", "CANCELLED", "SAVED"}

ACTIVE_STATUSES_PAST = {"APPROVED", "STUDYING", "VISA GRANTED", "CANCELLED", "FINISHED"}
EXCLUDED_STATUSES_PAST = {"SAVED"}

ACTIVE_STATUSES_FUTURE = {"APPROVED", "STUDYING", "VISA GRANTED"}
EXCLUDED_STATUSES_FUTURE = {"FINISHED", "CANCELLED", "SAVED"}


# --- Helper Functions ---

def validate_columns(df):
    missing = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    if missing:
        return False, missing
    return True, None


def parse_dates(df, date_cols):
    for col in date_cols:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors='coerce')
    return df


def filter_active_students(df, date, duration_weeks, mode):
    """
    mode: 'today', 'past', 'future'
    Filters active students by COE Status and date range.
    """
    # Filter based on COE Status
    if mode == 'today':
        status_include = ACTIVE_STATUSES_TODAY
        status_exclude = EXCLUDED_STATUSES_TODAY
    elif mode == 'past':
        status_include = ACTIVE_STATUSES_PAST
        status_exclude = EXCLUDED_STATUSES_PAST
    elif mode == 'future':
        status_include = ACTIVE_STATUSES_FUTURE
        status_exclude = EXCLUDED_STATUSES_FUTURE
    else:
        raise ValueError("Invalid mode")

    filtered = df[
        df["COE Status"].isin(status_include) &
        ~df["COE Status"].isin(status_exclude)
    ].copy()

    # Filter based on date range (considering Proposed Start/End Date)
    start_limit = date - pd.Timedelta(weeks=duration_weeks)
    end_limit = date + pd.Timedelta(weeks=duration_weeks)

    filtered = filtered[
        (filtered["Proposed Start Date"] <= end_limit) &
        (filtered["Proposed End Date"] >= start_limit)
    ]

    return filtered


def highlight_duplicates(df):
    """
    Find duplicates by COE Code, and also by Student Name + Course Code.
    Handle extensions by checking overlapping dates.
    Returns a DataFrame with an added column 'Is Duplicate' (True/False)
    """
    df = df.copy()
    df['Is Duplicate'] = False

    # Mark exact duplicate COE Codes
    dup_coe = df[df.duplicated(subset=["COE Code"], keep=False)]
    df.loc[dup_coe.index, 'Is Duplicate'] = True

    # Group by student name + course code
    name_course_groups = df.groupby(["First Name", "Second Name", "Family Name", "Course Code"])

    for _, group in name_course_groups:
        if len(group) <= 1:
            continue
        # Sort by Proposed Start Date
        group = group.sort_values("Proposed Start Date")

        # Check overlapping or gaps that indicate extensions
        # We'll consider COEs overlapping or back-to-back as one "active"
        # So mark as duplicates only if they overlap without gap or if multiple COEs
        prev_end = None
        for idx, row in group.iterrows():
            start = row["Proposed Start Date"]
            end = row["Proposed End Date"]
            if prev_end is not None and start <= prev_end:
                # Overlapping or continuous -> mark duplicate
                df.at[idx, 'Is Duplicate'] = True
            prev_end = end

    return df


# --- Streamlit App ---

def main():
    st.title("📊 Student COE Data Analyzer")

    uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])
    if not uploaded_file:
        st.info("Please upload an Excel file with required fields.")
        return

    df = pd.read_excel(uploaded_file)
    st.write(f"File loaded with {len(df)} rows and {len(df.columns)} columns.")

    # Validate columns
    valid, missing = validate_columns(df)
    if not valid:
        st.error(f"Uploaded file is missing required columns: {missing}")
        return

    # Parse relevant date columns
    df = parse_dates(df, ["Proposed Start Date", "Proposed End Date"])

    # Show a checkbox to show raw data
    if st.checkbox("Show Raw Data Preview"):
        st.dataframe(df.head(50))

    # Input filters
    st.sidebar.header("Filters")

    analysis_type = st.sidebar.selectbox("Select Analysis Type", [
        "Active Students by Qualification",
        "Export for CoE and Student Details"
    ])

    # Common inputs for active student analysis
    if analysis_type == "Active Students by Qualification":
        # Input for date and duration weeks
        input_date = st.sidebar.date_input("Select reference date", datetime.today())
        input_duration = st.sidebar.number_input("Duration in weeks (+/-)", min_value=1, max_value=520, value=26)

        st.header("Active Students Analysis")

        # Buttons for 3 modes
        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("Active Students - Today"):
                filtered = filter_active_students(df, pd.Timestamp(input_date), input_duration, "today")
                filtered = highlight_duplicates(filtered)
                st.write(f"Active students as of {input_date} (Status in APPROVED, STUDYING, VISA GRANTED):")
                st.dataframe(filtered[ANALYSIS_FIELDS + ["Is Duplicate"]])
                st.write(f"Count (unique COE Code): {filtered['COE Code'].nunique()}")

        with col2:
            if st.button("Active Students - Past"):
                filtered = filter_active_students(df, pd.Timestamp(input_date), input_duration, "past")
                filtered = highlight_duplicates(filtered)
                st.write(f"Active students in past relative to {input_date} (Status including CANCELLED, FINISHED):")
                st.dataframe(filtered[ANALYSIS_FIELDS + ["Is Duplicate"]])
                st.write(f"Count (unique COE Code): {filtered['COE Code'].nunique()}")

        with col3:
            if st.button("Active Students - Future"):
                filtered = filter_active_students(df, pd.Timestamp(input_date), input_duration, "future")
                filtered = highlight_duplicates(filtered)
                st.write(f"Active students in future relative to {input_date} (Status in APPROVED, STUDYING, VISA GRANTED):")
                st.dataframe(filtered[ANALYSIS_FIELDS + ["Is Duplicate"]])
                st.write(f"Count (unique COE Code): {filtered['COE Code'].nunique()}")

    elif analysis_type == "Export for CoE and Student Details":
        st.header("Export for CoE and Student Details")

        # For simplicity: just export relevant columns
        export_fields = EXPECTED_COLUMNS  # all columns as per your list

        if st.button("Show Export Data Preview"):
            st.dataframe(df[export_fields].head(50))

        if st.button("Download Export CSV"):
            csv_data = df[export_fields].to_csv(index=False)
            st.download_button(label="Download CSV", data=csv_data, file_name="coe_student_export.csv", mime="text/csv")

if __name__ == "__main__":
    main()
