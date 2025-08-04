import streamlit as st
import pandas as pd
from datetime import datetime
import numpy as np

EXPECTED_COLUMNS = [
    "Provider Code","COE Code","COE Status","COE Type","Principal CoE","Immigration Post",
    "Provider Student ID","Courtesy Title","First Name","Second Name","Family Name","Gender",
    "Date Of Birth","Country Of Birth","Nationality","Country Of Passport","Passport Number",
    "Email Address","Mobile","Phone","Student Address Line 1","Student Address Line 2",
    "Student Address Line 3","Student Address Line 4","Student Address Locality",
    "Student Address State","Student Address Country","Student Address Post Code",
    "ProviderArrangedHealthCover (OSHC)","OSHC Start Date","OSHC End Date","OSHC Provider",
    "English Test Exemption Reason","English Test Type","English Test Score","English Test Date",
    "Other Form Of Testing","Other Form Of Testing Comments","Provider Confirmed Study Commencement",
    "Course Code","Course Name","Course Sector","Is Dual Qualification",
    "Broad Field Of Education 1","Narrow Field Of Education 1","Detailed Field Of Education 1",
    "Broad Field Of Education 2","Narrow Field Of Education 2","Detailed Field Of Education 2",
    "Course Level","VET National Code","Foundation Studies","Work Component","Work Component Hrs/ Wk",
    "Work Component Weeks","Work Component Total Hours","Course Language", Duration In Weeks",
    "Current Total Course Fee","Proposed Start Date","Proposed End Date","Actual Start Date",
    "Actual End Date","Total Course Fee","Prepaid Course Fee","Visa Granted","Visa Grant Status",
    "Visa Grant Number","Visa Non  Grant Status","Visa Non Grant Action Date","Visa Effective Date",
    "Visa End Date","COE Created Date","Created By","COE Last Updated","Enrolment Comments",
    "Student Comments","Location Name","Active Location"
]

# COE Status categories
ACTIVE_STATUSES_TODAY_FUTURE = {"APPROVED", "STUDYING", "VISA GRANTED"}
EXCLUDE_STATUSES_TODAY_FUTURE = {"FINISHED", "CANCELLED", "SAVED"}

ACTIVE_STATUSES_PAST = {"APPROVED", "STUDYING", "VISA GRANTED", "CANCELLED", "FINISHED"}
EXCLUDE_STATUSES_PAST = {"SAVED"}

# Utility functions

def check_columns(df):
    missing_cols = [col for col in EXPECTED_COLUMNS if col not in df.columns]
    return missing_cols

def parse_dates(df, cols):
    for c in cols:
        df[c] = pd.to_datetime(df[c], errors='coerce')
    return df

def remove_duplicate_coes(df):
    # Each COE Code should be unique
    return df.drop_duplicates(subset=["COE Code"])

def flag_duplicate_students(df):
    # Flag students with same names + course code having overlapping COEs
    df = df.copy()
    df["Full Name"] = (df["First Name"].str.strip().fillna("") + " " +
                       df["Second Name"].str.strip().fillna("") + " " +
                       df["Family Name"].str.strip().fillna("")).str.upper()
    df["Course Code"] = df["Course Code"].fillna("").str.upper()
    
    # Sort by student and Proposed Start Date
    df = df.sort_values(by=["Full Name", "Course Code", "Proposed Start Date"])
    
    df["Duplicate Student"] = False
    
    # Group by student + course
    for (name, course), group in df.groupby(["Full Name", "Course Code"]):
        # Compare date ranges for overlap
        dates = group[["Proposed Start Date", "Proposed End Date"]].to_numpy()
        for i in range(len(dates)):
            for j in range(i+1, len(dates)):
                start_i, end_i = dates[i]
                start_j, end_j = dates[j]
                if pd.isnull(start_i) or pd.isnull(end_i) or pd.isnull(start_j) or pd.isnull(end_j):
                    continue
                latest_start = max(start_i, start_j)
                earliest_end = min(end_i, end_j)
                delta = (earliest_end - latest_start).days
                if delta >= 0:
                    # Overlapping COEs, mark duplicates except if one starts exactly after the other ends (allow extension)
                    # Here, we allow extensions if one starts the day after another ends
                    if (start_j - end_i).days > 1 and (start_i - end_j).days > 1:
                        # No overlap, do nothing
                        pass
                    else:
                        idx_i = group.index[i]
                        idx_j = group.index[j]
                        df.loc[idx_i, "Duplicate Student"] = True
                        df.loc[idx_j, "Duplicate Student"] = True
    return df

def filter_active_students(df, ref_date, duration_weeks, mode):
    # Filter by COE Status & Date range according to mode

    ref_date = pd.to_datetime(ref_date)
    duration_days = duration_weeks * 7
    date_from = ref_date
    date_to = ref_date + pd.Timedelta(days=duration_days)
    
    # Prepare status filters
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

    # Filter by COE Status inclusion and exclusion
    filtered = df[df["COE Status"].isin(status_incl) & ~df["COE Status"].isin(status_excl)]

    # Filter by date range overlap with Proposed Start Date and Proposed End Date
    cond = (
        (filtered["Proposed Start Date"] <= date_to) &
        (filtered["Proposed End Date"] >= date_from)
    )
    filtered = filtered[cond]

    return filtered

def display_active_students(df):
    st.write(f"Total active students found: {len(df)}")
    st.dataframe(df[[
        "COE Code","COE Status","First Name","Second Name","Family Name",
        "Course Code","Course Name","Course Language Duration In Weeks",
        "Proposed Start Date","Proposed End Date"
    ]])

st.title("Excel Data Analyzer with Multiple Analyzers")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    df = pd.read_excel(uploaded_file)
    missing_cols = check_columns(df)
    if missing_cols:
        st.error(f"Uploaded file is missing these required columns: {missing_cols}")
        st.stop()

    # Parse dates we need
    date_cols = ["Proposed Start Date", "Proposed End Date"]
    df = parse_dates(df, date_cols)
    
    df = remove_duplicate_coes(df)
    df = flag_duplicate_students(df)
    
    st.success("✅ File uploaded and verified!")

    # Show duplicate student flag count
    dup_count = df["Duplicate Student"].sum()
    if dup_count > 0:
        st.warning(f"⚠️ {dup_count} duplicate students with overlapping COEs detected.")

    # Choose analyzer
    analyzer = st.selectbox("Choose analyzer", ["Export for CoE and Student Details", "Active Students by Qualification"])

    if analyzer == "Export for CoE and Student Details":
        st.subheader("Export CoE and Student Details")
        st.dataframe(df[EXPECTED_COLUMNS])

    elif analyzer == "Active Students by Qualification":
        st.subheader("Active Students by Qualification")

        mode = st.radio("Select mode:", ["today", "past", "future"], index=0)

        ref_date = st.date_input("Reference Date (for filtering)", datetime.today())

        duration_weeks = st.number_input("Duration in weeks", min_value=1, max_value=520, value=52)

        if st.button("Run Analysis"):
            result_df = filter_active_students(df, ref_date, duration_weeks, mode)

            # Remove duplicate students flagged to avoid double counting
            result_df = result_df[~result_df["Duplicate Student"]]

            display_active_students(result_df)

            # Extra: Highlight students with same name + course code but multiple COEs
            grouped = result_df.groupby(
                ["First Name", "Second Name", "Family Name", "Course Code"]
            ).size().reset_index(name="Count")
            duplicates = grouped[grouped["Count"] > 1]

            if not duplicates.empty:
                st.info(f"⚠️ Found {len(duplicates)} students with multiple COEs (possible extensions).")
                st.dataframe(duplicates)

else:
    st.info("Please upload the Excel file with required fields to proceed.")
