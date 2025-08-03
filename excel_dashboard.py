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
    "Work Component Weeks","Work Component Total Hours","Course Language Duration In Weeks",
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
    # Filter by COE Status & Date range according to m
