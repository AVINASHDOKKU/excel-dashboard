# ==============================
# COE Student Analyzer - Part 1
# ==============================

import streamlit as st
import pandas as pd
import numpy as np
from datetime import datetime, timedelta

# ------------------------------
# Default App Settings (can be changed in Settings tab)
# ------------------------------
if "app_icon" not in st.session_state:
    st.session_state.app_icon = "🎓"
if "app_title" not in st.session_state:
    st.session_state.app_title = "COE Student Analyzer"

# Apply page config
st.set_page_config(
    page_title=st.session_state.app_title,
    page_icon=st.session_state.app_icon,
    layout="wide"
)

# ------------------------------
# File uploader
# ------------------------------
st.title(f"{st.session_state.app_icon} {st.session_state.app_title}")
uploaded_file = st.file_uploader("📂 Upload your Excel file", type=["xlsx"])

# ------------------------------
# Required Columns
# ------------------------------
REQUIRED_COLUMNS = [
    'COE STATUS', 'FIRST NAME', 'SECOND NAME', 'FAMILY NAME', 
    'STUDENT ID', 'PROVIDER STUDENT ID', 'COURSE NAME',
    'DURATION IN WEEKS', 'PROPOSED START DATE', 'PROPOSED END DATE'
]

def validate_columns(df):
    missing_cols = [col for col in REQUIRED_COLUMNS if col not in df.columns]
    if missing_cols:
        st.error(f"❌ Missing one or more required columns: {missing_cols}")
        return False
    return True
# ==============================
# COE Student Analyzer - Part 2
# ==============================

def load_and_prepare_data(file):
    """Load Excel, convert dates to dd/mm/yyyy, and validate."""
    try:
        df = pd.read_excel(file)
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        return None

    if not validate_columns(df):
        return None

    # Convert dates to datetime
    for col in ['PROPOSED START DATE', 'PROPOSED END DATE']:
        df[col] = pd.to_datetime(df[col], errors='coerce')

    return df


def style_duplicates(filtered_df):
    """Highlight duplicate PROVIDER STUDENT ID in filtered data only."""
    if 'PROVIDER STUDENT ID' not in filtered_df.columns:
        return filtered_df  # Fail-safe

    # Find duplicates in the filtered dataset only
    duplicates = filtered_df['PROVIDER STUDENT ID'].duplicated(keep=False)

    def highlight_row(row):
        if duplicates.loc[row.name]:
            return ['background-color: lightyellow'] * len(row)
        return [''] * len(row)

    return filtered_df.style.apply(highlight_row, axis=1)


# ------------------------------
# Main logic - Load file
# ------------------------------
if uploaded_file:
    df = load_and_prepare_data(uploaded_file)

    if df is not None:
        # Date filter (dd/mm/yyyy)
        st.subheader("📅 Filter by Proposed Start Date")
        min_date = df['PROPOSED START DATE'].min()
        max_date = df['PROPOSED START DATE'].max()

        start_date, end_date = st.date_input(
            "Select Proposed Start Date Range (dd/mm/yyyy)",
            value=[min_date, max_date],
            min_value=min_date,
            max_value=max_date,
            format="DD/MM/YYYY"
        )

        # Filter data
        mask = (df['PROPOSED START DATE'] >= pd.to_datetime(start_date)) & \
               (df['PROPOSED START DATE'] <= pd.to_datetime(end_date))
        filtered_df = df.loc[mask]

        st.success(f"🔍 {len(filtered_df)} records found between {start_date.strftime('%d/%m/%Y')} and {end_date.strftime('%d/%m/%Y')}")

        # Show styled dataframe (highlight duplicates in filtered data only)
        pd.set_option("styler.render.max_elements", len(filtered_df) * len(filtered_df.columns))
        st.dataframe(style_duplicates(filtered_df), use_container_width=True)
# ==============================
# COE Student Analyzer - Part 3
# ==============================
import datetime as dt

tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
    "📅 Date Filter",
    "🛂 Visa Expiry",
    "📄 COE Expiry",
    "📏 Duration Validator",
    "📆 Weekly Starts",
    "📇 Contact Sheet",
    "⚙ Settings"
])

# -------- TAB 1: Date Filter --------
with tab1:
    st.subheader("📅 Filter by Proposed Start Date")
    min_date = df['PROPOSED START DATE'].min()
    max_date = df['PROPOSED START DATE'].max()

    start_date, end_date = st.date_input(
        "Select Proposed Start Date Range",
        value=[min_date, max_date],
        min_value=min_date,
        max_value=max_date,
        format="DD/MM/YYYY"
    )
    mask = (df['PROPOSED START DATE'] >= pd.to_datetime(start_date)) & \
           (df['PROPOSED START DATE'] <= pd.to_datetime(end_date))
    filtered_df = df.loc[mask]

    pd.set_option("styler.render.max_elements", len(filtered_df) * len(filtered_df.columns))
    st.dataframe(style_duplicates(filtered_df), use_container_width=True)

# -------- TAB 2: Visa Expiry --------
with tab2:
    st.subheader("🛂 Visa Expiry Tracker")
    days_notice = st.slider("Show students whose visa expires within X days", 0, 365, 90)
    today = pd.to_datetime(dt.date.today())
    if 'VISA EXPIRY DATE' in df.columns:
        mask = (df['VISA EXPIRY DATE'] - today).dt.days <= days_notice
        visa_df = df.loc[mask]
        st.dataframe(visa_df, use_container_width=True)
    else:
        st.warning("⚠ No 'VISA EXPIRY DATE' column found.")

# -------- TAB 3: COE Expiry --------
with tab3:
    st.subheader("📄 COE Expiry Tracker")
    days_notice = st.slider("Show students whose COE expires within X days", 0, 365, 60)
    if 'COE END DATE' in df.columns:
        mask = (df['COE END DATE'] - today).dt.days <= days_notice
        coe_df = df.loc[mask]
        st.dataframe(coe_df, use_container_width=True)
    else:
        st.warning("⚠ No 'COE END DATE' column found.")

# -------- TAB 4: Duration Validator --------
with tab4:
    st.subheader("📏 Course Duration Validator")
    df['CALCULATED_DURATION'] = (df['PROPOSED END DATE'] - df['PROPOSED START DATE']).dt.days / 7
    duration_df = df[df['CALCULATED_DURATION'] != df['DURATION IN WEEKS']]
    st.dataframe(duration_df, use_container_width=True)

# -------- TAB 5: Weekly Starts --------
with tab5:
    st.subheader("📆 Weekly Start Count")
    weekly_count = df.groupby(df['PROPOSED START DATE'].dt.to_period('W')).size().reset_index(name='Count')
    weekly_count['PROPOSED START DATE'] = weekly_count['PROPOSED START DATE'].astype(str)
    st.bar_chart(weekly_count.set_index('PROPOSED START DATE')['Count'])

# -------- TAB 6: Contact Sheet --------
with tab6:
    st.subheader("📇 Student Contact Sheet")
    if st.button("📥 Download Contact Sheet as CSV"):
        contact_df = df[['FIRST NAME', 'SECOND NAME', 'FAMILY NAME', 'STUDENT ID', 'PROVIDER STUDENT ID']]
        csv = contact_df.to_csv(index=False).encode('utf-8')
        st.download_button(label="Download CSV", data=csv, file_name="contact_sheet.csv", mime="text/csv")

# -------- TAB 7: Settings --------
with tab7:
    st.subheader("⚙ App Settings")
    app_icon = st.text_input("Change App Icon (Emoji)", value="🎓")
    app_title = st.text_input("Change App Title", value="COE Student Analyzer")
    if st.button("Apply Settings"):
        st.session_state['app_icon'] = app_icon
        st.session_state['app_title'] = app_title
        st.success("✅ Settings saved! Please reload the page to apply changes.")
# ==============================
# COE Student Analyzer - Part 4
# ==============================
import streamlit as st
import pandas as pd
import datetime as dt

# Initialize Session State Defaults
if 'app_icon' not in st.session_state:
    st.session_state['app_icon'] = "🎓"
if 'app_title' not in st.session_state:
    st.session_state['app_title'] = "COE Student Analyzer"

# UI Header
st.markdown(
    f"""
    <div style="text-align: center; padding: 10px;">
        <span style="font-size: 50px;">{st.session_state['app_icon']}</span>
        <h1 style="margin-bottom: 0;">{st.session_state['app_title']}</h1>
        <p style="color: grey;">Analyze and track COE student data effortlessly</p>
    </div>
    """,
    unsafe_allow_html=True
)

# File Uploader
uploaded_file = st.file_uploader("📂 Upload COE Student Data (Excel)", type=["xlsx"])

if uploaded_file:
    try:
        df = pd.read_excel(uploaded_file, dayfirst=True)

        # Ensure required columns exist
        required_columns = [
            'COE STATUS', 'FIRST NAME', 'SECOND NAME', 'FAMILY NAME',
            'STUDENT ID', 'PROVIDER STUDENT ID', 'COURSE NAME',
            'DURATION IN WEEKS', 'PROPOSED START DATE', 'PROPOSED END DATE'
        ]
        missing = [col for col in required_columns if col not in df.columns]
        if missing:
            st.error(f"❌ Error: Missing one or more required columns: {missing}")
            st.stop()

        # Convert date columns
        date_columns = ['PROPOSED START DATE', 'PROPOSED END DATE', 'VISA EXPIRY DATE', 'COE END DATE']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce', dayfirst=True)

        # Style function for duplicates
        def style_duplicates(dataframe):
            duplicate_mask = dataframe.duplicated(subset=['STUDENT ID'], keep=False)
            return dataframe.style.apply(
                lambda x: ['background-color: yellow' if duplicate_mask.iloc[i] else '' for i in range(len(x))],
                axis=0
            )

        # 🎯 Launcher Button
        if st.button("🚀 Launch COE Analyzer"):
            st.success("✅ Analyzer launched! Use the tabs below to explore data.")

    except Exception as e:
        st.error(f"❌ Failed to process file: {str(e)}")
else:
    st.info("📌 Please upload an Excel file to get started.")
