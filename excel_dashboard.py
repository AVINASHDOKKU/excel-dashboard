import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# Setup
st.set_page_config(page_title="COE Data Analyzer", layout="wide")
pd.set_option("styler.render.max_elements", None)

st.title("📊 COE Data Analyzer")

# File uploader
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Helper functions
@st.cache_data
def load_data(file):
    try:
        df = pd.read_excel(file)
        df.columns = df.columns.str.strip().str.upper()
        df = df.rename(columns={
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
            "PROVIDER STUDENT ID": "PROVIDER STUDENT ID"
        })
        df["Proposed Start Date"] = pd.to_datetime(df["Proposed Start Date"], errors='coerce')
        df["Proposed End Date"] = pd.to_datetime(df["Proposed End Date"], errors='coerce')
        df["DURATION IN WEEKS"] = pd.to_numeric(df["DURATION IN WEEKS"], errors='coerce')
        return df
    except Exception as e:
        st.error(f"❌ Error reading file: {e}")
        return pd.DataFrame()

def detect_duplicates_by_id(data):
    duplicates = data.duplicated(subset=["PROVIDER STUDENT ID"], keep=False)
    data["IS_DUPLICATE"] = duplicates
    return data

def style_duplicates(data):
    def highlight(row):
        if row.get("IS_DUPLICATE"):
            return ['background-color: gold'] * len(row)
        return [''] * len(row)
    return data.style.apply(highlight, axis=1)

# Main logic
if uploaded_file:
    df = load_data(uploaded_file)

    if not df.empty:
        st.subheader("📁 Data Preview")
        st.dataframe(df.head(20), use_container_width=True)

        st.subheader("⚙️ Filter Settings")
        
        # Dynamic COE Status list
        available_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status", options=available_statuses, default=available_statuses)

        # Date range filter for Proposed Start Date
        st.subheader("📆 Students by Start Date Range")
        min_date = df["Proposed Start Date"].min()
        max_date = df["Proposed Start Date"].max()

        start_date, end_date = st.date_input(
            "Select Proposed Start Date Range",
            [min_date, max_date],
            format="DD/MM/YYYY"
        )

        st.caption(f"Selected: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")

        # Filtered DataFrame
        filtered_df = df[
            (df["COE STATUS"].isin(selected_statuses)) &
            (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
            (df["Proposed Start Date"] <= pd.to_datetime(end_date))
        ]

        filtered_df = detect_duplicates_by_id(filtered_df)

        st.write(f"🔍 {len(filtered_df)} records found")
        st.dataframe(style_duplicates(filtered_df), use_container_width=True)

        # Download filtered data
        output = BytesIO()
        filtered_df.to_excel(output, index=False)
        st.download_button("📥 Download Filtered Data", output.getvalue(), file_name="filtered_coe_data.xlsx")

else:
    st.info("📤 Please upload a valid Excel file to begin.")
