import streamlit as st
import pandas as pd
import datetime
from io import BytesIO

# Streamlit config
st.set_page_config(page_title="COE Data Analyzer", layout="wide")

# Pandas option to allow large styling
pd.set_option("styler.render.max_elements", None)

st.title("📊 COE Data Analyzer")

# Upload section
uploaded_file = st.file_uploader("Upload Excel File", type=["xlsx"])

# Caching data load
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

# Duplicate detection
def detect_duplicates_by_id(data):
    duplicates = data.duplicated(subset=["PROVIDER STUDENT ID"], keep=False)
    data["IS_DUPLICATE"] = duplicates
    return data

# Style for duplicates
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

        available_statuses = df["COE STATUS"].dropna().unique().tolist()
        selected_statuses = st.multiselect("Select COE Status", options=available_statuses, default=available_statuses)

        st.subheader("📆 Filter by Proposed Start Date")
        min_date = df["Proposed Start Date"].min()
        max_date = df["Proposed Start Date"].max()

        start_date, end_date = st.date_input(
            "Select Proposed Start Date Range",
            [min_date, max_date],
            format="DD/MM/YYYY"
        )

        st.caption(f"Selected: {start_date.strftime('%d/%m/%Y')} to {end_date.strftime('%d/%m/%Y')}")

        # Filter
        filtered_df = df[
            (df["COE STATUS"].isin(selected_statuses)) &
            (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
            (df["Proposed Start Date"] <= pd.to_datetime(end_date))
        ]

        st.write(f"🔍 {len(filtered_df)} records found")

        # Optional Duplicate Highlight
        enable_highlight = st.checkbox("🟡 Highlight Duplicate Students by ID", value=True)

        if enable_highlight:
            filtered_df = detect_duplicates_by_id(filtered_df)
            styled_df = style_duplicates(filtered_df)
            st.dataframe(styled_df, use_container_width=True)
        else:
            st.dataframe(filtered_df, use_container_width=True)

        # Download Button
        output = BytesIO()
        filtered_df.to_excel(output, index=False)
        st.download_button("📥 Download Filtered Data", output.getvalue(), file_name="filtered_coe_data.xlsx")

else:
    st.info("📤 Please upload an Excel file to get started.")
