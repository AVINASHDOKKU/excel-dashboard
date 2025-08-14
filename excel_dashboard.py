import streamlit as st
import pandas as pd
import datetime
from io import BytesIO
import fitz  # PyMuPDF

st.set_page_config(page_title="COE Student Analyzer", layout="wide")
st.markdown("""
<div style='text-align: center'>
    <h2>COE Student Analyzer</h2>
    <p><em>Powered by TEK4DAY</em></p>
</div>
""", unsafe_allow_html=True)

if 'launch' not in st.session_state:
    st.session_state.launch = False

if not st.session_state.launch:
    if st.button("🚀 Launch Analyzer"):
        st.session_state.launch = True
    st.stop()

uploaded_file = st.file_uploader("📤 Upload your Excel file", type=["xlsx"])
columns_to_hide = [ "IMMIGRATION POST", "COURTESY TITLE", "COUNTRY OF BIRTH", "NATIONALITY", "COUNTRY OF PASSPORT",
    "EMAIL ADDRESS", "MOBILE", "PHONE", "STUDENT ADDRESS LINE 1", "STUDENT ADDRESS LINE 2", "STUDENT ADDRESS LINE 3",
    "STUDENT ADDRESS LINE 4", "STUDENT ADDRESS LOCALITY", "STUDENT ADDRESS STATE", "STUDENT ADDRESS COUNTRY",
    "STUDENT ADDRESS POST CODE", "PROVIDERARRANGEDHEALTHCOVER (OSHC)", "OSHC START DATE", "OSHC END DATE",
    "OSHC PROVIDER", "ENGLISH TEST EXEMPTION REASON", "ENGLISH TEST TYPE", "ENGLISH TEST SCORE", "ENGLISH TEST DATE",
    "OTHER FORM OF TESTING", "OTHER FORM OF TESTING COMMENTS", "PROVIDER CONFIRMED STUDY COMMENCEMENT",
    "COURSE SECTOR", "IS DUAL QUALIFICATION", "BROAD FIELD OF EDUCATION 1", "NARROW FIELD OF EDUCATION 1",
    "DETAILED FIELD OF EDUCATION 1", "BROAD FIELD OF EDUCATION 2", "NARROW FIELD OF EDUCATION 2",
    "DETAILED FIELD OF EDUCATION 2", "COURSE LEVEL", "VET NATIONAL CODE", "FOUNDATION STUDIES", "WORK COMPONENT",
    "WORK COMPONENT HRS/ WK", "WORK COMPONENT WEEKS", "WORK COMPONENT TOTAL HOURS", "COURSE LANGUAGE",
    "DURATION IN WEEKS", "CURRENT TOTAL COURSE FEE", "ACTUAL END DATE", "TOTAL COURSE FEE", "PREPAID COURSE FEE",
    "COE CREATED DATE", "CREATED BY", "COE LAST UPDATED", "ENROLMENT COMMENTS", "STUDENT COMMENTS",
    "LOCATION NAME", "ACTIVE LOCATION"
]

def preprocess_data(df):
    df.columns = df.columns.str.strip()
    for col in ["Proposed Start Date", "Proposed End Date", "Visa End Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    df.dropna(subset=["Proposed Start Date", "Proposed End Date"], inplace=True)
    return df

def visa_expiry_tracker(df, days=30):
    today = pd.to_datetime(datetime.date.today())
    future_limit = today + pd.to_timedelta(days, unit="d")
    return df[(df["Visa End Date"] >= today) & (df["Visa End Date"] <= future_limit)]

def coe_expiry_tracker(df, within_days=30):
    today = pd.to_datetime(datetime.date.today())
    future_limit = today + pd.to_timedelta(within_days, unit="d")
    return df[(df["Proposed End Date"] >= today) & (df["Proposed End Date"] <= future_limit)]

def course_duration_validator(df):
    if "Proposed Start Date" in df.columns and "Proposed End Date" in df.columns and "Duration In Weeks" in df.columns:
        df["Actual Weeks"] = (df["Proposed End Date"] - df["Proposed Start Date"]).dt.days // 7
        df["Duration Mismatch"] = df["Actual Weeks"] != df["Duration In Weeks"]
        return df[df["Duration Mismatch"]]
    return pd.DataFrame()

def weekly_start_count(df):
    if "Proposed Start Date" in df.columns:
        df["Start Week"] = df["Proposed Start Date"].dt.isocalendar().week
        return df.groupby("Start Week")["Provider Student ID"].count().reset_index(name="Number of Starts")
    return pd.DataFrame()

def agent_summary(df):
    if "AGENT" in df.columns:
        return df.groupby("AGENT").agg(
            Total_Students=("Provider Student ID", "count"),
            Active_Students=("COE STATUS", lambda x: (x == "Active").sum())
        ).reset_index()
    return pd.DataFrame()

def generate_pdf(df):
    doc = fitz.open()
    for _, row in df.iterrows():
        text = "\n".join([f"{col}: {row[col]}" for col in df.columns])
        page = doc.new_page()
        page.insert_text((72, 72), text, fontsize=10)
    pdf_bytes = doc.write()
    return pdf_bytes
if uploaded_file:
    try:
        df_raw = pd.read_excel(uploaded_file, engine="openpyxl")
        df = preprocess_data(df_raw)

        tab1, tab2, tab3, tab4, tab5, tab6, tab7 = st.tabs([
            "📅 Start Date Filter", "🛂 Visa Expiry", "📄 COE Expiry",
            "⏳ Duration Check", "📈 Weekly Starts", "🤝 Agent Summary", "📥 Download Contacts"
        ])

        with tab1:
            statuses = df["COE STATUS"].dropna().unique().tolist()
            selected_statuses = st.multiselect("🎯 Select COE Status", statuses, default=statuses)
            min_date = df["Proposed Start Date"].min()
            max_date = df["Proposed Start Date"].max()
            start_date, end_date = st.date_input("📆 Start Date Range", [min_date, max_date])
            filtered_df = df[
                (df["COE STATUS"].isin(selected_statuses)) &
                (df["Proposed Start Date"] >= pd.to_datetime(start_date)) &
                (df["Proposed Start Date"] <= pd.to_datetime(end_date))
            ]
            st.dataframe(filtered_df.drop(columns=[col for col in columns_to_hide if col in filtered_df.columns]), use_container_width=True)

        with tab2:
            visa_days = st.slider("📆 Visa expiring in next X days", 7, 180, 30)
            df_visa_expiring = visa_expiry_tracker(df, visa_days)
            display_df = df_visa_expiring.drop(columns=[col for col in columns_to_hide if col in df_visa_expiring.columns])
            st.dataframe(display_df, use_container_width=True)

            csv = display_df.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Download Visa Expiry CSV", csv, file_name="visa_expiry_filtered.csv", mime="text/csv")

            pdf_bytes = generate_pdf(display_df)
            st.download_button("📄 Download Visa Expiry PDF", pdf_bytes, file_name="visa_expiry_filtered.pdf", mime="application/pdf")

            st.subheader("❌ Visa Refused Students")
            if "Visa Non Grant Status" in df.columns:
                df_visa_refused = df[df["Visa Non Grant Status"].str.lower() == "refused"]
                st.dataframe(df_visa_refused.drop(columns=[col for col in columns_to_hide if col in df_visa_refused.columns]), use_container_width=True)

        with tab3:
            df_coe = coe_expiry_tracker(df)
            st.dataframe(df_coe.drop(columns=[col for col in columns_to_hide if col in df_coe.columns]), use_container_width=True)

        with tab4:
            df_mismatch = course_duration_validator(df)
            st.dataframe(df_mismatch.drop(columns=[col for col in columns_to_hide if col in df_mismatch.columns]), use_container_width=True)

        with tab5:
            weekly_counts = weekly_start_count(df)
            st.bar_chart(weekly_counts, x="Start Week", y="Number of Starts")

        with tab6:
            df_agent = agent_summary(df)
            st.dataframe(df_agent, use_container_width=True)

        with tab7:
            contact_df = df[["Provider Student ID", "FIRST NAME", "SECOND NAME", "FAMILY NAME"]].drop_duplicates()
            csv = contact_df.to_csv(index=False).encode("utf-8")
            st.download_button("📥 Download Contact Sheet CSV", csv, file_name="contact_sheet.csv", mime="text/csv")

    except Exception as e:
        st.error(f"❌ Error: {e}")
else:
    st.info("📤 Upload an Excel file to begin analysis.")
