import streamlit as st
import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

# 🔐 Define login credentials
users = {
    "admin": "password123",
    "user1": "excel2025"
}

st.set_page_config(page_title="Excel Analyzer", layout="wide")

if "logged_in" not in st.session_state:
    st.session_state.logged_in = False

def login():
    st.title("🔐 Login")
    username = st.text_input("Username")
    password = st.text_input("Password", type="password")
    if st.button("Login"):
        if username in users and users[username] == password:
            st.session_state.logged_in = True
            st.success("✅ Login successful!")
        else:
            st.error("❌ Invalid credentials.")

if not st.session_state.logged_in:
    login()
    st.stop()

st.title("📊 Excel Data Analyzer")

uploaded_file = st.file_uploader("Upload your Excel file", type=["xlsx", "xls"])

if uploaded_file is not None:
    try:
        df = pd.read_excel(uploaded_file)
        st.success("✅ Excel file loaded successfully!")

        with st.expander("🔍 Preview Data", expanded=True):
            st.dataframe(df)

        if st.button("🔎 Analyze Data"):
            st.subheader("📋 Summary Statistics")
            st.dataframe(df.describe())

            st.subheader("🚫 Missing Values")
            st.dataframe(df.isnull().sum())

            numeric_cols = df.select_dtypes(include="number").columns.tolist()
            if numeric_cols:
                st.subheader("📊 Histogram")
                selected_col = st.selectbox("Choose a column", numeric_cols)
                fig, ax = plt.subplots()
                sns.histplot(df[selected_col].dropna(), kde=True, ax=ax)
                st.pyplot(fig)

                st.subheader("🧠 Correlation Heatmap")
                corr = df[numeric_cols].corr()
                fig2, ax2 = plt.subplots(figsize=(10, 6))
                sns.heatmap(corr, annot=True, cmap="coolwarm", ax=ax2)
                st.pyplot(fig2)
            else:
                st.warning("⚠️ No numeric columns found.")
    except Exception as e:
        st.error(f"❌ Error reading Excel file: {e}")
else:
    st.info("📤 Please upload an Excel file to get started.")
