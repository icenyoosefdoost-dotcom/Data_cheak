# app.py
import streamlit as st
import pandas as pd
import os

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="Data Cleaning App",
    page_icon="🧹",
    layout="wide"
)

# ---------------------------
# Sidebar
# ---------------------------
st.sidebar.title("Navigation")
page = st.sidebar.radio("Go to:", ["Home", "Upload Data", "Cleaning / Processing", "Analysis / Output"])

# ---------------------------
# Home
# ---------------------------
if page == "Home":
    st.title("🧹 Data Cleaning & Processing App")
    st.write("""
    Welcome!  
    This application runs your combined code in an easy-to-use Streamlit interface.  
    Use the sidebar to navigate between steps:
    - **Upload Data**: Load your Excel/CSV file  
    - **Cleaning / Processing**: Run your custom cleaning code  
    - **Analysis / Output**: View and download results  
    """)

# ---------------------------
# Upload Data
# ---------------------------
elif page == "Upload Data":
    st.title("📤 Upload Data File")
    uploaded_file = st.file_uploader("Choose a CSV or Excel file", type=["csv", "xlsx"])

    if uploaded_file:
        if uploaded_file.name.endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)

        st.session_state["data"] = df
        st.success("File uploaded successfully!")
        st.dataframe(df.head())

# ---------------------------
# Cleaning / Processing
# ---------------------------
elif page == "Cleaning / Processing":
    st.title("🧹 Data Cleaning & Processing")

    if "data" not in st.session_state:
        st.warning("Please upload a file first in the 'Upload Data' section.")
    else:
        df = st.session_state["data"]

        st.write("### Original Data Preview")
        st.dataframe(df.head())

        # --- HERE: paste your merged_code.py logic ---
        st.info("Your cleaning code will run here (integrated from merged_code.py).")

        # Example placeholder: (replace with your code logic)
        df_clean = df.dropna()
        st.session_state["cleaned"] = df_clean

        st.success("Cleaning completed!")
        st.write("### Cleaned Data Preview")
        st.dataframe(df_clean.head())

# ---------------------------
# Analysis / Output
# ---------------------------
elif page == "Analysis / Output":
    st.title("📊 Analysis & Output")

    if "cleaned" not in st.session_state:
        st.warning("Please clean data first in the 'Cleaning / Processing' section.")
    else:
        df_clean = st.session_state["cleaned"]
        st.write("### Final Cleaned Data")
        st.dataframe(df_clean.head())

        # Download option
        csv = df_clean.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="💾 Download Cleaned Data (CSV)",
            data=csv,
            file_name="cleaned_data.csv",
            mime="text/csv",
        )
