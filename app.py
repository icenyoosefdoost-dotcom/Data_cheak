import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from docx import Document
from datetime import datetime
import io

# ============================
# Helper functions (from your code)
# ============================

def normalize(name):
    return re.sub(r'[^a-z0-9]', '', name.lower())

standard_names = [
    "Dissolved Oxygen (DO)", "pH", "Water Temperature", "Water Transparency",
    "Flow Level", "Water Color", "Water Clarity", "Water Surface", "Water Odor",
    "Present Weather", "Days Since Last Significant Precipitation", "Rainfall Accumulation",
    "Tide Stage", "E. coli", "Enterococci", "Nitrate-Nitrogen (NO₃-N)", "Nitrite-Nitrogen (NO₂-N)",
    "Orthophosphate (O-P)", "Total Dissolved Solids (TDS)", "Total Suspended Solids (TSS)",
    "Total Kjeldahl Nitrogen (TKN)", "Total Maximum Daily Load (TMDL)", "Watershed",
    "Tributary", "Segment", "Air Temperature", "Conductivity", "Algae Cover", "Turbidity",
    "Streamflow", "Chlorophyll-a", "Ammonia-Nitrogen (NH₃-N)", "Total Phosphorus",
    "Base Flow", "First Flush", "Eutrophication", "Hypoxia", "Riparian Zone",
    "Nonpoint Source Pollution", "Point Source Pollution", "Best Management Practices (BMPs)",
    "Watershed Protection Plan (WPP)", "Water Quality Standards", "Bioaccumulation",
    "Biological Integrity", "Channelization", "Chloride (Cl⁻)", "Chronic Toxicity"
]
standard_map = {normalize(name): name for name in standard_names}


def run_pipeline(df):
    """
    اینجا تمام مراحل پردازش و ولیدیشن کدی که داری پشت سر هم میاد
    الان نمونه ساده نوشتم → باید کد اصلیت رو بذاریم
    """
    issues = []
    
    # مثال: بررسی ستون‌ها
    for col in df.columns:
        norm = normalize(col)
        if norm in standard_map:
            df.rename(columns={col: standard_map[norm]}, inplace=True)
        else:
            issues.append(f"Column '{col}' not recognized as standard.")

    # یک ستون نمونه برای ولیدیشن
    if "pH" in df.columns:
        df["pH_flag"] = df["pH"].apply(lambda x: "Invalid (<0 or >14)" if (x < 0 or x > 14) else "OK")

    return df, issues


def generate_report(issues):
    """
    ساخت گزارش Word با متن
    """
    doc = Document()
    doc.add_heading("Validation Report", level=0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    doc.add_heading("Issues Found:", level=1)

    if not issues:
        doc.add_paragraph("No issues found ✅")
    else:
        for i, issue in enumerate(issues, 1):
            doc.add_paragraph(f"{i}. {issue}")

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

# ============================
# Streamlit App
# ============================
st.set_page_config(page_title="Data Cleaning App", page_icon="🧹", layout="wide")

st.title("🧹 Data Cleaning & Validation App")

uploaded_file = st.file_uploader("📤 Upload Excel/CSV file", type=["csv", "xlsx"])

if uploaded_file:
    # خواندن فایل
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("👀 Raw Data Preview")
    st.dataframe(df.head(20))

    # اجرای خط لوله
    df_clean, issues = run_pipeline(df.copy())

    st.subheader("✅ Cleaned Data Preview")
    st.dataframe(df_clean.head(20))

    # ساخت فایل اکسل خروجی
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        df_clean.to_excel(writer, index=False, sheet_name="CleanedData")
    excel_buf.seek(0)

    # ساخت گزارش ورد
    report_buf = generate_report(issues)

    # دکمه‌های دانلود
    st.subheader("⬇️ Download Outputs")
    st.download_button("💾 Download Cleaned Excel", data=excel_buf, file_name="cleaned_data.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📝 Download Validation Report", data=report_buf, file_name="Validation_Report.docx", mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document")
