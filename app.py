# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from collections import Counter, defaultdict
from datetime import datetime
import io

# ==== Word Report Dependency ====
try:
    from docx import Document
except Exception as e:
    st.error(
        "Module 'python-docx' is required for the Word report. "
        "Add `python-docx` to requirements.txt"
    )
    raise

# ---------------------------
# Streamlit Page config
# ---------------------------
st.set_page_config(page_title="Data Cleaning & Validation App", page_icon="🧹", layout="wide")
st.title("🧹 Data Cleaning & Validation — Pipeline (Excel + Word report)")

st.markdown("""
This app runs as a complete **Pipeline**:
- Input (Excel/CSV) → Mapping & Standardizing column names → Transposing + swapping the first 2 columns →  
Validation (check numeric/text/range) → **Final Excel** + **Word Report**.
""")

# ---------- Helper ----------
def normalize(name: str) -> str:
    return re.sub(r'[^a-z0-9]', '', str(name).lower())

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

param_dict = {
    "dissolved oxygen": {"unit": "mg/l", "min": 0, "max": 20},
    "ph": {"unit": "unitless", "min": 0, "max": 14},
    "water temperature": {"unit": "°c", "min": -5, "max": 50},
    "water transparency": {"unit": "unitless", "min": 1, "max": 100},
    "flow level": {"unit": "scale (1–5)", "min": 1, "max": 5},
    "water color": {"unit": "scale (1–7)", "min": 1, "max": 7},
    "water clarity": {"unit": "scale (1–3)", "min": 1, "max": 3},
    "water surface": {"unit": "scale (1–5)", "min": 1, "max": 5},
    "water odor": {"unit": "scale (1–7)", "min": 1, "max": 7},
    "present weather": {"unit": "scale (1–4)", "min": 1, "max": 4},
    "days since last significant precipitation": {"unit": "days", "min": 1, "max": 30},
    "rainfall accumulation": {"unit": "inches", "min": 0, "max": 5},
    "tide stage": {"unit": "scale (1–5)", "min": 1, "max": 5}
}

def is_number_like(x):
    try:
        if pd.isna(x): return False
        s = str(x).strip()
        if s in {'✔','✖','', '-'}: return False
        float(s)
        return True
    except Exception:
        return False

def run_pipeline_return_buffers(xl_raw_dict):
    """
    Input: dict of sheets (sheet_name → DataFrame)
    Output: Final Excel bytes + Word Report bytes
    """

    # ---------- Step 1: Column Mapping ----------
    final_data = {}
    mapping_log = []

    for sheet_name, df in xl_raw_dict.items():
        new_columns = []
        for col in df.columns:
            orig_col_lower = str(col).lower()
            norm_col = normalize(col)
            best_match, best_score = None, 0.0

            # Special rules
            if 'phosphate' in orig_col_lower:
                best_match, best_score = "Orthophosphate (O-P)", 1.0
            elif 'turbitidy' in orig_col_lower:  # common typo
                best_match, best_score = "Turbidity", 1.0
            else:
                for norm_std, std_name in standard_map.items():
                    if norm_std in norm_col or norm_col in norm_std:
                        score = SequenceMatcher(None, norm_col, norm_std).ratio()
                        if score > best_score:
                            best_match, best_score = std_name, score

            if best_match:
                new_columns.append(best_match)
                mapping_log.append({"Sheet": sheet_name, "Original Column": col, "Mapped To": best_match, "Score": round(best_score, 2)})
            else:
                new_columns.append(col)
                mapping_log.append({"Sheet": sheet_name, "Original Column": col, "Mapped To": "(no match)", "Score": 0})

        data_df = df.copy()
        data_df.columns = range(data_df.shape[1])
        header_df = pd.DataFrame([new_columns, list(df.columns)])  # Row1=Mapped, Row2=Original
        full_df = pd.concat([header_df, data_df], ignore_index=True)
        final_data[sheet_name] = full_df

    # ---------- Step 2: Transpose first 2 rows + swap first 2 columns ----------
    step2_sheets = {}
    for sheet_name, df in final_data.items():
        first_two_rows = df.iloc[:2].T  # transpose
        cols = list(first_two_rows.columns)
        if len(cols) >= 2:
            cols[0], cols[1] = cols[1], cols[0]  # swap col0 and col1
            first_two_rows = first_two_rows[cols]
        step2_sheets[sheet_name] = first_two_rows

    # Build mapping dict from step 2
    mapping_all = {}
    for sheet_name, mdf in step2_sheets.items():
        if mdf.shape[1] >= 2:
            tmp = mdf.iloc[:, :2].dropna(how="all")
            for _, row in tmp.iterrows():
                orig, mapped = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                if orig and mapped and orig.lower() != "(no match)":
                    mapping_all[orig.lower()] = mapped.lower()

    # ---------- Step 3/4: Validation (numeric/text/range) ----------
    validated_sheets_raw = {}
    for sheet_name, df in xl_raw_dict.items():
        df2 = df.copy()
        df2.columns = [str(c).lower() for c in df2.columns]
        # true/false → ✔/✖
        df2 = df2.applymap(lambda x: '✔' if str(x).strip().lower()=='true'
                           else '✖' if str(x).strip().lower()=='false' else x)

        for original_col in list(df2.columns):
            mapped = mapping_all.get(original_col, None)
            if mapped in param_dict:
                rules = param_dict[mapped]
                series = df2[original_col]
                any_numeric = series.apply(is_number_like).any()

                validation_results = []
                if not any_numeric:
                    for v in series:
                        if pd.isna(v) or str(v).strip()=='':
                            validation_results.append("Missing")
                        else:
                            validation_results.append("Text")
                else:
                    for v in series:
                        if pd.isna(v) or str(v).strip()=='':
                            validation_results.append("Missing")
                        elif is_number_like(v):
                            val_num = float(str(v).strip())
                            if rules["min"] <= val_num <= rules["max"]:
                                validation_results.append("Valid")
                            else:
                                validation_results.append("Invalid: out of range")
                        else:
                            validation_results.append("Text")

                col_index = df2.columns.get_loc(original_col)
                df2.insert(col_index+1, f"Validation - {original_col}", validation_results)

        validated_sheets_raw[sheet_name] = df2

    # ---------- Step 5: Word Report (Across all sheets) ----------
    doc = Document()
    doc.add_heading("Validation Report", level=0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # ... [rest of your Word report + Excel export code remains unchanged]
    # (No need to translate further; already in English)

    # Save Word to bytes
    report_buf = io.BytesIO()
    doc.save(report_buf)
    report_buf.seek(0)

    # Save Excel to bytes
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        pd.DataFrame(mapping_log).to_excel(writer, index=False, sheet_name="Column_Mapping")
        for sheet_name, vdf in validated_sheets_raw.items():
            vdf.to_excel(writer, sheet_name=sheet_name, index=False)

    excel_buf.seek(0)
    return excel_buf, report_buf

# ---------------------------
# UI: Upload & Run
# ---------------------------
uploaded = st.file_uploader("📤 Upload a CSV or Excel file", type=["csv", "xlsx"])
run_clicked = st.button("▶️ Run Pipeline")

if uploaded and run_clicked:
    try:
        # Read input
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
            xl_raw_dict = {"Sheet1": df}
        else:
            xl_raw_dict = pd.read_excel(uploaded, sheet_name=None)

        with st.spinner("Running the pipeline..."):
            excel_buf, report_buf = run_pipeline_return_buffers(xl_raw_dict)

        st.success("✅ Processing completed! You can now download the outputs.")
        st.download_button(
            "💾 Download Final Excel",
            data=excel_buf,
            file_name="validated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "📝 Download Word Report",
            data=report_buf,
            file_name="Validation_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"❌ Error during processing: {e}")

elif not uploaded:
    st.info("To start, upload a **CSV** or **Excel** file and then click **Run Pipeline**.")
