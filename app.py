# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from collections import Counter, defaultdict
from datetime import datetime
import io

# ==== وابستگی گزارش Word ====
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
این اپ به صورت یک **Pipeline** کامل کار می‌کند:
- ورودی (Excel/CSV) → نگاشت و استانداردسازی نام ستون‌ها → افقی‌سازی و جابجایی ۲ ستون اول →  
اعتبارسنجی (تشخیص عدد/متن/رنج) → **Excel نهایی** + **گزارش Word**.
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
    ورودی: dict از شیت‌ها (نام شیت → DataFrame)
    خروجی: بایت‌های Excel نهایی + بایت‌های Word Report
    """

    # ---------- مرحله 1: نگاشت ستون‌ها ----------
    final_data = {}
    mapping_log = []

    for sheet_name, df in xl_raw_dict.items():
        new_columns = []
        for col in df.columns:
            orig_col_lower = str(col).lower()
            norm_col = normalize(col)
            best_match, best_score = None, 0.0

            # قوانین خاص
            if 'phosphate' in orig_col_lower:
                best_match, best_score = "Orthophosphate (O-P)", 1.0
            elif 'turbitidy' in orig_col_lower:  # تایپ رایج
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
        header_df = pd.DataFrame([new_columns, list(df.columns)])  # ردیف1=Mapped، ردیف2=Original
        full_df = pd.concat([header_df, data_df], ignore_index=True)
        final_data[sheet_name] = full_df

    # ---------- مرحله 2: افقی‌سازی ۲ ردیف اول + جابجایی دو ستون اول ----------
    step2_sheets = {}
    for sheet_name, df in final_data.items():
        first_two_rows = df.iloc[:2].T  # ترانهاده
        cols = list(first_two_rows.columns)
        if len(cols) >= 2:
            cols[0], cols[1] = cols[1], cols[0]  # col0=original, col1=mapped (یا برعکس)
            first_two_rows = first_two_rows[cols]
        step2_sheets[sheet_name] = first_two_rows

    # ساخت dict نگاشت از مرحله 2
    mapping_all = {}
    for sheet_name, mdf in step2_sheets.items():
        if mdf.shape[1] >= 2:
            tmp = mdf.iloc[:, :2].dropna(how="all")
            for _, row in tmp.iterrows():
                orig, mapped = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                if orig and mapped and orig.lower() != "(no match)":
                    mapping_all[orig.lower()] = mapped.lower()

    # ---------- مرحله 3/4: اعتبارسنجی (عدد/متن/رنج) ----------
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

    # ---------- مرحله 5: ساخت گزارش Word (Across all sheets) ----------
    doc = Document()
    doc.add_heading("Validation Report", level=0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

    # تجمیع شمارنده‌ها
    grand_total_records = 0
    grand_validation_cols = set()
    grand_total_checks = 0
    grand_issues = 0

    grand_per_issue_type_counter = Counter()
    grand_per_param_counter = Counter()

    grand_per_site_issue_counter = defaultdict(Counter)
    grand_per_site_param_counter = defaultdict(Counter)
    grand_per_site_valid_counts = Counter()
    grand_per_site_total_counts = Counter()

    def normalize_issue(val):
        s = str(val).strip()
        if s in ("Valid", "Text"):  # Text را معتبر حساب می‌کنیم
            return None
        if s in ("✔", "✖"):
            return None
        if s == "" or s.lower() == "nan" or s == "None":
            return "Missing"
        if s.startswith("Invalid"):
            return "Out of Range"
        if s == "Missing":
            return "Missing"
        return "Other"

    for sheet_name, validated_df in validated_sheets_raw.items():
        # ستون‌های کمکی (case-insensitive contains)
        def pick(col_part):
            cols = [c for c in validated_df.columns if col_part in str(c).lower()]
            return cols[0] if cols else None

        site_name_col = pick("site name")
        site_id_col   = pick("site id")
        date_col      = pick("sample date")

        validation_cols = [c for c in validated_df.columns if str(c).startswith("Validation - ")]
        total_records   = len(validated_df)
        total_checks    = total_records * len(validation_cols)

        grand_total_records += total_records
        grand_validation_cols.update(validation_cols)
        grand_total_checks += total_checks

        issues = []
        per_param_counter = Counter()
        per_issue_type_counter = Counter()
        per_site_issue_counter = defaultdict(Counter)
        per_site_param_counter = defaultdict(Counter)
        per_site_valid_counts = Counter()
        per_site_total_counts = Counter()

        for _, row in validated_df.iterrows():
            site_id   = row[site_id_col] if site_id_col else ""
            site_name = row[site_name_col] if site_name_col else ""
            sample_dt = row[date_col] if date_col else ""

            site_key = (str(site_id), str(site_name))

            for col in validation_cols:
                param_name = col.replace("Validation - ", "")
                result = row[col]

                per_site_total_counts[site_key] += 1
                if str(result).strip() in ("Valid", "Text"):
                    per_site_valid_counts[site_key] += 1

                issue_type = normalize_issue(result)
                if issue_type:
                    issues.append((str(site_id), str(site_name), sample_dt, param_name, issue_type))
                    per_param_counter[param_name] += 1
                    per_issue_type_counter[issue_type] += 1
                    per_site_issue_counter[site_key][issue_type] += 1
                    per_site_param_counter[site_key][param_name] += 1

        total_issues = len(issues)
        grand_issues += total_issues
        grand_per_issue_type_counter.update(per_issue_type_counter)
        grand_per_param_counter.update(per_param_counter)

        for sk, cnt in per_site_valid_counts.items():
            grand_per_site_valid_counts[sk] += cnt
        for sk, cnt in per_site_total_counts.items():
            grand_per_site_total_counts[sk] += cnt
        for sk, c in per_site_issue_counter.items():
            grand_per_site_issue_counter[sk].update(c)
        for sk, c in per_site_param_counter.items():
            grand_per_site_param_counter[sk].update(c)

        # گزارش هر شیت
        doc.add_heading(f"[Sheet: {sheet_name}] Summary", level=1)
        doc.add_paragraph(f"- Total records reviewed: {total_records:,}")
        doc.add_paragraph(f"- Parameters validated: {len(validation_cols):,}")
        doc.add_paragraph(f"- Total checks (rows × parameters): {total_checks:,}")
        doc.add_paragraph(f"- Total issues detected: {total_issues:,}")

        if per_issue_type_counter:
            doc.add_paragraph("Issues by type:")
            for itype, cnt in per_issue_type_counter.most_common():
                doc.add_paragraph(f"   • {itype}: {cnt:,}")
        else:
            doc.add_paragraph("No issues detected across all checks.")

        doc.add_heading("Issues by Parameter", level=2)
        if per_param_counter:
            for param, cnt in per_param_counter.most_common():
                doc.add_paragraph(f"• {param}: {cnt:,} issue(s)")
        else:
            doc.add_paragraph("No parameter-level issues detected.")

    # Overall
    doc.add_heading("Overall Summary (All Sheets)", level=1)
    doc.add_paragraph(f"- Total records reviewed: {grand_total_records:,}")
    doc.add_paragraph(f"- Parameters validated: {len(grand_validation_cols):,}")
    doc.add_paragraph(f"- Total checks (rows × parameters): {grand_total_checks:,}")
    doc.add_paragraph(f"- Total issues detected: {grand_issues:,}")

    if grand_per_issue_type_counter:
        doc.add_paragraph("Issues by type:")
        for itype, cnt in grand_per_issue_type_counter.most_common():
            doc.add_paragraph(f"   • {itype}: {cnt:,}")
    else:
        doc.add_paragraph("No issues detected across all checks.")

    doc.add_heading("Issues by Parameter (All Sheets)", level=1)
    if grand_per_param_counter:
        for param, cnt in grand_per_param_counter.most_common():
            doc.add_paragraph(f"• {param}: {cnt:,} issue(s)")
    else:
        doc.add_paragraph("No parameter-level issues detected.")

    # Station-by-Station
    doc.add_heading("Station-by-Station Summary (All Sheets)", level=1)
    all_site_keys = set(grand_per_site_total_counts.keys()) | set(grand_per_site_issue_counter.keys())

    def station_sort_key(item):
        site_key, issue_cnt = item
        total_issue = sum(issue_cnt.values())
        return (-total_issue, site_key[1], site_key[0])

    for (sid, sname), issue_cnt in sorted(
        ((sk, grand_per_site_issue_counter.get(sk, Counter())) for sk in all_site_keys),
        key=station_sort_key
    ):
        total_c = grand_per_site_total_counts.get((sid, sname), 0) or 1
        valid_c = grand_per_site_valid_counts.get((sid, sname), 0)
        overall_validity = 100.0 * valid_c / total_c

        doc.add_paragraph(f"■ Station '{sname}' (ID: {sid})")
        doc.add_paragraph(f"  - Total issues: {sum(issue_cnt.values()):,} | Overall Validity: {overall_validity:.1f}%")

        if issue_cnt:
            doc.add_paragraph("  - Issues by type:")
            for itype, cnt in issue_cnt.most_common():
                doc.add_paragraph(f"     • {itype}: {cnt:,}")

        top_params_site = grand_per_site_param_counter.get((sid, sname), Counter())
        if top_params_site:
            doc.add_paragraph("  - Most frequent problematic parameters:")
            for p, c in top_params_site.most_common():
                doc.add_paragraph(f"     • {p}: {c:,} issue(s)")

    # Row-Level Details: (می‌تونی در صورت نیاز اضافه کنی)

    # ذخیره Word به بایت‌ها
    report_buf = io.BytesIO()
    doc.save(report_buf)
    report_buf.seek(0)

    # ساخت Excel نهایی از validated_sheets_raw (بدون استایل برای سازگاری کامل Cloud)
    excel_buf = io.BytesIO()
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        # یک شیت خلاصه نگاشت هم اضافه می‌کنیم:
        pd.DataFrame(mapping_log).to_excel(writer, index=False, sheet_name="Column_Mapping")

        for sheet_name, vdf in validated_sheets_raw.items():
            vdf.to_excel(writer, sheet_name=sheet_name, index=False)

    excel_buf.seek(0)
    return excel_buf, report_buf

# ---------------------------
# UI: Upload & Run
# ---------------------------
uploaded = st.file_uploader("📤 فایل CSV یا Excel را انتخاب کنید", type=["csv", "xlsx"])
run_clicked = st.button("▶️ Run Pipeline")

if uploaded and run_clicked:
    try:
        # خواندن ورودی
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
            xl_raw_dict = {"Sheet1": df}
        else:
            # Excel چند شیتی
            xl_raw_dict = pd.read_excel(uploaded, sheet_name=None)

        with st.spinner("در حال اجرای Pipeline ..."):
            excel_buf, report_buf = run_pipeline_return_buffers(xl_raw_dict)

        st.success("✅ پردازش تکمیل شد! حالا می‌توانید خروجی‌ها را دانلود کنید.")
        st.download_button(
            "💾 دانلود Excel نهایی",
            data=excel_buf,
            file_name="validated_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.download_button(
            "📝 دانلود گزارش Word",
            data=report_buf,
            file_name="Validation_Report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    except Exception as e:
        st.error(f"❌ خطا در پردازش: {e}")

elif not uploaded:
    st.info("برای شروع، یک فایل **CSV** یا **Excel** آپلود کنید و سپس روی **Run Pipeline** بزنید.")
