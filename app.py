# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import re
from difflib import SequenceMatcher
from collections import Counter, defaultdict
from datetime import datetime
import io

# ==== Word report dependency ====
try:
    from docx import Document
except Exception:
    st.error(
        "The 'python-docx' package is required to generate the Word report.\n"
        "Please add `python-docx` to your requirements.txt and install dependencies."
    )
    raise

# ---------------------------
# Page config
# ---------------------------
st.set_page_config(
    page_title="Data Cleaning & Validation",
    page_icon="🧹",
    layout="wide",
    menu_items={
        "Get Help": "https://docs.streamlit.io/",
        "About": "Data Cleaning & Validation • Excel + Word report"
    }
)

# ---------------------------
# Global CSS (light aesthetic polish)
# ---------------------------
st.markdown("""
<style>
:root { --radius: 16px; }
.block-container { padding-top: 1.2rem; padding-bottom: 2rem; }
h1, h2, h3 { letter-spacing: .2px; }
.small-note { color: #6b7280; font-size: 0.9rem; }
.app-subtitle { color:#6b7280; margin-top:-12px; margin-bottom: 20px; }
.card {
  background: linear-gradient(180deg, #ffffff 0%, #fafafa 100%);
  border: 1px solid #eee;
  border-radius: var(--radius);
  padding: 16px 18px;
  box-shadow: 0 1px 2px rgba(16,24,40,.04);
}
.card h3 { margin: 0 0 6px 0; font-size: 1.05rem; }
.stat {
  background: #fff; border: 1px solid #eee; border-radius: 14px;
  padding: 14px; text-align: center;
}
.stat .value { font-size: 1.4rem; font-weight: 700; margin-top: 4px; }
.stat .label { font-size: .85rem; color:#6b7280; }
hr.soft { border: none; height: 1px; background: #eee; margin: 10px 0 18px; }
.download-wrap > div button { width: 100%; border-radius: 12px; padding: 10px 14px; }
.stTabs [data-baseweb="tab-list"] { gap: 6px; }
.stTabs [data-baseweb="tab"] { background: #f7f7f8; border-radius: 10px; padding: 10px 14px; }
.st-emotion-cache-1r6slb0 { padding-top: 0 !important; } /* remove extra top spacing on some themes */
</style>
""", unsafe_allow_html=True)

# ---------------------------
# Header
# ---------------------------
st.markdown("<h1>🧹 Data Cleaning & Validation</h1>", unsafe_allow_html=True)
st.markdown(
    '<div class="app-subtitle">End-to-end pipeline: Column mapping → Header transpose & swap → Validation → Final Excel + Word report.</div>',
    unsafe_allow_html=True
)

# ---------------------------
# Helpers
# ---------------------------
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

# ---------------------------
# Core pipeline
# ---------------------------
def run_pipeline_return_buffers(xl_raw_dict):
    final_data = {}
    mapping_log = []

    # Step 1: mapping
    for sheet_name, df in xl_raw_dict.items():
        new_columns = []
        for col in df.columns:
            orig_col_lower = str(col).lower()
            norm_col = normalize(col)
            best_match, best_score = None, 0.0

            if 'phosphate' in orig_col_lower:
                best_match, best_score = "Orthophosphate (O-P)", 1.0
            elif 'turbitidy' in orig_col_lower:
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
        header_df = pd.DataFrame([new_columns, list(df.columns)])
        full_df = pd.concat([header_df, data_df], ignore_index=True)
        final_data[sheet_name] = full_df

    # Step 2: transpose first two rows and swap first two columns
    step2_sheets = {}
    for sheet_name, df in final_data.items():
        first_two_rows = df.iloc[:2].T
        cols = list(first_two_rows.columns)
        if len(cols) >= 2:
            cols[0], cols[1] = cols[1], cols[0]
            first_two_rows = first_two_rows[cols]
        step2_sheets[sheet_name] = first_two_rows

    mapping_all = {}
    for sheet_name, mdf in step2_sheets.items():
        if mdf.shape[1] >= 2:
            tmp = mdf.iloc[:, :2].dropna(how="all")
            for _, row in tmp.iterrows():
                orig, mapped = str(row.iloc[0]).strip(), str(row.iloc[1]).strip()
                if orig and mapped and orig.lower() != "(no match)":
                    mapping_all[orig.lower()] = mapped.lower()

    # Step 3/4: validation
    validated_sheets_raw = {}
    for sheet_name, df in xl_raw_dict.items():
        df2 = df.copy()
        df2.columns = [str(c).lower() for c in df2.columns]
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
                        validation_results.append("Missing" if pd.isna(v) or str(v).strip()=='' else "Text")
                else:
                    for v in series:
                        if pd.isna(v) or str(v).strip()=='':
                            validation_results.append("Missing")
                        elif is_number_like(v):
                            val_num = float(str(v).strip())
                            validation_results.append("Valid" if rules["min"] <= val_num <= rules["max"] else "Invalid: out of range")
                        else:
                            validation_results.append("Text")

                col_index = df2.columns.get_loc(original_col)
                df2.insert(col_index+1, f"Validation - {original_col}", validation_results)

        validated_sheets_raw[sheet_name] = df2

    # Word report
    doc = Document()
    doc.add_heading("Validation Report", level=0)
    doc.add_paragraph(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")

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
        if s in ("Valid", "Text"): return None
        if s in ("✔", "✖"): return None
        if s == "" or s.lower() == "nan" or s == "None": return "Missing"
        if s.startswith("Invalid"): return "Out of Range"
        if s == "Missing": return "Missing"
        return "Other"

    for sheet_name, validated_df in validated_sheets_raw.items():
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

    report_buf = io.BytesIO()
    doc.save(report_buf)
    report_buf.seek(0)

    excel_buf = io.BytesIO()
    # Add mapping log as a sheet + validated sheets
    with pd.ExcelWriter(excel_buf, engine="openpyxl") as writer:
        pd.DataFrame(mapping_log).to_excel(writer, index=False, sheet_name="Column_Mapping")
        for sheet_name, vdf in validated_sheets_raw.items():
            vdf.to_excel(writer, sheet_name=sheet_name, index=False)

    excel_buf.seek(0)

    # Quick stats for UI cards
    stats = {
        "records": grand_total_records,
        "params": len(grand_validation_cols),
        "checks": grand_total_checks,
        "issues": grand_issues,
        "mapping_rows": len(mapping_log)
    }

    # Also provide lightweight previews for tabs
    preview_map = pd.DataFrame(mapping_log).head(40) if mapping_log else pd.DataFrame()
    any_sheet = next(iter(xl_raw_dict)) if xl_raw_dict else None
    preview_valid = None
    if any_sheet:
        preview_valid = validated_sheets_raw[any_sheet].head(20)

    return excel_buf, report_buf, stats, preview_map, preview_valid

# ---------------------------
# Sidebar (upload + run + help)
# ---------------------------
with st.sidebar:
    st.header("⚙️ Pipeline")
    uploaded = st.file_uploader("Upload CSV or Excel", type=["csv", "xlsx"])
    run_clicked = st.button("▶️ Run Pipeline", use_container_width=True)

    st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
    with st.expander("What this app does"):
        st.markdown("""
- **Maps** raw column names to standardized names (fuzzy).
- **Transposes** first two header rows and **swaps** the first two columns.
- **Validates** numeric ranges (configurable per parameter).
- Exports **Final Excel** and a **Word report** with summaries.
        """)
    st.markdown("<div class='small-note'>Tip: Add more rules in <code>param_dict</code> to expand validation.</div>", unsafe_allow_html=True)

# ---------------------------
# Main Area
# ---------------------------
tabs = st.tabs(["How it works", "Mapping preview", "Validation preview"])

with tabs[0]:
    st.markdown("""
<div class="card">
<h3>Pipeline Steps</h3>
<ol>
<li><b>Column Mapping</b> – fuzzy match raw headers to a controlled list.</li>
<li><b>Header Transpose & Swap</b> – rotate the first 2 rows, then swap first 2 columns to align <i>(Original, Mapped)</i>.</li>
<li><b>Validation</b> – for mapped parameters that have rules: check numeric/text/missing/out-of-range.</li>
<li><b>Reports</b> – download a clean Excel (with “Validation - …” columns) and a readable Word report.</li>
</ol>
</div>
""", unsafe_allow_html=True)

progress_placeholder = st.empty()
result_placeholder = st.container()

if uploaded and run_clicked:
    try:
        # Progress bar simulation
        pb = progress_placeholder.progress(10, text="Reading input...")
        if uploaded.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded)
            xl_raw_dict = {"Sheet1": df}
        else:
            xl_raw_dict = pd.read_excel(uploaded, sheet_name=None)
        pb.progress(35, text="Mapping & header processing...")
        # Run pipeline
        excel_buf, report_buf, stats, preview_map, preview_valid = run_pipeline_return_buffers(xl_raw_dict)
        pb.progress(85, text="Generating reports...")
        pb.progress(100, text="Done")

        with result_placeholder:
            st.success("✅ Processing completed!")

            # Stats cards
            c1, c2, c3, c4, c5 = st.columns(5)
            c1.markdown(f"<div class='stat'><div class='label'>Records</div><div class='value'>{stats['records']:,}</div></div>", unsafe_allow_html=True)
            c2.markdown(f"<div class='stat'><div class='label'>Parameters</div><div class='value'>{stats['params']:,}</div></div>", unsafe_allow_html=True)
            c3.markdown(f"<div class='stat'><div class='label'>Checks</div><div class='value'>{stats['checks']:,}</div></div>", unsafe_allow_html=True)
            c4.markdown(f"<div class='stat'><div class='label'>Issues</div><div class='value'>{stats['issues']:,}</div></div>", unsafe_allow_html=True)
            c5.markdown(f"<div class='stat'><div class='label'>Mapping rows</div><div class='value'>{stats['mapping_rows']:,}</div></div>", unsafe_allow_html=True)

            st.markdown("<hr class='soft'/>", unsafe_allow_html=True)
            d1, d2 = st.columns([1,1])
            with d1:
                st.markdown("<div class='card'><h3>📥 Final Excel</h3><div class='small-note'>Validated sheets + Column_Mapping tab.</div>", unsafe_allow_html=True)
                st.download_button(
                    "💾 Download Excel",
                    data=excel_buf,
                    file_name="validated_output.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
                st.markdown("</div>", unsafe_allow_html=True)
            with d2:
                st.markdown("<div class='card'><h3>📝 Word Report</h3><div class='small-note'>Readable summaries by sheet and overall.</div>", unsafe_allow_html=True)
                st.download_button(
                    "📝 Download Report",
                    data=report_buf,
                    file_name="Validation_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    use_container_width=True
                )
                st.markdown("</div>", unsafe_allow_html=True)

            # Fill previews in tabs
            with tabs[1]:
                st.subheader("Column Mapping (sample)")
                if not preview_map.empty:
                    st.dataframe(preview_map, use_container_width=True, height=360)
                else:
                    st.info("No mapping rows to preview.")
            with tabs[2]:
                st.subheader("Validated Sheet (sample)")
                if preview_valid is not None and not preview_valid.empty:
                    st.dataframe(preview_valid, use_container_width=True, height=420)
                else:
                    st.info("No validated data to preview.")

    except Exception as e:
        progress_placeholder.empty()
        st.error(f"❌ Processing error: {e}")
elif not uploaded:
    st.info("Upload a **CSV** or **Excel** file from the sidebar, then click **Run Pipeline**.")

# Footer
st.markdown("<div class='small-note' style='margin-top:30px;'>© 2025 • Data Cleaning & Validation • Built with Streamlit</div>", unsafe_allow_html=True)
