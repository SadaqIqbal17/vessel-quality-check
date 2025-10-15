import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import get_close_matches
import re

# ---------------- Helper Functions ----------------

def find_table_start(df):
    for i, row in df.iterrows():
        row_str = " ".join(str(x) for x in row.values if pd.notna(x)).lower()
        if "description" in row_str or "method" in row_str:
            return i
    return None

def extract_product_from_sheetname(sheet_name):
    sheet_upper = sheet_name.upper()
    if "MOGAS" in sheet_upper:
        return "MOGAS 92 RON" if "92" in sheet_upper else "MOGAS 95 RON"
    elif "HSD" in sheet_upper or "DIESEL" in sheet_upper:
        return "HSD"
    elif "JET" in sheet_upper:
        return "JET FUEL"
    elif "HOBC" in sheet_upper or "OCTANE" in sheet_upper:
        return "HOBC"
    else:
        match = re.search(r"(MOGAS.*RON|JET.*FUEL|HSD|DIESEL|HOBC|OCTANE)", sheet_upper)
        return match.group(1).strip() if match else "UNKNOWN"

def clean_numeric(val):
    if isinstance(val, str):
        val = val.replace("<", "").replace(">", "").strip()
    try:
        return float(val)
    except:
        return None

def compare_value(test_val, min_val, max_val):
    if min_val is not None and test_val < min_val:
        return False
    if max_val is not None and test_val > max_val:
        return False
    return True

def clean_table(df):
    start_row = find_table_start(df)
    if start_row is None:
        return None
    table = df.iloc[start_row:]
    table.columns = table.iloc[0]
    table = table.drop(table.index[0]).reset_index(drop=True)
    table.columns = [str(c).strip().replace("\n", " ") for c in table.columns]
    table = table.dropna(how="all").reset_index(drop=True)
    return table

# ---------------- Core Analysis Logic ----------------

def process_vessel_file(vessel_file, standards_file, month_year):
    std_xls = pd.read_excel(standards_file, sheet_name=None)
    vessel_xls = pd.ExcelFile(vessel_file)
    summary = []
    detailed_results = {}

    for sheet_name in vessel_xls.sheet_names:
        raw = pd.read_excel(vessel_xls, sheet_name=sheet_name, header=None)
        table = clean_table(raw)
        if table is None:
            continue

        product = extract_product_from_sheetname(sheet_name)
        match = get_close_matches(product, list(std_xls.keys()), n=1, cutoff=0.4)
        if not match:
            continue
        std_df = std_xls[match[0]]

        hdip_cols = [c for c in table.columns if "hdip" in c.lower() and "result" in c.lower()]
        load_cols = [c for c in table.columns if "load" in c.lower() and "result" in c.lower()]
        if not hdip_cols:
            possible_cols = [c for c in table.columns if "result" in c.lower()]
            if len(possible_cols) >= 2:
                hdip_col, load_col = possible_cols[0], possible_cols[1]
            elif len(possible_cols) == 1:
                hdip_col, load_col = possible_cols[0], None
            else:
                continue
        else:
            hdip_col = hdip_cols[0]
            load_col = load_cols[0] if load_cols else None

        passed, failed = 0, 0
        failed_tests = []
        detailed_data = []

        for _, row in table.iterrows():
            test_name = str(row.get("Tests Description", "")).strip()
            if not test_name:
                continue

            hdip_val = clean_numeric(row.get(hdip_col))
            load_val = clean_numeric(row.get(load_col)) if load_col else None
            match_test = get_close_matches(test_name, std_df["Parameter"].astype(str), n=1, cutoff=0.6)
            if not match_test:
                continue

            std_row = std_df[std_df["Parameter"] == match_test[0]].iloc[0]
            min_val = clean_numeric(std_row.get("Min"))
            max_val = clean_numeric(std_row.get("Max"))
            unit = std_row.get("Unit / Remarks", "")

            if isinstance(hdip_val, (int, float)) and (min_val is not None or max_val is not None):
                ok = compare_value(hdip_val, min_val, max_val)
                status = "PASS" if ok else "FAIL"
                detailed_data.append([test_name, hdip_val, load_val, min_val, max_val, unit, status])
                if ok:
                    passed += 1
                else:
                    failed += 1
                    failed_tests.append(test_name)

        overall = "PASS" if failed == 0 else "FAIL"
        summary.append([sheet_name, product, passed + failed, passed, failed, overall, ", ".join(failed_tests)])
        detailed_results[sheet_name] = pd.DataFrame(
            detailed_data,
            columns=["Test", "HDIP Result", "Load Port Result", "Min", "Max", "Unit/Remarks", "Status"]
        )

    summary_df = pd.DataFrame(summary, columns=[
        "Sheet Name", "Product", "Tests Checked", "Passed", "Failed", "Overall Result", "Failed Tests"
    ])

    # Prepare Excel output in memory
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        summary_df.to_excel(writer, index=False, sheet_name="Summary")
        for name, df in detailed_results.items():
            df.to_excel(writer, index=False, sheet_name=name[:31])
    output.seek(0)

    return summary_df, output


# ---------------- Streamlit UI ----------------

st.set_page_config(page_title="Vessel Quality Verification", layout="wide")

st.title("🛢️ Vessel Quality Verification System")
st.markdown("#### Automated QA check for OGRA product testing reports")

with st.expander("📋 Instructions"):
    st.markdown("""
    1. Upload **Standards File** (e.g. `standards.xlsx`)  
    2. Upload **Vessel Report File** (e.g. `vessel_reports.xlsx`)  
    3. Enter **Month and Year** (e.g. `September 2025`)  
    4. Click **Run Analysis** to generate the summary and detailed report.
    """)

standards_file = st.file_uploader("📘 Upload Standards Excel", type=["xlsx"])
vessel_file = st.file_uploader("📄 Upload Vessel Report Excel", type=["xlsx"])
month_year = st.text_input("🗓️ Enter Month and Year (e.g., September 2025)")

if st.button("▶️ Run Analysis"):
    if not standards_file or not vessel_file or not month_year:
        st.error("Please upload both files and enter month/year.")
    else:
        with st.spinner("Processing files... Please wait."):
            summary_df, output = process_vessel_file(vessel_file, standards_file, month_year)

        st.success("✅ Analysis Complete!")

        st.subheader("📊 Summary Results")
        st.dataframe(summary_df.style.applymap(
            lambda v: 'color:red;font-weight:bold;' if v == 'FAIL' else 'color:green;' if v == 'PASS' else ''
        ))

        st.download_button(
            label="📥 Download Full Quality Report",
            data=output,
            file_name=f"Quality_Report_{month_year.replace(' ', '_')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
