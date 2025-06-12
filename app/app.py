import streamlit as st
import pandas as pd
import zipfile
import os
import io
import shutil
import re

st.title("Aircraft Support Report Extractor")

uploaded_zip = st.file_uploader("Upload ZIP file of Daily Ops Excel Reports", type="zip")

if uploaded_zip:
    with st.spinner("Processing ZIP file..."):

        # Derive month-year for sheet name from file name
        zip_filename = uploaded_zip.name
        match = re.search(r'([A-Za-z]+)[_\-]?(\d{4})', zip_filename)
        if match:
            month_str = match.group(1).upper()
            year_str = match.group(2)
            sheet_suffix = f"{month_str}_{year_str}"
        else:
            sheet_suffix = "REPORT"

        # Extract uploaded zip
        extract_path = "/tmp/extracted_reports"
        if os.path.exists(extract_path):
            shutil.rmtree(extract_path)
        os.makedirs(extract_path, exist_ok=True)

        with zipfile.ZipFile(uploaded_zip, 'r') as zip_ref:
            zip_ref.extractall(extract_path)

        lot_rows = []
        rj_rows = []

        for filename in os.listdir(extract_path):
            if filename.endswith(".xlsx"):
                file_path = os.path.join(extract_path, filename)
                df = pd.read_excel(file_path, header=None)
                header_row_index = None
                for i in range(min(10, len(df))):
                    row = df.iloc[i].astype(str).str.lower().tolist()
                    if any("aircraft" in str(cell) or "flt" in str(cell) for cell in row):
                        header_row_index = i
                        break

                if header_row_index is None:
                    continue

                df = pd.read_excel(file_path, header=header_row_index)
                cols = [str(c).strip().lower() for c in df.columns]
                df.columns = cols

                reg_col = next((c for c in cols if "reg" in c), None)
                date_col = next((c for c in cols if "date" in c), None)
                flt_col = next((c for c in cols if "flt" in c or "flight number" in c), None)
                type_col = next((c for c in cols if "type" in c), "")
                airline_col = next((c for c in cols if "airline" in c), "")
                remarks_col = next((c for c in cols if "remark" in c or "services" in c), cols[-1])

                if not reg_col or not date_col or not flt_col:
                    continue

                for _, row in df.iterrows():
                    reg = str(row.get(reg_col, "")).strip().upper()
                    support_text = str(row.get(remarks_col, "")).strip().lower()
                    flight_num = str(row.get(flt_col, "")).strip()

                    if support_text not in ["on call", "on call - needed engineer support"]:
                        continue

                    try:
                        date_parsed = pd.to_datetime(row.get(date_col))
                        formatted_date = date_parsed.strftime("%d-%b").upper()
                    except Exception:
                        continue  # skip row if date parsing fails

                    support = "YES" if support_text == "on call - needed engineer support" else "NO"

                    if reg.startswith("SP-L"):
                        clean_flight_num = re.sub(r"\bLO\s*", "", flight_num, flags=re.IGNORECASE)
                        lot_rows.append({
                            "FLT NUMBER": clean_flight_num,
                            "DATE": formatted_date,
                            "DATE_SORT": date_parsed,
                            "AIRCRAFT REG": reg,
                            "ENG SUPPORT": support
                        })
                    elif reg.startswith("JY-"):
                        rj_rows.append({
                            "DATE_SORT": date_parsed,
                            "NUMBERS OF FLIGHT": len(rj_rows) + 1,
                            "DATE OF FLIGHT": formatted_date,
                            "AIRLINES": row.get(airline_col, "ROYAL JORDANIAN"),
                            "AIRCRAFT TYPE": row.get(type_col, ""),
                            "AIRCRAFT REGISTRATION": reg,
                            "FLIGHT NUMBER": flight_num,
                            "TECH SUPT.  YES/NO": support
                        })

        # Create DataFrames
        lot_df = pd.DataFrame(lot_rows)
        rj_df = pd.DataFrame(rj_rows)

        # Sort and drop helper column if it exists
        if not lot_df.empty and "DATE_SORT" in lot_df.columns:
            lot_df = lot_df.sort_values("DATE_SORT").drop(columns=["DATE_SORT"])
        if not rj_df.empty and "DATE_SORT" in rj_df.columns:
            rj_df = rj_df.sort_values("DATE_SORT").drop(columns=["DATE_SORT"])

        # Write to Excel
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
            if not lot_df.empty:
                lot_df.to_excel(writer, sheet_name=f"LOT_{sheet_suffix}", index=False)
            if not rj_df.empty:
                rj_df.to_excel(writer, sheet_name=f"RJ_{sheet_suffix}", index=False)
        output.seek(0)

    # Final success message and download button
    st.success("Report is ready!")
    st.download_button("Download Excel File", output.getvalue(), file_name="Aircraft_Support_Report.xlsx")
