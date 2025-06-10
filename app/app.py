
import streamlit as st
import pandas as pd
import zipfile
import os
import io

st.title("Aircraft Support Report Extractor")

uploaded_zip = st.file_uploader("Upload ZIP file of Daily Ops Excel Reports", type="zip")

if uploaded_zip:
    st.success("File uploaded. Processing...")

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

                if support_text not in ["on call", "on call - needed engineer support"]:
                    continue

                formatted_date = pd.to_datetime(row.get(date_col)).strftime("%d-%b").upper()

                if reg.startswith("SP-L"):
                    support = "YES" if support_text == "on call - needed engineer support" else "NO"
                    lot_rows.append({
                        "FLT NUMBER": row.get(flt_col, ""),
                        "DATE": formatted_date,
                        "AIRCRAFT REG": reg,
                        "ENG SUPPORT": support
                    })
                elif reg.startswith("JY-"):
                    support = "YES" if support_text == "on call - needed engineer support" else "NO"
                    rj_rows.append({
                        "NUMBERS OF FLIGHT": len(rj_rows) + 1,
                        "DATE OF FLIGHT": formatted_date,
                        "AIRLINES": row.get(airline_col, "ROYAL JORDANIAN"),
                        "AIRCRAFT TYPE": row.get(type_col, ""),
                        "AIRCRAFT REGISTRATION": reg,
                        "FLIGHT NUMBER": row.get(flt_col, ""),
                        "TECH SUPT.  YES/NO": support
                    })

    # Convert to DataFrames
    lot_df = pd.DataFrame(lot_rows)
    rj_df = pd.DataFrame(rj_rows)

    # Save to Excel in memory
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        lot_df.to_excel(writer, sheet_name="LOT_MAY_2025", index=False)
        rj_df.to_excel(writer, sheet_name="RJ_MAY_2025", index=False)

    st.download_button("Download Excel File", output.getvalue(), file_name="Aircraft_Support_Report.xlsx")
