# Aircraft Support Report Extractor

This Streamlit app extracts aircraft records for LOT and Royal Jordanian from daily operational Excel reports.
Only aircraft entries with 'ON CALL' or 'ON CALL - NEEDED ENGINEER SUPPORT' in the remarks are included.

## How to Use

1. Upload a ZIP file containing the daily `.xlsx` reports.
2. The app will process and filter the data.
3. Download a single Excel file with two sheets: one for LO (LOT) and one for RJ (Royal Jordanian).

## Run the App

```bash
pip install -r requirements.txt
streamlit run app/app.py
```