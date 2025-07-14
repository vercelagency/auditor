import streamlit as st
import pandas as pd
import openpyxl
import io

st.set_page_config(page_title="Excel Discrepancy Checker", layout="centered")
st.title("Excel Discrepancy Checker")

st.write("Upload two Excel files. Each must contain a sheet named 'ControlSheet' (case-insensitive). The app will compare key sales data using the provided rubric.")

uploaded_file1 = st.file_uploader("Upload Subway Excel file", type=["xlsx"], key="file1")
uploaded_file2 = st.file_uploader("Upload Sunthesis Excel file", type=["xlsx"], key="file2")

if uploaded_file1 and uploaded_file2:
    st.success("Both files uploaded. Processing...")
    
    def load_control_sheet(file):
        # Read all sheet names (case-insensitive match)
        xls = pd.ExcelFile(file, engine="openpyxl")
        control_sheet_name = None
        for sheet in xls.sheet_names:
            if sheet.lower() == "controlsheet":
                control_sheet_name = sheet
                break
        if control_sheet_name is None:
            return None, xls.sheet_names
        df = pd.read_excel(xls, sheet_name=control_sheet_name, header=None, engine="openpyxl")
        return df, xls.sheet_names

    subway_df, subway_sheets = load_control_sheet(uploaded_file1)
    sunthesis_df, sunthesis_sheets = load_control_sheet(uploaded_file2)

    if subway_df is None:
        st.error(f"'ControlSheet' not found in Subway file. Sheets found: {subway_sheets}")
    if sunthesis_df is None:
        st.error(f"'ControlSheet' not found in Sunthesis file. Sheets found: {sunthesis_sheets}")
    if subway_df is not None and sunthesis_df is not None:
        st.success("Both ControlSheet tabs loaded successfully.")

        # Define the row and column indices (0-based)
        # R = 17, S = 18, ..., X = 23 (Excel columns, 0-based index)
        col_start, col_end = 17, 23  # R:X
        label_rows = {
            "GROSS SALES": (12, 2, 8),            # C13:I13 (row 13, 0-based index 12, columns C=2 to I=8)
            "FOOT LONG SALES (30cm)": 8,        # R9:X9 (row 9, 0-based index 8)
            "=+SALAD/WRAPS SALES": 9,           # R10:X10 (row 10, 0-based index 9)
            "+6-INCH SALES (15cm)": 10,         # R11:X11 (row 11, 0-based index 10)
            "+ADD-ON SALES": 12,                # R13:X13 (row 13, 0-based index 12)
            "SALES TAX": 31,                    # R32:X32 (row 32, 0-based index 31)
            "ADJ. FOOT LONG UNITS (30cm)": 38,  # R39:X39 (row 39, 0-based index 38)
            "=+ADJ. SALAD/WRAPS UNITS": 39,     # R40:X40 (row 40, 0-based index 39)
            "+ADJ. 6-INCH UNITS (15cm)": 40,    # R41:X41 (row 41, 0-based index 40)
            "+FREE UNITS": 43                   # R44:X44 (row 44, 0-based index 43)
        }
        # Extract date/day headers for GROSS SALES (C11:I11 and C12:I12)
        gross_dates_subway = subway_df.iloc[10, 2:9].tolist()
        gross_days_subway = subway_df.iloc[11, 2:9].tolist()
        gross_dates_sunthesis = sunthesis_df.iloc[10, 2:9].tolist()
        gross_days_sunthesis = sunthesis_df.iloc[11, 2:9].tolist()

        # Extract date/day headers for all other labels (R7:X7 and R8:X8)
        dates_subway = subway_df.iloc[6, col_start:col_end+1].tolist()
        days_subway = subway_df.iloc[7, col_start:col_end+1].tolist()
        dates_sunthesis = sunthesis_df.iloc[6, col_start:col_end+1].tolist()
        days_sunthesis = sunthesis_df.iloc[7, col_start:col_end+1].tolist()

        # Prepare to extract label rows (handle missing rows gracefully)
        extracted_subway = {}
        extracted_sunthesis = {}
        for label, row_info in label_rows.items():
            if label == "GROSS SALES":
                row_idx, col_start_gross, col_end_gross = row_info
                extracted_subway[label] = subway_df.iloc[row_idx, col_start_gross:col_end_gross+1].tolist() if row_idx < len(subway_df) else None
                extracted_sunthesis[label] = sunthesis_df.iloc[row_idx, col_start_gross:col_end_gross+1].tolist() if row_idx < len(sunthesis_df) else None
            else:
                row_idx = row_info if isinstance(row_info, int) else row_info[0]
                extracted_subway[label] = subway_df.iloc[row_idx, col_start:col_end+1].tolist() if row_idx < len(subway_df) else None
                extracted_sunthesis[label] = sunthesis_df.iloc[row_idx, col_start:col_end+1].tolist() if row_idx < len(sunthesis_df) else None

        # Prepare comparison results
        results = []
        for label, row_info in label_rows.items():
            if label == "GROSS SALES":
                row_idx, col_start_gross, col_end_gross = row_info
                # Use Subway headers by default, fallback to Sunthesis if missing
                if all(pd.notnull(gross_dates_subway)) and all(pd.notnull(gross_days_subway)):
                    headers = [f"{str(day)} - {str(date)}" for date, day in zip(gross_dates_subway, gross_days_subway)]
                else:
                    headers = [f"{str(day)} - {str(date)}" for date, day in zip(gross_dates_sunthesis, gross_days_sunthesis)]
                row_subway = extracted_subway[label]
                row_sunthesis = extracted_sunthesis[label]
                for i, day_date in enumerate(headers):
                    entry = {
                        "Label": label,
                        "Day / Date": str(day_date),
                        "Subway Value": None,
                        "Sunthesis Value": None,
                        "Difference": None
                    }
                    if row_subway is None and row_sunthesis is None:
                        entry["Subway Value"] = "Label Missing"
                        entry["Sunthesis Value"] = "Label Missing"
                    elif row_subway is None:
                        entry["Subway Value"] = "Label Missing"
                        entry["Sunthesis Value"] = str(row_sunthesis[i]) if row_sunthesis else None
                    elif row_sunthesis is None:
                        entry["Subway Value"] = str(row_subway[i]) if row_subway else None
                        entry["Sunthesis Value"] = "Label Missing"
                    else:
                        val1 = row_subway[i]
                        val2 = row_sunthesis[i]
                        entry["Subway Value"] = str(val1)
                        entry["Sunthesis Value"] = str(val2)
                        try:
                            num1 = float(val1)
                            num2 = float(val2)
                            diff = num2 - num1
                            if num1 != num2:
                                entry["Difference"] = diff
                            else:
                                entry["Difference"] = 0
                        except (ValueError, TypeError):
                            entry["Difference"] = "Mismatch" if val1 != val2 else 0
                    results.append(entry)
            else:
                # Use Subway headers by default, fallback to Sunthesis if missing
                if all(pd.notnull(dates_subway)) and all(pd.notnull(days_subway)):
                    headers = [f"{str(day)} - {str(date)}" for date, day in zip(dates_subway, days_subway)]
                else:
                    headers = [f"{str(day)} - {str(date)}" for date, day in zip(dates_sunthesis, days_sunthesis)]
                row_idx = row_info if isinstance(row_info, int) else row_info[0]
                row_subway = extracted_subway[label]
                row_sunthesis = extracted_sunthesis[label]
                for i, day_date in enumerate(headers):
                    entry = {
                        "Label": label,
                        "Day / Date": str(day_date),
                        "Subway Value": None,
                        "Sunthesis Value": None,
                        "Difference": None
                    }
                    if row_subway is None and row_sunthesis is None:
                        entry["Subway Value"] = "Label Missing"
                        entry["Sunthesis Value"] = "Label Missing"
                    elif row_subway is None:
                        entry["Subway Value"] = "Label Missing"
                        entry["Sunthesis Value"] = str(row_sunthesis[i]) if row_sunthesis else None
                    elif row_sunthesis is None:
                        entry["Subway Value"] = str(row_subway[i]) if row_subway else None
                        entry["Sunthesis Value"] = "Label Missing"
                    else:
                        val1 = row_subway[i]
                        val2 = row_sunthesis[i]
                        entry["Subway Value"] = str(val1)
                        entry["Sunthesis Value"] = str(val2)
                        try:
                            num1 = float(val1)
                            num2 = float(val2)
                            diff = num2 - num1
                            if num1 != num2:
                                entry["Difference"] = diff
                            else:
                                entry["Difference"] = 0
                        except (ValueError, TypeError):
                            entry["Difference"] = "Mismatch" if val1 != val2 else 0
                    results.append(entry)
        results_df = pd.DataFrame(results)
        # Show only rows with discrepancies or missing labels
        display_df = results_df[(results_df['Difference'] != 0) | (results_df['Subway Value'] == 'Label Missing') | (results_df['Sunthesis Value'] == 'Label Missing')]
        st.subheader("Discrepancy Report")
        if not display_df.empty:
            st.dataframe(display_df, use_container_width=True)
        else:
            st.success("No discrepancies found between the selected ranges.")
        # Placeholder for export logic
else:
    st.info("Please upload both Excel files to proceed.") 

# Add custom CSS for green background on Subway file uploader
st.markdown(
    """
    <style>
    div[data-testid="stFileUploader"]:first-of-type > div:first-child {
        background-color: #2ecc40 !important;
        border-radius: 0.5rem;
        padding: 1.5rem 1rem;
    }
    </style>
    """,
    unsafe_allow_html=True
) 