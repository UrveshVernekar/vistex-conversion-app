import os
import zipfile
import pandas as pd
import script1
import script2


def auto_detect_script(input_file_path):
    """
    Inspects input Excel or CSV file to determine whether to run script1 (Upload Format -> Flat Output)
    or script2 (Flat Output -> Reverse Upload).
    """
    try:
        if input_file_path.lower().endswith('.csv'):
            df = pd.read_csv(input_file_path, header=None, nrows=10)
        else:
            excel_in = pd.ExcelFile(input_file_path)
            if not excel_in.sheet_names:
                return "script1"
            df = pd.read_excel(excel_in, sheet_name=0, header=None, nrows=10)
            
        if df.empty:
            return "script1"

        # Search top 10 rows for section tags typical of Upload Format (script1) vs Flat Output (script2)
        all_values = [str(val).strip() for val in df.values.flatten() if pd.notna(val)]
        
        # Script 2 output format usually has "Master Agreement External ID" in column 1 (row 2 or 3)
        # or flat 175-column template structure
        has_subagrmt_tag = any(val in ['SUBAGRMT', 'ELCUST', 'ELGBPRDGRP'] for val in all_values)
        has_agrmt_tag = 'AGRMT' in all_values
        
        # Check if first row has Request Header / Master Agreement Header (script2 input)
        row0_str = " ".join([str(val) for val in df.iloc[0].values if pd.notna(val)])
        row1_str = " ".join([str(val) for val in df.iloc[1].values if pd.notna(val)]) if len(df) > 1 else ""

        if "Request Header" in row0_str and "Master Agreement Header" in row0_str:
            return "script2"
        if "HEADER" in row1_str and "AGRMT" in row1_str and "SUBAGRMT" in row1_str:
            return "script2"

        # Upload Format (script1) has vertical block tags in data rows
        if has_subagrmt_tag and has_agrmt_tag:
            return "script1"

        # Check column 2 for tags
        if df.shape[1] > 2:
            col2_vals = [str(v).strip() for v in df.iloc[:, 2].dropna().values]
            if any(v in ['HEADER', 'AGRMT', 'SUBAGRMT', 'ELCUST'] for v in col2_vals):
                return "script1"

        return "script1"
    except Exception as e:
        print(f"Auto-detection fallback to script1 due to error: {e}")
        return "script1"


def process_conversion(input_file_path, output_dir, mode="auto"):
    """
    Processes the uploaded input Excel file using script1 or script2 based on mode.
    Generates both .xlsx and .csv files, packages them into a .zip archive, and returns metadata.
    """
    os.makedirs(output_dir, exist_ok=True)
    
    if mode == "auto":
        mode_used = auto_detect_script(input_file_path)
    elif mode in ["script1", "script2"]:
        mode_used = mode
    else:
        mode_used = "script1"

    base_name = os.path.splitext(os.path.basename(input_file_path))[0]
    
    if mode_used == "script1":
        mode_label = "Upload Format -> Flat Output (Script 1)"
        xlsx_filename = f"{base_name}_converted_output.xlsx"
        xlsx_path = os.path.join(output_dir, xlsx_filename)
        script1.convert_upload_to_output(input_upload_file=input_file_path, output_file=xlsx_path)
    else:
        mode_label = "Flat Output -> Reverse Upload Format (Script 2)"
        xlsx_filename = f"{base_name}_reverse_upload.xlsx"
        xlsx_path = os.path.join(output_dir, xlsx_filename)
        script2.convert_output_to_upload(input_output_file=input_file_path, output_reverse_file=xlsx_path)

    # Collect generated .xlsx and .csv files
    base_output_prefix = os.path.splitext(xlsx_path)[0]
    generated_files = []
    
    if os.path.exists(xlsx_path):
        generated_files.append(xlsx_path)
        
    # Look for .csv file(s) generated alongside xlsx
    parent_dir = os.path.dirname(xlsx_path)
    out_basename_no_ext = os.path.splitext(os.path.basename(xlsx_path))[0]
    
    csv_files = []
    for f in os.listdir(parent_dir):
        if f.startswith(out_basename_no_ext) and f.endswith(".csv"):
            csv_files.append(os.path.join(parent_dir, f))
            
    generated_files.extend(csv_files)

    # Package into ZIP archive for simultaneous download
    zip_filename = f"{out_basename_no_ext}_package.zip"
    zip_path = os.path.join(output_dir, zip_filename)
    
    with zipfile.ZipFile(zip_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
        for file_path in generated_files:
            zipf.write(file_path, arcname=os.path.basename(file_path))

    # Read sheet names for summary
    try:
        excel_out = pd.ExcelFile(xlsx_path)
        sheets = excel_out.sheet_names
    except Exception:
        sheets = ["Sheet1"]

    return {
        "mode_used": mode_used,
        "mode_label": mode_label,
        "xlsx_filename": os.path.basename(xlsx_path),
        "csv_filenames": [os.path.basename(c) for c in csv_files],
        "zip_filename": zip_filename,
        "sheets": sheets,
        "xlsx_path": xlsx_path,
        "csv_paths": csv_files,
        "zip_path": zip_path
    }
