import openpyxl
import os
import json
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import datetime
import pytz # Import pytz for timezone handling

# --- Core Data Extraction Functions (Unchanged) ---

def get_cell_data(cell_value, cell_formula):
    formula_str = ""
    if isinstance(cell_formula, str) and cell_formula.startswith('='):
        formula_str = cell_formula
    display_value = cell_value if cell_value is not None else '[empty]'
    return {"value": display_value, "formula": formula_str}

def read_worksheet_data(ws_values, ws_formulas):
    sheet_data = {}
    max_row = ws_formulas.max_row
    max_col = ws_formulas.max_column
    if max_row is None or max_col is None or max_row == 0 or max_col == 0:
         return {}

    for row_idx in range(1, max_row + 1):
        for col_idx in range(1, max_col + 1):
            cell_coord = f"{get_column_letter(col_idx)}{row_idx}"
            cell_formulas_obj = ws_formulas.cell(row=row_idx, column=col_idx)
            formula_content = cell_formulas_obj.value
            cell_values_obj = ws_values.cell(row=row_idx, column=col_idx)
            evaluated_value = cell_values_obj.value

            if evaluated_value is not None or (isinstance(formula_content, str) and formula_content.startswith('=')):
                 cell_info = get_cell_data(evaluated_value, formula_content)
                 sheet_data[cell_coord] = cell_info

    return sheet_data

def read_excel_file_data(filepath):
    try:
        # Added retries for reliability, especially if files are large or just created
        retries = 3
        delay = 0.5 # seconds
        for i in range(retries):
            try:
                wb_values = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
                wb_formulas = openpyxl.load_workbook(filepath, data_only=False, read_only=True)
                break # Success
            except Exception as e_inner:
                if i < retries - 1:
                    print(f"  Retrying read for {filepath} after error: {e_inner}")
                    import time
                    time.sleep(delay)
                else:
                    raise e_inner # Raise last error if all retries fail
    except FileNotFoundError:
        print(f"Error: File not found at '{filepath}'")
        return None
    except Exception as e:
        print(f"Error reading Excel file '{filepath}' after retries: {e}")
        return None

    all_excel_data = {}
    sheet_names = wb_formulas.sheetnames
    if not sheet_names:
        print(f"  Warning: No sheets found in '{filepath}'.")
        wb_values.close()
        wb_formulas.close()
        return {}

    for sheet_name in sheet_names:
        if sheet_name in wb_values.sheetnames:
            ws_values = wb_values[sheet_name]
            ws_formulas = wb_formulas[sheet_name]
            print(f"  Processing sheet: '{sheet_name}'...")
            sheet_content = read_worksheet_data(ws_values, ws_formulas)
            if sheet_content:
                 all_excel_data[sheet_name] = sheet_content
            else:
                 print(f"  Sheet '{sheet_name}' appears empty or contains no data/formulas. Adding as empty.")
                 all_excel_data[sheet_name] = {}
        else:
             print(f"  Warning: Sheet '{sheet_name}' found in formula view but not value view. Skipping.")

    wb_values.close()
    wb_formulas.close()
    return all_excel_data

# --- Comparison Function (Unchanged) ---

def compare_excel_data(data1, data2):
    if data1 is None or data2 is None:
        print("Error: Cannot compare None data.")
        return {"_metadata": {"error": "Input data missing"}}

    differences = {}
    sheets1 = set(data1.keys())
    sheets2 = set(data2.keys())

    sheets_only_in_file1 = list(sheets1 - sheets2)
    sheets_only_in_file2 = list(sheets2 - sheets1)
    common_sheets = list(sheets1 & sheets2)

    meta = {
        "sheets_common": sorted(common_sheets),
        "sheets_only_in_file1": sorted(sheets_only_in_file1),
        "sheets_only_in_file2": sorted(sheets_only_in_file2)
    }
    if meta["sheets_only_in_file1"] or meta["sheets_only_in_file2"] or common_sheets:
         differences["_metadata"] = meta

    for sheet_name in sorted(common_sheets):
        sheet_diff = {}
        sheet_data1 = data1.get(sheet_name, {})
        sheet_data2 = data2.get(sheet_name, {})
        all_cells = set(sheet_data1.keys()) | set(sheet_data2.keys())

        for cell_coord in sorted(list(all_cells)):
            default_cell = {"value": "[missing]", "formula": ""}
            cell_info1 = sheet_data1.get(cell_coord, default_cell)
            cell_info2 = sheet_data2.get(cell_coord, default_cell)

            values_differ = str(cell_info1["value"]) != str(cell_info2["value"])
            formulas_differ = cell_info1["formula"] != cell_info2["formula"]

            if values_differ or formulas_differ:
                sheet_diff[cell_coord] = {
                    "file1": sheet_data1.get(cell_coord, default_cell),
                    "file2": sheet_data2.get(cell_coord, default_cell)
                }

        if sheet_diff:
            differences[sheet_name] = sheet_diff

    return differences

# --- Demo File Creation (Unchanged) ---

def create_demo_excel_file(filepath, file_data):
    # --- (Function content remains exactly the same as before) ---
    wb = Workbook()
    if "Sheet" in wb.sheetnames:
        wb.remove(wb["Sheet"])

    for sheet_name, sheet_contents in file_data.items():
        ws = wb.create_sheet(title=sheet_name)
        if not sheet_contents:
             print(f"  Creating empty sheet '{sheet_name}' in '{filepath}'")
             continue
        for cell_coord, content in sheet_contents.items():
            value_to_write = content
            if isinstance(content, dict):
                 if 'formula' in content and content['formula']:
                      value_to_write = content['formula']
                 elif 'value' in content:
                      value_to_write = content['value'] if content['value'] != '[empty]' else None
                 else:
                      value_to_write = None
            elif content == '[empty]':
                 value_to_write = None
            ws[cell_coord] = value_to_write
    try:
        wb.save(filepath)
        print(f"Successfully created demo file: '{filepath}'")
    except Exception as e:
        print(f"Error creating demo file '{filepath}': {e}")
    wb.close()


# --- Text File Output Functions (Unchanged Function Logic, path handled in main) ---

def write_excel_data_to_txt(source_filepath, excel_data, output_filepath): # Changed arg name
    """Writes the structured Excel data to a human-readable text file at the specified path."""
    print(f"  Writing content dump to '{output_filepath}'...")
    try:
        # Ensure the directory exists (though it should have been created in main)
        os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
        with open(output_filepath, 'w', encoding='utf-8') as f:
            # --- (Rest of the function content remains exactly the same as before) ---
            f.write(f"Content Dump for Excel File: {source_filepath}\n")
            f.write(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 40 + "\n\n")
            if not excel_data:
                f.write("No sheets or data found in the file.\n")
                return
            sheet_names = sorted(excel_data.keys())
            f.write(f"Sheets found ({len(sheet_names)}): {', '.join(sheet_names)}\n\n")
            for sheet_name in sheet_names:
                f.write(f"--- Sheet: {sheet_name} ---\n")
                sheet_content = excel_data.get(sheet_name, {})
                if not sheet_content:
                    f.write("  [Sheet is empty or contains no tracked data]\n\n")
                    continue
                sorted_cells = sorted(sheet_content.keys(), key=lambda x: (int(''.join(filter(str.isdigit, x))), ''.join(filter(str.isalpha, x))))
                for cell_coord in sorted_cells:
                    cell_info = sheet_content[cell_coord]
                    value_str = cell_info.get('value', '[error retrieving value]')
                    formula_str = cell_info.get('formula', '')
                    f.write(f"  {cell_coord:<8}: Value = {value_str}")
                    if formula_str: f.write(f", Formula = {formula_str}")
                    f.write("\n")
                f.write("\n")
        print(f"  Successfully wrote content dump to '{output_filepath}'")
    except Exception as e:
        print(f"Error writing content dump to '{output_filepath}': {e}")

def write_comparison_summary_to_txt(file1_path, file2_path, comparison_results, output_filepath): # Changed arg name
    """Writes the comparison results to a human-readable text file at the specified path."""
    print(f"  Writing comparison summary to '{output_filepath}'...")
    try:
        # Ensure the directory exists
        os.makedirs(os.path.dirname(output_filepath), exist_ok=True)
        with open(output_filepath, 'w', encoding='utf-8') as f:
            # --- (Rest of the function content remains exactly the same as before) ---
            f.write("Excel File Comparison Summary\n")
            f.write(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write(f"File 1: {file1_path}\n")
            f.write(f"File 2: {file2_path}\n")
            f.write("=" * 40 + "\n\n")
            if not comparison_results:
                f.write("Comparison data is empty or invalid.\n")
                return
            meta = comparison_results.get("_metadata", {})
            if not meta and len(comparison_results) == 0:
                 f.write("No differences found between the files (including sheet structure).\n")
                 return
            elif not meta and len(comparison_results) > 0:
                 f.write("Comparison metadata missing, but differences found in sheets:\n")
                 common_sheets_inferred = sorted([k for k in comparison_results.keys() if k != "_metadata"])
                 meta = {"sheets_common": common_sheets_inferred, "sheets_only_in_file1": [], "sheets_only_in_file2": []}
            f.write("--- I. Sheet Structure Overview ---\n")
            common = meta.get("sheets_common", [])
            only1 = meta.get("sheets_only_in_file1", [])
            only2 = meta.get("sheets_only_in_file2", [])
            if common: f.write(f"Sheets with the same name in BOTH files ({len(common)}): {', '.join(common)}\n")
            else: f.write("No sheets found with the same name in both files.\n")
            if only1: f.write(f"Sheets found ONLY in File 1 ({len(only1)}): {', '.join(only1)}\n")
            else: f.write("No sheets found only in File 1.\n")
            if only2: f.write(f"Sheets found ONLY in File 2 ({len(only2)}): {', '.join(only2)}\n")
            else: f.write("No sheets found only in File 2.\n")
            f.write("\n")
            f.write("--- II. Detailed Cell Differences (in Common Sheets) ---\n")
            diff_sheets = sorted([sheet for sheet in comparison_results if sheet != "_metadata"])
            if not diff_sheets:
                if not only1 and not only2:
                     f.write("No differences found in cell values or formulas within common sheets.\n")
                     if common: f.write(f"(Compared sheets: {', '.join(common)})\n")
                else:
                     f.write("No common sheets to compare cells within, or common sheets were identical.\n")
                # Removed return here to ensure summary file is always created even if no cell diffs
            else:
                for sheet_name in diff_sheets:
                    sheet_diff_data = comparison_results[sheet_name]
                    f.write(f"\n--- Differences in Sheet: {sheet_name} ---\n")
                    sorted_cells = sorted(sheet_diff_data.keys(), key=lambda x: (int(''.join(filter(str.isdigit, x))), ''.join(filter(str.isalpha, x))))
                    for cell_coord in sorted_cells:
                        diff_info = sheet_diff_data[cell_coord]
                        info1 = diff_info.get('file1', {'value': '[error]', 'formula': ''})
                        info2 = diff_info.get('file2', {'value': '[error]', 'formula': ''})
                        f.write(f"  Cell: {cell_coord}\n")
                        f.write(f"    File 1: Value = {info1['value']}")
                        if info1['formula']: f.write(f", Formula = {info1['formula']}")
                        f.write("\n")
                        f.write(f"    File 2: Value = {info2['value']}")
                        if info2['formula']: f.write(f", Formula = {info2['formula']}")
                        f.write("\n\n")
        print(f"  Successfully wrote comparison summary to '{output_filepath}'")
    except Exception as e:
        print(f"Error writing comparison summary to '{output_filepath}': {e}")

# --- Main Execution & Demonstration ---

def main():
    print("Starting Excel Comparison Utility Demo...")

    # --- Define Base Filenames ---
    # Excel files remain in root
    file1_name = "demo_excel_1.xlsx"
    file2_name = "demo_excel_2.xlsx"
    # Base names for text files
    file1_txt_dump_base = "demo_excel_1_contents.txt"
    file2_txt_dump_base = "demo_excel_2_contents.txt"
    comparison_summary_txt_base = "comparison_summary.txt"

    # --- Create Timestamped Output Directory ---
    try:
        jakarta_tz = pytz.timezone('Asia/Jakarta')
        now_jakarta = datetime.datetime.now(jakarta_tz)
        # Format timestamp suitable for directory names (YYYYMMDD_HHMMSS)
        timestamp_str = now_jakarta.strftime('%Y%m%d_%H%M%S')
        output_dir_name = f"output_{timestamp_str}"
        # Create the directory (exist_ok=True prevents error if it already exists)
        os.makedirs(output_dir_name, exist_ok=True)
        print(f"Created output directory: '{output_dir_name}'")

        # --- Construct Full Paths for Output Text Files ---
        file1_txt_dump_path = os.path.join(output_dir_name, file1_txt_dump_base)
        file2_txt_dump_path = os.path.join(output_dir_name, file2_txt_dump_base)
        comparison_summary_txt_path = os.path.join(output_dir_name, comparison_summary_txt_base)

    except Exception as e:
        print(f"Error creating output directory or paths: {e}")
        print("Text outputs will be saved in the current directory instead.")
        # Fallback to current directory if error occurs
        output_dir_name = "." # Use current directory
        file1_txt_dump_path = file1_txt_dump_base
        file2_txt_dump_path = file2_txt_dump_base
        comparison_summary_txt_path = comparison_summary_txt_base


    # --- Demo Data Structure Definitions (Unchanged) ---
    file1_data_structure = {
        "DataSheet": { # Common sheet, will have differences
            "A1": {"value": "Name", "formula": ""}, "B1": {"value": "Value1", "formula": ""}, "C1": {"value": "Value2", "formula": ""}, "D1": {"value": "Total", "formula": ""},
            "A2": {"value": "Item A", "formula": ""}, "B2": {"value": 10, "formula": ""}, "C2": {"value": 20, "formula": ""}, "D2": {"value": 30, "formula": "=B2+C2"},
            "A3": {"value": "Item B", "formula": ""}, "B3": {"value": 15, "formula": ""}, "C3": {"value": 25, "formula": ""}, "D3": {"value": 40, "formula": "=SUM(B3:C3)"},
            "E5": {"value": "[empty]", "formula":""}
        },
        "Summary": { # Sheet only in File 1
            "A1": {"value": "Grand Total", "formula": ""}, "B1": {"value": 70, "formula": "=DataSheet!D2+DataSheet!D3"}, "C5": {"value": "Only in File 1 Sheet", "formula": ""}
        },
        "IdenticalSheet": { # Common sheet, should have no differences
             "A1": {"value": "Status", "formula": ""}, "B1": {"value": "OK", "formula": ""}
        },
        "EmptySheet1": {} # Empty sheet only in File 1
    }
    file2_data_structure = {
        "DataSheet": { # Common sheet, will have differences
            "A1": {"value": "Name", "formula": ""}, "B1": {"value": "Value One", "formula": ""}, "C1": {"value": "Value2", "formula": ""}, "D1": {"value": "Total Sum", "formula": ""},
            "A2": {"value": "Item A", "formula": ""}, "B2": {"value": 10.0, "formula": ""}, "C2": {"value": 20, "formula": ""}, "D2": {"value": 30, "formula": "=B2+C2"},
            "A3": {"value": "Item B New", "formula": ""}, "B3": {"value": 15, "formula": ""}, "C3": {"value": 99, "formula": ""}, "D3": {"value": 114, "formula": "=B3+C3"},
            "F6": {"value": "Extra Cell F6", "formula": ""}
        },
        "Details": { # Sheet only in File 2
            "X1": {"value": "Detail Info", "formula": ""}, "Y10": {"value": 12345, "formula": ""}
        },
        "IdenticalSheet": { # Common sheet, should have no differences
             "A1": {"value": "Status", "formula": ""}, "B1": {"value": "OK", "formula": ""}
        }
    }


    # 1. Create Demo Files (in root)
    print("\n--- Creating Demo Files ---")
    create_demo_excel_file(file1_name, file1_data_structure)
    create_demo_excel_file(file2_name, file2_data_structure)

    # 2. Read Data from Files
    print(f"\n--- Reading Data from '{file1_name}' ---")
    excel1_contents = read_excel_file_data(file1_name)

    print(f"\n--- Reading Data from '{file2_name}' ---")
    excel2_contents = read_excel_file_data(file2_name)

    # 3. Write Individual Content Dumps (to output directory)
    print("\n--- Writing Content Dumps ---")
    if excel1_contents is not None:
        # Pass the full path including the directory
        write_excel_data_to_txt(file1_name, excel1_contents, file1_txt_dump_path)
    if excel2_contents is not None:
        # Pass the full path including the directory
        write_excel_data_to_txt(file2_name, excel2_contents, file2_txt_dump_path)

    # 4. Compare the Data
    print("\n--- Comparing Files ---")
    comparison_results = compare_excel_data(excel1_contents, excel2_contents)

    # 5. Write Comparison Summary Report (to output directory)
    print("\n--- Writing Comparison Summary ---")
    # Pass the full path including the directory
    write_comparison_summary_to_txt(file1_name, file2_name, comparison_results, comparison_summary_txt_path)


    # 6. Cleanup Demo Files
    print("\n--- Cleaning Up Demo Files ---")
    # List files to clean, including the text files *inside* the output directory
    files_to_clean = [
        file1_name,                     # Excel file in root
        file2_name,                     # Excel file in root
        file1_txt_dump_path,            # Text file in output dir
        file2_txt_dump_path,            # Text file in output dir
        comparison_summary_txt_path     # Text file in output dir
    ]
    for f_path in files_to_clean:
        try:
            if os.path.exists(f_path):
                os.remove(f_path)
                print(f"Removed '{f_path}'")
            # else: # Optional: print if file was already gone
            #    print(f"File '{f_path}' not found for cleanup.")
        except Exception as e:
            print(f"Error during cleanup of '{f_path}': {e}")

    # Optionally remove the now empty output directory
    try:
        # Check if it exists and is a directory before attempting removal
        if output_dir_name != "." and os.path.exists(output_dir_name) and os.path.isdir(output_dir_name):
             # Check if directory is empty
             if not os.listdir(output_dir_name):
                  os.rmdir(output_dir_name)
                  print(f"Removed empty output directory '{output_dir_name}'")
             else:
                  print(f"Output directory '{output_dir_name}' not empty, not removed.")
    except Exception as e:
        print(f"Error removing output directory '{output_dir_name}': {e}")


    print("\nDemo finished.")

# --- Script Entry Point ---
if __name__ == "__main__":
    # Check for required libraries
    try:
        import openpyxl
    except ImportError:
        print("Error: The 'openpyxl' library is required but not installed.")
        print("Please install it using: pip install openpyxl")
        exit(1)
    try:
        import pytz
    except ImportError:
        print("Error: The 'pytz' library is required but not installed.")
        print("Please install it using: pip install pytz")
        exit(1)

    main()