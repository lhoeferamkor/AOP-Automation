import win32com.client
import os
import pythoncom # For CoInitialize (sometimes needed in threads or complex scenarios)
import pandas as pd

# --- Configuration ---
# Define the column headers to be set in row 2
NEW_HEADERS = ["Customer", "Sold-To", "Legal Name", "Pkg", "Plant", "PDL"]
# Assuming "Platn" is the "Plant" column, which is the 5th column (E)
PLANT_COLUMN_LETTER = 'E' # After deletions and header setting, this is E2, E3, etc.
PLANT_COLUMN_INDEX = 5 # 1-based index for Excel, corresponds to 'E'

# Keywords for red highlighting in the "Plant" column
RED_PLANT_KEYWORDS_SET1 = ["MEMORY", "SIP", "FPS", "MOLDED MEMS"] # Case-insensitive search
RED_PLANT_KEYWORDS_SET2 = ["CABGA", "BGA", "SCSP"] # Case-insensitive search

# RGB Colors
COLOR_RED = 255         # R=255, G=0, B=0
COLOR_YELLOW = 65535    # R=255, G=255, B=0
COLOR_GREEN = 5287936   # R=0, G=255, B=0 (Excel has different ways to represent this)
# Or use RGB values:
# COLOR_RED_RGB = (255, 0, 0)
# COLOR_YELLOW_RGB = (255, 255, 0)
# COLOR_GREEN_RGB = (0, 128, 0) # Darker green often looks better

def process_excel_file(input_filepath, output_filepath):
    """
    Processes an Excel 97-2003 Worksheet (.xls) according to specified rules.
    """
    # Ensure absolute paths
    input_filepath = os.path.abspath(input_filepath)
    output_filepath = os.path.abspath(output_filepath)

    if not os.path.exists(input_filepath):
        print(f"Error: Input file not found at {input_filepath}")
        return

    excel = None
    workbook = None

    try:
        pythoncom.CoInitialize() # Initialize COM for the current thread

        excel = win32com.client.DispatchEx("Excel.Application") # DispatchEx for better isolation
        excel.Visible = False  # Run in background. Set to True for debugging.
        excel.DisplayAlerts = False # Suppress alerts like "Save changes?"

        workbook = excel.Workbooks.Open(input_filepath)
        # Assuming we work on the first sheet. Modify if needed.
        sheet = workbook.Worksheets(1)
        sheet.Activate() # Make sure it's the active sheet

        print("Step 1: Deleting the top five rows...")
        sheet.Rows("1:5").Delete(Shift=-4162) # xlUp

        print("Step 2: Deleting specified columns (1st, 3rd, 5th, 6th, 8th, 10th)...")
        # Columns to delete (1-based index). Delete from right to left to avoid index shifting.
        # Original: A(1), C(3), E(5), F(6), H(8), J(10)
        # Convert to actual column letters/numbers for deletion
        cols_to_delete_indices = [10, 8, 6, 5, 3, 1] # Delete J, then H, then F, E, C, A
        for col_idx in cols_to_delete_indices:
            print(f"  Deleting original column {col_idx}...")
            sheet.Columns(col_idx).Delete()

        print("Step 3: Removing values from A1:F2 (after column deletions)...")
        sheet.Range("A1:F2").ClearContents()

        print("Step 4: Filling in header values in A2:F2...")
        for i, header_text in enumerate(NEW_HEADERS):
            sheet.Cells(2, i + 1).Value = header_text # Cells are 1-indexed (row, col)

        # --- REMOVAL STAGE (Highlighting) ---
        print("Step 5 & 6: Highlighting rows red based on 'Plant' column keywords...")
        last_row = sheet.Cells(sheet.Rows.Count, PLANT_COLUMN_INDEX).End(-4162).Row # xlUp from last cell
        
        # Store which rows were highlighted red for step 7
        red_highlighted_rows = set()

        for r in range(3, last_row + 1): # Start from row 3 (data starts below headers)
            plant_cell = sheet.Cells(r, PLANT_COLUMN_INDEX)
            plant_value = str(plant_cell.Value).upper() if plant_cell.Value else ""

            is_red_row = False
            # Check Set 1
            for keyword in RED_PLANT_KEYWORDS_SET1:
                if keyword.upper() in plant_value:
                    is_red_row = True
                    break
            # Check Set 2 (if not already found in Set 1)
            if not is_red_row:
                for keyword in RED_PLANT_KEYWORDS_SET2:
                    if keyword.upper() in plant_value:
                        is_red_row = True
                        break
            
            if is_red_row:
                # Highlight entire row red
                sheet.Rows(r).Interior.Color = COLOR_RED
                red_highlighted_rows.add(r)


        print("Step 7: Highlighting 'test' cells yellow/green based on preceding red row...")
        # Determine the last used column dynamically
        # Find the last column in the header row (row 2)
        last_col = sheet.Cells(2, sheet.Columns.Count).End(-4159).Column # xlToLeft from last cell in row 2

        for r in range(3, last_row + 1): # Start from row 3
            if r%30 == 0:
                print(r)
            if r > 120:
                break
            # Check if the *previous* row (r-1) was highlighted red
            previous_row_was_red = (r - 1) in red_highlighted_rows

            for c in range(1, last_col + 1): # Iterate through columns A to last_col
                current_cell = sheet.Cells(r, c)
                cell_value = str(current_cell.Value).lower() if current_cell.Value else ""

                if "test" in cell_value:
                    if previous_row_was_red:
                        current_cell.Interior.Color = COLOR_YELLOW
                    else:
                        current_cell.Interior.Color = COLOR_GREEN

        print(f"Saving processed file to: {output_filepath}")
        # Save As to avoid modifying original, and to ensure it's saved in a potentially newer format if desired
        # For saving as old .xls format (Excel 97-2003 Workbook):
        # FileFormat=56 is for xlExcel8 or Excel97-2003 Workbook (.xls)
        # If you want to save as .xlsx, use FileFormat=51 (xlOpenXMLWorkbook)
        try:
            workbook.SaveAs(output_filepath, FileFormat=51) # 56 for .xls
            print("Processing complete.")
        except Exception as e:
            print(f"An error occured: {e}")

    except Exception as e:
        print(f"An error occurred: {e}")
        import traceback
        traceback.print_exc()
    finally:
        if workbook:
            workbook.Close(SaveChanges=False) # Close without saving again (already saved with SaveAs)
        if excel:
            excel.Quit()
        # Release COM objects
        sheet = None
        workbook = None
        excel = None
        pythoncom.CoUninitialize() # Uninitialize COM for the current thread

# --- Main execution ---
if __name__ == "__main__":
    actual_input_file = "input_data/test_excel.xlsx"
    actual_output_file = "output_data/AOP Sales Attained.xlsx"
    if os.path.exists(actual_input_file):
        process_excel_file(actual_input_file, actual_output_file)
    else:
        print(f"Actual input file not found: {actual_input_file}")