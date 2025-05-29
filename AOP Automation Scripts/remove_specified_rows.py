import pandas as pd
from openpyxl import load_workbook # Not strictly needed for writing new, but good to know
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows # Efficiently write df to openpyxl
import re # For case-insensitive "test" matching


def highlight_rows(column_name : str, df : pd.DataFrame):
    
    # Define keywords for red highlighting
    red_keywords_group = ["MEMORY", "SIP", "FPS", "Molded MEMS"] # Case-sensitive as per examples
    green_keywords_group = ["CABGA", "BGA", "SCSP"]              # Case-sensitive

    # Define fills
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Light Red
    yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # Light Yellow
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Light Green


    # --- Logic to determine cell colors ---
    # We'll store color instructions: (row_index, col_index, fill_object)
    cell_formats_to_apply = []
    remove_activated = False
    plant_col_idx = df.columns.get_loc(column_name) + 1 # 1-based index for openpyxl

    for r_idx, row in df.iterrows():
        # openpyxl rows are 1-based, and ExcelWriter usually writes headers, so add 2
        # (1 for 1-based, 1 for header row)
        excel_row_num = r_idx + 2
        current_row_override = False
        cell_value_plant = str(row[column_name]) # Ensure it's a string

        # --- Rule 1 & 2: Red Highlighting for whole row based on 'Plant' column ---
        is_red = any(re.search(keyword, cell_value_plant, re.IGNORECASE) for keyword in red_keywords_group)
        is_green = any(re.search(keyword, cell_value_plant, re.IGNORECASE) for keyword in green_keywords_group)
        
        #RED Rule
        if is_red:
            remove_activated = True
            current_row_override = True
            for c_idx in range(1, len(df.columns) + 1): # Apply to all cells in the row
                cell_formats_to_apply.append({'row': excel_row_num, 'column': c_idx, 'fill': red_fill})

        #Green Rule
        elif is_green:
            remove_activated = False
            current_row_override = True
            for c_idx in range(1, len(df.columns) + 1):
                cell_formats_to_apply.append({'row': excel_row_num, 'column': c_idx, 'fill': green_fill})

        if re.search(r'test', cell_value_plant, re.IGNORECASE):
            if remove_activated and not current_row_override:
                for c_idx in range(1, len(df.columns) + 1):
                    cell_formats_to_apply.append({'row': excel_row_num, 'column': c_idx, 'fill': yellow_fill, 'overwrite_previous_for_cell': True})
            elif not remove_activated:
                for c_idx in range(1, len(df.columns) + 1):
                    cell_formats_to_apply.append({'row': excel_row_num, 'column': c_idx, 'fill': green_fill, 'overwrite_previous_for_cell': True})
            elif current_row_override:
                print(f"ERROR! Override in the Test on row {excel_row_num}")

    return cell_formats_to_apply, df

def remove_rows(column_name : str, df : pd.DataFrame):
    red_keywords_group = ["MEMORY", "SIP", "FPS", "Molded MEMS"] # Case-sensitive as per examples
    green_keywords_group = ["CABGA", "BGA", "SCSP"]              # Case-sensitive

    # Define fills
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Light Red
    yellow_fill = PatternFill(start_color="FFFFCC", end_color="FFFFCC", fill_type="solid") # Light Yellow
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid") # Light Green


    # --- Logic to determine cell colors ---
    # We'll store color instructions: (row_index, col_index, fill_object)
    remove_activated = False
    remove_index = []

    for r_idx, row in df.iterrows():
        # openpyxl rows are 1-based, and ExcelWriter usually writes headers, so add 2
        # (1 for 1-based, 1 for header row)
        excel_row_num = r_idx + 2
        current_row_override = False
        cell_value_plant = str(row[column_name]) # Ensure it's a string

        # --- Rule 1 & 2: Red Highlighting for whole row based on 'Plant' column ---
        is_red = any(re.search(keyword, cell_value_plant, re.IGNORECASE) for keyword in red_keywords_group)
        is_green = any(re.search(keyword, cell_value_plant, re.IGNORECASE) for keyword in green_keywords_group)
            
        if is_red:
            remove_activated = True
            current_row_override = True
            remove_index.append(r_idx)

        elif is_green:
            remove_activated = False
            current_row_override = True
        else:
            if re.search(r'test', cell_value_plant, re.IGNORECASE):
                if remove_activated and not current_row_override:
                    remove_index.append(r_idx)
                elif not remove_activated:
                    pass
                elif current_row_override:
                    print(f"ERROR! Override in the Test on row {excel_row_num}")

    #Flip the remove index and then remove unecessary rows from Dataframe                 
    remove_index.reverse()
    print(df.shape)
    df = df.drop(index=remove_index).reset_index(drop=True)
    print(df.shape)

    return remove_index, df 


def apply_conditional_formatting(input_excel_path, output_excel_path, highlight_mode = False, remove_mode = True, column_name="Unnamed: 2", sheet_name=0):
    """
    Applies conditional highlighting to an Excel sheet based on complex rules.

    Args:
        input_excel_path (str): Path to the input Excel file.
        output_excel_path (str): Path to save the output Excel file with formatting.
        column_name (str): The name of the column to check for keywords.
        sheet_name (int or str): Sheet index (0-based) or name to process.
    """
    try:
        df = pd.read_excel(input_excel_path, sheet_name=sheet_name, keep_default_na=False)
    except FileNotFoundError:
        print(f"Error: Input file '{input_excel_path}' not found.")
        return
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return

    if column_name not in df.columns:
        print(f"Error: Column '{column_name}' not found in the sheet.")
        print(f"Available columns are: {df.columns.tolist()}")
        return

    cell_formats_to_apply = []
    remove_activated = False
    remove_index = []
    #remove_index, df = remove_rows(column_name, df)
    cell_formats_to_apply, df = highlight_rows(column_name, df)

    #Reconfigure excel headers: 
    df.iloc[0,0:4]=['Date', 'Legal Name', 'Pkg', 'PDL']
    df.rename(columns={'Unnamed: 0' : ' ', 'Unnamed: 1' : ' ', 'Unnamed: 2' : ' ', 'Unnamed: 3' : ' '}, inplace=True)

    # --- Write to Excel and Apply Formatting ---
    try:
        # Create a new workbook or overwrite if output_excel_path is same as input
        # It's often safer to write to a new file first.
        df.to_excel(output_excel_path, sheet_name=sheet_name, index=False, engine='openpyxl')

        # Now, load the workbook we just wrote to apply styles
        wb = load_workbook(output_excel_path)
        if isinstance(sheet_name, int):
            ws = wb.worksheets[sheet_name]
        else:
            ws = wb[sheet_name]

        # Apply stored formats, handling potential overwrites for the specific yellow "test" cell
        applied_fills = {} # To keep track of what's applied to a cell to handle overwrites

        for fmt_instruction in cell_formats_to_apply:
            cell_coord = (fmt_instruction['row'], fmt_instruction['column'])
            
            if fmt_instruction.get('overwrite_previous_for_cell') or cell_coord not in applied_fills:
                cell = ws.cell(row=fmt_instruction['row'], column=fmt_instruction['column'])
                cell.fill = fmt_instruction['fill']
                applied_fills[cell_coord] = fmt_instruction['fill']
            # else: cell already has a fill from a row-level red rule and this rule doesn't overwrite

        wb.save(output_excel_path)
        print(f"Successfully applied formatting and saved to '{output_excel_path}'")

    except Exception as e:
        print(f"Error writing Excel file or applying styles: {e}")
        import traceback
        traceback.print_exc()


# --- Example Usage ---
if __name__ == "__main__":
    input_file = "AOP Automation Scripts/output_data/converted_report.xlsx"
    output_file = "AOP Automation Scripts/output_data/formatted_report.xlsx"
    apply_conditional_formatting(input_file, output_file, highlight_mode = False, remove_mode=True, column_name="Unnamed: 2", sheet_name="CleanedData")