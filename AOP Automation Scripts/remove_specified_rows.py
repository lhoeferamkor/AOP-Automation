import pandas as pd
from openpyxl import load_workbook # Not strictly needed for writing new, but good to know
from openpyxl.styles import PatternFill, Font
from openpyxl.utils.dataframe import dataframe_to_rows # Efficiently write df to openpyxl
import re # For case-insensitive "test" matching
import os

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
    df = df.drop(index=remove_index).reset_index(drop=True)

    return remove_index, df 


def apply_conditional_formatting(input_excel_path, output_excel_path, task='remove', column_name="Unnamed: 2", sheet_name="CleanedData"):
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

    # Determine if file exists
    file_exists = os.path.exists(output_excel_path)

    if task == 'both':
        print("Running Both Commands")
        # Get both DataFrames and formatting instructions
        remove_index, remove_df = remove_rows(column_name, df)
        remove_df.iloc[0,0:4]=['Date', 'Legal Name', 'Pkg', 'PDL']
        remove_df.rename(columns={'Unnamed: 0' : ' ', 'Unnamed: 1' : ' ', 'Unnamed: 2' : ' ', 'Unnamed: 3' : ' '}, inplace=True)
        cell_formats_to_apply, highlight_df = highlight_rows(column_name, df)
        try:
            if file_exists:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
            else:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w')
            with writer:
                highlight_df.to_excel(writer, sheet_name='highlight', index=False)
                remove_df.to_excel(writer, sheet_name='remove', index=False)
            wb = load_workbook(output_excel_path)
            ws = wb['highlight']
            applied_fills = {}
            for fmt_instruction in cell_formats_to_apply:
                cell_coord = (fmt_instruction['row'], fmt_instruction['column'])
                if fmt_instruction.get('overwrite_previous_for_cell') or cell_coord not in applied_fills:
                    cell = ws.cell(row=fmt_instruction['row'], column=fmt_instruction['column'])
                    cell.fill = fmt_instruction['fill']
                    applied_fills[cell_coord] = fmt_instruction['fill']
            wb.save(output_excel_path)
            print(f"Successfully wrote both sheets and applied formatting to '{output_excel_path}'")
        except Exception as e:
            print(f"Error writing Excel file or applying styles: {e}")
            import traceback
            traceback.print_exc()
        return

    if task == 'remove':
        print("running remove command")
        remove_index, remove_df = remove_rows(column_name, df)
        remove_df.iloc[0,0:4]=['Date', 'Legal Name', 'Pkg', 'PDL']
        remove_df.rename(columns={'Unnamed: 0' : ' ', 'Unnamed: 1' : ' ', 'Unnamed: 2' : ' ', 'Unnamed: 3' : ' '}, inplace=True)
        try:
            if file_exists:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
            else:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w')
            with writer:
                remove_df.to_excel(writer, sheet_name='remove', index=False)
            print(f"Successfully updated 'remove' sheet in '{output_excel_path}'")
        except Exception as e:
            print(f"Error writing Excel file: {e}")
            import traceback
            traceback.print_exc()

    if task == 'highlight':
        print("running Highlight Command")
        cell_formats_to_apply, highlight_df = highlight_rows(column_name, df)
        try:
            if file_exists:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='a', if_sheet_exists='replace')
            else:
                writer = pd.ExcelWriter(output_excel_path, engine='openpyxl', mode='w')
            with writer:
                highlight_df.to_excel(writer, sheet_name='highlight', index=False)
            wb = load_workbook(output_excel_path)
            ws = wb['highlight']
            applied_fills = {}
            for fmt_instruction in cell_formats_to_apply:
                cell_coord = (fmt_instruction['row'], fmt_instruction['column'])
                if fmt_instruction.get('overwrite_previous_for_cell') or cell_coord not in applied_fills:
                    cell = ws.cell(row=fmt_instruction['row'], column=fmt_instruction['column'])
                    cell.fill = fmt_instruction['fill']
                    applied_fills[cell_coord] = fmt_instruction['fill']
            wb.save(output_excel_path)
            print(f"Successfully updated and formatted 'highlight' sheet in '{output_excel_path}'")
        except Exception as e:
            print(f"Error writing Excel file or applying styles: {e}")
            import traceback
            traceback.print_exc()



# --- Example Usage ---
if __name__ == "__main__":
    input_file = "AOP Automation Scripts/output_data/converted_report.xlsx"
    output_file = "AOP Automation Scripts/output_data/formatted_report.xlsx"
    apply_conditional_formatting(input_file, output_file, task = "remove")