import pandas as pd
import os
from email import message_from_bytes
from email.parser import BytesParser
from email.policy import default as default_policy
import re
import io # For using StringIO with pd.read_html

# --- Configuration ---
# Ensure these are installed:
# pip install pandas lxml html5lib openpyxl (for .xlsx) xlwt (for .xls)

def extract_html_from_mhtml(mhtml_file_path):
    """
    Extracts the primary HTML content from an MHTML file.
    """
    try:
        with open(mhtml_file_path, 'rb') as fp:
            headers = BytesParser(policy=default_policy).parse(fp)
            for part in headers.walk():
                if part.get_content_type() == 'text/html':
                    charset = part.get_content_charset() or 'utf-8'
                    html_content = part.get_payload(decode=True)
                    try:
                        return html_content.decode(charset)
                    except UnicodeDecodeError:
                        try:
                            return html_content.decode('utf-8')
                        except UnicodeDecodeError:
                            return html_content.decode('latin-1', errors='replace')
            print(f"No text/html part found in '{mhtml_file_path}'.")
            return None
    except FileNotFoundError:
        print(f"Error: MHTML file not found at '{mhtml_file_path}'.")
        return None
    except Exception as e:
        print(f"Error reading or parsing MHTML file '{mhtml_file_path}': {e}")
        return None

def clean_and_convert_numeric(df):
    """
    Attempts to clean and convert columns in a DataFrame to numeric types.
    """
    if df.empty:
        return df # Return empty if nothing to process
    df_converted = df.copy()
    for col in df_converted.columns:
        if df_converted[col].dtype == 'object':
            try:
                cleaned_series = df_converted[col].astype(str).str.strip()
                is_negative = cleaned_series.str.startswith('(') & cleaned_series.str.endswith(')')
                cleaned_series = cleaned_series.str.replace(r'[\(\)]', '', regex=True)
                cleaned_series = cleaned_series.str.replace(r'[$,%]', '', regex=True)
                cleaned_series = cleaned_series.str.replace(r',(?=\d{3})', '', regex=True)
                numeric_series = pd.to_numeric(cleaned_series, errors='coerce')

                if not numeric_series.isnull().all() or not cleaned_series.empty: # Check if any conversion happened or if series wasn't empty
                    numeric_series.loc[is_negative[is_negative].index] = -numeric_series.loc[is_negative[is_negative].index] # Apply negative carefully
                    df_converted[col] = numeric_series
                    print(f"Column '{col}' processed for numeric conversion.")
                else:
                    print(f"Column '{col}' kept as object (all values NaN after trying or empty).")
            except Exception as e:
                print(f"Could not convert column '{col}' to numeric: {e}. Kept as object.")
    return df_converted

def convert_mhtml_to_excel(mhtml_file_path, output_excel_path, sheet_name="Sheet1"):
    """
    Extracts tables from an MHTML file, preserves top info rows, converts the main table
    to a DataFrame, attempts numeric conversion, and saves it all to an Excel file.
    """
    if not os.path.exists(mhtml_file_path):
        print(f"Error: Input MHTML file not found at '{mhtml_file_path}'")
        return False

    base_name = os.path.basename(mhtml_file_path)
    print(f"Processing MHTML file: '{base_name}'")
    html_string = extract_html_from_mhtml(mhtml_file_path)

    if html_string:
        print(f"Successfully extracted HTML content from '{base_name}'.")
        try:
            print("Attempting to parse tables with pd.read_html (header=None)...")
            # Use io.StringIO to pass the HTML string directly
            # header=None ensures all rows are read as data initially
            list_of_dfs = pd.read_html(io.StringIO(html_string), header=None, thousands=',', decimal='.')

            if not list_of_dfs:
                print(f"No tables found in the HTML content of '{base_name}'.")
                return False

            # Assume the largest table is the target one
            df_full = max(list_of_dfs, key=lambda d: d.size if d.size > 0 else -1)

            if df_full.empty:
                print(f"The largest table found in '{base_name}' is empty.")
                # Optionally, write an empty Excel file or just return False
                # For now, let's try to write what we have, even if it's just info rows
                # return False

            # --- Logic to separate info rows, data headers, and data body ---
            num_total_rows = df_full.shape[0]
            num_info_rows = 2
            num_data_header_rows = 1 # The row immediately after info rows

            info_row_dfs_to_write = [] # List to hold DataFrames for each info row

            for i in range(min(num_info_rows, num_total_rows)):
                # Create a single-row DataFrame for each info row
                # Fill NaN with empty strings for cleaner output in info rows
                info_row_series = df_full.iloc[i].fillna('')
                info_row_dfs_to_write.append(pd.DataFrame([info_row_series.tolist()]))

            data_headers_row_index = num_info_rows
            actual_data_start_index = data_headers_row_index + num_data_header_rows
            main_data_df = pd.DataFrame() # Initialize as empty

            if num_total_rows > data_headers_row_index: # Enough rows for data headers
                # Get data headers from the row after info rows
                data_column_headers = df_full.iloc[data_headers_row_index].fillna('').tolist()

                if num_total_rows >= actual_data_start_index: # Actual data rows exist
                    main_data_df = df_full.iloc[actual_data_start_index:].copy()
                else: # Only header row for data, no data rows themselves
                    main_data_df = pd.DataFrame(columns=data_column_headers) # Empty DF with these headers
                main_data_df.columns = data_column_headers
                main_data_df.reset_index(drop=True, inplace=True)
            else:
                print(f"Warning: Table has {num_total_rows} rows. Not enough rows for data headers after {num_info_rows} info rows.")
                # main_data_df remains empty

            # Clean and convert numeric types for the main_data_df
            if not main_data_df.empty or (not main_data_df.columns.empty and main_data_df.shape[0] == 0):
                print("Attempting to clean and convert numeric data for the main table...")
                df_cleaned_data = clean_and_convert_numeric(main_data_df)
            else:
                print("No main data body to process or clean.")
                df_cleaned_data = main_data_df # Use as is (likely empty)


            # --- Saving to Excel ---
            try:
                excel_engine = None
                if output_excel_path.lower().endswith('.xlsx'):
                    excel_engine = 'openpyxl'
                elif output_excel_path.lower().endswith('.xls'):
                    excel_engine = 'xlwt' # Note: xlwt has limitations (e.g., >256 cols, specific formats)
                else:
                    print(f"Warning: Output file '{output_excel_path}' has an unrecognized Excel extension. Defaulting to .xlsx.")
                    output_excel_path += ".xlsx"
                    excel_engine = 'openpyxl'

                with pd.ExcelWriter(output_excel_path, engine=excel_engine) as writer:
                    current_excel_row = 0
                    # Write info rows
                    for i, df_info in enumerate(info_row_dfs_to_write):
                        print(f"Writing info row {i+1} to Excel...")
                        df_info.to_excel(writer, sheet_name=sheet_name,
                                         startrow=current_excel_row,
                                         header=False, index=False)
                        current_excel_row += df_info.shape[0] # Should be 1

                    # Write the main data DataFrame (with its own headers)
                    if not df_cleaned_data.columns.empty: # Check if there are columns to write
                        print("Writing main data (with headers) to Excel...")
                        df_cleaned_data.to_excel(writer, sheet_name=sheet_name,
                                                 startrow=current_excel_row,
                                                 index=False) # index=False to not write pandas index
                    elif not df_cleaned_data.empty : # A failsafe, if it's 0x0 but not caught by columns.empty
                        print("Writing main data (empty but not 0x0) to Excel...")
                        df_cleaned_data.to_excel(writer, sheet_name=sheet_name,
                                                 startrow=current_excel_row,
                                                 index=False)
                    else:
                        print("No actual data (neither headers nor rows) to write for the main data section.")

                print(f"Data successfully saved to '{output_excel_path}' in sheet '{sheet_name}'.")
                return True

            except ImportError as ie:
                if 'openpyxl' in str(ie).lower(): print("Error: 'openpyxl' library is required to save to .xlsx files. Please install it: pip install openpyxl")
                elif 'xlwt' in str(ie).lower(): print("Error: 'xlwt' library is required to save to .xls files. Please install it: pip install xlwt")
                else: print(f"ImportError while saving to Excel: {ie}")
                return False
            except Exception as e_save:
                print(f"Error saving DataFrame to Excel '{output_excel_path}': {e_save}")
                return False

        except ImportError:
            print("Error: 'lxml' or 'html5lib' not installed for pd.read_html. Please install them: pip install lxml html5lib pandas")
            return False
        except ValueError as ve: # This can be "No tables found" or other parsing errors
            print(f"ValueError with pd.read_html for '{base_name}': {ve}")
            return False
        except Exception as e_parse:
            print(f"Error parsing HTML tables from '{base_name}' with pd.read_html: {e_parse}")
            return False
    else:
        print(f"Could not extract HTML content from '{base_name}'.")
        return False

# --- Example Usage ---
if __name__ == "__main__":
    input_mhtml_file = "AOP Automation Scripts/ZANALYSIS_PATTERN.xls" # Replace with your file path
    output_directory = os.path.dirname(input_mhtml_file) if os.path.dirname(input_mhtml_file) else "."
    input_filename_without_ext = os.path.splitext(os.path.basename(input_mhtml_file))[0]
    output_excel_file = os.path.join(output_directory, f"{input_filename_without_ext}_layout_converted.xlsx")

    print(f"Input MHTML (or misnamed .xls): {input_mhtml_file}")
    print(f"Output Excel file: {output_excel_file}")

    if convert_mhtml_to_excel(input_mhtml_file, output_excel_file, sheet_name="FormattedData"):
        print("\nConversion and saving process completed successfully.")
    else:
        print("\nConversion and saving process failed.")