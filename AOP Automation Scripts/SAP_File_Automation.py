import pandas as pd
import os
from email import message_from_bytes
from email.parser import BytesParser
from email.policy import default as default_policy
import re # For more complex string cleaning

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
    df_converted = df.copy()
    for col in df_converted.columns:
        # Attempt to convert only if the column is of object type (likely strings)
        if df_converted[col].dtype == 'object':
            try:
                cleaned_series = df_converted[col].astype(str).str.strip()
                is_negative = cleaned_series.str.startswith('(') & cleaned_series.str.endswith(')')
                cleaned_series = cleaned_series.str.replace(r'[\(\)]', '', regex=True) # Remove parentheses
                cleaned_series = cleaned_series.str.replace(r'[$,%]', '', regex=True)
                cleaned_series = cleaned_series.str.replace(r',(?=\d{3})', '', regex=True) # Remove thousand separator commas
                numeric_series = pd.to_numeric(cleaned_series, errors='coerce')

                # If conversion was largely successful (not all NaN), update the column
                if not numeric_series.isnull().all():
                    numeric_series.loc[is_negative] = -numeric_series.loc[is_negative]
                    df_converted[col] = numeric_series
                    print(f"Column '{col}' converted to numeric.")
                else:
                    print(f"Column '{col}' could not be converted to numeric (all values became NaN after trying). Kept as object.")

            except Exception as e:
                print(f"Could not convert column '{col}' to numeric: {e}. Kept as object.")
    return df_converted


def convert_mhtml_to_excel(mhtml_file_path, output_excel_path, sheet_name="Sheet1"):
    """
    Extracts tables from an MHTML file, converts the primary table to a DataFrame,
    attempts numeric conversion, and saves it to an Excel file.
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
            print("Attempting to parse tables with pd.read_html...")
            list_of_dfs = pd.read_html(html_string, thousands=',', decimal='.') # Tell read_html about separators
                                                                                # This helps sometimes, but manual cleaning is often safer
            
            if list_of_dfs:
                df_raw = max(list_of_dfs, key=lambda d: d.size) if len(list_of_dfs) > 1 else list_of_dfs[0]
                print(f"Successfully parsed HTML table from '{base_name}' into DataFrame.")

                # Attempt to clean and convert numeric columns
                print("Attempting to clean and convert numeric data...")

                info_row_dfs_to_write = []
                df_raw = df_raw.drop(columns=[7, 4, 2, 0]).reset_index(drop=True)
                for i in range(2):
                    info_row_series = df_raw.iloc[i].fillna('')
                    info_row_dfs_to_write.append(pd.DataFrame([info_row_series.tolist()]))
    

                
                df_cleaned_data = clean_and_convert_numeric(df_raw)
                df_cleaned_data = df_cleaned_data.drop(index=[0, 1, 2]).reset_index(drop=True)
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
                                                    index=False, header=False) # index=False to not write pandas index
                            

                        
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
                    # ... (error handling for saving as before)
                    if 'openpyxl' in str(ie).lower(): print("Error: 'openpyxl' library is required to save to .xlsx files. Please install it: pip install openpyxl")
                    elif 'xlwt' in str(ie).lower(): print("Error: 'xlwt' library is required to save to .xls files. Please install it: pip install xlwt")
                    else: print(f"ImportError while saving to Excel: {ie}")
                    return False
                except Exception as e_save:
                    print(f"Error saving DataFrame to Excel '{output_excel_path}': {e_save}")
                    return False
            else:
                print(f"No tables found in the HTML content of '{base_name}' using pd.read_html.")
                return False
        except ImportError:
            print("Error: 'lxml' or 'html5lib' not installed for pd.read_html. Please install them: pip install lxml html5lib pandas")
            return False
        except ValueError as ve:
            print(f"ValueError with pd.read_html for '{base_name}' (e.g., no tables or parse error): {ve}")
            return False
        except Exception as e_parse:
            print(f"Error parsing HTML tables from '{base_name}' with pd.read_html: {e_parse}")
            return False
    else:
        print(f"Could not extract HTML content from '{base_name}'.")
        return False

# --- Example Usage ---
if __name__ == "__main__":
    input_mhtml_file = "AOP Automation Scripts/input_data/ZANALYSIS_PATTERN.xls" # Replace with your file
    output_directory = "AOP Automation Scripts/output_data"
    output_excel_file = os.path.join(output_directory, "temporary_file.xlsx")

    print(f"Input MHTML (or misnamed .xls): {input_mhtml_file}")
    print(f"Output Excel file: {output_excel_file}")

    if convert_mhtml_to_excel(input_mhtml_file, output_excel_file, sheet_name="CleanedData"):
        print("\nConversion and saving process completed successfully.")
    else:
        print("\nConversion and saving process failed.")