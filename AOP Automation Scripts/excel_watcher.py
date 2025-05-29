import time
import os
import pandas as pd # For the conversion function
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler

# --- Configuration ---
FOLDER_TO_WATCH = "AOP Automation Scripts/input_data"

PROCESSED_FILES_LOG = "processed_files.txt" # To keep track of processed files

# --- Helper: Load/Save Processed Files ---
def load_processed_files():
    if not os.path.exists(PROCESSED_FILES_LOG):
        return set()
    with open(PROCESSED_FILES_LOG, 'r') as f:
        return set(line.strip() for line in f)

def add_to_processed_files(file_path):
    with open(PROCESSED_FILES_LOG, 'a') as f:
        f.write(file_path + "\n")

# --- Conversion Function (same as Part 1) ---
def xls_to_dataframe(file_path):
    if not os.path.exists(file_path):
        print(f"Error: File not found at '{file_path}'")
        return None
    if not file_path.lower().endswith('.xls'):
        # This check might be redundant if the handler filters, but good for direct calls
        print(f"Error: File '{file_path}' is not a .xls file.")
        return None
    try:
        df = pd.read_excel(file_path, engine='xlrd')
        print(f"Successfully converted '{os.path.basename(file_path)}' to DataFrame.")
        return df
    except ImportError:
        print("Error: The 'xlrd' library is required for .xls files. "
              "Please install it: pip install xlrd")
        return None
    except Exception as e:
        print(f"Error converting '{os.path.basename(file_path)}': {e}")
        return None

# --- File System Event Handler ---
class ExcelFileHandler(FileSystemEventHandler):
    def __init__(self, processed_files_set):
        super().__init__()
        self.processed_files = processed_files_set

    def on_created(self, event):
        """
        Called when a file or directory is created.
        """
        if event.is_directory:
            return  # Ignore directory creation events

        file_path = event.src_path
        if file_path.lower().endswith('.xls'):
            # Check if it's already processed (e.g., if script restarted quickly)
            # Or if multiple events fire for the same file creation (can happen)
            time.sleep(1) # Wait a bit for the file to be fully written

            # A more robust check for file readiness might be needed on some systems
            # e.g., try to open in exclusive mode, or check file size stability.
            # For simplicity, a small delay is often sufficient.

            if file_path in self.processed_files:
                print(f"File '{os.path.basename(file_path)}' already processed or processing initiated. Skipping.")
                return

            # Add to processed set immediately to prevent reprocessing by other events/restarts
            self.processed_files.add(file_path)
            # Log it to disk after successful processing attempt (or before, depending on strategy)

            print(f"New .xls file detected: {os.path.basename(file_path)}")
            df = xls_to_dataframe(file_path)

            if df is not None:
                print(f"--- DataFrame from {os.path.basename(file_path)} ---")
                print(df.head())
                # --- DO SOMETHING WITH THE DATAFRAME HERE ---
                # e.g., save to CSV, database, further processing, etc.
                # output_csv = os.path.splitext(file_path)[0] + ".csv"
                # df.to_csv(output_csv, index=False)
                # print(f"Saved DataFrame to {output_csv}")
                # ---------------------------------------------
                add_to_processed_files(file_path) # Log after successful processing
            else:
                # If conversion failed, perhaps remove from processed_files if you want to retry later?
                # For now, we assume a failed conversion is still "processed" in terms of attempt.
                # Or, you could remove it from self.processed_files if retry is desired.
                # self.processed_files.remove(file_path)
                print(f"Failed to process {os.path.basename(file_path)}. It will not be re-attempted automatically by this run.")
                # If you want to log failed attempts to a different file, do so here.


# --- Main Watcher Function ---
def start_watching(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Error: Folder '{folder_path}' does not exist. Please create it or check the path.")
        try:
            print(f"Attempting to create directory: {folder_path}")
            os.makedirs(folder_path, exist_ok=True)
            print(f"Directory '{folder_path}' created successfully.")
        except Exception as e:
            print(f"Could not create directory '{folder_path}': {e}")
            return

    processed_files_set = load_processed_files()
    print(f"Loaded {len(processed_files_set)} previously processed files.")

    # Initial scan for existing .xls files that haven't been processed
    print(f"Performing initial scan of '{folder_path}'...")
    for filename in os.listdir(folder_path):
        if filename.lower().endswith('.xls'):
            full_path = os.path.join(folder_path, filename)
            if full_path not in processed_files_set:
                print(f"Found unprocessed existing file: {filename}. Processing now.")
                # Simulate an event for consistency
                mock_event = type('Event', (object,), {'src_path': full_path, 'is_directory': False})
                handler_instance = ExcelFileHandler(processed_files_set) # Create instance to call method
                handler_instance.on_created(mock_event) # Call the handler method


    event_handler = ExcelFileHandler(processed_files_set)
    observer = Observer()
    observer.schedule(event_handler, folder_path, recursive=False) # Set recursive=True to watch subfolders

    print(f"Watching folder: {folder_path} for new .xls files...")
    observer.start()

    try:
        while True:
            time.sleep(5) # Keep the main thread alive, check for observer health if needed
    except KeyboardInterrupt:
        print("Watcher stopped by user.")
        observer.stop()
    except Exception as e:
        print(f"An error occurred in the watcher: {e}")
        observer.stop()
    finally:
        observer.join()
        print("Observer shut down.")


if __name__ == "__main__":
    # Make sure FOLDER_TO_WATCH is correctly set above!
    if FOLDER_TO_WATCH == "/path/to/your/excel_files":
         print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
         print("!!! CRITICAL: You MUST set the FOLDER_TO_WATCH variable in the script.    !!!")
         print("!!! Edit the script and change '/path/to/your/excel_files' to a real path.!!!")
         print("!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!")
    else:
        # Ensure xlrd and pandas are installed for the conversion
        try:
            import pandas
            import xlrd
        except ImportError as ie:
            print(f"Missing required library: {ie}. Please install it.")
            print("Try: pip install pandas xlrd watchdog")
        else:
            start_watching(FOLDER_TO_WATCH)