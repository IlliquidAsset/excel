import os
import pandas as pd
import xlwings as xw
from datetime import datetime
from zappy import run_with_debug

def main_logic():
    """
    This script processes budget files from a source folder, merges them with an Excel
    template, and saves them to an output folder with consistent formatting. It logs
    all processed filenames for easy tracking and reporting.
    
    HOW IT WORKS:
    1. Reads each .xlsx file from the source folder.
    2. Opens a master template file.
    3. Copies data to a specified sheet/range (e.g., Import!A1:N300).
    4. Hides all sheets except the ones you specify to keep visible.
    5. Pulls a new file name from another sheet/range (e.g., BudgetModel!A3).
    6. Saves the resulting workbook as a .xlsm file in the output folder.
    7. Logs the file and property name for future reference.

    WHY USE IT:
    - Saves time by automating repetitive budget creation tasks.
    - Keeps file names consistent and based on data (like property names).
    - Prevents accidental overwrites or missed sheets by hiding unneeded tabs.
    - Maintains an ongoing log to track which files were processed and when.
    
    SETUP NOTES:
    - Make sure to enable macros in Excel.
    - Update the variables below (paths, sheet names, cell ranges) to match your workflow.
    - Confirm that your template has an "Import" sheet, a "Budget Model" sheet, and possibly an "OBR" sheet.
    - Install the required libraries: pandas, xlwings, etc.
    """

    # >>> EDIT THESE VALUES AS NEEDED <<<

    # Paths
    folder_path = r'PATH_TO_SOURCE_FOLDER'            # Source folder containing .xlsx files
    template_path = r'PATH_TO_TEMPLATE_FILE'           # Master template .xltm file
    output_folder = r'PATH_TO_OUTPUT_FOLDER'           # Destination for generated .xlsm files
    log_file_path = r'PATH_TO_RUN_LOG_FILE'            # Log file to store property names, timestamps, etc.
    processed_files_log = r'PATH_TO_PROCESSED_FILES'   # Tracks which .xlsx files have been processed

    # Sheet names
    IMPORT_SHEET_NAME = "Import"
    BUDGET_MODEL_SHEET_NAME = "Budget Model"
    SHEETS_TO_KEEP_VISIBLE = ["Budget Model", "OBR"]  # Only these sheets remain unhidden

    # Cell references
    IMPORT_CLEAR_RANGE = "A1:N300"      # Range to clear before pasting new data
    IMPORT_PASTE_START = "A1"           # Where the pandas DataFrame will be pasted
    NEW_FILE_NAME_CELL = "A3"           # Cell containing the new file name (in Budget Model sheet)

    def get_processed_files(log_file):
        """Return a set of previously processed file names from the log."""
        if os.path.exists(log_file):
            with open(log_file, "r") as lf:
                return set(lf.read().splitlines())
        return set()

    def update_log_file(log_file, file_name):
        """Append a new file name to the processed files log."""
        with open(log_file, "a") as lf:
            lf.write(file_name + "\n")

    print("Budget creation process started.")

    # Initialize sets and lists
    processed_files = get_processed_files(processed_files_log)
    missing_a1_files = []
    failed_files = []

    # Append a run header to the main log file
    with open(log_file_path, "a") as main_log:
        main_log.write(f"\n\nRun Date and Time: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
        main_log.write("Property Register:\n")

    # Start Excel in the background
    app = xw.App(visible=False)

    try:
        # Open the template workbook once
        template_wb = app.books.open(template_path)

        # Process each .xlsx file in the source folder
        for filename in os.listdir(folder_path):
            if filename.endswith(".xlsx") and filename not in processed_files:
                try:
                    file_path = os.path.join(folder_path, filename)
                    print(f"Processing file: {filename}")

                    # Read the entire Excel file into a DataFrame
                    df = pd.read_excel(file_path, header=None)

                    # Create a new workbook based on the template
                    wb = template_wb.api.Application.Workbooks.Add(template_wb.fullname)

                    # Clear contents in the "Import" sheet and place new data
                    import_sheet = wb.Sheets[IMPORT_SHEET_NAME]
                    import_sheet.Range(IMPORT_CLEAR_RANGE).ClearContents()
                    import_sheet.Range(IMPORT_PASTE_START).Value = df.values

                    # Hide all sheets except the ones we want to keep visible
                    for sheet in wb.Sheets:
                        if sheet.Name not in SHEETS_TO_KEEP_VISIBLE:
                            sheet.Visible = False

                    # Pull the new file name from the Budget Model sheet
                    budget_model_sheet = wb.Sheets[BUDGET_MODEL_SHEET_NAME]
                    new_file_name = budget_model_sheet.Range(NEW_FILE_NAME_CELL).Value

                    if new_file_name:
                        # Clean up invalid filename characters
                        invalid_chars = '<>:"/\\|?*'
                        for char in invalid_chars:
                            new_file_name = new_file_name.replace(char, "")

                        new_file_path = os.path.join(output_folder, f"{new_file_name}.xlsm")
                        wb.SaveAs(new_file_path, FileFormat=52)  # .xlsm
                        wb.Close(SaveChanges=False)

                        # Log the property name (assumed in Import!A1)
                        property_name = import_sheet.Range("A1").Value
                        with open(log_file_path, "a") as main_log:
                            main_log.write(f"{property_name}\n")

                        # Update processed files
                        update_log_file(processed_files_log, filename)
                        print(f"Processed: {filename} -> {new_file_name}")
                    else:
                        print(f"Missing file name in '{BUDGET_MODEL_SHEET_NAME}' {NEW_FILE_NAME_CELL} for {filename}")
                        missing_a1_files.append(filename)

                except Exception as e:
                    print(f"Error processing {filename}: {e}")
                    failed_files.append(filename)

    finally:
        # Always close the template workbook and quit the app to prevent memory leaks
        if "template_wb" in locals():
            template_wb.close()
        app.quit()

    # Basic validation to ensure all source files were processed
    source_files = [f for f in os.listdir(folder_path) if f.endswith(".xlsx")]
    output_files = os.listdir(output_folder)

    if len(output_files) + len(missing_a1_files) + len(failed_files) == len(source_files):
        print("All source files processed successfully.")
    else:
        print("Some source files were not processed.")

    # Print a summary of the run
    summary = (
        f"Total files in source folder: {len(source_files)}\n"
        f"Total budgets created: {len(output_files)}\n"
        f"Files with missing A1 values: {len(missing_a1_files)}\n"
        f"Failed files: {len(failed_files)}"
    )
    print(summary)

# Wrap in zappy's debug structure
if __name__ == "__main__":
    run_with_debug("Budget Creation Script", main_logic)
