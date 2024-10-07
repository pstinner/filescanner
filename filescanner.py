import os
import openpyxl
from datetime import datetime

def update_tracker(root_folder, tracker_file):
    # Load the existing workbook and select the active sheet
    wb = openpyxl.load_workbook(tracker_file)
    ws = wb.active

    # Read existing data into a dictionary for quick lookup
    existing_files = {}
    for row in ws.iter_rows(min_row=2, max_col=6, values_only=True):
        if row[1] and row[2]:  # Assuming the file name is in the second column and created date in the third
            key = (row[1], row[2])
            existing_files[key] = row

    # Traverse the root folder and update the tracker
    current_time = datetime.now().strftime('%Y-%m-%d %H:%M')
    found_files = set()

    for foldername, subfolders, filenames in os.walk(root_folder):
        for filename in filenames:
            if filename == 'desktop.ini' or filename.startswith('.'):
                continue  # Skip hidden files and 'desktop.ini'
            
            file_path = os.path.join(foldername, filename)
            created_date = datetime.fromtimestamp(os.path.getctime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
            modified_date = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
            found_files.add((filename, created_date))

            status = "Active"
            key = (filename, created_date)
            if key in existing_files:
                # Update the existing row
                for row in ws.iter_rows(min_row=2, max_col=8):
                    if (row[1].value, row[2].value) == key:
                        row[0].value = foldername
                        row[3].value = modified_date
                        row[4].value = current_time
                        row[5].value = status
                        if foldername.endswith("könyvelve"):
                            row[7].value = "könyvelve"
                        elif row[7].value == "könyvelve":
                            row[7].value = ""
                        break
            else:
                # Append a new row
                new_row = [foldername, filename, created_date, modified_date, current_time, status]
                if foldername.endswith("könyvelve"):
                    new_row.append("könyvelve")
                else:
                    new_row.append("")
                ws.append(new_row)

    # Mark files as "Deleted" if they were not found in the current run
    for row in ws.iter_rows(min_row=2, max_col=8):
        key = (row[1].value, row[2].value)
        if key not in found_files:
            row[5].value = "Deleted"
        elif not row[0].value.endswith("könyvelve") and row[7].value == "könyvelve":
            row[7].value = ""

    # Save the workbook
    wb.save(tracker_file)
    print(f'Excel file "{tracker_file}" updated successfully.')


root_folder = 'G:\\My Drive\\Konyveles\\'
tracker_file = root_folder + 'tracker.xlsx'
update_tracker(root_folder, tracker_file)
