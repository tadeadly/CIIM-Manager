import os
import sys
import re
import shutil
import time
import datetime as dt
from datetime import datetime
from fileinput import filename
from pathlib import Path
from tkinter import *
from tkinter import filedialog, messagebox
import pandas as pd
from openpyxl import load_workbook
from PIL import ImageTk, Image
import ttkbootstrap as ttk
from pandas import date_range
from ttkbootstrap.tooltip import ToolTip
from ttkbootstrap.utility import enable_high_dpi_awareness
from ttkbootstrap.validation import add_numeric_validation


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base_path, relative_path)


background_image_path = resource_path('images/background.png')


def define_related_paths():
    """Defines all paths relative to the global CIIM_FOLDER_PATH."""

    paths = {
        "faults": base_path / "Faults" / "Electrification Control Center Fault "
                                                                              "Report Management 2.0.xlsx",
        "templates": base_path / "Important" / "Templates",
        "procedure": base_path / "Important" / "Protocols" / "CIIM procedure test2.0.xlsx",
        "delays": base_path / "Delays+Cancelled works"
    }

    return paths


def get_ww_delay_file():
    current_date = datetime.now()
    year = str(current_date.year)
    week_num = calculate_week_num(current_date)

    paths = define_related_paths()

    filename = f"Weekly Delay table WW{week_num}.xlsx"
    path = paths["delays"] / year / f"WW{week_num}" / f"Weekly Delay table WW{week_num}.xlsx"

    print(path)
    return path, filename


def get_base_path_from_file(file_path):
    """Retrieve the CIIM folder path from the given file path."""
    return file_path.parent.parent.parent


def select_const_wp():
    """
    Opens a file dialog for the user to select an Excel file.
    """
    global construction_wp_var, construction_wp_path, cancel_wp_path, cancel_wp_var
    pattern = "WW*Construction Work Plan*.xlsx"
    path = filedialog.askopenfilename(filetypes=[("Excel Files", pattern)])

    if path:  # Check if a path was actually selected
        construction_wp_path = Path(path)
        construction_wp_var.set(construction_wp_path.name)  # Set the StringVar to just the filename
        # After path has been set, update the dates
        update_dates_based_on_file()
        # If no path was selected, simply do nothing (i.e., leave the entry as is)

    return construction_wp_path if path else None


def update_dates_based_on_file():
    """
    Update the unique dates based on the selected construction work plan file.
    """
    global construction_wp_path, base_path, cp_dates

    if not construction_wp_path or construction_wp_path == Path("/"):
        return

    construction_wp_workbook = load_workbook(filename=construction_wp_path)
    base_path = get_base_path_from_file(construction_wp_path)
    cp_dates = extract_unique_dates_from_worksheet(
        construction_wp_workbook["Const. Plan"]
    )
    construction_wp_workbook.close()



def extract_unique_dates_from_worksheet(worksheet):
    """
    Extract unique dates from a given worksheet column.

    Returns:
        list: List of unique dates extracted from the worksheet, sorted in ascending order.
    """
    unique_dates = set()
    for cell in worksheet["D"]:
        date_value_str = process_date_cell(cell)
        if date_value_str:
            unique_dates.add(date_value_str)

    return sorted(list(unique_dates))


def process_date_cell(cell):
    """
    Processes a given cell's value to extract the date.
    """

    if not cell.value:
        return None

    if isinstance(cell.value, datetime):
        return cell.value.date().isoformat()

    try:
        date_value = datetime.strptime(cell.value, "%d/%m/%Y").date()
        return date_value.isoformat()
    except ValueError:
        return None


def extract_date(date_str):
    dt_date = datetime.strptime(date_str, "%Y-%m-%d")
    formatted_str_date = dt_date.strftime("%Y-%m-%d")
    week_num = calculate_week_num(dt_date)
    return formatted_str_date, dt_date, week_num


def extract_src_path_from_date(str_date, dt_date, week_num):
    """
    Constructs source file paths based on provided date information.
    """
    paths, c_formatted_dates = derive_paths_from_date(dt_date)

    # Creating the CIIM Daily Report Table file path
    daily_report_name = derive_report_name(c_formatted_dates["dot"])
    daily_report_f_path = paths["day"]
    daily_report_path = daily_report_f_path / daily_report_name

    print(daily_report_path)

    return Path(daily_report_path)


def are_files_locked(src_filepath: Path, dest_filepath: Path) -> bool:
    return is_file_locked(src_filepath) or is_file_locked(dest_filepath)


def is_file_locked(filepath: Path) -> bool:
    locked = None
    file_object = None
    if filepath.exists():
        try:
            # Try to open and close the file in append mode.
            # If this fails, the file is locked.
            file_object = filepath.open("a")
            if file_object:
                locked = False
        except IOError:
            locked = True
        finally:
            if file_object:
                file_object.close()
    return locked


def format_time(time_obj):
    if time_obj:
        return time_obj.strftime("%H:%M")
    else:
        return "None"


def transfer_data_to_cancelled(source_file, destination_file, mappings, dest_start_row):
    """
    Transfers data from a source file to a destination file based on column mappings provided.
    """
    src_wb = load_workbook(source_file, read_only=True)
    src_ws = src_wb["Const. Plan"]

    dest_wb = load_workbook(destination_file)
    dest_ws = dest_wb["Cancellations"]

    # Print all headers from the source file
    print("Source headers:", [cell.value for cell in src_ws[2]])
    print("Destination headers:", [cell.value for cell in dest_ws[2]])

    src_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(src_ws[2])
        if cell.value in mappings
           or any(cell.value in key for key in mappings if isinstance(key, tuple))
    }

    dest_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(dest_ws[2])
        if cell.value in mappings.values()
    }

    dest_row_counter = dest_start_row
    observation_col = src_header.get("Observations", None)

    transferred_rows = 0
    for row_num, row in enumerate(src_ws.iter_rows(values_only=True), 2):
        if observation_col:

            # Check if the row is blank by looking at certain key columns
            key_column_indexes = [
                src_header['T.P Start [Time]'] - 1,
                src_header['Team Leader\nName (Phone)'] - 1,
                src_header['Date'] - 1,
            ]
            if all(not row[idx] for idx in key_column_indexes):
                continue  # Skip the row as it is considered blank

            observation_value = row[observation_col - 1]  # -1 because row is 0-indexed
            # It will skip the rows that are blank or those who doesn't have 'Cancel' in the Observation cell
            if not observation_value or "cancel" not in observation_value.lower():
                continue
            # It will skip the rows that were cancelled by OCS/Scada/TS
            if any(word in observation_value.lower() for word in ["scada", "ocs", "ocs-l", "ocs-d"]):
                print(f"Skipping row {row_num} due to observation value: {observation_value}")
                continue

        for src_col, dest_col in mappings.items():

            if src_col in src_header and dest_col in dest_header:
                dest_ws.cell(
                    row=dest_row_counter, column=dest_header[dest_col]
                ).value = row[src_header[src_col] - 1]
            else:
                print(
                    f"Missing columns '{src_col}' or '{dest_col}' in source or destination for row {row_num}.")

        dest_row_counter += 1
        transferred_rows += 1

    dest_wb.save(destination_file)
    dest_wb.close()
    src_wb.close()

    return transferred_rows


def transfer_data_to_delay(source_file, destination_file, mappings, dest_start_row):
    """
    Transfers data from a source file to a destination file based on column mappings provided.
    Skips rows where 'Observations' column contains 'cancel'.
    """

    # Load the workbooks and worksheets
    src_wb = load_workbook(source_file, read_only=True)
    src_ws = src_wb.active

    dest_wb = load_workbook(destination_file)
    dest_ws = dest_wb["Delays"]

    # Print headers for debugging
    print("Source headers:", [cell.value for cell in src_ws[3]])
    print("Destination headers:", [cell.value for cell in dest_ws[2]])

    # Mapping source and destination headers to their respective column numbers
    src_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(src_ws[3])
        if cell.value in mappings or any(cell.value in key for key in mappings if isinstance(key, tuple))
    }

    dest_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(dest_ws[2])
        if cell.value in mappings.values()
    }


    # Find the column number for "Activity Summary" in the source file
    sum_col_num = None
    for col_num, cell in enumerate(src_ws[3]):
        if cell.value == "Activity Summary":
            sum_col_num = col_num + 1
            break

    if sum_col_num is None:
        print("Warning: 'Activity Summary' column not found in the source file.")
        return

    dest_row_counter = dest_start_row
    transferred_rows = 0

    # Iterating through each row in the source worksheet
    for row_num, row in enumerate(src_ws.iter_rows(min_row=4, values_only=True), 4):
        # Checks if the row is blank by looking at certain key columns
        key_column_indexes = [
            src_header['Planned Start'] - 1,
            src_header['Team Leader Name'] - 1,
            src_header['Date'] - 1
        ]
        if all(not row[idx] for idx in key_column_indexes):
            continue  # Skip the row as it is considered blank

        if sum_col_num and row[sum_col_num - 1] and "cancel" in row[sum_col_num - 1].lower():
            print(f"Skipping row {row_num} due to 'cancel' in Summary value : {row[sum_col_num - 1]}")
            continue

        # Transferring data based on mappings
        for src_col, dest_col in mappings.items():

            if src_col in src_header and dest_col in dest_header:
                dest_ws.cell(row=dest_row_counter, column=dest_header[dest_col]).value = row[
                    src_header[src_col] - 1]

            else:
                print(f"Missing columns '{src_col}' or '{dest_col}' in source or destination for row {row_num}.")

        dest_row_counter += 1
        transferred_rows += 1  # Increment only if data is actually transferred in this row

    # Save and close workbooks
    dest_wb.save(destination_file)
    dest_wb.close()
    src_wb.close()

    return transferred_rows


def calculate_week_num(date):
    # Adjust the date by moving any Sunday to the next day (Monday)
    adjusted_date = date + dt.timedelta(days=1) if date.weekday() == 6 else date

    # Use isocalendar to get the ISO week number
    iso_year, iso_week, iso_weekday = adjusted_date.isocalendar()

    return iso_week


def derive_paths_from_date(selected_date):
    """
    Constructs various related paths based on a given date.
    """

    c_day, c_month, c_year = [selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]]

    c_week = calculate_week_num(selected_date)

    c_formatted_dates = {
        "slash": f"{c_day}/{c_month}/{c_year[-2:]}",
        "dot": f"{c_day}.{c_month}.{c_year[-2:]}",
        "compact": f"{c_year[-2:]}{c_month}{c_day}",
    }

    paths = {
        "year": base_path / c_year,
        "week": base_path / c_year / f"WW{c_week}",
        "day": base_path
               / c_year
               / f"WW{c_week}"
               / f"{c_year[-2:]}{c_month}{c_day}",

    }

    return paths, c_formatted_dates



def derive_report_name(date, template="CIIM Report Table {}.xlsx"):

    return template.format(date)


def create_folders_for_entries(path, entry, prefix):
    """Utility to create folders for the given prefix and entry."""

    for index in range(int(entry.get() or 0)):
        (path / f"{prefix}{index + 1}" / "Pictures").mkdir(parents=True, exist_ok=True)
        (path / f"{prefix}{index + 1}" / "Worklogs").mkdir(parents=True, exist_ok=True)


def create_folders():
    # Prompt user to select a date
    date_str = cal_entry.entry.get()
    date_obj = datetime.strptime(date_str, '%Y-%m-%d')
    paths, c_formatted_dates = derive_paths_from_date(date_obj)

    day_message_exist = f'{c_formatted_dates["compact"]} folder already exists'

    if paths["day"].exists():
        overwrite = messagebox.askyesno("Folder Exists", f"{day_message_exist}\nDo you want to overwrite?")
        if not overwrite:
            return  # Exit the function if the user does not want to overwrite

    # If the folder does not exist or the user chooses to overwrite, continue
    main_paths = define_related_paths()

    # Creating main paths
    for key in ["year", "week", "day"]:
        Path(paths[key]).mkdir(parents=True, exist_ok=True)

    if paths["day"].exists():
        day_created_message = (
            f'{c_formatted_dates["compact"]} folder was created successfully!'
        )
        messagebox.showinfo(None, day_created_message)

    # Derive report name and set paths
    ciim_daily_report = derive_report_name(c_formatted_dates["dot"])
    new_report_path = paths["day"] / ciim_daily_report
    template_in_dest = paths["day"] / DAILY_REPORT_TEMPLATE

    # Copy and rename the template only if the report file does not already exist
    if not new_report_path.exists():
        templates_path = main_paths["templates"]
        fc_ciim_template_path = templates_path / DAILY_REPORT_TEMPLATE

        print(f'Copying template to: {paths["day"]}')
        shutil.copy(fc_ciim_template_path, paths["day"])

        if template_in_dest.exists():
            template_in_dest.rename(new_report_path)
        else:
            print(f'Template not found in {paths["day"]}!')
    else:
        print(f'{new_report_path} already exists, skipping template copying.')

    # # Creating folders for entries
    # create_folders_for_entries(paths["day"], ocs_entry, "W")
    # create_folders_for_entries(paths["day"], scada_entry, "S")

    # Creating other necessary folders
    folders_to_create = [
        "Nominations",
        "Pictures",
        "Worklogs",
        "Toolboxes",
        "Pump Documents",
    ]
    for folder in folders_to_create:
        (paths["day"] / folder).mkdir(exist_ok=True)


    # Handle data report writing and copying
    if Path(construction_wp_path).parent != Path(paths["week"]):
        print("Not copying works to the selected date")
        return

    write_data_to_excel(
        construction_wp_path,
        c_formatted_dates["slash"],
        paths["day"],
        TO_DAILY_REPORT_MAPPINGS,
    )

    # # Reset and configure other widgets
    # ocs_entry.delete(0, END)
    # scada_entry.delete(0, END)


def write_data_to_excel(src_path, target_date, target_directory, mappings, start_row=4):
    """
    Write data from the source Excel to a target report based on given mappings.
    """
    target_datetime = pd.to_datetime(target_date, format="%d/%m/%y", errors="coerce")
    formatted_target_date = target_datetime.strftime("%d.%m.%y")
    report_filename = derive_report_name(formatted_target_date)
    target_report_path = Path(target_directory / report_filename)

    # Handle the case that the user chooses to overwrite and the file is open
    if is_file_locked(target_report_path):
        messagebox.showwarning(
            "File Locked",
            f"Please close {target_report_path.name} and try again."
        )
        print("Not overwriting since the file is open")
        return

    usecols_value = list(mappings.keys())
    df = pd.read_excel(src_path, skiprows=1, usecols=usecols_value)

    df["Date"] = pd.to_datetime(df["Date"], format="%d/%m/%Y", dayfirst=True, errors="coerce")
    target_df = df[df["Date"] == target_datetime]

    # Variables
    total_works_num = None
    planned_works_num = 0

    # Open the target workbook
    wb = load_workbook(filename=target_report_path)
    ws = wb.active

    try:
        # Insert the report filename in cell A1
        ws.cell(row=1, column=1, value=report_filename[:-5])

        # Write headers (using the mappings values as headers)
        for col, header in enumerate(mappings.values(), 2):  # Starting from column B
            ws.cell(row=start_row - 1, column=col, value=header)

        # Write data
        for row_idx, (index, row_data) in enumerate(target_df.iterrows(), start=start_row):
            for col_idx, (src_header, dest_header) in enumerate(mappings.items(), 2):  # Starting from column B
                ws.cell(row=row_idx, column=col_idx, value=row_data[src_header])
            total_works_num = row_idx

        # Delete irrelevant rows
        if total_works_num is not None:
            for row_idx in range(total_works_num + 1, ws.max_row + 1):
                ws.delete_rows(total_works_num + 1)

        # Iterate through column N (col=14) starting from row 4
        for row_idx in range(4, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=14).value

            if isinstance(cell_value, float):
                # Convert float value to a string
                cell_value = str(cell_value)

            if not re.search(r"Cancel*", cell_value, re.IGNORECASE):
                # Replace the cell content with text
                ws.cell(
                    row=row_idx,
                    column=14,
                    value=""
                )
                planned_works_num += 1

        cancelled_works_num = total_works_num - planned_works_num - 3  # (-3) because it starts from that row num
        messagebox.showinfo(title="Information",
                            message=f"Num of planned works (excluding cancelled): {planned_works_num}"
                                    f"\nNum of cancelled works: {cancelled_works_num}")

        wb.save(target_report_path)
        print(f"Report for {formatted_target_date} has been updated and saved.")

    except ValueError as e:
        messagebox.showerror("Error", f"Failed to read Excel file: {e}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")

    finally:
        wb.close()



def show_frame(frame_name):
    global current_frame

    for name, frame in frames.items():
        frame.pack_forget()
        if name == frame_name:
            frame.pack(fill="both", expand=True)
            current_frame = frame_name


def clock():
    time = datetime.now()
    hour = time.strftime(" %H:%M ")
    weekday = time.strftime("%A")
    day = time.day
    month = time.strftime("%m")
    year = time.strftime("%Y")

    hour_label.configure(text=hour)
    hour_label.after(6000, clock)

    day_label.configure(text=weekday + ", " + str(day) + "/" + str(month) + "/" + str(year))


def display_dist_list():
    global dist_list_populated

    show_frame("Dist list")
    paths = define_related_paths()
    proc_path = paths["procedure"]

    if not dist_list_populated:
        try:
            # Read the Excel file into a pandas DataFrame
            df = pd.read_excel(proc_path, sheet_name='Dist. List', usecols='A, C, E')

            # Iterate over the DataFrame and the text widgets at the same time
            for col, text_widget in zip(df.columns, text_widgets):
                # Clear the text widget first
                text_widget.delete('1.0', END)
                # Insert the data into the text widget
                column_data = '\n'.join(df[col].dropna().astype(str))
                text_widget.insert('1.0', column_data)
                # Highlight lines containing "cc" after inserting the text
                highlight_lines_containing_cc(text_widget)
                text_widget.config(state="disabled")

        except ValueError as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

        else:
            dist_list_populated = True


def display_phone_list():
    global is_phone_tree_populated

    show_frame("Phone")
    # Read the Excel file

    if not is_phone_tree_populated:
        try:
            phone_tree.heading("#0", text="Phone Numbers")
            df = pd.read_excel(construction_wp_path, sheet_name="SEMI List", usecols="B:E, G")

            # Populate "Team Leaders" from DataFrame
            for department in organization["Team Leaders"]:
                if department in df:
                    workers = df[department].dropna().tolist()
                    workers = sorted(workers)  # sorts the name from A -> Z
                    organization["Team Leaders"][department].extend(workers)

            # Populate "Foremen" if the column exists in DataFrame
            if "Foremen" in df:
                foremen = df["Foremen"].dropna().tolist()
                foremen = sorted(foremen)  # sorts the name from A -> Z
                organization["Foremen"].extend(foremen)

            # Populate the tree with the organizational data
            for category, data in organization.items():
                category_id = phone_tree.insert('', 'end', text=category, open=False)
                if isinstance(data, list):
                    # If data is a list (like for "Foremen"), add items directly under the category
                    for item in data:
                        phone_tree.insert(category_id, 'end', text=item)
                else:
                    # If data is a dictionary (like for "Team Leaders"), iterate through departments
                    for dept, workers in data.items():
                        dept_id = phone_tree.insert(category_id, 'end', text=dept, open=False)
                        for worker in workers:
                            phone_tree.insert(dept_id, 'end', text=worker)

        except ValueError as e:
            messagebox.showerror("Error", f"Failed to read Excel file: {e}")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")

        else:
            is_phone_tree_populated = True


# Function to toggle between template and original content
def toggle_content(text_widget, template, original_contents, column):
    current_content = text_widget.get('1.0', 'end-1c')
    if current_content.strip() == template.strip():
        text_widget.config(state="normal")
        # If current content is the template, replace with original email
        text_widget.delete('1.0', 'end')
        text_widget.insert('1.0', original_contents[column])
        text_widget.config(state="disabled")
    else:
        # If current content is not the template, store it and insert template
        text_widget.config(state="normal")
        original_contents[column] = current_content
        text_widget.delete('1.0', 'end')
        text_widget.insert('1.0', template)
        text_widget.config(state="disabled")

    # Reapply the highlight to the text widget
    highlight_lines_containing_cc(text_widget)


def highlight_lines_containing_cc(text_widget):
    text_widget.tag_add("default_color", "1.0", "end")
    text_widget.tag_configure("highlight", underline=True, justify="center")
    text_widget.tag_configure("cc_highlight", font=("Roboto", 9, "bold"), background='#FB6C83')
    words_to_highlight = ["email", "whatsapp", "preview"]
    text_widget.tag_configure("email_color", foreground="#0489c9")

    # Searchs for lines containing '@' and apply the tag
    start_index = '1.0'
    while True:
        # Finds the next occurrence of '@'
        start_index = text_widget.search("@", start_index, 'end', nocase=True)
        if not start_index:
            break

        # Finds the end of the line containing '@'
        end_index = f"{start_index} lineend"
        # Applies the tag to the entire line
        text_widget.tag_add("email_color", f"{start_index} linestart", end_index)
        # Moves to the next line
        start_index = f"{end_index}+1c"

    # Insert new lines before "Preview Report (ISR)" if not already present
    start_index = '1.0'
    while True:
        start_index = text_widget.search("Preview Report (ISR)", start_index, 'end', nocase=True)
        if not start_index:
            break

        # Checks if the preceding characters are already '\n\n\n'
        if text_widget.get(f"{start_index}-3c", start_index) != '\n\n\n':
            text_widget.insert(start_index, '\n\n\n')
            start_index = f"{start_index}+{len('Preview Report (ISR)')}c"
        else:
            # Moves past the current match to continue searching
            start_index = f"{start_index}+{len('Preview Report (ISR)')}c"

    # Iterates over the list of words and highlight them with the 'highlight' tag
    for word in words_to_highlight:
        start_index = '1.0'
        while True:
            start_index = text_widget.search(word, start_index, 'end', nocase=True)
            if not start_index:
                break
            end_index = f"{start_index} lineend"
            text_widget.tag_add("highlight", start_index, end_index)
            start_index = f"{end_index}+1c"

    # Searchs and highlights 'cc' with a different tag 'cc_highlight'
    start_index = '1.0'
    while True:
        start_index = text_widget.search("cc", start_index, 'end', nocase=True)
        if not start_index:
            break
        end_index = f"{start_index} lineend"
        text_widget.tag_add("cc_highlight", start_index, end_index)
        start_index = f"{end_index}+1c"


def copy_to_clipboard(event):
    try:
        selected_item = phone_tree.selection()
        if selected_item:  # Check if something is selected
            parent = phone_tree.parent(selected_item[0])
            # Ensure the item is not a top-level category or department
            if parent and phone_tree.parent(parent):
                name = phone_tree.item(selected_item[0], 'text')
                frames["Phone"].clipboard_clear()
                frames["Phone"].clipboard_append(name)
                messagebox.showinfo("Info", f"Copied to clipboard: {name}")

    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def open_procedure_file():
    paths = define_related_paths()
    proc_path = paths["procedure"]
    os.startfile(proc_path)


def open_wp_file():
    global construction_wp_path
    os.startfile(construction_wp_path)


def open_faults():
    paths = define_related_paths()
    faults_path = paths["faults"]
    os.startfile(faults_path)



def open_ww_delay():
    path, filename = get_ww_delay_file()

    if path.exists():
        os.startfile(path)
    else:
        paths = define_related_paths()  # Assuming this returns a dict with 'templates' key
        template = paths["templates"] / WEEKLY_DELAY_TEMPLATE

        # Ensure the parent directory exists
        path.parent.mkdir(parents=True, exist_ok=True)

        # Copy the template file to the new location
        destination_file = path.parent / filename
        shutil.copy(template, destination_file)

        print(f"Created: {destination_file.name}")

        # Open the new file
        os.startfile(destination_file)


def clear_destination_sheet(destination_file, sheet_name):
    """
    Clears the data in the specified sheet of the destination file, keeping only the header row (row 2).
    Clears only if there is data beyond the header row.
    """
    wb = load_workbook(destination_file)
    ws = wb[sheet_name]

    # Check if there are more than two rows to clear
    if ws.max_row > 2:
        # Only clear rows starting from row 3
        ws.delete_rows(3, ws.max_row - 2)

    wb.save(destination_file)
    wb.close()



def open_options_window():
    global cp_dates

    def on_confirm_delay():
        nonlocal daily_report_path
        global current_combobox_index

        delay_transferred_total = 0
        dest_start_row = 3

        # Define paths and filenames
        main_paths = define_related_paths()
        date_str = cp_dates[0]
        formatted_str_date, dt_date, week_num = extract_date(date_str)
        paths, _ = derive_paths_from_date(dt_date)

        # Set the delay filename and path
        delay_filename = f"Weekly Delay table WW{week_num}.xlsx"
        new_report_path = paths["week"] / delay_filename
        template_in_dest = paths["week"] / WEEKLY_DELAY_TEMPLATE

        # Check if the file exists
        if not new_report_path.exists():
            # File doesn't exist, copy and rename the template
            templates_path = main_paths["templates"]
            temp_delay_template_path = templates_path / WEEKLY_DELAY_TEMPLATE

            print(f'Copying template to: {paths["week"]}')
            shutil.copy(temp_delay_template_path, paths["week"])

            if template_in_dest.exists():
                template_in_dest.rename(new_report_path)
            else:
                print(f'Template not found in {paths["week"]}!')
        else:
            print(f'{new_report_path} already exists, skipping template copying.')

            # Clear the destination sheet if the file already exists
            # clear_destination_sheet(new_report_path, "Delays")

        # Transfer data to the newly created or cleared file
        try:
            for date in cp_dates:
                formatted_str_date, dt_date, week_num = extract_date(date)
                daily_report_path = extract_src_path_from_date(formatted_str_date, dt_date, week_num)

                # Check if the daily report path exists
                if not os.path.exists(daily_report_path):
                    error_message = f"File not found: {daily_report_path}. Stopping the transfer process."
                    messagebox.showerror("File Not Found", error_message)
                    break

                delay_transferred = transfer_data_to_delay(
                    daily_report_path,
                    new_report_path,
                    TO_DELAY_MAPPINGS,
                    dest_start_row
                )
                dest_start_row += delay_transferred  # Update the starting row for the next date
                delay_transferred_total += delay_transferred

            transferred_message = f"{delay_transferred_total} rows transferred in total"
            messagebox.showinfo("Success", transferred_message)

            # Clean up empty rows after transfer
            delete_empty_rows(
                new_report_path,
                "Delays",
                delay_transferred_total + 2  # 2 accounted for the title and the Headers column
            )

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")
            top_level.destroy()


    def delete_empty_rows(file_path, sheet_name, num_rows_to_keep):
        from openpyxl import load_workbook

        wb = load_workbook(file_path)
        ws = wb[sheet_name]

        # Determine the total number of rows
        max_row = ws.max_row

        # Calculate the starting row to delete
        rows_to_delete_start = num_rows_to_keep + 1  # +1 to account for the first row to delete

        # Ensure we don't attempt to delete rows before the start of the worksheet
        if rows_to_delete_start <= max_row:
            # Delete rows from the end up to the calculated starting row
            for row in range(max_row, rows_to_delete_start - 1, -1):
                ws.delete_rows(row)

        wb.save(file_path)
        wb.close()


    def on_confirm_cancel():

        dest_start_row = 3

        # Define paths and filenames
        main_paths = define_related_paths()
        date_str = cp_dates[0]
        formatted_str_date, dt_date, week_num = extract_date(date_str)
        paths, _ = derive_paths_from_date(dt_date)

        # Set the delay filename and path
        delay_filename = f"Weekly Delay Table WW{week_num}.xlsx"
        new_report_path = paths["week"] / delay_filename
        template_in_dest = paths["week"] / WEEKLY_DELAY_TEMPLATE

        # Check if the file exists
        if not new_report_path.exists():
            # File doesn't exist, copy and rename the template
            templates_path = main_paths["templates"]
            temp_delay_template_path = templates_path / WEEKLY_DELAY_TEMPLATE

            print(f'Copying template to: {paths["week"]}')
            shutil.copy(temp_delay_template_path, paths["week"])

            if template_in_dest.exists():
                template_in_dest.rename(new_report_path)
            else:
                print(f'Template not found in {paths["week"]}!')
        else:
            print(f'{new_report_path} already exists, skipping template copying.')

        try:
            # clear_destination_sheet(new_report_path, "Cancellations Data")

            cancelled_transferred = transfer_data_to_cancelled(
                construction_wp_path,
                new_report_path,
                TO_CANCELLED_MAPPING,
                dest_start_row
            )

            # Updated message to show how many rows were transferred
            transferred_message = f"{cancelled_transferred} rows transferred." if cancelled_transferred is not None \
                else "No rows were transferred."
            messagebox.showinfo("Success", transferred_message)

            delete_empty_rows(
                new_report_path,
                "Cancellations",
                cancelled_transferred + 2  # 2 accounted for the title and the Headers column
            )

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {e}")


    daily_report_path = None

    top_level = ttk.Toplevel()
    top_level.withdraw()
    top_level.title("Transfer to Weekly Delay table")
    top_level.geometry('360x200')
    top_level.resizable(False, False)
    top_level.place_window_center()
    top_level.deiconify()

    transfer_frame = ttk.Frame(master=top_level)
    transfer_frame.pack(fill="both", expand=True)

    transfer_frame.rowconfigure(0, weight=1)
    transfer_frame.columnconfigure(0, weight=1)

    transfer_top_frame = ttk.Frame(master=transfer_frame)
    transfer_top_frame.grid(row=0, column=0, sticky="nsew")

    transfer_toolbar = ttk.Frame(master=transfer_frame)
    transfer_toolbar.grid(row=1, column=0, sticky="nsew")

    ttk.Label(transfer_top_frame, text="Select to which Sheet you want to transfer", anchor=CENTER).pack(pady=35)

    separator = ttk.Separator(transfer_toolbar)
    separator.pack(fill=X)
    delay_button = ttk.Button(transfer_toolbar, text="Delays", command=on_confirm_delay, width=8)
    delay_button.pack(side=RIGHT, padx=5, pady=10)
    cancelled_button = ttk.Button(transfer_toolbar, text="Cancelled", command=on_confirm_cancel, width=8)
    cancelled_button.pack(side=RIGHT, padx=5, pady=10)
    cancel_button = ttk.Button(transfer_toolbar, text="Cancel", command=top_level.destroy, width=8, style="secondary")
    cancel_button.pack(side=RIGHT, padx=5, pady=10)

    # Binding function
    def on_button_click(event):
        button = event.widget
        if button == delay_button:
            on_confirm_delay()
        elif button == cancelled_button:
            on_confirm_cancel()
        top_level.destroy()

    # Bind all children of transfer_toolbar to on_button_click
    for child in transfer_toolbar.winfo_children():
        child.bind("<Button-1>", on_button_click)


# ========================= Root config =========================
# Set DPI Awareness
enable_high_dpi_awareness()

app = ttk.Window(themename="litera")
app.resizable(0, 0)
app.title("CIIM Manager")

# Grid
app.grid_columnconfigure(0, weight=1)
app.grid_rowconfigure(0, weight=1)
# Geometry
app_width = 790
app_height = 530
screen_width = app.winfo_screenwidth()
screen_height = app.winfo_screenheight()
x = (screen_width / 2) - (app_width / 2)
y = (screen_height / 2) - app_height
app.geometry(f"{app_width}x{app_height}+{int(x)}+{int(y)}")

# ============================ Style ============================
style = ttk.Style()
style.configure("TButton", font=("Roboto", 9, "bold"), takefocus=False)
style.configure("TMenubutton", font=("Roboto", 9, "bold"))
style.configure("light.Treeview.Heading", font=("Roboto", 9, "bold"), rowheight=40)
style.configure("light.Treeview", rowheight=20, indent=50)

# =========================== Variables ===========================
current_frame = None
# Transfer cancelled variables
cancel_wp_path = Path("/")
cancel_wp_var = StringVar()
# Paths
base_path = Path("/")
ww_delay_path = Path("/")
construction_wp_path = Path("/")
construction_wp_var = StringVar()
delay_report_path = Path("/")
current_combobox_index = StringVar()
# Lists and associated data
cp_dates = []
# Miscellaneous variables
DELAY_TEMPLATE = "Delay Report template v.02.xlsx"
DAILY_REPORT_TEMPLATE = "CIIM Report Table v.1.xlsx"
WEEKLY_DELAY_TEMPLATE = "Weekly Delay table template v.2.xlsx"

# List of Delay reasons
delay_reasons = [
    "Delay due to no TP",
    "Delay due to track vehicle maneuvers",
    "Delay due to waiting for the ISR/WSP Supervisor",
    "Delay due to waiting for the ISR Safety/ISR Comm. Supervisor",
    "Delay due to in the release of the electrified area/rail",
    "Delay due to coordination with the control center for the TP",
    "Delay due to track vehicle maneuvers",
    "Delay due to real hours are different for the 612",
    "--Other--"
]
# ============================ Mappings ============================
TO_DAILY_REPORT_MAPPINGS = {
    "Discipline": "Discipline",
    "Date": "Date",
    "T.P Start [Time]": "Planned Start",
    "T.P End [Time]": "Planned End",
    "T.P Start [K.P]": "Start KP",
    "T.P End [K.P]": "End KP",
    "ISR Start Section [Name]": "Start Section",
    "ISR  End Section [Name]": "End Section",
    "EP": "EP",
    "Foremen [Israel]": "Foreman Name",
    "Team Leader\nName (Phone)": "Team Leader Name",
    "Work Description (Baseline)": "Activity Description",
    "Observations": "Activity Summary"
}

TO_DELAY_MAPPINGS = {
    "Discipline": "Discipline",
    "Date": "Date",
    "Start Section":"Start Section",
    "End Section": "End Section",
    "Delay Cause": "Delay Cause",
    "Team Leader Name": "Team leader Name",
    "EP": "EP",
    "Activity Description": "Activity Description",
    "Toolbox": "Toolbox",
    "Worklog": "Worklog",
    "Planned Start": "Planned Start",
    "Actual Start": "Actual Start",
    "Planned End": "Planned End",
    "Actual End": "Actual End",
}

TO_CANCELLED_MAPPING = {
    "Date": "Date",
    "Discipline": "Discipline",
    "T.P Start [Time]": "Planned Start",
    "T.P End [Time]": "Planned End",
    "EP": "EP",
    "Team Leader\nName (Phone)": "Team leader Name",
    "Work Description (Baseline)": "Activity Description",
    "Observations": "Cancellation Cause",
}

# =========================== Frames ===========================
frames = {
    "Login": ttk.Frame(master=app),
    "Home": ttk.Frame(master=app),
    "Folder": ttk.Frame(master=app),
    "Phone": ttk.Frame(master=app),
    "Dist list": ttk.Frame(master=app),
    "Faults": ttk.Frame(master=app)
}

# ====================== Images ======================
images_dict = {
    "Home": 'images/home_l.png',
    "Folder": 'images/folder_l.png',
    "Phone": 'images/phone_l.png',
    "Dist list": 'images/mail_l.png',
    "Faults": 'images/faults_l.png',
    "Transfer": 'images/transfer_l.png',
}

photo_images = {}  # Dictionary to store the PhotoImage objects

# Convert each image path to a PhotoImage object and resize them
for key, path in images_dict.items():
    corrected_path = resource_path(path)  # Get the correct path
    image = Image.open(corrected_path)
    photo_image = ImageTk.PhotoImage(image)
    photo_images[key] = photo_image

# ====================== Side Frame ======================
side_frame = ttk.Frame(master=app, bootstyle="dark")
side_frame.pack(side=LEFT, fill=Y)

tab1_empty = ttk.Label(master=side_frame, bootstyle="inverse.dark")
tab1_empty.pack(fill='x', pady=35)

tab1_button = ttk.Button(master=side_frame, command=lambda: show_frame("Home"), bootstyle="dark",
                         image=photo_images["Home"], takefocus=False)
tab1_button.pack(fill='x', ipady=7)


tab3_button = ttk.Button(master=side_frame, command=lambda: show_frame("Folder"),
                         bootstyle="dark",
                         image=photo_images["Folder"], takefocus=False)
tab3_button.pack(fill='x', ipady=7)

distlist_button = ttk.Button(master=side_frame, command=lambda: display_dist_list(),
                             bootstyle="dark",
                             image=photo_images["Dist list"], takefocus=False)
distlist_button.pack(fill='x', ipady=7)

phone_button = ttk.Button(master=side_frame, text="Phones", command=lambda: display_phone_list(),
                          bootstyle="dark",
                          image=photo_images["Phone"], takefocus=False)
phone_button.pack(fill='x', ipady=7)

faults_button = ttk.Button(master=side_frame, text="Faults", command=lambda: show_frame("Faults"),
                          bootstyle="dark",
                          image=photo_images["Faults"], takefocus=False)
faults_button.pack(fill='x', ipady=7)

transfer_button = ttk.Button(side_frame, text="Transfer Works", command=open_options_window,
                             image=photo_images["Transfer"], bootstyle="dark", takefocus=False)
transfer_button.pack(fill='x', ipady=7)

# ====================== Tab 1 - Home ======================

tab1 = frames["Home"]

tab1.columnconfigure(0, weight=0)
tab1.columnconfigure(1, weight=1)
tab1.rowconfigure(1, weight=1)

# ====================== Tab 1 - Top Frame ======================
top_frame = ttk.Frame(master=tab1)
top_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)


# Packing the hour and day labels at the top first
hour_label = ttk.Label(master=top_frame, text="12:49", font="digital-7 120")
hour_label.pack(anchor="center")

day_label = ttk.Label(master=top_frame, text="Saturday 22/01/2023", font="digital-7 35", style="secondary")
day_label.pack(padx=5, pady=5)

# ====================== Tab 1 - Bottom Frame ======================
bottom_frame = ttk.Frame(master=tab1)
bottom_frame.grid(row=2, column=1, sticky='nsew', padx=5, pady=5)

home_browse_button = ttk.Button(master=bottom_frame, text="Change", command=select_const_wp, width=10,
                                takefocus=False, bootstyle="secondary")
home_browse_button.pack(anchor='sw', side='left', pady=5)

# Open files menu
open_mb = ttk.Menubutton(top_frame, text="Open file", width=10)
open_mb.pack(pady=50)

open_menu = ttk.Menu(open_mb, tearoff=0)

open_menu.add_command(label="Construction Work Plan", command=open_wp_file)
open_menu.add_command(label="Weekly Delay Table", command=open_ww_delay)
open_menu.add_command(label="Electrification Control Center", command=open_faults)
open_menu.add_command(label="Procedure", command=open_procedure_file)

open_mb["menu"] = open_menu

open_menu.config(relief="raised")

path_entry = ttk.Entry(master=bottom_frame, textvariable=construction_wp_var)
path_entry.pack(anchor='s', side='left', fill='x', expand=True, pady=5)

# ====================== Tab 1 -Phones frame ======================
is_phone_tree_populated = False  # it will ensure it runs only once and not each time we launch the frame

phones_frame = frames["Phone"]
phones_frame.pack(fill="both", expand=True)

phones_frame.rowconfigure(0, weight=1)
phones_frame.columnconfigure(0, weight=1)

phone_tree_scroll = ttk.Scrollbar(phones_frame, style="round")
phone_tree_scroll.grid(row=0, column=1, sticky="nsw")

phone_tree = ttk.Treeview(phones_frame, cursor="hand2", yscrollcommand=phone_tree_scroll.set,
                          style="light.Treeview", padding=10)
phone_tree.grid(row=0, column=0, sticky="nsew")

organization = {
    "Team Leaders": {
        "OCS": [],
        "SCADA": [],
        "SURICATA": [],
        "OCS-D": [],
    },
    "Foremen": []
}

phone_tree_scroll.config(command=phone_tree.yview)
# Bindings
phone_tree.bind('<Double-1>', copy_to_clipboard)

# ====================== Tab 1 - Dist. list frame ======================
def get_dates():
    today = dt.date.today()
    tonight = today
    tomorrow = today + dt.timedelta(days=1)
    return tonight.strftime('%d.%m.%y'), tomorrow.strftime('%d.%m.%y')


def fill_template(template, tonight_date, tomorrow_date):
    # Replace all occurrences of dd.mm.yy with the correct dates
    return template.replace("dd.mm.yy", tonight_date, 1).replace("dd.mm.yy", tomorrow_date, 1).replace("dd.mm.yy", tonight_date, 1).replace("dd.mm.yy", tomorrow_date, 1)


def populate_templates_with_dates(templates):
    tonight_date, tomorrow_date = get_dates()
    for key, template in templates.items():
        templates[key] = fill_template(template, tonight_date, tomorrow_date)
    return templates

# Initialize frame configuration
dist_list_populated = False  # it will ensure it runs only once and not each time we launch the frame
dist_frame = frames["Dist list"]

# Configure the frame to give equal weight to all columns
for i in range(4):
    dist_frame.columnconfigure(i, weight=1)
    dist_frame.rowconfigure(1, weight=1)

templates = {
    "Preview":
               "              Email:"
               "\n\nDear All,"
               "\n\nFind attached the draft of the CIIM Report.",
    "Not Approved": "                 Email (12:00):"
                    "\n\nHi Randall,"
                    "\n\nFind attached the updated plan for tonight (dd.mm.yy) and tomorrow morning (dd.mm.yy)."
                    "\nPlease add the WSP supervisors, ISR working charges and ISR communication supervisors names in "
                    "the file."
                    "\n\n\n              Whatsapp (16:00):"
                    "\n\nGood afternoon,\nAttached is the updated work file for tonight (dd.mm.yy) and "
                    "tomorrow morning (dd.mm.yy)."
                    "\nPlease note that the hours listed are the starting hours of the T.P. Please keep in touch with "
                    "your managers about the time you should be in the field."
                    "\nGood luck."
                    "\n*TPs and supervisors in charge will be updated by ISR as soon as possible.*",
    "Approved": "      Email (17:00~20:00):"
                "\n\nDear All,"
                "\n\nPlease find the approved Construction Plan for tonight (dd.mm.yy) and tomorrow morning ("
                "dd.mm.yy)."
}

templates = populate_templates_with_dates(templates)

# Stores the original content of the text widgets
original_contents = ['' for _ in range(4)]

# Text widgets list
text_widgets = [Text(dist_frame) for _ in range(4)]


def make_command(col, tw, temp):
    return lambda: toggle_content(tw, temp, original_contents, col)


# Creates buttons and text widgets, and place them in the frame
for column, (label_text, template) in enumerate(templates.items()):
    button = ttk.Button(dist_frame, text=label_text, command=make_command(column, text_widgets[column], template),
                        bootstyle="link", takefocus=False)
    button.grid(row=0, column=column, pady=5)
    text_widget = text_widgets[column]
    # text_widget.config(highlightbackground="#d3d3d3")
    text_widget.grid(row=1, column=column, sticky="nsew", padx=2)
    ToolTip(button, text="Click for template/emails", delay=500)


# ====================== Tab 1 - Faults ======================

# Function to generate the current week number and formatted date
def get_save_path():
    today = datetime.today()
    week_number = calculate_week_num(today)
    formatted_date = today.strftime("%d.%m.%y")  # Date in dd.mm.yy format
    return f"WW{week_number}>{formatted_date}"

# Function to generate the email content
def generate_faults_email():
    # Get fault number and department selection
    fault_number = fault_number_entry.get()
    department = department_var.get()

    # Validate fault number (must be 7 digits)
    if len(fault_number) != 7 or not fault_number.isdigit():
        messagebox.showerror("Invalid Input", "Fault number must be 7 digits.")
        return

    # Validate department selection
    if department not in department_options:
        messagebox.showerror("Invalid Input", "Please select a valid department.")
        return

    # Prepare the email content based on department
    confirmation_mail = "We confirm the reception of the Fault Report No. {}.\n".format(fault_number)

    email_content = ""

    if department == "OCS":
        email_content += "ja.sierra@syneox.com\naturk@gruposemi.com\n\n"
        email_content += "Subject: OCS Fault report No. {}\n\n".format(fault_number)
        email_content += "Dear All,\n\nAttached is the Fault Report No. {}.\n".format(fault_number)

    elif department == "SCADA":
        email_content += "yshoshany@gruposemi.com\nmmiran@gruposemi.com\ngmaskalchi@gruposemi.com\n\n"
        email_content += "Subject: SCADA Fault report No. {}\n\n".format(fault_number)
        email_content += "Dear All,\n\nAttached is the Fault Report No. {}.\n".format(fault_number)

    elif department == "TS":
        email_content += "aturk@gruposemi.com\nCC:\narodriguez@gruposemi.com\nygutmacher@gruposemi.com\n\n"
        email_content += "Subject: TS Fault report No. {}\n\n".format(fault_number)
        email_content += "Dear All,\n\nAttached is the Fault Report No. {}.\n".format(fault_number)

    # Update the text widgets with the generated email content
    confirmation_text_widget.delete(1.0, END)
    confirmation_text_widget.insert(END, confirmation_mail)

    email_text_widget.delete(1.0, END)
    email_text_widget.insert(END, email_content)

    # Apply highlighting to the text widgets
    apply_highlighting(confirmation_text_widget, fault_number)
    apply_highlighting(email_text_widget, fault_number)

# Function to highlight emails, 'cc', and specific fault report lines in a text widget
def apply_highlighting(text_widget, fault_number):
    # Configure text tags for highlighting
    text_widget.tag_configure("email_color", foreground="#0489c9")
    text_widget.tag_configure("cc_highlight", font=("Roboto", 9, "bold"), background='#FB6C83')
    text_widget.tag_configure("bold_text", font=("Roboto", 9, "bold"))

    # Highlight email addresses (containing '@')
    start_index = '1.0'
    while True:
        start_index = text_widget.search("@", start_index, 'end', nocase=True)
        if not start_index:
            break
        end_index = f"{start_index} lineend"
        text_widget.tag_add("email_color", f"{start_index} linestart", end_index)
        start_index = f"{end_index}+1c"

    # Highlight 'cc' occurrences
    start_index = '1.0'
    while True:
        start_index = text_widget.search("CC:", start_index, 'end', nocase=True)
        if not start_index:
            break
        end_index = f"{start_index} lineend"
        text_widget.tag_add("cc_highlight", start_index, end_index)
        start_index = f"{end_index}+1c"

    # Highlight specific fault report lines
    patterns = [
        f"OCS Fault report No. {fault_number}",
        f"SCADA Fault report No. {fault_number}",
        f"TS Fault report No. {fault_number}"
    ]

    for pattern in patterns:
        start_index = '1.0'
        while True:
            start_index = text_widget.search(pattern, start_index, 'end', nocase=True)
            if not start_index:
                break
            end_index = f"{start_index}+{len(pattern)}c"
            text_widget.tag_add("bold_text", start_index, end_index)
            start_index = end_index

# Create a frame for the content
faults_frame = frames["Faults"]
faults_frame.pack(pady=20, padx=20, fill="both", expand=True)

# Configure grid for layout
faults_frame.columnconfigure(1, weight=1)

# First step label for saving file
save_label = ttk.Label(faults_frame, text=f"1. Save the file to {get_save_path()}")
save_label.grid(row=0, column=0, columnspan=2, sticky="w", pady=5, padx=5)

# Entry widget for fault number
fault_number_label = ttk.Label(faults_frame, text="Fault Number (7 digits):")
fault_number_label.grid(row=1, column=0, sticky="w", pady=5, padx=5)

fault_number_entry = ttk.Entry(faults_frame, width=12)
fault_number_entry.grid(row=1, column=1, sticky="w", pady=5, padx=5)

# Dropdown menu for department selection
department_label = ttk.Label(faults_frame, text="Select Department:")
department_label.grid(row=2, column=0, sticky="w", pady=5, padx=5)

department_options = ["OCS", "SCADA", "TS"]
department_var = StringVar()
department_dropdown = ttk.Combobox(faults_frame, textvariable=department_var, values=department_options,
                                   state="readonly", width=10)
department_dropdown.grid(row=2, column=1, sticky="w", pady=5, padx=5)

# Generate button
generate_button = ttk.Button(faults_frame, text="Generate", command=generate_faults_email)
generate_button.grid(row=3, column=0, columnspan=2, pady=20, padx=5)

# Label and text widget for confirmation mail
confirmation_label = ttk.Label(faults_frame, text="2. Click 'Reply' with the following:")
confirmation_label.grid(row=4, column=0, columnspan=2, sticky="w", pady=5, padx=5)

confirmation_text_widget = Text(faults_frame, height=3, width=70)
confirmation_text_widget.grid(row=5, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

# Label and text widget for department-specific email
email_label = ttk.Label(faults_frame, text="3. Send an email with the Fault Report attached and write the following:")
email_label.grid(row=6, column=0, columnspan=2, sticky="w", pady=5, padx=5)

email_text_widget = Text(faults_frame, height=13, width=70)
email_text_widget.grid(row=7, column=0, columnspan=2, sticky="ew", pady=5, padx=5)

# ====================== Tab 3 - Folder ======================

tab3 = frames["Folder"]

tab3.rowconfigure(0, weight=1)
tab3.columnconfigure(0, weight=1)
tab3.columnconfigure(2, weight=1)

tab3_mid_frame = ttk.Frame(master=tab3)
tab3_mid_frame.grid(row=0, column=1, sticky='nsew')

select_folder_label = ttk.Label(master=tab3_mid_frame, text="   Select date:  ")
select_folder_label.grid(row=0, column=0, padx=5, pady=43, sticky="e")

cal_entry = ttk.DateEntry(tab3_mid_frame, bootstyle="danger", dateformat="%Y-%m-%d")
cal_entry.grid(row=0, column=1, padx=5, pady=5, sticky="w")

discipline_frame = ttk.Frame(master=tab3_mid_frame)
discipline_frame.grid(row=1, column=0, sticky="nsew", columnspan=2, pady=40)
fc_ocs_label = ttk.Label(master=discipline_frame, text="Num of OCS works")
fc_ocs_label.grid(row=0, column=0, sticky="e", padx=5, pady=30, )
fc_scada_label = ttk.Label(master=discipline_frame, text="Num of SCADA works")
fc_scada_label.grid(row=1, column=0, sticky="e", padx=5, pady=15)

ocs_entry = ttk.Entry(master=discipline_frame, width=10)
ocs_entry.grid(row=0, column=1, sticky="e", padx=5)
ocs_entry_val = add_numeric_validation(ocs_entry, when="key")
ocs_entry.configure(validatecommand=ocs_entry_val)

scada_entry = ttk.Entry(master=discipline_frame, width=10)
scada_entry.grid(row=1, column=1, sticky="e", padx=5, pady=15)
scada_entry_val = add_numeric_validation(scada_entry, when="key")
ocs_entry.configure(validatecommand=scada_entry_val)

tab3_toolbar = ttk.Frame(master=tab3)
tab3_toolbar.grid(row=1, columnspan=3, sticky="nsew")

tab2_seperator = ttk.Separator(tab3_toolbar, orient="horizontal")
tab2_seperator.pack(side=TOP, fill=BOTH)

# Button
create_button = ttk.Button(master=tab3_toolbar, text="Create", command=create_folders, width=10)
create_button.pack(side=RIGHT, padx=10, pady=10)


show_frame("Home")
clock()
app.mainloop()
