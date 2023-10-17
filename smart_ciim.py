import re
import shutil
from pathlib import Path
from tkinter import *
from tkinter import filedialog
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
from ttkbootstrap.dialogs import Querybox
import ttkbootstrap as ttk
from datetime import timedelta, datetime, date
import time
from openpyxl.utils.exceptions import InvalidFileException
from tkinter import simpledialog, messagebox


def define_related_paths():
    """Define all paths relative to the global CIIM_FOLDER_PATH."""
    base_path = Path(
        CIIM_FOLDER_PATH
    )  # Convert the global CIIM_FOLDER_PATH to a Path object

    paths = {
        "delays": base_path / "General Updates" / "Delays+Cancelled works",
        "passdown": base_path / "Pass Down",
        "templates": base_path / "Important doc" / "Empty reports (templates)",
    }

    return paths


def get_latest_username_from_file():
    global username
    paths = define_related_paths()
    passdown_path = paths["passdown"]

    # Getting all .xlsx files sorted by modification time
    files = sorted(
        passdown_path.glob("*.xlsx"),
        key=lambda x: x.stat().st_mtime,
        reverse=True,
    )

    if files:
        filename = files[0].stem  # Gets the file name without the extension
        match = re.search(r"\d{6}\.\d+\s+(\w+)", filename)
        username = match.group(1)
        return username if match else None
    return None


def get_ciim_folder_path_from_file(file_path):
    """Retrieve the CIIM folder path from the given file path."""
    path = Path(file_path)
    return path.parent.parent.parent


def get_potential_week_num():
    # Constructing the initial path for the filedialog
    paths = define_related_paths()  # Get the Paths dictionary
    delays_path = paths["delays"]
    today = date.today()  # Get today's date

    # Adjust so that Sunday is the start of the new week
    adjusted_day = today + timedelta(days=1)

    current_yr = adjusted_day.year
    curr_week_num = adjusted_day.isocalendar()[1]

    # Check for the existence of the folder for the current week
    # and decrement the week number until it finds an existing folder.
    while curr_week_num > 0:  # Ensures the loop doesn't go below week 1
        potential_path = delays_path / str(current_yr) / f"WW{curr_week_num:02}"
        if potential_path.exists():
            print(curr_week_num)
            return potential_path  # Return the existing path
        curr_week_num -= 1  # Decrement the week number to check the previous week

    return (
        None  # Return None if no path is found. You can handle this case as required.
    )


def open_delays_folder():
    global delays_dir_path, tl_list

    delays_dir_path = filedialog.askdirectory(
        title="Select the Delays folder", initialdir=get_potential_week_num()
    )
    delays_dir_path = Path(delays_dir_path)
    print(f"The Delays folder Path is : {delays_dir_path}")

    # Check if the selected folder name matches the desired pattern
    pattern = re.compile(r"^\d{2}\.\d{2}\.\d{2}$")
    folder_name = delays_dir_path.name

    if not pattern.match(folder_name):
        messagebox.showerror("Error", "Please select a the delays folder")
        return  # exits

    tl_list = []
    tl_listbox.delete(0, END)

    for child in sorted(delays_dir_path.iterdir(), key=lambda x: x.stem):
        if child.is_file():
            tl_name = child.stem
            tl_listbox.insert(END, tl_name)

    set_config(save_button, state="normal")
    set_config(refresh_button, state="normal")
    set_config(transfer_to_cancelled_button, state="normal")
    set_config(transfer_to_delay_button, state="normal")


def open_const_wp():
    global construction_wp_path, CIIM_FOLDER_PATH, cp_dates

    # Pattern for the filename
    pattern = "WW*Construction Work Plan*.xlsx"

    construction_wp_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", pattern)]
    )
    construction_wp_path = Path(construction_wp_path)
    if not construction_wp_path:
        return  # Exit the function if no file was chosen

    try:
        construction_wp_workbook = load_workbook(filename=construction_wp_path)
        print(f"The Construction Plan Path is : {construction_wp_path}")

        CIIM_FOLDER_PATH = get_ciim_folder_path_from_file(construction_wp_path)
        print(f"The CIIM folder Path is : {CIIM_FOLDER_PATH}")

        construction_wp_worksheet = construction_wp_workbook["Const. Plan"]
        unique_dates = set()  # Use a set to keep track of unique dates

        for cell in construction_wp_worksheet["D"]:
            date_value_str = None
            if cell.value:
                if isinstance(cell.value, datetime):
                    date_value_str = cell.value.date().isoformat()
                else:
                    try:
                        date_value = datetime.strptime(cell.value, "%d/%m/%Y").date()
                        date_value_str = date_value.isoformat()
                    except ValueError:
                        # If it can't be parsed as a date, we'll just continue to the next cell
                        continue
            if date_value_str:
                unique_dates.add(date_value_str)

        # Convert set to a list for further use
        cp_dates = list(unique_dates)
        cp_dates.sort()
        print(f"Dates : {cp_dates}")
        construction_wp_workbook.close()

        date_pick_state = "disabled" if not CIIM_FOLDER_PATH else "normal"
        set_config(calendar_button, stat=date_pick_state)

        show_frame(frames["Delays Creator"])
        return construction_wp_path, CIIM_FOLDER_PATH

    except InvalidFileException:
        # Handling the exception
        messagebox.showerror("Error", "Please select the Construction Work Plan.")


def get_filtered_team_leaders(construction_wp_worksheet, date):
    global TL_BLACKLIST, tl_index

    maxrow = construction_wp_worksheet.max_row
    team_leaders_list, tl_index = [], []

    for i in range(3, maxrow):
        cell_obj = construction_wp_worksheet.cell(row=i, column=4)
        if pd.Timestamp(cell_obj.value) == date:
            tl_cell_value = construction_wp_worksheet.cell(row=i, column=13).value
            if tl_cell_value:
                tl_name = re.sub("[-0123456789)(.]", "", str(tl_cell_value)).strip()
                if tl_name not in TL_BLACKLIST:
                    team_leaders_list.append(tl_name)
                    tl_index.append(i)

    return team_leaders_list, tl_index


def combo_selected(event):
    global dc_year, dc_month, dc_week, dc_day, tl_index, dc_selected_date

    dc_selected_date = pd.Timestamp(dates_combobox.get())
    dc_day, dc_month, dc_year = [
        dc_selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]
    ]
    dc_week = dc_selected_date.strftime("%U")

    construction_wp_workbook = load_workbook(filename=construction_wp_path)
    construction_wp_worksheet = construction_wp_workbook["Const. Plan"]

    team_leaders_list, tl_index = get_filtered_team_leaders(
        construction_wp_worksheet, dc_selected_date
    )

    dc_tl_listbox.delete(0, END)
    for tl_name in team_leaders_list:
        dc_tl_listbox.insert(END, tl_name)
    # TODO : Fix the username lag
    # get_latest_username_from_file()
    # print(username)

    construction_wp_workbook.close()


def dc_on_listbox_double_click(event):
    global dc_selected_team_leader, tl_num
    dc_listbox_selection_index = dc_tl_listbox.curselection()
    dc_tl_listbox.itemconfig(dc_listbox_selection_index, bg="#ED969D")
    dc_selected_team_leader = str(dc_tl_listbox.get(dc_listbox_selection_index))
    tl_num = tl_index[dc_listbox_selection_index[0]]
    create_delay_wb()


def create_delay_wb():
    paths = define_related_paths()
    delays_path = paths["delays"]

    dc_year_path = delays_path / dc_year
    dc_week_path = dc_year_path / f"WW{dc_week}"
    dc_day_dir = f"{dc_day}.{dc_month}.{dc_year[2:]}"
    dc_day_path = dc_week_path / dc_day_dir

    # Create required paths if they don't exist
    dc_day_path.mkdir(parents=True, exist_ok=True)

    dc_delay_report_file = f"Delay Report {dc_selected_team_leader} {dc_day_dir}.xlsx"
    dc_delay_report_path = dc_day_path / dc_delay_report_file

    # Handle Template Copy & Rename
    if dc_delay_report_path.exists():
        status_msg = f"Delay Report {dc_selected_team_leader} {dc_day_dir} already exists!\n{dc_day_path}"
        messagebox.showerror("Error", status_msg)
        print(status_msg)
    else:
        print(f"Creating Delay Report {dc_selected_team_leader}")
        copy_and_rename_template(
            paths["templates"] / DELAY_TEMPLATE,
            dc_day_path,
            dc_delay_report_file,
        )

        cp_wb = load_workbook(filename=construction_wp_path, read_only=True)
        cp_ws = cp_wb["Const. Plan"]
        dc_delay_wb = load_workbook(filename=dc_delay_report_path)
        dc_delay_ws = dc_delay_wb.active

        copy_from_cp_to_delay(cp_ws, dc_delay_ws, tl_num, dc_day_dir)
        fill_delay_ws_cells(dc_delay_ws, cp_ws, tl_num)

        dc_delay_wb.save(str(dc_delay_report_path))

        status_msg = f"Delay Report {dc_selected_team_leader} {dc_day_dir} created!\n{dc_day_path}"
        messagebox.showinfo(None, status_msg)


def copy_and_rename_template(src_path, dest_path, new_name):
    shutil.copy(src_path, dest_path / src_path.name)
    (dest_path / src_path.name).rename(dest_path / new_name)


def set_cell(wb_sheet, row, column, value, fill=None):
    """Utility function to set cell values and, optionally, a fill pattern."""
    cell = wb_sheet.cell(row=row, column=column)
    cell.value = value
    if fill:
        cell.fill = PatternFill(bgColor=fill)


def copy_from_cp_to_delay(cp_ws, delay_ws, team_leader_num, day_folder):
    """Copy values from cp_ws to delay_ws based on a mapping."""
    mapping = {
        # (delay_row, delay_col): (cp_col, transform_fn)
        (3, 2): (None, lambda _: day_folder),
        (5, 6): (None, lambda _: ""),
        (6, 6): (None, lambda _: ""),
        (7, 7): (13, None),
        (7, 3): (11, None),
        (5, 2): (5, None),
        (6, 2): (6, None),
        (32, 2): (22, None),
        (34, 2): (23, None),
        (33, 2): (24, None),
        (16, 1): (11, None),
        (17, 1): (13, None),
        (8, 2): (21, None),
    }

    for delay_coords, (cp_col, transform) in mapping.items():
        delay_row, delay_col = delay_coords
        if cp_col is None:
            value = transform(None)
        else:
            value = cp_ws.cell(row=int(team_leader_num), column=cp_col).value
            if transform:
                value = transform(value)
        set_cell(delay_ws, delay_row, delay_col, value)


def fill_delay_ws_cells(delay_ws, cp_ws, team_leader_index):
    """Fill specific cells of delay_ws with fixed values or patterns."""
    cells_to_fill = {
        (8, 8): username,
        (16, 5): "Foreman",
        (17, 5): "Team Leader",
        (16, 7): "SEMI",
        (17, 7): "SEMI",
        (28, 2): "Y",
        (29, 2): "Y",
        (
            8,
            6,
        ): f"{cp_ws.cell(row=int(team_leader_index), column=7).value} to {cp_ws.cell(row=int(team_leader_index), column=8).value}",
        (
            8,
            4,
        ): f"{cp_ws.cell(row=int(team_leader_index), column=9).value} - {cp_ws.cell(row=int(team_leader_index), column=10).value}",
    }
    for (row, col), value in cells_to_fill.items():
        set_cell(delay_ws, row, col, value)

    # Set fill patterns for specific cells
    pattern_fill_cells = [
        "B3",
        "G7",
        "C7",
        "B5",
        "B6",
        "F8",
        "B8",
        "F5",
        "F6",
    ]
    for cell in pattern_fill_cells:
        delay_ws[cell].fill = PatternFill(bgColor="FFFFFF")


def refresh_delays_folder():
    global delays_dir_path, tl_list

    if not delays_dir_path:
        return

    # Generate sorted list of file stems in the directory
    tl_list = sorted(
        child.stem for child in delays_dir_path.iterdir() if child.is_file()
    )

    print(tl_list)
    tl_listbox.delete(0, END)
    for tl_name in tl_list:
        tl_listbox.insert(END, tl_name)


def clear_cells():
    global ENTRIES_CONFIG

    # Dynamically get the entries using their names
    entries = [globals()[entry_name] for entry_name in ENTRIES_CONFIG.keys()]

    # Clear all the entries
    for entry in entries:
        entry.delete(0, "end")

    # Reset variables
    frame4_workers_var.set(0)
    frame4_vehicles_var.set(0)


def get_cell_mapping():
    mapping = {}
    for entry_name, config in ENTRIES_CONFIG.items():
        mapping[globals()[entry_name]] = {
            "row": config["row"],
            "col": config["col"],
            "time_format": config.get("time_format", False),
        }
    return mapping


def load_delay_wb():
    global delay_report_wb, delay_report_path, delay_report_ws

    def insert_value(row, col, widget, time_format=False):
        cell_value = delay_report_ws.cell(row=row, column=col).value
        if cell_value:
            if time_format and not isinstance(cell_value, str):
                cell_value = cell_value.strftime("%H:%M")
            if isinstance(cell_value, str):
                widget.insert(0, cell_value)

    try:
        delay_report_wb.close()

        delay_wb_name = delay_report_path.name.replace(".xlsx", "")
        print(f"Previous file closed: {delay_wb_name}")

        delay_report_path = delay_report_path.parent
    except AttributeError:
        pass

    delay_report_path = delays_dir_path / f"{team_leader_name}.xlsx"
    delay_report_wb = load_workbook(filename=delay_report_path)
    delay_report_ws = delay_report_wb["Sheet1"]

    mapping = get_cell_mapping()
    for widget, details in mapping.items():
        row = details["row"]
        col = details["col"]
        time_format = details.get("time_format", False)
        insert_value(row, col, widget, time_format)


def on_tl_listbox_left_double_click(event):
    global team_leader_name
    cs = tl_listbox.curselection()
    tl_name_selected.config(text=tl_listbox.get(cs))
    team_leader_name = tl_listbox.get(cs)
    print(f"Loading : {team_leader_name}")
    clear_cells()
    load_delay_wb()
    line_status()


# def open_delay_file(event):  # opens the delay report file manually
#     tl_listbox.curselection()
#     os.startfile(delay_report_path)


def on_tl_listbox_right_double_click(event):
    tl_listbox.curselection()
    new_team_leader_name = simpledialog.askstring(
        "Input",
        "Enter the new TL name:",
    )
    new_team_leader_name = new_team_leader_name.strip()
    # Check if the user cancelled the simpledialog
    if new_team_leader_name is None:
        return

    new_delay_report_path = (
        delays_dir_path
        / f"Delay Report {new_team_leader_name} {delays_dir_path.name}.xlsx"
    )

    # 1. Update the Excel cell
    delay_ws = delay_report_wb["Sheet1"]
    set_cell(delay_ws, 7, 7, new_team_leader_name)

    # Save the changes to the current file before renaming
    delay_report_wb.save(delay_report_path)
    tl_name_selected.config(text="None")

    # 2. Rename the file
    if not new_delay_report_path.exists():
        delay_report_path.rename(new_delay_report_path)
        messagebox.showinfo(
            title="Success", message=f"Renamed to {new_delay_report_path.name}"
        )
    else:
        messagebox.showwarning("Warning", "A file with that name already exists!")

    refresh_delays_folder()


def save_delay_wb():
    global delay_report_path

    if not team_leader_name:
        return

    temp_delay_report_wb = load_workbook(filename=delay_report_path)
    delay_report_ws = temp_delay_report_wb["Sheet1"]

    # Direct assignments using ENTRIES_CONFIG
    for entry_name, config in ENTRIES_CONFIG.items():
        cell_address = config["cell"]
        entry = globals()[entry_name]
        delay_report_ws[cell_address] = entry.get()

    v1_entry_widget = globals().get("v1_entry")
    w1_entry_widget = globals().get("w1_entry")
    v1_entry_widget.insert(0, ".")
    w1_entry_widget.insert(0, "No vehicle")

    temp_delay_report_wb.save(str(delay_report_path))
    clear_cells()
    load_delay_wb()
    line_status()
    print(f"Saved successfully : {team_leader_name}")


def status_check():
    global status_color

    if (
        start_time == 1
        and end_time == 1
        and reason_var == 1
        and worker1_var == 1
        and vehicle1_var == 1
    ):
        set_config(frame3_status, text="Completed", foreground="green")

        status_color = 1
    else:
        set_config(frame3_status, text="Not completed", foreground="#E83845")
        status_color = 0


def set_entry_status(entry, var_name, default_val=0):
    if entry.get() == "":
        entry.config(style="danger.TEntry")
        globals()[var_name] = default_val
    else:
        entry.config(style="success.TEntry")
        globals()[var_name] = 1

    # TODO : No vehicle + "."
    #     v1_entry_widget = globals().get("v1_entry")
    #     w1_entry_widget = globals().get("w1_entry")
    #     v1_entry_widget.insert(0, ".")
    #     w1_entry_widget.insert(0, "No vehicle")


def line_status():
    for entry_name, config in ENTRIES_CONFIG.items():
        entry = globals()[entry_name]
        var_name = config["var"]
        set_entry_status(entry, var_name)

    if w1_entry.get() == "" and frame4_workers_var.get() == 0:
        for entry_name in WORKER_ENTRIES:
            entry = globals()[entry_name]
            entry.config(style="danger.TEntry")
        globals()["worker1_var"] = 0
    else:
        for entry_name in WORKER_ENTRIES:
            entry = globals()[entry_name]
            entry.config(style="success.TEntry")
        globals()["worker1_var"] = 1

    status_check()


def set_config(widget, **options):
    """Utility function to configure widget properties."""
    widget.config(**options)


def check_path_exists(path):
    """Check if a given path exists and print a message."""
    try:
        if path.exists():
            print(f"Path exists: {path}")
            return True
        else:
            print(f"Path does not exist: {path}")
            return False
    except Exception as e:
        print(f"An error occurred while checking the path: {e}")
        return False


def extract_date_from_path(path):
    # Get the last 2 components, which should be the date and week num
    str_date = path.name  # The last component of the path (should be the date)

    week_info = (
        path.parent.name
    )  # The second last component (should be 'Working Week Nxx')

    # Extract week number from week_info
    week_num = week_info.split("N")[-1]  # Assumes format 'Working Week Nxx'

    # Convert the string date to a datetime object
    dt_date = datetime.strptime(str_date, "%d.%m.%y")

    return str_date, dt_date, week_num


def extract_src_path_from_date(str_date, dt_date, week_num):
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(dt_date)

    # Creating the WW Delay Table
    weekly_delay_name = f"Weekly Delay table {week_num}.xlsx"
    weekly_delay_f_path = delays_dir_path.parent
    weekly_delay_path = weekly_delay_f_path / weekly_delay_name

    # Creating the CIIM Daily Report Table file path
    daily_report_name = derive_report_name(str_date)
    daily_report_f_path = paths["day"]
    daily_report_path = daily_report_f_path / daily_report_name

    print(daily_report_path)
    print(weekly_delay_path)

    return daily_report_path, weekly_delay_path


def transfer_data(
    source_file, destination_file, mappings, dest_start_row=4, dest_sheet_name=None
):
    # Load the workbooks and worksheets in read_only mode for the source file
    src_wb = load_workbook(source_file, read_only=True)
    src_ws = src_wb.active

    dest_wb = load_workbook(destination_file)
    dest_ws = dest_wb[dest_sheet_name] if dest_sheet_name else dest_wb.active

    src_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(src_ws[3])
        if cell.value in mappings
    }
    dest_header = {
        cell.value: col_num + 1
        for col_num, cell in enumerate(dest_ws[3])
        if cell.value in mappings.values()
    }

    dest_row_counter = dest_start_row
    observation_col = src_header.get("Observations", None)

    # Stream through rows using an iterator to minimize memory consumption
    for row_num, row in enumerate(src_ws.iter_rows(min_row=4, values_only=True), 4):
        if observation_col:
            print("Filtering Cancelled works")
            observation_value = row[observation_col - 1]  # -1 because row is 0-indexed
            if observation_value and "cancel" not in observation_value.lower():
                continue  # Skip this row

        for src_col, dest_col in mappings.items():
            if src_col in src_header and dest_col in dest_header:
                dest_ws.cell(
                    row=dest_row_counter, column=dest_header[dest_col]
                ).value = row[
                    src_header[src_col] - 1
                ]  # -1 because row is 0-indexed

        dest_row_counter += 1

    dest_wb.save(destination_file)


def transfer_data_generic(mapping, dest_sheet, filter_observation=None):
    if not delays_dir_path:
        messagebox.showerror(
            title="Error", message="Please select the delays folder and try again."
        )
        return

    # Prompt the user for the starting row
    dest_start_row = simpledialog.askinteger(
        "Input", "Enter the starting row:", minvalue=4
    )

    if not dest_start_row:
        return  # Exits if the dialog was closed without entering a value or if it's zero

    # Ask the user for confirmation
    confirm_transfer = messagebox.askyesno(
        "Confirm Transfer",
        f"Are you sure you want to transfer the data to row {dest_start_row}?",
    )
    if not confirm_transfer:
        return

    str_date, dt_date, week_num = extract_date_from_path(delays_dir_path)
    daily_report_path, weekly_delay_path = extract_src_path_from_date(
        str_date, dt_date, week_num
    )

    try:
        transfer_data(
            daily_report_path,
            weekly_delay_path,
            mapping,
            dest_start_row,
            dest_sheet_name=dest_sheet,
        )
        messagebox.showinfo("Success", "Data transferred successfully!")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred: {e}")


def transfer_data_to_weekly_delay():
    transfer_data_generic(TO_WEEKLY_DELAY_MAPPINGS, "Work Delay")


def transfer_data_to_weekly_cancelled():
    transfer_data_generic(
        TO_WEEKLY_CANCELLED_MAPPING, "Work Cancelled", filter_observation="Cancel"
    )


def derive_dates(selected_date):
    """Derive all related paths from a given date including multiple date formats."""
    day, month, year = [
        selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]
    ]
    week = selected_date.strftime(
        "%U"
    )  # returns the week number considering the first day of the week as Sunday

    formatted_dates = {
        "slash": f"{day}/{month}/{year[-2:]}",
        "dot": f"{day}.{month}.{year[-2:]}",
        "compact": f"{year[-2:]}{month}{day}",
    }

    return formatted_dates, week


def derive_paths_from_date(selected_date):
    """Derive all related paths from a given date including multiple date formats."""
    c_day, c_month, c_year = [
        selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]
    ]
    p_date_datetime = selected_date - timedelta(days=1)
    p_day, p_month, p_year = [
        p_date_datetime.strftime(pattern) for pattern in ["%d", "%m", "%Y"]
    ]

    c_week = selected_date.strftime(
        "%U"
    )  # returns the week number considering the first day of the week as Sunday

    c_formatted_dates = {
        "slash": f"{c_day}/{c_month}/{c_year[-2:]}",
        "dot": f"{c_day}.{c_month}.{c_year[-2:]}",
        "compact": f"{c_year[-2:]}{c_month}{c_day}",
    }

    p_formatted_dates = {
        "slash": f"{p_day}/{p_month}/{p_year[-2:]}",
        "dot": f"{p_day}.{p_month}.{p_year[-2:]}",
        "compact": f"{p_year[-2:]}{p_month}{p_day}",
    }

    paths = {
        "year": CIIM_FOLDER_PATH / f"Working Week {c_year}",
        "week": CIIM_FOLDER_PATH / f"Working Week {c_year}" / f"Working Week N{c_week}",
        "day": CIIM_FOLDER_PATH
        / f"Working Week {c_year}"
        / f"Working Week N{c_week}"
        / f"{c_year[-2:]}{c_month}{c_day}",
        "previous_year": CIIM_FOLDER_PATH / f"Working Week {p_year}",
        "previous_week": CIIM_FOLDER_PATH
        / f"Working Week {p_year}"
        / f"Working Week N{c_week}",
        "previous_day": CIIM_FOLDER_PATH
        / f"Working Week {p_year}"
        / f"Working Week N{c_week}"
        / f"{p_year[-2:]}{p_month}{p_day}",
    }

    return paths, c_formatted_dates, p_formatted_dates


def pick_date():
    global fc_selected_date
    cal = Querybox()
    fc_selected_date = cal.get_date(bootstyle="danger")
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(
        fc_selected_date
    )

    # Feedback using button's text
    calendar_button.config(
        text=f"WW: {fc_selected_date.strftime('%U')}     Date: {fc_selected_date.strftime('%d.%m.%Y')} "
    )

    day_message_exist = f'{c_formatted_dates["compact"]} folder already exists'
    if paths["day"].exists():
        messagebox.showerror("Error", day_message_exist)

    entries_state = "disabled" if paths["day"].exists() else "normal"
    set_config(fc_ocs_entry, state=entries_state)
    set_config(fc_scada_entry, state=entries_state)
    set_config(create_button, state=entries_state)

    return paths


def check_and_create_path(path):
    """If path doesn't exist, create it."""
    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)


def derive_report_name(date, template="CIIM Report Table {}.xlsx"):
    """Derive report name from a given date."""
    return template.format(date)


def create_folders_for_entries(path, entry, prefix):
    """Utility to create folders for the given prefix and entry."""
    for i in range(int(entry.get() or 0)):
        (path / f"{prefix}{i + 1}" / "Pictures").mkdir(parents=True, exist_ok=True)
        (path / f"{prefix}{i + 1}" / "Worklogs").mkdir(parents=True, exist_ok=True)


def create_folders():
    # Importing the paths and the formatted dates
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(
        fc_selected_date
    )
    main_paths = define_related_paths()

    # Creating main paths
    for key in ["year", "week", "day"]:
        Path(paths[key]).mkdir(parents=True, exist_ok=True)

    if paths["day"].exists():
        day_created_message = (
            f'{c_formatted_dates["compact"]} folder was created successfully'
        )
        messagebox.showinfo(None, day_created_message)

    # Derive report name and handle file copying and renaming
    ciim_daily_report = derive_report_name(c_formatted_dates["dot"])
    print(f"Generated report name: {ciim_daily_report}")

    templates_path = main_paths["templates"]
    fc_ciim_template_path = templates_path / DAILY_REPORT_TEMPLATE

    # Copy and rename
    print(f'Copying template to: {paths["day"]}')
    shutil.copy(fc_ciim_template_path, paths["day"])

    new_report_path = paths["day"] / ciim_daily_report
    print(f"Renaming file to: {new_report_path}")
    template_in_dest = paths["day"] / DAILY_REPORT_TEMPLATE
    if template_in_dest.exists():
        template_in_dest.rename(new_report_path)

        # Introduce a slight delay
        time.sleep(1)
    else:
        print(f'Template not found in {paths["day"]}!')

    # Creating folders for entries
    create_folders_for_entries(paths["day"], fc_ocs_entry, "W")
    create_folders_for_entries(paths["day"], fc_scada_entry, "S")

    # Creating other necessary folders
    folders_to_create = [
        "Foreman",
        "Track possession",
        "TS Worklogs",
        "PDF Files",
        "Worklogs",
    ]
    for folder in folders_to_create:
        (paths["day"] / folder).mkdir(exist_ok=True)

    # Reset and configure other widgets
    fc_ocs_entry.delete(0, END)
    fc_scada_entry.delete(0, END)
    set_config(fc_ocs_entry, state="disabled")
    set_config(fc_scada_entry, state="disabled")
    set_config(create_button, state="disabled")

    # Handle data report writing and copying
    if Path(construction_wp_path).parent != Path(paths["week"]):
        print("Not copying works to the selected date")
        return

    write_data_to_report(
        construction_wp_path,
        c_formatted_dates["slash"],
        paths["day"],
        TO_DAILY_REPORT_MAPPINGS,
    )

    # Only show the popup if previous day path exists
    if paths["previous_day"].exists():
        result = messagebox.askyesno(
            title=None,
            message=f"Transfer data to {derive_report_name(p_formatted_dates['dot'])}?",
        )
        print(result)
        if result is True:
            write_data_to_previous_report(
                construction_wp_path,
                p_formatted_dates["slash"],
                paths["previous_day"],
                TO_DAILY_REPORT_MAPPINGS,
            )


def format_datetime_column(worksheet, column_idx, row_start, row_end):
    """Format datetime column in an openpyxl worksheet."""
    for row_idx in range(row_start, row_end + 1):
        cell_value = worksheet.cell(row=row_idx, column=column_idx).value
        # Format the cell_value here ...


def format_observations_column(worksheet, column_idx, row_start, row_end):
    """Handle the observations column in an openpyxl worksheet."""
    for row_idx in range(row_start, row_end + 1):
        cell_value = worksheet.cell(row=row_idx, column=column_idx).value
        # Handle the cell_value here ...


def write_data_to_excel(src_path, target_date, target_directory, mappings, start_row=4):
    target_datetime = pd.to_datetime(target_date, format="%d/%m/%y", errors="coerce")
    formatted_target_date = target_datetime.strftime("%d.%m.%y")
    report_filename = derive_report_name(formatted_target_date)
    target_report_path = target_directory / report_filename

    # Using mappings to determine columns to load from the source
    usecols_value = list(mappings.values())
    df = pd.read_excel(src_path, skiprows=1, usecols=usecols_value)

    # Convert the 'Date [DD/MM/YY]' column to datetime format with day first
    df["Date [DD/MM/YY]"] = pd.to_datetime(
        df["Date [DD/MM/YY]"], format="%d/%m/%Y", dayfirst=True, errors="coerce"
    )

    # Filter data
    target_df = df[df["Date [DD/MM/YY]"] == target_datetime]

    # Open the target workbook
    target_workbook = load_workbook(
        filename=target_report_path,
    )
    target_worksheet = target_workbook.active

    # Map columns for efficiency outside loop
    col_mapping = {k: (list(mappings.keys()).index(k) + 2) for k in mappings.keys()}

    # Write data
    for row_idx, (_, row_data) in enumerate(target_df.iterrows(), start=start_row):
        for header, col_idx in col_mapping.items():
            target_worksheet.cell(
                row=row_idx, column=col_idx, value=row_data[mappings[header]]
            )

    # Format columns
    format_datetime_column(
        target_worksheet,
        col_mapping["Date [DD/MM/YY]"],
        start_row,
        target_worksheet.max_row,
    )
    format_observations_column(
        target_worksheet,
        col_mapping["Observations"],
        start_row,
        target_worksheet.max_row,
    )

    target_workbook.save(target_report_path)
    print(f"Report for {formatted_target_date} has been updated and saved.")


def write_data_to_report(src_path, target_date, target_directory, mappings):
    write_data_to_excel(src_path, target_date, target_directory, mappings)


def write_data_to_previous_report(src_path, target_date, target_directory, mappings):
    # Prompt the user for the starting row
    start_row_delay = simpledialog.askinteger(
        "Input", "Enter the starting row:", minvalue=4
    )
    start_row_delay = int(start_row_delay)
    # If the user cancels the prompt or doesn't enter a valid number, exit the function
    if not start_row_delay:
        return

    write_data_to_excel(
        src_path, target_date, target_directory, mappings, start_row=start_row_delay
    )


def hide_all_frames():
    for frame in frames.values():
        frame.pack_forget()


def show_frame(frame):
    app.update_idletasks()
    hide_all_frames()
    frame.pack(fill="both", expand=True)
    # If the frame being shown is not "Start Page", configure the menubar
    if frame != frames["Start Page"]:
        menu_options_list = list(frames.keys())[
            :3
        ]  # Get only the first 3 keys from the frames dictionary

        menubar = Menu(app)
        app.config(menu=menubar)
        for option in menu_options_list:
            menubar.add_command(
                label=option, command=lambda option=option: show_frame((frames[option]))
            )


def on_closing():
    # handle any cleanup here
    app.destroy()


def create_and_grid_label(parent, text, row, col, sticky="w", padx=None, pady=None):
    label = ttk.Label(parent, text=text)
    label.grid(
        row=row,
        column=col,
        sticky=sticky,
        padx=(padx if padx is not None else 0),
        pady=(pady if pady is not None else 0),
    )
    return label


def create_and_grid_entry(
    parent, row, col, sticky=None, padx=None, pady=None, **kwargs
):
    # Separate grid arguments from entry initialization arguments
    grid_args = {k: kwargs.pop(k) for k in ["columnspan"] if k in kwargs}

    entry = ttk.Entry(parent, **kwargs)
    entry.grid(
        row=row,
        column=col,
        sticky=(sticky if sticky is not None else ""),
        padx=(padx if padx is not None else 0),
        pady=(pady if pady is not None else 0),
        **grid_args,
    )  # Pass the grid arguments here
    return entry


# Root config
app = ttk.Window(
    themename="cosmo", size=(768, 552), resizable=(0, 0), title="Smart CIIM"
)

# Variables
username = ""
# Paths
CIIM_FOLDER_PATH = Path("/")
delays_dir_path = Path("/")
construction_wp_path = Path("/")
delay_report_path = Path("/")
selected_date = ""
# Tkinter variables
team_leader_name = ""
status_color = IntVar()
previous_day_entry = IntVar()
dc_day, dc_month, dc_week, dc_year = StringVar(), StringVar(), StringVar(), StringVar()
start_time, end_time, reason_var, worker1_var, vehicle1_var = 0, 0, 0, 0, 0
dc_selected_date = ""
fc_selected_date = ""
# Lists and associated data
tl_list = []
cp_dates = []
tl_index = []
# Miscellaneous variables
dc_selected_team_leader = ""
tl_num = ""
delay_report_wb = ""
delay_report_ws = ""
DELAY_TEMPLATE = "Delay Report template v.02.xlsx"
DAILY_REPORT_TEMPLATE = "CIIM Report Table v.1.xlsx"
# Those TLs won't appear in the Listbox that creates delays
TL_BLACKLIST = [
    "Eliyau Ben Zgida",
    "Emerson Gimenes Freitas",
    "Emilio Levy",
    "Samuel Lakko",
    "Ofer Akian",
    "Wissam Hagay",
    "Rami Arami",
]
# EXCEL TO CSV Columns to copy
TIME_COLUMNS = [
    "T.P Start [Time]",
    "T.P End [Time]",
    "Actual Start Time (TL):",
    "Actual Finish Time (TL):",
    "Difference",
    "Actual work time",
]
CONSTRUCTION_WP_HEADERS = [
    "Discipline [OCS/Old Bridges/TS/Scada]",
    "WW [Nº]",
    "Date [DD/MM/YY]",
    "T.P Start [Time]",
    "T.P End [Time]",
    "T.P Start [K.P]",
    "T.P End [K.P]",
    "EP",
    "ISR Start Section [Name]",
    "ISR  End Section [Name]",
    "Foremen [Israel]",
    "Team Name",
    "Team Leader\nName (Phone)",
    "Work Description (Baseline)",
    "ISR Safety Request",
    "ISR Comm&Rail:",
    "ISR T.P request (All/Track number)",
    "Observations",
]
HEADER_TO_INDEX = {
    header: index for index, header in enumerate(CONSTRUCTION_WP_HEADERS)
}
# All the Headers from the Construction Work Plan match the CIIM Report Table
TO_DAILY_REPORT_MAPPINGS = {
    "Discipline [OCS/Old Bridges/TS/Scada]": "Discipline [OCS/Old Bridges/TS/Scada]",
    "WW [Nº]": "WW [Nº]",
    "Date [DD/MM/YY]": "Date [DD/MM/YY]",
    "T.P Start [Time]": "T.P Start [Time]",
    "T.P End [Time]": "T.P End [Time]",
    "T.P Start [K.P]": "T.P Start [K.P]",
    "T.P End [K.P]": "T.P End [K.P]",
    "ISR Start Section [Name]": "ISR Start Section [Name]",
    "ISR  End Section [Name]": "ISR  End Section [Name]",
    "EP": "EP",
    "Foremen [Israel]": "Foremen [Israel]",
    "Team Name": "Team Name",
    "Team Leader\nName (Phone)": "Team Leader\nName (Phone)",
    "Work Description (Baseline)": "Work Description",
    "ISR Safety Request": "ISR Safety Request",
    "ISR Comm&Rail:": "ISR Comm&Rail:",
    "ISR T.P request (All/Track number)": "ISR T.P request (All/Track number)",
    "Observations": "Observations",
}

DAILY_REPORT_HEADERS = [
    "WW [Nº]",
    "Discipline [OCS/Old Bridges/TS/Scada]",
    "Date [DD/MM/YY]",
    "Delay details (comments + description)",
    "Team Name",
    "Team Leader\nName (Phone)",
    "EP",
    "T.P Start [Time]",
    "Actual Start Time (TL):",
    "T.P End [Time]",
    "Actual Finish Time (TL):",
    "Number of workers",
    "Work Description",
    "Observations",
]
DAILY_REPORT_HEADERS_INDEX = {
    header: index for index, header in enumerate(DAILY_REPORT_HEADERS)
}
TO_WEEKLY_DELAY_MAPPINGS = {
    "WW [Nº]": "WW",
    "Discipline [OCS/Old Bridges/TS/Scada]": "Discipline [OCS, Scada, TS]",
    "Date [DD/MM/YY]": "Date",
    "Delay details (comments + description)": "Reason",
    "Team Name": "Team Name",
    "Team Leader\nName (Phone)": "Team leader",
    "EP": "ISR section {EP}",
    "T.P Start [Time]": "TP Start Time (TAK)",
    "Work Description": "Work Description",
    "Actual Start Time (TL):": "Actual Start Time (Real Start time - TL)",
    "T.P End [Time]": "TP Finish Time (TAK)",
    "Actual Finish Time (TL):": "Actual Finish Time (Real Finish time - TL)",
    "Number of workers": "Workers",
}
TO_WEEKLY_CANCELLED_MAPPING = {
    "WW [Nº]": "WW",
    "Discipline [OCS/Old Bridges/TS/Scada]": "Discipline [OCS, Scada, TS]",
    "Date [DD/MM/YY]": "Date",
    "Observations": "Reason",
    "Team Leader\nName (Phone)": "Team leader",
    "Work Description": "Work Description",
    ("T.P Start [Time]", "T.P End [Time]"): "Planned hour per shift",
    "EP": "ISR section {EP}",
}


# Centralized list of entries and their configurations
ENTRIES_CONFIG = {
    "frame4_stime_entry": {
        "cell": "F5",
        "var": "start_time",
        "time_format": True,
        "row": 5,
        "col": 6,
    },
    "frame4_endtime_entry": {
        "cell": "F6",
        "var": "end_time",
        "time_format": True,
        "row": 6,
        "col": 6,
    },
    "frame4_reason_entry": {"cell": "A11", "var": "reason_var", "row": 11, "col": 1},
    "w1_entry": {"cell": "A18", "var": "worker1_var", "row": 18, "col": 1},
    "w2_entry": {"cell": "A19", "var": "worker2_var", "row": 19, "col": 1},
    "w3_entry": {"cell": "A20", "var": "worker3_var", "row": 20, "col": 1},
    "w4_entry": {"cell": "A21", "var": "worker4_var", "row": 21, "col": 1},
    "w5_entry": {"cell": "A22", "var": "worker5_var", "row": 22, "col": 1},
    "w6_entry": {"cell": "A23", "var": "worker6_var", "row": 23, "col": 1},
    "w7_entry": {"cell": "A24", "var": "worker7_var", "row": 24, "col": 1},
    "w8_entry": {"cell": "A25", "var": "worker8_var", "row": 25, "col": 1},
    "v1_entry": {"cell": "D28", "var": "vehicle1_var", "row": 28, "col": 4},
}

WORKER_ENTRIES = [
    "w1_entry",
    "w2_entry",
    "w3_entry",
    "w4_entry",
    "w5_entry",
    "w6_entry",
    "w7_entry",
    "w8_entry",
]
# Dictionary of frames
frame_names = [
    "Delays Creator",
    "Folders Creator",
    "Delays Manager",
    "Start Page",
]
frames = {}
# Frames
for name in frame_names:
    frames[name] = ttk.Frame(app, name=name.lower())

# Configuration for each frame's row and column weights
frame_configs = {
    "Delays Creator": {"columns": [(0, 1), (1, 5)], "rows": [(0, 1), (1, 8)]},
    "Folders Creator": {"columns": [(0, 1), (1, 5)], "rows": [(0, 1), (1, 2)]},
    "Delays Manager": {"columns": [(1, 1)], "rows": [(0, 1), (1, 8)]},
}

# Adjust the frame configurations
for frame_name, config in frame_configs.items():
    frame = frames[frame_name]

    for col, weight in config["columns"]:
        frame.columnconfigure(col, weight=weight)

    for row, weight in config["rows"]:
        frame.rowconfigure(row, weight=weight)

# Start page
welcome_label = ttk.Label(
    frames["Start Page"], text="Welcome to Smart CIIM!", font=("Helvetica", 26, "bold")
)
welcome_label.pack(pady=100)

start_button = ttk.Button(
    frames["Start Page"],
    text="Get Started!",
    command=open_const_wp,
    width=20,
    style="success",
)
start_button.pack(pady=20)
show_frame(frames["Start Page"])

# Menu 1 - Create Delays
# Frame 1 - Date select
menu1_frame1 = ttk.LabelFrame(frames["Delays Creator"], text="", style="light")
menu1_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=5)
dc_select_date_label = ttk.Label(menu1_frame1, text="   Select date:  ")
dc_select_date_label.pack(side="left")
dates_combobox = ttk.Combobox(
    menu1_frame1, values=cp_dates, postcommand=update_combo_list
)
dates_combobox.set("Date")
dates_combobox.bind("<<ComboboxSelected>>", combo_selected)
dates_combobox.pack(side="left")
# Frame 2 - TLs Listbox
menu1_frame2 = ttk.LabelFrame(frames["Delays Creator"], text="Team Leaders")
menu1_frame2.grid(row=1, column=0, sticky="wens", padx=5, pady=5)
dc_tl_listbox = Listbox(menu1_frame2)
dc_tl_listbox.pack(fill="both", expand=True)
dc_tl_listbox.bind("<Double-1>", dc_on_listbox_double_click)

# Menu 2 - Create Folders
# Frame 1 - Calendar
menu2_frame1 = ttk.LabelFrame(frames["Folders Creator"], style="light")
menu2_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=5)
select_folder_label = ttk.Label(menu2_frame1, text="   Select date:  ")
select_folder_label.pack(side="left")
calendar_button = ttk.Button(
    menu2_frame1,
    text="Create folder",
    command=pick_date,
    style="danger.",
    width=25,
    state="disabled",
)
calendar_button.pack(side="left")
# Frame 2- OCS AND SCADA WORKS
menu2_frame2 = ttk.LabelFrame(frames["Folders Creator"], text="Discipline")
menu2_frame2.grid(row=1, column=0, sticky="wens", padx=5, pady=5)
create_and_grid_label(menu2_frame2, "", 0, 0, "e", 10, 15)
create_and_grid_label(menu2_frame2, "OCS works:", 1, 1, "e", 20, 20)
fc_ocs_entry = create_and_grid_entry(menu2_frame2, 1, 2, "w", 10)
fc_ocs_entry.config(state="disabled", width=8)
create_and_grid_label(menu2_frame2, "SCADA works:", 2, 1, "e", 20, 20)
fc_scada_entry = create_and_grid_entry(menu2_frame2, 2, 2, "w", 10)
fc_scada_entry.config(state="disabled", width=8)

create_button = ttk.Button(
    menu2_frame2, text="Create", command=create_folders, state="disabled", width=8
)
create_button.grid(row=4, column=2, sticky="es", pady=10)

# Menu 3 - Delays Manager
# Frame 1 - Folder select
menu3_frame1 = ttk.LabelFrame(frames["Delays Manager"], style="light")
menu3_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=15)
delay_folder_button = ttk.Button(
    menu3_frame1,
    text="Select Delays Folder",
    command=open_delays_folder,
    width=25,
    style="success.Outline",
)
delay_folder_button.pack()
# Frame 2 - Team Leaders Listbox
menu3_frame2 = ttk.LabelFrame(
    frames["Delays Manager"],
    text="Team Leaders",
)
menu3_frame2.grid(row=1, column=0, sticky="wens", padx=5)
tl_listbox = Listbox(menu3_frame2, bd=0, width=40)
tl_listbox.pack(fill="both", expand=True)
tl_listbox.bind("<Double-1>", on_tl_listbox_left_double_click)
tl_listbox.bind("<Double-3>", on_tl_listbox_right_double_click)
# Frame 3 - Name + Status
menu3_frame3 = ttk.LabelFrame(
    frames["Delays Manager"],
    text="Status",
)
menu3_frame3.grid(
    row=0,
    column=1,
    sticky="wens",
    padx=5,
    pady=15,
)
ttk.Label(menu3_frame3, text="Selected: ").grid(
    row=0, column=0, sticky="e", pady=5, padx=5
)
tl_name_selected = ttk.Label(
    menu3_frame3, text="None", width=43, font=("Helvetica", 9, "bold")
)
tl_name_selected.grid(row=0, column=1, sticky="w")
ttk.Label(
    menu3_frame3,
    text="Status: ",
).grid(row=0, column=2, sticky="e", pady=5)
frame3_status = ttk.Label(
    menu3_frame3,
    text="Not completed",
    foreground="#E83845",
    font=("Helvetica", 9, "bold"),
)
frame3_status.grid(row=0, column=3, sticky="e")
# Frame 4 - Manager
menu3_frame4 = ttk.LabelFrame(frames["Delays Manager"], style="light")
menu3_frame4.grid(row=1, column=1, sticky="nsew", padx=5)
menu3_frame4.columnconfigure(0, weight=1)
menu3_frame4.columnconfigure(2, weight=1)
menu3_frame4.columnconfigure(3, weight=1)
menu3_frame4.columnconfigure(4, weight=1)
menu3_frame4.columnconfigure(5, weight=1)
create_and_grid_label(menu3_frame4, "Start time", 0, 0, "w", 15)
frame4_stime_entry = create_and_grid_entry(menu3_frame4, 0, 1, "e", 0, 2)
create_and_grid_label(menu3_frame4, "End time", 1, 0, "w", 15)
frame4_endtime_entry = create_and_grid_entry(menu3_frame4, 1, 1, "e", 0, 2)
create_and_grid_label(menu3_frame4, "Reason", 2, 0, "w", 15, 2)
frame4_reason_entry = create_and_grid_entry(
    menu3_frame4, 2, 1, "we", 0, 2, columnspan=3
)

sep = ttk.Separator(menu3_frame4)
sep.grid(row=3, column=0, columnspan=5, sticky="we", pady=5)

# Workers
create_and_grid_label(menu3_frame4, "Workers", 4, 0, "w", 15)

for i, entry_name in enumerate(WORKER_ENTRIES, start=4):
    globals()[entry_name] = create_and_grid_entry(menu3_frame4, i, 1, "e", pady=2)

# Add the extra label
create_and_grid_label(menu3_frame4, "", 12, 1, "we")

# Vehicles
create_and_grid_label(menu3_frame4, "      Vehicles", 4, 2, "e")
v1_entry = create_and_grid_entry(menu3_frame4, 4, 3, "e")

# Check Boxes
frame4_workers_var = IntVar()
frame4_workers_cb = ttk.Checkbutton(
    menu3_frame4,
    text="No workers",
    variable=frame4_workers_var,
    command=line_status,
    style="primary.TCheckbutton",
)
frame4_workers_cb.grid(
    row=12,
    column=1,
    sticky="e",
    pady=5,
)

frame4_vehicles_var = IntVar()
frame4_vehicles_cb = ttk.Checkbutton(
    menu3_frame4,
    text="No vehicles",
    variable=frame4_vehicles_var,
    command=line_status,
    style="primary.TCheckbutton",
)
frame4_vehicles_cb.grid(row=5, column=3, sticky="e")


# Toolbar Frame
toolbar_frame = ttk.Frame(frames["Delays Manager"])

# Create buttons (or any other widgets) for the toolbar
transfer_to_cancelled_button = ttk.Button(
    toolbar_frame,
    text="Transfer delays",
    command=transfer_data_to_weekly_delay,
    style="secondary",
    state="disabled",
)
transfer_to_cancelled_button.pack(side=LEFT, fill="both", expand=True)

transfer_to_delay_button = ttk.Button(
    toolbar_frame,
    text="Transfer cancelled",
    command=transfer_data_to_weekly_cancelled,
    style="secondary",
    state="disabled",
)
transfer_to_delay_button.pack(side=LEFT, fill="both", expand=True)

# Position the toolbar frame at the bottom of "Delays Manager" frame

save_button = ttk.Button(
    toolbar_frame, text="Save", command=save_delay_wb, style="success", state="disabled"
)
save_button.pack(side=RIGHT, fill="both", expand=True)
refresh_button = ttk.Button(
    toolbar_frame, text="Refresh", command=refresh_delays_folder, state="disabled"
)
refresh_button.pack(side=RIGHT, fill="both", expand=True)


toolbar_frame.grid(row=999, column=0, sticky="sew", columnspan=2)


app.protocol("WM_DELETE_WINDOW", on_closing)
app.mainloop()
