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
from datetime import timedelta, datetime
import time
from tkinter import simpledialog, messagebox
from ttkbootstrap import Style


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
    """
    Fetch the latest username from the filename in the "passdown" directory.

    Returns:
        str: The extracted username from the most recently modified file in the directory.
             Returns None if no suitable file or match is found.
    """
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
        temp_username = match.group(1)
        return temp_username if match else None
    return None


def get_ciim_folder_path_from_file(file_path):
    """Retrieve the CIIM folder path from the given file path."""
    path = Path(file_path)
    return path.parent.parent.parent


def open_const_wp():
    """
    Handle the opening and reading of the construction work plan file.
    Fetches paths for the Construction Plan and CIIM folder, and extracts unique dates from the worksheet.
    At the end, prompts the user to provide their username.

    Returns:
        tuple: Contains paths to the Construction Plan and CIIM folder.
    """
    global construction_wp_path, CIIM_FOLDER_PATH, cp_dates, username

    construction_wp_path = select_file_path()
    if not construction_wp_path:
        return

    construction_wp_workbook = load_workbook(filename=construction_wp_path)
    print(f"The Construction Plan Path is : {construction_wp_path}")

    CIIM_FOLDER_PATH = get_ciim_folder_path_from_file(construction_wp_path)
    print(f"The CIIM folder Path is : {CIIM_FOLDER_PATH}")

    cp_dates = extract_unique_dates_from_worksheet(
        construction_wp_workbook["Const. Plan"]
    )
    print(f"Dates : {cp_dates}")
    construction_wp_workbook.close()

    # username = prompt_for_username()

    return construction_wp_path, CIIM_FOLDER_PATH


def select_file_path():
    """
    Opens a file dialog for the user to select an Excel file.

    Returns:
        Path: Path object of the selected file. If no file is selected, returns None.
    """
    pattern = "WW*Construction Work Plan*.xlsx"
    path = filedialog.askopenfilename(filetypes=[("Excel Files", pattern)])
    return Path(path) if path else None


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
    Handles both datetime objects and strings representing dates.

    Returns:
        str: String representation of the date in the format YYYY-MM-DD.
             Returns None if no valid date is found.
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


def prompt_for_username():
    """
    Asks the user for their username.
    Initially, it tries to verify the most recent username and if it's not a match, prompts the user to input it.
    """
    global username
    current_username = get_latest_username_from_file()

    if messagebox.askyesno(title="Confirmation", message=f"Is it {current_username}?"):
        username = current_username
    else:
        while True:
            username = simpledialog.askstring("Input", "Please enter your name:")
            if username and username.strip():
                username = username.strip()
                break
            else:
                messagebox.showwarning("Warning", "Name cannot be empty!")

    show_frame(frames["Delays Creator"])
    print(username)
    return username


def get_filtered_team_leaders(construction_wp_worksheet, date):
    """
    Extracts team leader names and their corresponding indexes from a worksheet for a specific date.
    Excludes team leaders that match a predefined blacklist.

    Returns:
        tuple: Contains a list of team leader names and a list of their corresponding indexes in the worksheet.
    """
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


def dc_combo_selected(event):
    """
    Handle date selection from the dates_combobox and update relevant variables.

    Note:
        Also reads the construction worksheet to get the relevant list of team leaders for the selected date.
    """

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

    construction_wp_workbook.close()

    dates_combobox.configure(bootstyle="default")
    menu1_frame2.configure(bootstyle="primary")


def update_combo_list():
    """
    Update the values of dates_combobox and dm_dates_combobox with cp_dates values.
    """

    dates_combobox["values"] = cp_dates
    dm_dates_combobox["values"] = cp_dates


def dc_on_listbox_double_click(event):
    """
    Handle the event of a double click on the list box of team leaders.

    Note:
        Also calls the create_delay_wb() function to process the delay report workbook.
    """

    global dc_selected_team_leader, tl_num
    dc_listbox_selection_index = dc_tl_listbox.curselection()
    dc_tl_listbox.itemconfig(dc_listbox_selection_index, bg="#ED969D")
    dc_selected_team_leader = str(dc_tl_listbox.get(dc_listbox_selection_index))
    tl_num = tl_index[dc_listbox_selection_index[0]]
    create_delay_wb()


def create_delay_wb():
    """
    Define paths for delays and create a delay report for the selected date and team leader.

    Note:
        Copies the delay report template, populates it, and saves it to the appropriate path.
    """

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

        dc_delay_wb.save(dc_delay_report_path)

        status_msg = f"Delay Report {dc_selected_team_leader} {dc_day_dir} created!\n{dc_day_path}"
        messagebox.showinfo(None, status_msg)


def copy_and_rename_template(src_path, dest_path, new_name):
    """
    Copy a source file to a destination and rename it.
    """

    shutil.copy(src_path, dest_path / src_path.name)
    (dest_path / src_path.name).rename(dest_path / new_name)


def set_cell(wb_sheet, row, column, value, fill=None):
    """
    Set a specific cell's value and optionally its fill pattern in a workbook sheet.

    Args:
        wb_sheet: Target workbook sheet.
        row (int): Row number of the cell.
        column (int): Column number of the cell.
        value: The value to be set for the cell.
        fill (optional): Fill pattern for the cell.
    """

    """Utility function to set cell values and, optionally, a fill pattern."""
    cell = wb_sheet.cell(row=row, column=column)
    cell.value = value
    if fill:
        cell.fill = PatternFill(bgColor=fill)


def copy_from_cp_to_delay(cp_ws, delay_ws, team_leader_num, day_folder):
    """
    Copy data from a construction plan worksheet to a delay worksheet.

    Args:
        cp_ws: Source construction plan worksheet.
        delay_ws: Target delay worksheet.
        team_leader_num (int): Index of the team leader to consider.
        day_folder (str): String representation of the day folder.
    """

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
    """
    Fill specific cells of a delay worksheet with pre-defined values or patterns.
    """

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
    pattern_fill_cells = ["B3", "G7", "C7", "B5", "B6", "F8", "B8", "F5", "F6", "H8"]
    for cell in pattern_fill_cells:
        delay_ws[cell].fill = PatternFill(bgColor="FFFFFF")


def clear_cells():
    """
    Clear all the entry cells defined in the ENTRIES_CONFIG and reset related global variables.
    """

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
    """
    Generate a mapping of cell configurations based on the global ENTRIES_CONFIG.

    Globals:
        Uses ENTRIES_CONFIG to generate the mapping.

    Returns:
        dict: Mapping of widget configurations with row, column, and optional time_format details.
    """

    mapping = {}
    for entry_name, config in ENTRIES_CONFIG.items():
        mapping[globals()[entry_name]] = {
            "row": config["row"],
            "col": config["col"],
            "time_format": config.get("time_format", False),
        }
    return mapping


def load_delay_wb():
    """
    Load the delay report workbook and populate certain GUI components based on its contents

    Note:
        If a workbook is already open, it closes it before loading a new one.
    """

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


def clear_listbox():
    """
    Clear all items from the dm_tl_listbox.
    """
    dm_tl_listbox.delete(0, "end")


def populate_listbox():
    """
    Populate dm_tl_listbox with filenames present in the delays_dir_path directory.
    Filenames are sorted by their stem (name without extension).
    """

    check_and_create_path(delays_dir_path)
    for child in sorted(delays_dir_path.iterdir(), key=lambda x: x.stem):
        if child.is_file():
            tl_name = child.stem
            dm_tl_listbox.insert(END, tl_name)


def construct_delay_report_path(tl_name=None):
    """
    Construct and return the path to the delay report based on the selected date
    and optionally, the team leader's name.
    """

    global delays_dir_path

    dm_selected_date = pd.Timestamp(dm_dates_combobox.get())
    formatted_dates, week = derive_dates(dm_selected_date)
    m_formatted_date = formatted_dates["dot"]
    paths = define_related_paths()
    dir_path = (
        paths["delays"] / str(dm_selected_date.year) / f"WW{week}" / m_formatted_date
    )
    if tl_name:
        return dir_path / f"Delay Report {tl_name} {dir_path.name}.xlsx"
    else:
        return dir_path


def dm_combo_selected(event):
    """
    Handle the event when a new date is selected in the dm_dates_combobox.
    Updates the global delays directory path and repopulates the listbox with
    relevant filenames. It also enables the required configuration buttons.
    """
    global delays_dir_path

    delays_dir_path = construct_delay_report_path()
    clear_listbox()
    populate_listbox()

    # Set configurations
    set_config(save_button, state="normal")
    set_config(transfer_button, state="normal")
    dm_dates_combobox.configure(bootstyle="default")
    menu3_frame2.configure(bootstyle="primary")


def on_tl_listbox_left_double_click(event):
    """
    Handle the event when a team leader name in the listbox is double-clicked.
    Sets the displayed team leader name, clears previous data, loads new data,
    and sets line status.
    """
    global team_leader_name
    cs = dm_tl_listbox.curselection()
    if not cs:  # Check if cs is empty
        return
    team_leader_name = dm_tl_listbox.get(cs[0])

    tl_name_selected.config(text=team_leader_name)

    print(f"Loading : {team_leader_name}")
    clear_cells()
    load_delay_wb()
    line_status()


def on_tl_listbox_right_double_click(event):
    """
    Handle the event when a team leader name in the listbox is right double-clicked.
    Allows the user to rename a team leader and updates the related Excel file accordingly.
    """
    global team_leader_name

    # Request the new team leader name
    new_team_leader_name = simpledialog.askstring(
        "Input",
        "Enter the new TL name:",
    )
    if new_team_leader_name:
        new_team_leader_name = new_team_leader_name.strip()
    else:
        return

    # Confirmation of new file name
    new_delay_report_path = construct_delay_report_path(new_team_leader_name)
    confirm = messagebox.askokcancel(
        "Confirmation",
        f"Are you sure you want to rename to {new_delay_report_path.name}?",
    )
    if not confirm:
        return

    # Update the Excel cell with the new team leader's name
    delay_ws = delay_report_wb["Sheet1"]
    set_cell(delay_ws, 7, 7, new_team_leader_name)
    set_cell(delay_ws, 17, 1, new_team_leader_name)

    # Save changes and rename the file
    delay_report_wb.save(delay_report_path)

    if not new_delay_report_path.exists():
        delay_report_path.rename(new_delay_report_path)
        messagebox.showinfo("Success", f"Renamed to {new_delay_report_path.name}")
    else:
        messagebox.showwarning("Warning", "A file with that name already exists!")
        return

    team_leader_name = new_team_leader_name

    # Clear the listbox and repopulate it
    clear_listbox()
    populate_listbox()

    # Convert tuple to list
    list_items = list(dm_tl_listbox.get(0, "end"))

    # Find the index of the new team leader
    new_tl_index = None
    for i, item in enumerate(list_items):
        if new_team_leader_name in item:
            new_tl_index = i
            break

    # If the new team leader's name is found, select that item
    if new_tl_index is not None:
        dm_tl_listbox.selection_set(new_tl_index)
        on_tl_listbox_left_double_click(None)  # Passing None or a dummy event
    else:
        print(f"{new_team_leader_name} not found in the list box.")


def delete_selected_item(event):
    """
    Handle the event to delete a selected item from the listbox and the actual file.
    """

    # Get the selected item from the listbox
    cs = dm_tl_listbox.curselection()
    if not cs:
        return
    selected_item = dm_tl_listbox.get(cs[0])

    # Construct the full path to the file
    file_path = delays_dir_path / f"{selected_item}.xlsx"

    # Confirm deletion with the user
    confirm = messagebox.askyesno(
        "Confirmation", f"Do you really want to delete {selected_item}?"
    )
    if not confirm:
        return

    # Delete the file
    if file_path.exists():
        file_path.unlink()
        # Remove the item from the listbox
        dm_tl_listbox.delete(cs)
        print(f"Deleted: {file_path}")
    else:
        messagebox.showwarning("Warning", f"{selected_item} not found!")


def save_delay_wb():
    """
    Save the details from the GUI entries into the delay workbook related to the
    selected team leader.
    """
    global delay_report_path

    if not team_leader_name:
        return

    temp_delay_report_wb = load_workbook(filename=delay_report_path)
    temp_delay_report_ws = temp_delay_report_wb["Sheet1"]

    # Check for empty worker entries and update w1_entry if necessary
    if all([globals()[entry_name].get() == "" for entry_name in WORKER_ENTRIES]):
        globals()["w1_entry"].delete(0, "end")  # Clear existing content first
        globals()["w1_entry"].insert(0, ".")

    # Check for empty vehicle entry and update if necessary
    if globals()["v1_entry"].get() == "":
        globals()["v1_entry"].delete(0, "end")  # Clear existing content first
        globals()["v1_entry"].insert(0, "No vehicle")

    # Direct assignments using ENTRIES_CONFIG
    for entry_name, config in ENTRIES_CONFIG.items():
        cell_address = config["cell"]
        entry = globals()[entry_name]
        temp_delay_report_ws[cell_address] = entry.get()

    temp_delay_report_wb.save(delay_report_path)
    clear_cells()
    load_delay_wb()
    line_status()
    print(f"Saved successfully : {team_leader_name}")


def status_check():
    """
    Checks if all necessary conditions are met and updates the status displayed in the frame.
    Status will be set to "Completed" if all criteria are met, otherwise "Not completed".
    """
    global status_color

    if (
        start_time == 1
        and end_time == 1
        and reason_var == 1
        and worker1_var == 1
        and vehicle1_var == 1
    ):
        set_config(frame3_status, text="Completed", bootstyle="success")

        status_color = 1
    else:
        set_config(frame3_status, text="Not completed", bootstyle="danger")
        status_color = 0


def set_entry_status(entry, var_name, default_val=0):
    """
    Updates the style of a given entry based on its content and modifies a global variable accordingly.
    """
    if entry.get() == "":
        entry.config(style="danger.TEntry")
        globals()[var_name] = default_val
    else:
        entry.config(style="success.TEntry")
        globals()[var_name] = 1


def line_status():
    """
    Iterates over the pre-defined entry widgets to set their styles and states based on their contents.
    Also triggers the overall status check for the frame.
    """
    for entry_name, config in ENTRIES_CONFIG.items():
        entry = globals()[entry_name]
        var_name = config["var"]
        set_entry_status(entry, var_name)

    if globals()["w1_entry"].get() == "" and frame4_workers_var.get() == 0:
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
    """
    Configures properties for a given tkinter widget.

    Parameters:
    - widget: The widget to configure.
    - **options: Variable arguments representing widget properties and their values.
    """
    widget.config(**options)


def extract_date_from_path(path):
    """
    Extracts date and week information from a given directory path.

    Parameters:
    - path (Path): The path to extract date information from.

    Returns:
    - tuple: A tuple containing string format of date, datetime format of date, and week number.
    """
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
    """
    Constructs source file paths based on provided date information.

    Parameters:
    - str_date (str): String format of the date.
    - dt_date (datetime): Datetime format of the date.
    - week_num (str): Week number extracted from the date.

    Returns:
    - tuple: A tuple containing daily report path and weekly delay path.
    """
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
    """
    Transfers data from a source file to a destination file based on column mappings provided.

    Parameters:
    - source_file (str/Path): Path to the source Excel file.
    - destination_file (str/Path): Path to the destination Excel file.
    - mappings (dict): Mapping dictionary where keys are column headers in source and values are in destination.
    - dest_start_row (int, optional): The starting row in the destination file to write data. Defaults to 4.
    - dest_sheet_name (str, optional): The sheet name in the destination file to write data. Uses active sheet if None.
    """

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
    """
    Generic function to facilitate data transfer using a mapping and potentially applying a filter.

    Parameters:
    - mapping (dict): Column mapping between source and destination files.
    - dest_sheet (str): Sheet name in the destination file.
    - filter_observation (str, optional): A value to filter rows based on the "Observations" column in the source file.
    """
    if delays_dir_path == Path("/"):
        messagebox.showerror(
            title="Error", message="Please select the date of the delay."
        )
        return

    # Prompt the user for the starting row
    dest_start_row = simpledialog.askinteger(
        "Input", "Enter the starting row:", minvalue=4
    )

    if not dest_start_row:
        return  # Exits if the dialog was closed without entering a value or if it's zero

    str_date, dt_date, week_num = extract_date_from_path(delays_dir_path)
    daily_report_path, weekly_delay_path = extract_src_path_from_date(
        str_date, dt_date, week_num
    )

    # Ask the user for confirmation by entering "CONFIRM"
    user_input = simpledialog.askstring(
        "Confirmation",
        f"src: {daily_report_path.name}\ndsn: {weekly_delay_path.name}\nrow: {dest_start_row}\n\n\nType 'CONFIRM' to proceed.",
    )
    confirm_transfer = user_input == "CONFIRM"

    if not confirm_transfer:
        messagebox.showerror(title="Error", message="Data was not transferred!")
        return

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
    """
    Uses the generic data transfer function to specifically transfer data to the "Work Delay" sheet.
    """
    transfer_data_generic(TO_WEEKLY_DELAY_MAPPINGS, "Work Delay")


def transfer_data_to_weekly_cancelled():
    """
    Uses the generic data transfer function to specifically transfer data to the "Work Cancelled" sheet with a filter for cancelled works.
    """
    transfer_data_generic(
        TO_WEEKLY_CANCELLED_MAPPING, "Work Cancelled", filter_observation="Cancel"
    )


def derive_dates(selected_date):
    """
    Derive all related paths from a given date including multiple date formats.
    """

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
    """
    Constructs various related paths based on a given date.

    Parameters:
    - selected_date (datetime): The date to derive paths from.

    Returns:
    - tuple: A tuple containing a dictionary of paths, a dictionary of current formatted dates, and a dictionary of previous formatted dates.
    """

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
    """
    Prompt user to select a date using the Query box widget.

    - Updates the GUI (button's text) to display the chosen week and date.
    - Checks if a directory corresponding to the chosen date already exists.
    - Depending on directory existence, it updates the state of entry widgets and the create button.

    Returns:
        dict: Dictionary containing paths derived from the chosen date.
    """

    global fc_selected_date
    cal = Querybox()
    fc_selected_date = cal.get_date(bootstyle="primary")
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
    """
    Checks and creates a directory for the given path if it doesn't exist.
    """

    if not path.exists():
        path.mkdir(parents=True, exist_ok=True)


def derive_report_name(date, template="CIIM Report Table {}.xlsx"):
    """
    Derive a report filename based on the given date.

    Args:
        date (str): The date used for naming.
        template (str, optional): String template for report naming. Default is "CIIM Report Table {}.xlsx".

    Returns:
        str: Report name with the date inserted into the template.
    """

    return template.format(date)


def create_folders_for_entries(path, entry, prefix):
    """
    Create a set of folders based on the provided entry value.
    Each folder will have a unique name prefixed by the given prefix and will contain subfolders named "Pictures" and "Worklogs".
    """

    """Utility to create folders for the given prefix and entry."""
    for i in range(int(entry.get() or 0)):
        (path / f"{prefix}{i + 1}" / "Pictures").mkdir(parents=True, exist_ok=True)
        (path / f"{prefix}{i + 1}" / "Worklogs").mkdir(parents=True, exist_ok=True)


def create_folders():
    """
    Execute the process to:
    - Import paths and formatted dates.
    - Create main paths for year, week, and day.
    - Notify the user when a folder is successfully created.
    - Generate, copy, and rename report files.
    - Create additional folders based on entry values.
    - Create other necessary folders.
    - Reset and configure GUI widgets.
    - Handle data report writing and copying.
    """

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


def write_data_to_excel(src_path, target_date, target_directory, mappings, start_row=4):
    """
    Write data from the source Excel to a target report based on given mappings.
    """

    target_datetime = pd.to_datetime(target_date, format="%d/%m/%y", errors="coerce")
    formatted_target_date = target_datetime.strftime("%d.%m.%y")
    report_filename = derive_report_name(formatted_target_date)
    target_report_path = target_directory / report_filename

    usecols_value = list(mappings.values())
    df = pd.read_excel(src_path, skiprows=1, usecols=usecols_value)

    df["Date [DD/MM/YY]"] = pd.to_datetime(
        df["Date [DD/MM/YY]"], format="%d/%m/%Y", dayfirst=True, errors="coerce"
    )
    target_df = df[df["Date [DD/MM/YY]"] == target_datetime]

    # Open the target workbook
    target_workbook = load_workbook(
        filename=target_report_path,
    )
    target_worksheet = target_workbook.active

    col_mapping = {k: (list(mappings.keys()).index(k) + 2) for k in mappings.keys()}

    for row_idx, (_, row_data) in enumerate(target_df.iterrows(), start=start_row):
        for header, col_idx in col_mapping.items():
            target_worksheet.cell(
                row=row_idx, column=col_idx, value=row_data[mappings[header]]
            )

    target_workbook.save(target_report_path)
    print(f"Report for {formatted_target_date} has been updated and saved.")


def write_data_to_report(src_path, target_date, target_directory, mappings):
    """
    Write data to the current day's report.
    """

    write_data_to_excel(src_path, target_date, target_directory, mappings)


def write_data_to_previous_report(src_path, target_date, target_directory, mappings):
    """
    Write data to the previous day's report. Prompt the user to select a starting row.
    """

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


def change_theme(theme_name):
    """Changes the theme of the app to the specified theme name."""
    style.theme_use(theme_name)  # Set the theme
    current_theme = style.theme_use()  # Retrieve the current theme's name
    print(current_theme)

    dark_themes = THEMES[-3:]
    labelframes_to_change = [menu1_frame1, menu2_frame1, menu3_frame1, menu3_frame4]

    if current_theme in dark_themes:
        for labelframe in labelframes_to_change:
            labelframe.config(style="dark.TLabelframe")
    else:
        for labelframe in labelframes_to_change:
            labelframe.config(style="light.TLabelframe")


def show_frame(frame):
    hide_all_frames()
    frame.pack(fill="both", expand=True)

    # If the frame is not the "Start Page" frame, then create the menubar
    # if frame != frames["Start Page"]:
    # The menubar now has options for File, Manage, and Settings
    menubar = Menu(app)
    app.config(menu=menubar)

    create_menu = Menu(menubar, tearoff=0)
    create_menu.add_command(
        label="New file", command=lambda: show_frame(frames["Delays Creator"])
    )
    create_menu.add_command(
        label="New folder", command=lambda: show_frame(frames["Folders Creator"])
    )
    create_menu.add_separator()
    create_menu.add_command(label="Exit", command=app.quit)
    menubar.add_cascade(label="File", menu=create_menu)

    edit_menu = Menu(menubar, tearoff=0)
    menubar.add_command(
        label="Edit", command=lambda: show_frame(frames["Delays Manager"])
    )

    settings_menu = Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Settings", menu=settings_menu)
    theme_menu = Menu(settings_menu, tearoff=0)
    settings_menu.add_cascade(label="Appearance", menu=theme_menu)

    for theme in THEMES:
        theme_menu.add_command(
            label=theme, command=lambda theme_name=theme: change_theme(theme_name)
        )

    # transfer_menu.add_command(
    #     label="Weekly delay sheet", command=transfer_data_to_weekly_delay
    # )
    # transfer_menu.add_command(
    #     label="Weekly cancelled sheet", command=transfer_data_to_weekly_cancelled
    # )
    settings_menu.add_separator()
    settings_menu.add_command(label="Close", command=lambda: None)
    # else:
    #     # If it is the "Start page" frame, set an empty menu (remove the menubar)
    #     app.config(menu=Menu(app))
    #


def on_closing():
    # handle any cleanup here
    app.destroy()


# Root config
app = ttk.Window(
    themename="cosmo", size=(768, 522), resizable=(0, 0), title="Smart CIIM"
)

app.iconbitmap("icon.ico")


style = Style()

# Define a custom light style for Labelframe
style.configure("light.TLabelframe")  # add any other styling properties

# Define a custom dark style for Labelframe
style.configure("dark.TLabelframe")  # add any other styling properties

# app = Tk()
# app.resizable(0, 0)
# app.title("Smart CIIM")
# app.geometry("768x552")


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
frame4_workers_var = IntVar()
frame4_vehicles_var = IntVar()
DELAY_TEMPLATE = "Delay Report template v.02.xlsx"
DAILY_REPORT_TEMPLATE = "CIIM Report Table v.1.xlsx"
# Themes
THEMES = [
    "journal",
    "minty",
    "cosmo",
    "cerculean",
    "yeti",
    "solar",
    "superhero",
    "darkly",
]
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
# Frames
frames = {
    "Start Page": ttk.Frame(app),
    "Delays Creator": ttk.Frame(app),
    "Folders Creator": ttk.Frame(app),
    "Delays Manager": ttk.Frame(app),
}

frame_configs = {
    "Delays Creator": {"columns": [(0, 1), (1, 5)], "rows": [(0, 1), (1, 8)]},
    "Folders Creator": {"columns": [(0, 1), (1, 5)], "rows": [(0, 1), (1, 9)]},
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
    menu1_frame1, values=cp_dates, postcommand=update_combo_list, style="danger"
)
dates_combobox.set("Date")
dates_combobox.bind("<<ComboboxSelected>>", dc_combo_selected)
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
    menu2_frame1, text="Browse", command=pick_date, width=25, style="Outline"
)
calendar_button.pack(side="left")

# Frame 2- OCS AND SCADA WORKS
menu2_frame2 = ttk.LabelFrame(frames["Folders Creator"], text="Discipline")
menu2_frame2.grid(row=1, column=0, sticky="wens", padx=5, pady=5)
ttk.Label(menu2_frame2, text="OCS works:").grid(
    row=1, column=1, sticky="e", padx=20, pady=20
)
fc_ocs_entry = ttk.Entry(menu2_frame2, width=8)
fc_ocs_entry.grid(row=1, column=2, sticky="w", padx=10)
fc_ocs_entry.config(state="disabled")

ttk.Label(menu2_frame2, text="SCADA works:").grid(
    row=2, column=1, sticky="e", padx=20, pady=30
)
fc_scada_entry = ttk.Entry(menu2_frame2, width=8)
fc_scada_entry.grid(row=2, column=2, sticky="w", padx=10)
fc_scada_entry.config(state="disabled")

create_button = ttk.Button(
    menu2_frame2, text="Create", command=create_folders, width=8, state="disabled"
)
create_button.grid(row=4, column=2, sticky="es", pady=10)

# Menu 3 - Delays Manager
# Frame 1 - Date Select
menu3_frame1 = ttk.LabelFrame(frames["Delays Manager"], text="", style="light")
menu3_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=5)
dc_select_date_label = ttk.Label(menu3_frame1, text="   Select date:  ")
dc_select_date_label.pack(side="left")
dm_dates_combobox = ttk.Combobox(
    menu3_frame1, values=cp_dates, postcommand=update_combo_list, style="danger"
)
dm_dates_combobox.set("Date")
dm_dates_combobox.bind("<<ComboboxSelected>>", dm_combo_selected)
dm_dates_combobox.pack(side="left")

# Frame 2 - Team Leaders Listbox
menu3_frame2 = ttk.LabelFrame(frames["Delays Manager"], text="Team Leaders")
menu3_frame2.grid(row=1, column=0, sticky="wens", padx=5, pady=5)
dm_tl_listbox = Listbox(menu3_frame2, bd=0, width=40)
dm_tl_listbox.pack(fill="both", expand=True)
dm_tl_listbox.bind("<Double-1>", on_tl_listbox_left_double_click)
dm_tl_listbox.bind("<Double-3>", on_tl_listbox_right_double_click)
dm_tl_listbox.bind("<Delete>", delete_selected_item)

# Frame 3 - Name + Status
menu3_frame3 = ttk.LabelFrame(frames["Delays Manager"], text="Status")
menu3_frame3.grid(row=0, column=1, sticky="wens", padx=5, pady=5)
ttk.Label(menu3_frame3, text="Selected: ").grid(
    row=0, column=0, sticky="e", pady=5, padx=5
)
tl_name_selected = ttk.Label(
    menu3_frame3, text="None", width=43, font=("Helvetica", 9, "bold")
)
tl_name_selected.grid(row=0, column=1, sticky="w")
ttk.Label(menu3_frame3, text="Status: ").grid(row=0, column=2, sticky="e", pady=5)
frame3_status = ttk.Label(
    menu3_frame3,
    text="Not completed",
    # foreground="#ED254E",
    font=("Helvetica", 9, "bold"),
    style="danger",
)
frame3_status.grid(row=0, column=3, sticky="e")

# Frame 4 - Manager
menu3_frame4 = ttk.LabelFrame(frames["Delays Manager"], style="light")
menu3_frame4.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
# menu3_frame4.columnconfigure(0, weight=1)
menu3_frame4.columnconfigure(2, weight=1)
menu3_frame4.columnconfigure(3, weight=1)
menu3_frame4.columnconfigure(4, weight=1)
menu3_frame4.columnconfigure(5, weight=1)
ttk.Label(menu3_frame4, text="Start time:").grid(row=0, column=0, sticky="w", padx=15)
frame4_stime_entry = ttk.Entry(menu3_frame4)
frame4_stime_entry.grid(row=0, column=1, sticky="e", pady=2)

ttk.Label(menu3_frame4, text="End time:").grid(row=1, column=0, sticky="w", padx=15)
frame4_endtime_entry = ttk.Entry(menu3_frame4)
frame4_endtime_entry.grid(row=1, column=1, sticky="e", pady=2)

ttk.Label(menu3_frame4, text="Reason:").grid(
    row=2, column=0, sticky="w", padx=15, pady=2
)
frame4_reason_entry = ttk.Entry(menu3_frame4)
frame4_reason_entry.grid(row=2, column=1, sticky="we", pady=2, columnspan=3)

sep = ttk.Separator(menu3_frame4)
sep.grid(row=3, column=0, columnspan=5, sticky="we", pady=10)

# Workers
ttk.Label(menu3_frame4, text="Workers:").grid(row=4, column=0, sticky="w", padx=15)
for i, entry_name in enumerate(WORKER_ENTRIES, start=4):
    globals()[entry_name] = ttk.Entry(menu3_frame4)
    globals()[entry_name].grid(row=i, column=1, sticky="e", pady=2)

# Vehicles
ttk.Label(menu3_frame4, text="Vehicles:").grid(row=4, column=2, sticky="e")
v1_entry = ttk.Entry(menu3_frame4)
v1_entry.grid(row=4, column=3, sticky="e")


# Toolbar Frame
toolbar_frame = ttk.Frame(frames["Delays Manager"])
save_button = ttk.Button(
    toolbar_frame, text="Save", command=save_delay_wb, style="success", state="disabled"
)
save_button.pack(side=RIGHT, fill="both", expand=True)


toolbar_frame.grid(row=999, column=0, sticky="sew", columnspan=2)

transfer_button = ttk.Menubutton(
    toolbar_frame,
    text="Transfer",
    state="disabled",
)
transfer_button.pack(side=LEFT, fill="both", expand=True)
# Create Transfer menu
transfer_menu = ttk.Menu(transfer_button)
# Add items to our inside menu
item_var = StringVar()
transfer_menu.add_radiobutton(
    label="Weekly delay sheet", variable=item_var, command=transfer_data_to_weekly_delay
)
transfer_menu.add_radiobutton(
    label="Weekly cancelled sheet",
    variable=item_var,
    command=transfer_data_to_weekly_cancelled,
)
# Associate the inside menu with the menubutton
transfer_button["menu"] = transfer_menu

app.protocol("WM_DELETE_WINDOW", on_closing)
app.mainloop()
