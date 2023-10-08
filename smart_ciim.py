import os
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
from typing import Optional
from openpyxl.utils.exceptions import InvalidFileException
from tkinter import messagebox


def define_related_paths():
    """Define all paths relative to the global CIIM_FOLDER_PATH."""
    base_path = CIIM_FOLDER_PATH  # Use the global CIIM_FOLDER_PATH

    paths = {
        "general_updates": os.path.join(base_path, "General Updates"),
        "delays": os.path.join(base_path, "General Updates", "Delays+Cancelled works"),
        "passdown": os.path.join(base_path, "Pass Down"),
        "templates": os.path.join(
            base_path, "Important doc", "Empty reports (templates)"
        ),
        "ciim_report": os.path.join(base_path, "Important doc", "CIIM 2023 Report"),
    }
    # Normalize each path in the dictionary
    normalized_paths = {key: os.path.normpath(value) for key, value in paths.items()}
    return normalized_paths


def get_ciim_folder_path_from_file(file_path):
    """Retrieve the CIIM folder path from the given file path."""
    return os.path.dirname(os.path.dirname(os.path.dirname(file_path)))


def open_delays_folder():
    global delays_folder_path, tl_list, tl_list_internal

    # Constructing the initial path for the filedialog
    paths = define_related_paths()  # Get the Paths dictionary
    delays_path = paths["delays"]
    today = date.today()  # Get today's date
    current_year = today.year  # Get the current year
    current_week_number = today.isocalendar()[
        1
    ]  # Get the current week number (ISO week number)

    # Check for the existence of the folder for the current week
    # and decrement the week number until it finds an existing folder.
    while current_week_number > 0:  # Ensures the loop doesn't go below week 1
        initial_delays_path = os.path.join(
            delays_path, str(current_year), f"WW{current_week_number:02}"
        )

        if os.path.exists(initial_delays_path):
            break  # Exit the loop if the folder exists

        current_week_number -= 1  # Decrement the week number to check the previous week

    delays_folder_path = filedialog.askdirectory(
        title="Select the Delays folder", initialdir=initial_delays_path
    )
    delays_folder_path = os.path.normpath(delays_folder_path)

    # Check if the selected folder name matches the desired pattern
    pattern = re.compile(r"^\d{2}\.\d{2}\.\d{2}$")
    folder_name = os.path.basename(delays_folder_path)

    if not pattern.match(folder_name):
        messagebox.showerror(
            "Error", "Please select a folder with the pattern dd.mm.yy"
        )
        return  # exits

    tl_list = []
    tl_list_internal = os.listdir(delays_folder_path)
    for i in range(len(tl_list_internal)):
        tl_list.append(tl_list_internal[i][:-5])

    tl_list.sort()
    tl_listbox.delete(0, END)
    for tl_name in tl_list:
        tl_listbox.insert(END, tl_name)
        print(f"Loading : {tl_name}")

    print(f"The Delays folder Path is : {delays_folder_path}")


def open_const_wp():
    global cp_dates, construction_wp_path, CIIM_FOLDER_PATH

    # Pattern for the filename
    pattern = "WW*Construction Work Plan*.xlsx"

    construction_wp_path = filedialog.askopenfilename(
        filetypes=[("Excel Files", pattern)]
    )
    construction_wp_path = os.path.normpath(construction_wp_path)

    if not construction_wp_path:
        return  # Exit the function if no file was chosen

    try:
        construction_wp_workbook = load_workbook(filename=construction_wp_path)
        print(f"The Construction Plan Path is : {construction_wp_path}")

        CIIM_FOLDER_PATH = get_ciim_folder_path_from_file(construction_wp_path)
        print(f"The CIIM folder Path is : {CIIM_FOLDER_PATH}")

        construction_wp_worksheet = construction_wp_workbook["Const. Plan"]
        unique_dates = set()  # Use a set to keep track of unique dates
        for i in range(3, construction_wp_worksheet.max_row):
            cell_value = construction_wp_worksheet.cell(row=i, column=4).value
            if cell_value:
                # Add the date (minus the last 9 characters) to the set
                unique_dates.add(str(cell_value)[:-9])

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


def update_combo_list():
    dates_combobox["values"] = cp_dates


def get_filtered_team_leaders(construction_wp_worksheet, date):
    global TL_BLACKLIST

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
    global year, month, week, day, tl_index

    combo_selected_date = pd.Timestamp(dates_combobox.get())
    day, month, year = [combo_selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]]
    week = combo_selected_date.strftime("%U")

    construction_wp_workbook = load_workbook(
        filename=construction_wp_path, data_only=True
    )
    construction_wp_worksheet = construction_wp_workbook["Const. Plan"]

    team_leaders_list, tl_index = get_filtered_team_leaders(
        construction_wp_worksheet, combo_selected_date
    )

    dc_tl_listbox.delete(0, END)
    for tl_name in team_leaders_list:
        dc_tl_listbox.insert(END, tl_name)

    construction_wp_workbook.close()
    print(day, month, year)
    print(f'WW{week}')


def go(event):
    global team_leader_name
    cs = tl_listbox.curselection()
    tl_name_selected.config(text=tl_listbox.get(cs))
    team_leader_name = tl_listbox.get(cs)
    print(team_leader_name)
    clear_cells()
    load_from_excel()
    line_status()


def open_delay_file(delays):
    tl_listbox.curselection()
    os.startfile(delay_excel_path)


def dc_tl_selected(event):
    global dc_tl_name, tl_num
    dc_listbox_selection_index = dc_tl_listbox.curselection()
    dc_tl_listbox.itemconfig(dc_listbox_selection_index, bg="#ED969D")
    dc_tl_name = str(dc_tl_listbox.get(dc_listbox_selection_index))
    tl_num = tl_index[dc_listbox_selection_index[0]]
    create_delay()


def copy_and_rename_template(src_path, dest_path, new_name):
    shutil.copy(src_path, dest_path)
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


def fill_delay_ws_cells(delay_ws, cp_ws, team_leader_num):
    """Fill specific cells of delay_ws with fixed values or patterns."""
    cells_to_fill = {
        (16, 5): "Foreman",
        (17, 5): "Team Leader",
        (16, 7): "SEMI",
        (17, 7): "SEMI",
        (28, 2): "Y",
        (29, 2): "Y",
        (
            8,
            6,
        ): f"{cp_ws.cell(row=int(team_leader_num), column=7).value} to {cp_ws.cell(row=int(team_leader_num), column=8).value}",
        (
            8,
            4,
        ): f"{cp_ws.cell(row=int(team_leader_num), column=9).value} - {cp_ws.cell(row=int(team_leader_num), column=10).value}",
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


def create_delay():
    paths_dict = define_related_paths()
    dc_delays_f_path = Path(paths_dict["delays"])

    year_path = dc_delays_f_path / year
    week_path = year_path / f"WW{week}"
    day_folder = f"{day}.{month}.{year[2:]}"
    day_path = week_path / day_folder

    # Check for required paths and create if they don't exist
    for path in [year_path, week_path, day_path]:
        path.mkdir(parents=True, exist_ok=True)

    delay_report_template = "Delay Report template v.02.xlsx"
    dc_new_report_name = f"Delay Report {dc_tl_name} {day_folder}.xlsx"
    delay_wb_path = day_path / dc_new_report_name

    # Handle Template Copy & Rename
    if os.path.exists(delay_wb_path):
        status_msg = (
            f"Delay Report {dc_tl_name} {day_folder} already exists!\n{day_path}"
        )
        messagebox.showerror("Error", status_msg)
        print(f"Delay Report {dc_tl_name} already exists!")
    else:
        print(f"Creating Delay Report {dc_tl_name}")
        copy_and_rename_template(
            Path(paths_dict["templates"]) / delay_report_template,
            day_path,
            dc_new_report_name,
        )

        cp_wb = load_workbook(filename=construction_wp_path)
        cp_ws = cp_wb["Const. Plan"]
        delay_wb = load_workbook(filename=delay_wb_path)
        delay_ws = delay_wb["Sheet1"]

        copy_from_cp_to_delay(cp_ws, delay_ws, tl_num, day_folder)
        fill_delay_ws_cells(delay_ws, cp_ws, tl_num)

        delay_wb.save(str(delay_wb_path))

        status_msg = f"Delay Report {dc_tl_name} {day_folder} created!\n{day_path}"
        messagebox.showinfo(None, status_msg)

    # ciim_passdown_path = temp_ciim_folder_path / "Pass Down/*.xlsx"
    #
    # ciim_passdown_list = sorted(glob.glob(str(ciim_passdown_path)), key=os.path.getmtime,
    #                             reverse=True)
    # ciim_passdown_name = os.path.basename(ciim_passdown_list[0]).replace(".xlsx", "")[10:]
    # print(ciim_passdown_name)

    # delay_ws["D8"].fill = PatternFill(bgColor="FFFFFF")
    # delay_ws.cell(row=8, column=8).value = ciim_passdown_name
    # delay_ws["H8"].fill = PatternFill(bgColor="FFFFFF")


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


def load_from_excel():
    global delay_excel_workbook, delay_excel_path

    def insert_value(row, col, widget, time_format=False):
        cell_value = delay_excel_worksheet.cell(row=row, column=col).value
        if cell_value:
            if time_format and not isinstance(cell_value, str):
                cell_value = cell_value.strftime("%H:%M")
            if isinstance(cell_value, str):
                widget.insert(0, cell_value)

    try:
        delay_excel_workbook.close()
        print(
            f"Previous file closed: {delay_excel_path}"
        )  # Print the path of the closed workbook here
        delay_excel_path = delay_excel_path.rsplit("/", 1)[0]
        delay_excel_path = Path(delay_excel_path)
    except AttributeError:
        pass

    delay_excel_path = os.path.join(delays_folder_path, f"{team_leader_name}.xlsx")
    delay_excel_workbook = load_workbook(filename=delay_excel_path)
    delay_excel_worksheet = delay_excel_workbook["Sheet1"]

    mapping = get_cell_mapping()
    for widget, details in mapping.items():
        row = details["row"]
        col = details["col"]
        time_format = details.get("time_format", False)
        insert_value(row, col, widget, time_format)


def save_to_excel():
    global delay_excel_workbook
    full_file_path = os.path.join(delays_folder_path, f"{team_leader_name}.xlsx")
    delay_excel_workbook = load_workbook(filename=full_file_path)
    delay_excel_worksheet = delay_excel_workbook["Sheet1"]

    # Direct assignments using ENTRIES_CONFIG
    for entry_name, config in ENTRIES_CONFIG.items():
        cell_address = config["cell"]
        entry = globals()[entry_name]
        delay_excel_worksheet[cell_address] = entry.get()

    delay_excel_workbook.save(str(full_file_path))
    clear_cells()
    load_from_excel()
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


def line_status():
    for entry_name, config in ENTRIES_CONFIG.items():
        entry = globals()[entry_name]
        var_name = config["var"]
        set_entry_status(entry, var_name)

    if frame4_w1_entry.get() == "" and frame4_workers_var.get() == 0:
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
        if os.path.exists(path):
            print(f"Path exists: {path}")
            return True
        else:
            print(f"Path does not exist : {path}")
            return False
    except Exception as e:
        print(f"An error occurred while checking the path: {e}")
        return False


def create_path_if_not_exists(path, label=None, message=None, **config_options):
    """Utility function to create a directory if it doesn't exist and optionally update a label."""
    path = os.path.normpath(path)  # Normalize the path
    if not os.path.exists(path):
        os.makedirs(path)
        if label and message:
            set_config(label, text=message, **config_options)
    return path


def derive_paths_from_date(selected_date):
    """Derive all related paths from a given date including multiple date formats."""
    c_day, c_month, c_year = [selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]]
    p_date_datetime = selected_date - timedelta(days=1)
    p_day, p_month, p_year = [p_date_datetime.strftime(pattern) for pattern in ["%d", "%m", "%Y"]]

    c_week = selected_date.strftime("%U")  # returns the week number considering the first day of the week as Sunday

    c_formatted_dates = {
        "slash": f"{c_day}/{c_month}/{c_year[-2:]}",
        "dot": f"{c_day}.{c_month}.{c_year[-2:]}",
        "compact": f"{c_year[-2:]}{c_month}{c_day}"
    }

    p_formatted_dates = {
        "slash": f"{p_day}/{p_month}/{p_year[-2:]}",
        "dot": f"{p_day}.{p_month}.{p_year[-2:]}",
        "compact": f"{p_year[-2:]}{p_month}{p_day}"
    }

    paths = {
        "year": os.path.join(CIIM_FOLDER_PATH, f"Working Week {c_year}"),
        "week": os.path.join(CIIM_FOLDER_PATH, f"Working Week {c_year}", f"Working Week N{c_week}"),
        "day": os.path.join(CIIM_FOLDER_PATH, f"Working Week {c_year}", f"Working Week N{c_week}",
                            f"{c_year[-2:]}{c_month}{c_day}"),
        "previous_year": os.path.join(CIIM_FOLDER_PATH, f"Working Week {p_year}"),
        "previous_week": os.path.join(CIIM_FOLDER_PATH, f"Working Week {p_year}", f"Working Week N{c_week}"),
        "previous_day": os.path.join(CIIM_FOLDER_PATH, f"Working Week {p_year}", f"Working Week N{c_week}",
                                     f"{p_year[-2:]}{p_month}{p_day}"),
    }

    # Normalize the paths
    for key, value in paths.items():
        paths[key] = os.path.normpath(value)

    return paths, c_formatted_dates, p_formatted_dates


def pick_date():
    global selected_date
    cal = Querybox()
    selected_date = cal.get_date(bootstyle="danger")
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(selected_date)

    # Feedback using button's text
    calendar_button.config(text=f"WW: {selected_date.strftime('%U')}     Date: {selected_date.strftime('%d.%m.%Y')} ")

    day_message_exist = f'{c_formatted_dates["compact"]} folder already exists'
    if os.path.exists(paths["day"]):
        messagebox.showerror("Error", day_message_exist)

    entries_state = "disabled" if os.path.exists(c_formatted_dates["compact"]) else "normal"
    set_config(fc_ocs_entry, state=entries_state)
    set_config(fc_scada_entry, state=entries_state)
    set_config(create_button, state=entries_state)

    return paths


def create_folders():
    # Importing the paths and the formatted dates
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(selected_date)
    main_paths = define_related_paths()

    create_path_if_not_exists(paths["year"])
    create_path_if_not_exists(paths["week"])
    create_path_if_not_exists(paths["day"])

    if os.path.exists(paths["day"]):
        day_created_message = (
            f'{c_formatted_dates["compact"]} folder was created successfully'
        )
        messagebox.showinfo(None, day_created_message)

    fc_ciim_report_name = (
        f'CIIM Report Table {c_formatted_dates["dot"]}.xlsx'.strip()
    )
    print(f"Generated report name: {fc_ciim_report_name}")

    templates_path = Path(main_paths["templates"])
    fc_ciim_template_path = os.path.join(templates_path, "CIIM Report Table v.1.xlsx")

    # Copy and rename
    print(f'Copying template to: {paths["day"]}')
    shutil.copy(fc_ciim_template_path, paths["day"])

    new_report_path = os.path.join(paths["day"], fc_ciim_report_name)
    print(f"Renaming file to: {new_report_path}")
    if os.path.exists(os.path.join(paths["day"], "CIIM Report Table v.1.xlsx")):
        os.rename(os.path.join(paths["day"], "CIIM Report Table v.1.xlsx"), new_report_path)
        # Print the list of files in the directory for verification
        print("Files in directory after renaming:")
        print(os.listdir(paths["day"]))

        # Introduce a slight delay
        time.sleep(1)
    else:
        print(f'Template not found in {paths["day"]}!')

    for i in range(int(fc_ocs_entry.get() or 0)):
        create_path_if_not_exists(os.path.join(paths["day"], f"W{i + 1}", "Pictures"))
        create_path_if_not_exists(os.path.join(paths["day"], f"W{i + 1}", "Worklogs"))

    for i in range(int(fc_scada_entry.get() or 0)):
        create_path_if_not_exists(os.path.join(paths["day"], f"S{i + 1}", "Pictures"))
        create_path_if_not_exists(os.path.join(paths["day"], f"S{i + 1}", "Worklogs"))

    folders_to_create = [
        "Foreman",
        "Track possession",
        "TS Worklogs",
        "PDF Files",
        "Worklogs",
    ]
    for folder in folders_to_create:
        create_path_if_not_exists(os.path.join(paths["day"], folder))

    fc_ocs_entry.delete(0, END)
    fc_scada_entry.delete(0, END)

    set_config(fc_ocs_entry, state="disabled")
    set_config(fc_scada_entry, state="disabled")
    set_config(create_button, state="disabled")

    #
    if os.path.dirname(construction_wp_path) != paths["week"]:
        print("Not copying works to the selected date")
        return

    write_data_to_report(construction_wp_path, c_formatted_dates["slash"], paths["day"], TO_DAILY_REPORT_MAPPINGS)

    # Only show the popup if previous day path exists
    if check_path_exists(paths["previous_day"]):  # Use the new function here
        result = messagebox.askyesno(title=None,
                                     message=f'Copy to CIIM Report Table {p_formatted_dates["dot"]} as well?')
        print(result)
        if result is True:
            popup_question()


def popup_func():
    paths, c_formatted_dates, p_formatted_dates = derive_paths_from_date(selected_date)

    write_data_to_previous_report(
        construction_wp_path, p_formatted_dates["slash"], paths["previous_day"], TO_DAILY_REPORT_MAPPINGS
    )
    global top
    if top:
        top.destroy()
        top = None  # Reset it to None after destroying to avoid future issues


def popup_question():
    global previous_day_entry, top
    top = ttk.Toplevel(app)
    # Assuming 'app' is the instance of your main window
    x = app.winfo_x()
    y = app.winfo_y()
    top.geometry(f"+{x + 200}+{y + 100}")
    top.title(None)
    previous_day_label = ttk.Label(top, text="Row number")
    previous_day_label.pack(side="left", padx=5, pady=5)
    previous_day_entry = ttk.Entry(top, width=8)
    previous_day_entry.pack(side="left", padx=5, pady=5)
    previous_button = ttk.Button(top, text="Ok", command=popup_func, width=5, )
    previous_button.pack(side="bottom", padx=5, pady=5, anchor="center")


# Utility Functions
def filter_by_date(df, date_column, target_date):
    """Filter a DataFrame based on a target date."""
    return df[df[date_column] == target_date]


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
    report_filename = f"CIIM Report Table {formatted_target_date}.xlsx"
    target_report_path = os.path.join(target_directory, report_filename)

    # Using mappings to determine columns to load
    usecols_value = [mappings[header] for header in mappings.keys()]
    df = pd.read_excel(src_path, sheet_name="Const. Plan", skiprows=1, usecols=usecols_value)

    # Filter data
    target_df = filter_by_date(df, "Date [DD/MM/YY]", target_datetime)

    # Write to target workbook
    target_workbook = load_workbook(filename=target_report_path)
    target_worksheet = target_workbook.active

    # Write headers (using the mappings keys as headers)
    for col, header in enumerate(mappings.keys(), 2):  # Starting from column B
        target_worksheet.cell(row=start_row - 1, column=col, value=header)

    # Write data
    for row_idx, (index, row_data) in enumerate(target_df.iterrows(), start=start_row):
        for col_idx, header in enumerate(mappings.keys(), 2):  # Starting from column B
            target_worksheet.cell(row=row_idx, column=col_idx, value=row_data[header])

    # Format Date column
    date_col_idx = list(mappings.keys()).index("Date [DD/MM/YY]") + 2
    format_datetime_column(target_worksheet, date_col_idx, start_row, target_worksheet.max_row)

    # Format Observations column (assuming "Observations" is always a key in your mappings)
    observations_col_idx = list(mappings.keys()).index("Observations") + 2
    format_observations_column(target_worksheet, observations_col_idx, start_row, target_worksheet.max_row)

    target_workbook.save(target_report_path)
    print(f"Report for {formatted_target_date} has been updated and saved.")


def write_data_to_report(src_path, target_date, target_directory, mappings):
    write_data_to_excel(src_path, target_date, target_directory, mappings)


def write_data_to_previous_report(src_path, target_date, target_directory, mappings):
    user_input = (previous_day_entry.get())
    start_row = int(user_input)

    write_data_to_excel(src_path, target_date, target_directory, mappings, start_row=start_row)


def refresh_delays_folder():
    global delays_folder_path, tl_list, tl_list_internal
    tl_list = []
    tl_list_internal = os.listdir(delays_folder_path)
    for i in range(len(tl_list_internal)):
        tl_list.append(tl_list_internal[i][:-5])
        tl_list.sort()
    print(tl_list)
    tl_listbox.delete(0, END)
    for name in tl_list:
        tl_listbox.insert(END, name)


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
    themename="lumen", size=(768, 512), resizable=(0, 0), title="Smart CIIM"
)

# Variables
# Paths
CIIM_FOLDER_PATH: Optional[str] = None
# Tkinter variables
team_leader_name = ""
status_color = IntVar()
previous_day_entry = IntVar()
day, month, week, year = StringVar(), StringVar(), StringVar(), StringVar()
start_time, end_time, reason_var, worker1_var, vehicle1_var = 0, 0, 0, 0, 0
combo_selected_date = ""
# File and folder paths
delays_folder_path = ""
construction_wp_path = ""
delay_excel_path = ""
selected_date = None
# Lists and associated data
tl_list = []
cp_dates = []
tl_list_internal = []
tl_index = []
# Miscellaneous variables
dc_tl_name = ""
tl_num = ""
delay_excel_workbook = None
top = None
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
    "Discipline [OCS/Old Bridges/TS/SCADA & COM]",
    "WW [Nº]",
    "Date [DD/MM/YY]",
    "T.P Start [Time]",
    "T.P End [Time]",
    "T.P Start [K.P]",
    "T.P End [K.P]",
    "EP#",
    "ISR Start Section [Name]",
    "ISR  End Section [Name]",
    "Foremen [Israel]",
    "Team Name",
    "Team Leader Name (Phone)",
    "Work Description (Baseline)",
    "ISR Safety Request",
    "ISR Comm&Rail:",
    "ISR T.P request (All/Track number)",
    "Observations"
]
HEADER_TO_INDEX = {header: index for index, header in enumerate(CONSTRUCTION_WP_HEADERS)}
# All the Headers from the Construction Work Plan match the CIIM Report Table
TO_DAILY_REPORT_MAPPINGS = {
    header: header for header in CONSTRUCTION_WP_HEADERS}

DAILY_REPORT_HEADERS = [
    "WW [Nº]",
    "Discipline [OCS/Old Bridges/TS/SCADA & COM]",
    "Date [DD/MM/YY]",
    "Delay details (comments + description)",
    "Team Name",
    "Team Leader Name (Phone)",
    "EP",
    "T.P Start [Time]",
    "Actual Start Time (TL):",
    "T.P End [Time]",
    "Actual Finish Time (TL):",
    "Number of workers",
    "Work Description",
    "Observations"
]
DAILY_REPORT_HEADERS = {header: index for index, header in enumerate(DAILY_REPORT_HEADERS)}
TO_WEEKLY_DELAY_MAPPINGS = {
    "WW [Nº]": "WW",
    "Discipline [OCS/Old Bridges/TS/SCADA & COM]": "Discipline [OCS, Scada, TS]",
    "Delay details (comments + description)": "Reason",
    "Date [DD/MM/YY]": "Date",
    "Team Name": "Team Name",
    "Team Leader Name (Phone)": "Team leader ",
    "EP": "ISR section {EP}",
    "T.P Start [Time]": "TP Start Time (TAK)",
    "Actual Start Time (TL):": "Actual Start Time (Real Start time - TL)",
    "T.P End [Time]": "TP Finish Time (TAK)",
    "Actual Finish Time (TL):": "Actual Finish Time (Real Finish time - TL)",
    "Number of workers": "Workers"
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
    "frame4_w1_entry": {"cell": "A18", "var": "worker1_var", "row": 18, "col": 1},
    "frame4_w2_entry": {"cell": "A19", "var": "worker2_var", "row": 19, "col": 1},
    "frame4_w3_entry": {"cell": "A20", "var": "worker3_var", "row": 20, "col": 1},
    "frame4_w4_entry": {"cell": "A21", "var": "worker4_var", "row": 21, "col": 1},
    "frame4_w5_entry": {"cell": "A22", "var": "worker5_var", "row": 22, "col": 1},
    "frame4_w6_entry": {"cell": "A23", "var": "worker6_var", "row": 23, "col": 1},
    "frame4_w7_entry": {"cell": "A24", "var": "worker7_var", "row": 24, "col": 1},
    "frame4_w8_entry": {"cell": "A25", "var": "worker8_var", "row": 25, "col": 1},
    "frame4_v1_entry": {"cell": "D28", "var": "vehicle1_var", "row": 28, "col": 4},
}

WORKER_ENTRIES = [
    "frame4_w1_entry",
    "frame4_w2_entry",
    "frame4_w3_entry",
    "frame4_w4_entry",
    "frame4_w5_entry",
    "frame4_w6_entry",
    "frame4_w7_entry",
    "frame4_w8_entry",
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
    "Delays Manager": {"columns": [(1, 9)], "rows": [(0, 1), (1, 8)]},
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
    frames["Start Page"], text="Get Started!", command=open_const_wp, width=20, style='success'
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
dc_tl_listbox.bind("<Double-1>", dc_tl_selected)

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
create_and_grid_label(menu2_frame2, "OCS works:", 1, 1, "e", 10, 10)
fc_ocs_entry = create_and_grid_entry(menu2_frame2, 1, 2, "w", 10)
fc_ocs_entry.config(state="disabled", width=8)
create_and_grid_label(menu2_frame2, "SCADA works:", 2, 1, "e", 10, 20)
fc_scada_entry = create_and_grid_entry(menu2_frame2, 2, 2, "w", 10)
fc_scada_entry.config(state="disabled", width=8)

create_button = ttk.Button(
    menu2_frame2, text="Create", command=create_folders, state="disabled", width=8
)
create_button.grid(row=4, column=2, sticky="es", pady=10)

# Menu 2 - Delays Manager
# Frame 1 - Folder select
menu3_frame1 = ttk.LabelFrame(frames["Delays Manager"], style="light")
menu3_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=15)
delay_folder_button = ttk.Button(
    menu3_frame1, text="Select Delays Folder", command=open_delays_folder, width=25
)
delay_folder_button.pack()
# Frame 2 - Team Leaders Listbox
menu3_frame2 = ttk.LabelFrame(
    frames["Delays Manager"],
    text="Team Leaders",
)
menu3_frame2.grid(row=1, column=0, sticky="wens", padx=5, pady=5)
tl_listbox = Listbox(menu3_frame2, bd=0, width=40)
tl_listbox.pack(fill="both", expand=True)
tl_listbox.bind("<Double-1>", go)
tl_listbox.bind("<Double-3>", open_delay_file)
# Frame 3 - Name + Status
menu3_frame3 = ttk.LabelFrame(
    frames["Delays Manager"],
    text="Status",
    style="warning",
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
frame4_v1_entry = create_and_grid_entry(menu3_frame4, 4, 3, "e")

# check boxes
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

refresh_button = ttk.Button(
    frames["Delays Manager"], text="Refresh", command=refresh_delays_folder, width=8
)
refresh_button.place(anchor=CENTER, relx=0.82, rely=0.94)

save_button = ttk.Button(
    frames["Delays Manager"],
    text="Save",
    command=save_to_excel,
    style="success",
    width=8,
)
save_button.place(anchor=CENTER, relx=0.93, rely=0.94)

app.mainloop()
