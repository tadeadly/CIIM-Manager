def check_path_exists(path):
    """Check if a given path exists and print a message."""
    try:
        if os.path.exists(path):
            print(f"Path exists: {path}")
            return True
        else:
            print(f"Path does NOT exist: {path}")
            return False
    except Exception as e:
        print(f"An error occurred while checking the path: {e}")
        return False


def pick_date():
    global c_day, c_month, c_year, c_date, c_week, c_date, p_date, p_day, p_month, p_year
    cal = Querybox()

    selected_date = cal.get_date(bootstyle="danger")
    c_day = selected_date.strftime("%d")
    c_month = selected_date.strftime("%m")
    c_year = selected_date.strftime("%Y")
    c_date = selected_date.strftime("%d.%m.%Y")
    ts = pd.Timestamp(selected_date)

    # Calculate previous date which is a day before c_date
    p_date_datetime = selected_date - timedelta(days=1)
    p_date = p_date_datetime.strftime("%d.%m.%Y")
    p_day = p_date_datetime.strftime("%d")
    p_month = p_date_datetime.strftime("%m")
    p_year = p_date_datetime.strftime("%Y")
    print(p_date)
    (iso_year, c_week, iso_weekday) = ts.isocalendar()
    if iso_weekday == 7:
        c_week += 1

    # Using the button's text  to provide feedback
    calendar_button.config(text=f"WW: {c_week}     Date: {c_date} ")

    c_week = str(c_week).zfill(2)  # Add leading zero if necessary

    week_path = os.path.join(
        CIIM_FOLDER_PATH, f"Working Week {c_year}", f"Working Week N{c_week}"
    )
    day_path = os.path.join(week_path, f"{c_year[-2:]}{c_month}{c_day}")
    print(week_path)
    print(day_path)

    day_message_exist = f"{c_year[-2:]}{c_month}{c_day} folder already exists"
    if os.path.exists(day_path):
        messagebox.showerror("Error", day_message_exist)

    entries_state = "disabled" if os.path.exists(day_path) else "normal"
    set_config(fc_ocs_entry, state=entries_state)
    set_config(fc_scada_entry, state=entries_state)
    set_config(create_button, state=entries_state)


def create_path_if_not_exists(path, label=None, message=None, **config_options):
    """Utility function to create a directory if it doesn't exist and optionally update a label."""
    path = os.path.normpath(path)  # Normalize the path
    if not os.path.exists(path):
        os.makedirs(path)
        if label and message:
            set_config(label, text=message, **config_options)
    return path


def create_folders():
    paths_dict = define_related_paths()

    year_path = create_path_if_not_exists(
        os.path.join(CIIM_FOLDER_PATH, f"Working Week {c_year}")
    )

    week_path = create_path_if_not_exists(
        os.path.join(year_path, f"Working Week N{c_week}")
    )

    day_path = create_path_if_not_exists(
        os.path.join(week_path, f"{c_year[-2:]}{c_month}{c_day}")
    )
    if os.path.exists(day_path):
        day_created_message = (
            f"{c_year[-2:]}{c_month}{c_day} folder was created successfully"
        )
        messagebox.showinfo(None, day_created_message)

    fc_ciim_report_name = (
        f"CIIM Report Table {c_day}.{c_month}.{c_year[-2:]}.xlsx".strip()
    )

    print(f"Generated report name: {fc_ciim_report_name}")

    templates_path = Path(paths_dict["templates"])
    fc_ciim_template_path = os.path.join(templates_path, "CIIM Report Table v.1.xlsx")

    # Copy and rename
    print(f"Copying template to: {day_path}")
    shutil.copy(fc_ciim_template_path, day_path)

    new_report_path = os.path.join(day_path, fc_ciim_report_name)
    print(f"Renaming file to: {new_report_path}")
    if os.path.exists(os.path.join(day_path, "CIIM Report Table v.1.xlsx")):
        os.rename(os.path.join(day_path, "CIIM Report Table v.1.xlsx"), new_report_path)

        # Print the list of files in the directory for verification
        print("Files in directory after renaming:")
        print(os.listdir(day_path))

        # Introduce a slight delay
        time.sleep(1)
    else:
        print(f"Template not found in {day_path}!")

    for i in range(int(fc_ocs_entry.get() or 0)):
        create_path_if_not_exists(os.path.join(day_path, f"W{i + 1}", "Pictures"))
        create_path_if_not_exists(os.path.join(day_path, f"W{i + 1}", "Worklogs"))

    for i in range(int(fc_scada_entry.get() or 0)):
        create_path_if_not_exists(os.path.join(day_path, f"S{i + 1}", "Pictures"))
        create_path_if_not_exists(os.path.join(day_path, f"S{i + 1}", "Worklogs"))

    folders_to_create = [
        "Foreman",
        "Track possession",
        "TS Worklogs",
        "PDF Files",
        "Worklogs",
    ]
    for folder in folders_to_create:
        create_path_if_not_exists(os.path.join(day_path, folder))

    c_date_slash = c_date.replace(".", "/")

    if not os.path.exists(new_report_path):
        print(f"Expected file {new_report_path} not found!")
    write_data_to_report(construction_wp_path, c_date_slash, day_path)

    fc_ocs_entry.delete(0, END)
    fc_scada_entry.delete(0, END)

    set_config(fc_ocs_entry, state="disabled")
    set_config(fc_scada_entry, state="disabled")
    set_config(create_button, state="disabled")

    previous_year_path = create_path_if_not_exists(
        os.path.join(CIIM_FOLDER_PATH, f"Working Week {p_year}"))

    previous_week_path = create_path_if_not_exists(
        os.path.join(previous_year_path, f"Working Week N{c_week}"))

    previous_day_path = create_path_if_not_exists(
        os.path.join(previous_week_path, f"{p_year[-2:]}{p_month}{p_day}"))

    # Only show the popup if previous day path exists
    if check_path_exists(previous_day_path):  # Use the new function here
        result = messagebox.askyesno(title=None, message=f"Copy to CIIM Report Table {p_date} as well?")
        print(result)
        if result is True:
            popup_question()


# TODO : FIX IT AND THE LOGIC
def popup_func():
    year_path = create_path_if_not_exists(
        os.path.join(CIIM_FOLDER_PATH, f"Working Week {p_year}")
    )

    week_path = create_path_if_not_exists(
        os.path.join(year_path, f"Working Week N{c_week}")
    )

    day_path = create_path_if_not_exists(
        os.path.join(week_path, f"{p_year[-2:]}{p_month}{p_day}")
    )
    previous_date_slash = p_date.replace(".", "/")
    previous_day_path = os.path.join(week_path, f"{p_year[-2:]}{p_month}{p_day}")
    os.path.normpath(previous_day_path)

    write_data_to_previous_report(
        construction_wp_path, previous_date_slash, previous_day_path
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


def write_data_to_excel(src_path, target_date, target_directory, start_row=4):
    # Convert target_date string to datetime object
    target_datetime = pd.to_datetime(target_date, format="%d/%m/%Y", errors="coerce")

    if not pd.isna(
            target_datetime
    ):  # Check if target_datetime is a valid datetime object
        # Format the target_date to DD.MM.YY
        formatted_target_date = target_datetime.strftime("%d.%m.%y")

        # Construct the report filename
        report_filename = f"CIIM Report Table {formatted_target_date}.xlsx"

        # Combine the directory path with the filename to get the complete path
        target_report_path = os.path.join(target_directory, report_filename)

        print(f"Creating : {report_filename} at {target_directory}")

        if os.path.exists(src_path):
            df = pd.read_excel(
                src_path,
                sheet_name="Const. Plan",
                skiprows=1,
                usecols="B:J, K:N, P, R:S, U, AF",
            )
        else:
            print(f"File not found: {src_path}")
            return

        new_column_order = [
            0,
            1,
            2,
            3,
            4,
            5,
            6,
            7,
            8,
            16,
            9,
            10,
            11,
            12,
            13,
            14,
            15,
            17,
        ]
        df = df.iloc[:, new_column_order]

        # Filter rows where the date matches target_date
        target_df = df[df["Date [DD/MM/YY]"] == target_datetime]

        if not os.path.exists(target_report_path):
            print(f"Report file does not exist: {target_report_path}")
            return

        target_workbook = load_workbook(filename=target_report_path)
        target_worksheet = target_workbook.active
        data = [target_df.columns.tolist()] + target_df.values.tolist()

        for row_idx, row_data in enumerate(data[1:], start=start_row):
            for col_idx, cell_value in enumerate(
                    row_data, start=2
            ):  # Column B corresponds to index 2
                target_worksheet.cell(row=row_idx, column=col_idx, value=cell_value)

        # Iterate through column S starting from row 4
        for row_idx in range(4, target_worksheet.max_row + 1):
            cell_value = target_worksheet.cell(
                row=row_idx, column=19
            ).value  # Column S corresponds to index 19
            if isinstance(cell_value, float):
                # Convert float value to a string
                cell_value = str(cell_value)
            if cell_value and re.search(r"Need.*", cell_value):
                # Replace the cell content with text
                target_worksheet.cell(
                    row=row_idx,
                    column=19,
                    value="Work Details:\n- TL arrived to the field.\n- TL sent toolbox.\nSummary:",
                )

        target_workbook.save(target_report_path)

        print(f"Report for {formatted_target_date} has been updated and saved.")
    else:
        print(f"Invalid target_datetime: {target_date}")


def write_data_to_report(src_path, target_date, target_directory):
    write_data_to_excel(src_path, target_date, target_directory)


def write_data_to_previous_report(src_path, target_date, target_directory):
    user_input = (
        previous_day_entry.get().strip()
    )  # Get the user input and remove any whitespace

    # Check if the input is empty or not a number
    if not user_input.isdigit():
        print(f"Not editing {p_date}")
        return  # Exit the function

    start_row = int(user_input)  # Convert user input to integer

    write_data_to_excel(src_path, target_date, target_directory, start_row=start_row)

    # Those TLs won't appear in the Listbox that creates delays
    tl_blacklist = [
        "Eliyau Ben Zgida",
        "Emerson Gimenes Freitas",
        "Emilio Levy",
        "Samuel Lakko",
        "Ofer Akian",
        "Wissam Hagay",
        "Rami Arami",
    ]

    combo_selected_date = pd.Timestamp(dates_combobox.get())
    day, month, year = [combo_selected_date.strftime(pattern) for pattern in ["%d", "%m", "%Y"]]

    week = combo_selected_date.strftime("%U")  # returns the week number considering the first day of the week as Sunday

    construction_wp_workbook = load_workbook(
        filename=construction_wp_path, data_only=True
    )
    construction_wp_worksheet = construction_wp_workbook["Const. Plan"]

    team_leaders_list, tl_index = get_filtered_team_leaders(
        construction_wp_worksheet, combo_selected_date, tl_blacklist
    )

    dc_tl_listbox.delete(0, END)
    for tl_name in team_leaders_list:
        dc_tl_listbox.insert(END, tl_name)

    construction_wp_workbook.close()
