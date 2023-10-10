def transfer_data(
    source_file, destination_file, mappings, dest_start_row=4, dest_sheet_name=None
):
    # Load the workbooks and worksheets
    src_wb = load_workbook(source_file)
    src_ws = src_wb.active

    dest_wb = load_workbook(destination_file)

    # Set the destination worksheet
    if dest_sheet_name and dest_sheet_name in dest_wb.sheetnames:
        dest_ws = dest_wb[dest_sheet_name]
    else:
        dest_ws = dest_wb.active  # default to the active sheet if none specified

    # Find the header mapping in the source file
    src_header = {}
    for col_num, cell in enumerate(src_ws[3]):
        if cell.value in mappings:
            src_header[cell.value] = col_num + 1

    # Find the header mapping in the destination file
    dest_header = {}
    for col_num, cell in enumerate(dest_ws[3]):
        if cell.value in mappings.values():
            dest_header[cell.value] = col_num + 1

    dest_row_counter = (
        dest_start_row  # Initializing destination row counter with the user input
    )

    # Transfer the data based on mapping
    for row in range(4, src_ws.max_row + 1):  # Always start from 4th row in the source
        for src_col, dest_col in mappings.items():
            if src_col in src_header and dest_col in dest_header:
                src_cell = src_ws.cell(row=row, column=src_header[src_col])
                dest_cell = dest_ws.cell(
                    row=dest_row_counter, column=dest_header[dest_col]
                )
                dest_cell.value = src_cell.value
                print(
                    f"Copied from Source(R{row}C{src_header[src_col]}) to Dest(R{dest_row_counter}C{dest_header[dest_col]})"
                )

        dest_row_counter += (
            1  # Increment the destination row counter after each row of data
        )

    dest_wb.save(destination_file)


def transfer_data_generic(mapping, dest_sheet):
    if delays_folder_path == "":
        messagebox.showerror(
            title="error", message="Please select the delays folder and try again."
        )
        return

    # Prompt the user for the starting row
    dest_start_row = simpledialog.askinteger(
        "Input", "Enter the starting row:", minvalue=4
    )

    # Check if dest_start_row is None (i.e., the dialog was closed without entering a value)
    if dest_start_row is None:
        return  # Exit the function

    dest_start_row = int(dest_start_row)

    if not dest_start_row:
        return  # exits

    str_date, dt_date, week_num = extract_date_from_path(delays_folder_path)
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
    transfer_data_generic(TO_WEEKLY_CANCELLED_MAPPING, "Work Cancelled")









    for row in range(4, src_ws.max_row + 1):  # Always start from 4th row in the source
    observation_cell = src_ws.cell(
        row=row, column=src_header.get("Observations", None)
        for src_col, dest_col in mappings.items():
            if src_col in src_header and dest_col in dest_header:
                src_cell = src_ws.cell(row=row, column=src_header[src_col])
                dest_cell = dest_ws.cell(
                    row=dest_row_counter, column=dest_header[dest_col]
                )
                dest_cell.value = src_cell.value
                print(
                    f"Copied from Source(R{row}C{src_header[src_col]}) to Dest(R{dest_row_counter}C{dest_header[dest_col]})"
                )

            if observation_cell:
                observation_value = observation_cell.value)
                if observation_value and "cancel" in observation_value.lower():
                # Proceed with the data transfer for this row.
                    for src_col, dest_col in mappings.items():
                if src_col in src_header and dest_col in dest_header:
                    src_cell = src_ws.cell(row=row, column=src_header[src_col])
                dest_cell = dest_ws.cell(
                    row=dest_row_counter, column=dest_header[dest_col]
                )
                dest_cell.value = src_cell.value
                print(
                    f"Copied from Source(R{row}C{src_header[src_col]}) to Dest(R{dest_row_counter}C{dest_header[dest_col]})"
                )



        dest_row_counter += (
            1  # Increment the destination row counter after each row of data
        )