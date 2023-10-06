import os
import shutil
import fnmatch
import pandas as pd
import re
from tkinter import Tk, filedialog, messagebox
import ttkbootstrap as ttk


def convert_decimal_to_time(value):
    if pd.notna(value):  # Skip conversion for NaN values
        days = int(value)
        hours = int((value - days) * 24)
        minutes = int(((value - days) * 24 * 60) % 60)
        seconds = int(((value - days) * 24 * 60 * 60) % 60)
        return pd.Timestamp('1900-01-01').replace(day=1, hour=hours, minute=minutes, second=seconds).round(
            freq='T').time()
    else:
        return None


def copy_matching_files(src_folder, dst_folder, pattern):
    copied_files = []
    for dirpath, dirnames, filenames in os.walk(src_folder):
        for filename in fnmatch.filter(filenames, pattern):
            source_file_path = os.path.join(dirpath, filename)
            destination_file_path = os.path.join(dst_folder, filename)
            shutil.copy2(source_file_path, destination_file_path)
            copied_files.append(destination_file_path)
    return copied_files


def process_file(file_path, csv_dst):
    df = pd.read_excel(file_path, skiprows=2, usecols='B, D:O, S, U, W:AB, AD:AE')
    if pd.api.types.is_numeric_dtype(df['Date [DD/MM/YY]']):
        df['Date [DD/MM/YY]'] = pd.to_datetime(df['Date [DD/MM/YY]'], unit='D', origin='1899-12-30')
    else:
        df['Date [DD/MM/YY]'] = pd.to_datetime(df['Date [DD/MM/YY]'], format='%d/%m/%Y', errors='coerce')

    for column in TIME_COLUMNS:
        if pd.api.types.is_numeric_dtype(df[column]):
            df[column] = df[column].apply(convert_decimal_to_time)

    filename = os.path.basename(file_path).replace(".xlsx", ".csv")
    output_file = os.path.join(csv_dst, filename)
    df.to_csv(output_file, index=False, encoding='utf-8')


def extract_and_convert_to_csv():
    pattern = "Working Week N\d+"  # Corrected the pattern
    src = filedialog.askdirectory(title="Select the Working Week folder")
    dst = filedialog.askdirectory(title="Select Desktop")

    excel_dst = os.path.join(dst, "Excel Files")
    csv_dst = os.path.join(dst, "CSV Files")

    os.makedirs(excel_dst, exist_ok=True)
    os.makedirs(csv_dst, exist_ok=True)

    compiled_pattern = re.compile(pattern)
    folder_name = os.path.basename(src)

    if not compiled_pattern.match(folder_name):
        messagebox.showerror("Error", "Please select the Working Week folder")
        return

    copied_files = copy_matching_files(src, excel_dst, "CIIM Report Table *.xlsx")

    for file_path in copied_files:
        process_file(file_path, csv_dst)


TIME_COLUMNS = [
    'T.P Start [Time]',
    'T.P End [Time]',
    'Actual Start Time (TL):',
    'Actual Finish Time (TL):',
    'Difference',
    'Actual work time'
]

root = Tk()
root.geometry("200x200")

csv_conv_label = ttk.Label(root, text="Holla Senior")
csv_conv_label.pack(pady=20)
btn_convert = ttk.Button(root, text='Choose folder', command=extract_and_convert_to_csv, width=25, style='Success')
btn_convert.pack(pady=20)

root.mainloop()
