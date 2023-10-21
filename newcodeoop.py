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


class ConfigManager:
    # Initialize data attributes (previously global variables)
    def __init__(self):
        self.CIIM_FOLDER_PATH = "/"
        self.construction_wp_path = "/"
        self.delays_dir_path = "/"
        self.username = ""

    def define_related_paths(self):
        base_path = Path(self.CIIM_FOLDER_PATH)
        paths = {
            "delays": base_path / "General Updates" / "Delays+Cancelled works",
            "passdown": base_path / "Pass Down",
            "templates": base_path / "Important doc" / "Empty reports (templates)",
        }
        return paths

    def select_wp_path(self):
        pattern = "WW*Construction Work Plan*.xlsx"
        path = filedialog.askopenfilename(filetypes=[("Excel Files", pattern)])
        return Path(path) if path else None

    def get_latest_username_from_file(self):
        paths = self.define_related_paths()
        passdown_path = paths["passdown"]

        files = sorted(
            passdown_path.glob("*.xlsx"),
            key=lambda x: x.stat().st_mtime,
            reverse=True,
        )

        if files:
            filename = files[0].stem
            match = re.search(r"\d{6}\.\d+\s+(\w+)", filename)
            temp_username = match.group(1)
            return temp_username if match else None
        return None

    def get_ciim_folder_path_from_file(self, file_path):
        path = Path(file_path)
        return path.parent.parent.parent


class MainApp(ttk.Window):
    def __init__(self):
        super().__init__()
        self.title("Smart CIIM")
        self.geometry("776x522")
        self.resizable(0, 0)
        self.iconbitmap("icon.ico")

        # Create a dictionary of frames.
        self.frames = {
            "Create Delays": CreateDelaysFrame(self),
            "Create Folders": CreateFoldersFrame(self),
            "Manage Delays": ManageDelaysFrame(self),
        }

        # You could pack or grid one as a default, or leave it as per your design.
        self.frames["Create Delays"].pack(fill=BOTH, expand=True)

        self.config(menu=MenuBar(self))

    def show_frame(self, frame_key):
        for frame in self.frames.values():
            frame.pack_forget()
        self.frames[frame_key].pack(fill=BOTH, expand=True)


class CreateDelaysFrame(ttk.Frame):
    def __init__(self, parent):
        cp_dates = []

        super().__init__(parent)

        # Frame 1 - Date select
        menu1_frame1 = ttk.LabelFrame(self, text="", style="light")
        menu1_frame1.grid(row=0, column=0, sticky="wens", padx=5, pady=5)
        dc_select_date_label = ttk.Label(menu1_frame1, text="   Select date:  ")
        dc_select_date_label.pack(side="left")
        dates_combobox = ttk.Combobox(
            menu1_frame1, values=cp_dates, postcommand=update_combo_list
        )
        # dates_combobox.set("Date")
        # dates_combobox.bind("<<ComboboxSelected>>", dc_combo_selected)
        # dates_combobox.pack(side="left")

    def update_combo_list(self):
        pass


class CreateFoldersFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)


class ManageDelaysFrame(ttk.Frame):
    def __init__(self, parent):
        super().__init__(parent)


class MenuBar(ttk.Menu):
    def __init__(self, main_app):
        super().__init__(main_app)

        # Initialize menus with access to MainApp
        self.create_menu = CreateMenu(self, main_app)
        self.edit_menu = EditMenu(self, main_app)
        self.settings_menu = SettingsMenu(self)

        # Add cascades
        self.add_cascade(label="File", menu=self.create_menu)
        self.add_cascade(label="Edit", menu=self.edit_menu)
        self.add_cascade(label="Settings", menu=self.settings_menu)


class CreateMenu(ttk.Menu):
    def __init__(self, parent, main_app):
        super().__init__(parent, tearoff=0)
        self.add_command(
            label="New file", command=lambda: main_app.show_frame("Create Delays")
        )
        self.add_command(
            label="New folder", command=lambda: main_app.show_frame("Create Folders")
        )
        self.add_separator()
        self.add_command(label="Exit", command=parent.quit)


class EditMenu(ttk.Menu):
    def __init__(self, parent, main_app):
        super().__init__(parent, tearoff=0)
        self.add_command(
            label="Manage", command=lambda: main_app.show_frame("Manage Delays")
        )


class SettingsMenu(ttk.Menu):
    def __init__(self, parent):
        super().__init__(parent, tearoff=0)
        theme_menu = ttk.Menu(self, tearoff=0)
        self.add_cascade(label="Appearance", menu=theme_menu)


app = MainApp()
app.mainloop()
