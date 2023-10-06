import ttkbootstrap as ttk
from tkinter import *


def create_and_grid_label(parent, text, row, col):
    label = ttk.Label(parent, text=text)
    label.grid(row=row, column=col, sticky="w", padx=15)
    return label


def create_and_grid_entry(parent, row, col, sticky="e", pady=2, **kwargs):
    entry = ttk.Entry(parent, **kwargs)
    entry.grid(row=row, column=col, sticky=sticky, pady=pady)
    return entry


# Test these functions by themselves:
root = Tk()
frame = ttk.Frame(root)
frame.pack(pady=20, padx=20)
create_and_grid_label(frame, "Test Label", 0, 0)
test_entry = create_and_grid_entry(frame, 0, 1)
root.mainloop()
