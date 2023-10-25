import tkinter as tk
import ttkbootstrap as ttk


def simple_progress():
    progress_win = tk.Toplevel()
    progress_win.title("Simple Progress...")
    progress_win.geometry("300x50")

    progress_bar = ttk.Progressbar(
        progress_win,
        orient="horizontal",
        length=280,
        mode="determinate",
        maximum=10,
        style="success",
    )
    progress_bar.grid(row=0, column=0, padx=10, pady=20)
    progress_bar["value"] = 5


app = tk.Tk()
app.title("Test App")
app.geometry("300x300")

btn = ttk.Button(app, text="Show Progress", command=simple_progress)
btn.pack(pady=50)

app.mainloop()
