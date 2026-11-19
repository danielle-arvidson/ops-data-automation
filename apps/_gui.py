"""Small shared tkinter pieces so each app stays short."""
import os
import platform
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import messagebox, ttk

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def open_path(path):
    """Open a file or folder with the default application."""
    if platform.system() == "Windows":
        os.startfile(path)  # noqa: S606 (Windows only)
    elif platform.system() == "Darwin":
        subprocess.run(["open", path], check=False)
    else:
        subprocess.run(["xdg-open", path], check=False)


class ToolWindow:
    """Title, a frame for the tool's own controls, a run button and a log box."""

    def __init__(self, root, title, subtitle):
        self.root = root
        root.title(title)
        ttk.Style().theme_use("clam")
        frame = ttk.Frame(root, padding=20)
        frame.pack(fill=tk.BOTH, expand=True)
        ttk.Label(frame, text=title, font=("Arial", 16, "bold")).pack()
        ttk.Label(frame, text=subtitle).pack(pady=(0, 10))
        self.controls = ttk.Frame(frame)
        self.controls.pack(fill=tk.X)
        self.run_button = ttk.Button(frame, text="Run", width=40)
        self.run_button.pack(pady=15)
        self.progress = ttk.Progressbar(frame, mode="indeterminate", length=500)
        self.progress.pack()
        self.text = tk.Text(frame, height=15, width=90, state="disabled", font=("Courier", 9))
        self.text.pack(pady=10, fill=tk.BOTH, expand=True)

    def log(self, message):
        def write():
            self.text.config(state="normal")
            self.text.insert(tk.END, f"{message}\n")
            self.text.see(tk.END)
            self.text.config(state="disabled")
        self.root.after(0, write)

    def run_in_background(self, work, on_done):
        """Run work() off the UI thread; on_done(result) runs on the UI thread."""
        self.run_button.config(state="disabled")
        self.progress.start(10)

        def target():
            try:
                result = work()
                self.root.after(0, lambda: on_done(result))
            except Exception as exc:  # show any failure to the user, don't crash silently
                self.log(f"ERROR: {exc}")
                self.root.after(0, lambda: messagebox.showerror("Error", str(exc)))
            finally:
                self.root.after(0, self.progress.stop)
                self.root.after(0, lambda: self.run_button.config(state="normal"))

        threading.Thread(target=target, daemon=True).start()
