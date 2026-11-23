"""Pack-out tool: shares a 3PL packing list across the SO/Line/Release lines of one or more movement requests."""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from _gui import ToolWindow, open_path
from stocktools import packout


class PackoutApp:
    def __init__(self, root):
        self.exports = []
        self.win = ToolWindow(root, "Pack-out Process", "Fill SO, Line and Release on the packing list")
        box = ttk.LabelFrame(self.win.controls, text="ERP shipping exports", padding=10)
        box.pack(fill=tk.X)
        buttons = ttk.Frame(box)
        buttons.pack(fill=tk.X)
        ttk.Button(buttons, text="Add CSV", command=self.add).pack(side=tk.LEFT, padx=5)
        ttk.Button(buttons, text="Clear", command=self.clear).pack(side=tk.LEFT)
        self.listbox = tk.Listbox(box, height=4)
        self.listbox.pack(fill=tk.X, pady=8)
        ttk.Label(box, text="Add every movement request export on this shipment, oldest first.\n"
                  "Quantities are shared across releases; check the Check sheet for totals that don't tie.",
                  foreground="gray").pack()
        self.win.run_button.config(text="Select packing list and process", command=self.run)

    def add(self):
        path = filedialog.askopenfilename(filetypes=[("CSV", "*.csv")])
        if path and path not in self.exports:
            self.exports.append(path)
            self.listbox.insert(tk.END, os.path.basename(path))

    def clear(self):
        self.exports.clear()
        self.listbox.delete(0, tk.END)

    def run(self):
        if not self.exports:
            messagebox.showwarning("No exports", "Add at least one ERP export first.")
            return
        packing = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if not packing:
            return
        out = os.path.splitext(packing)[0] + "_completed.xlsx"
        exports = list(self.exports)
        self.win.run_in_background(
            lambda: packout.run(exports, packing, out, log=self.win.log),
            lambda result: open_path(result["output"]))


if __name__ == "__main__":
    root = tk.Tk()
    PackoutApp(root)
    root.mainloop()
