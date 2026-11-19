"""Movement Request Tool: ERP order export in, 3PL movement request out."""
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from _gui import ToolWindow, open_path
from stocktools import config, movement_request
from stocktools.inventory import Inventory


class MovementRequestApp:
    def __init__(self, root):
        self.cfg = config.load()
        self.inventory = None
        self.win = ToolWindow(root, "Movement Request Tool",
                              "Stock check and case calculation against the 3PL report")

        row = ttk.Frame(self.win.controls)
        row.pack(fill=tk.X, pady=5)
        ttk.Label(row, text="Owner ID:", width=12).pack(side=tk.LEFT)
        self.owner = tk.StringVar(value=self.cfg["default_owner_id"])
        ttk.Combobox(row, textvariable=self.owner, values=self.cfg["owner_ids"],
                     state="readonly", width=18).pack(side=tk.LEFT)
        ttk.Button(row, text="Load 3PL inventory", command=self.load_inventory).pack(side=tk.LEFT, padx=10)
        self.status = ttk.Label(self.win.controls, text="No inventory loaded", foreground="gray")
        self.status.pack()

        self.win.run_button.config(text="Select order CSV and build request", command=self.run)

    def load_inventory(self):
        folder = self.cfg["paths"]["inventory_folder"]
        path = Inventory.find_newest(folder)
        if path:
            answer = messagebox.askyesnocancel(
                "Inventory file found", f"Newest file:\n{os.path.basename(path)}\n\n"
                "Yes = use it, No = browse, Cancel = stop")
            if answer is None:
                return
            if answer is False:
                path = None
        if not path:
            path = filedialog.askopenfilename(initialdir=folder, filetypes=[("CSV", "*.csv")])
        if not path:
            return
        self.inventory = Inventory(path, self.owner.get())
        if not len(self.inventory):
            messagebox.showwarning("No data", f"Nothing found for owner {self.owner.get()}")
        self.status.config(
            text=f"Loaded {os.path.basename(path)}: {len(self.inventory)} lines, "
                 f"{self.inventory.pack_config_count} pack configs", foreground="green")

    def run(self):
        csv_path = filedialog.askopenfilename(title="ERP order export", filetypes=[("CSV", "*.csv")])
        if not csv_path:
            return
        if self.inventory is None and not messagebox.askyesno(
                "No inventory loaded", "Stock checks and case counts will be skipped. Continue?"):
            return
        paths = self.cfg["paths"]
        self.win.run_in_background(
            lambda: movement_request.build(
                csv_path, paths["movement_request_template"], paths["movement_request_output"],
                inventory=self.inventory, case_sku_pattern=self.cfg.get("case_sku_pattern"),
                log=self.win.log),
            self.done)

    def done(self, result):
        open_path(result["output"])
        note = ""
        if result["issues"]:
            note = (f"\n\n{len(result['issues'])} stock issue(s), highlighted in the file. "
                    f"3PL report time: {result['report_time']}.")
        messagebox.showinfo("Done", f"{result['lines']} lines for {result['customer']}.{note}")


if __name__ == "__main__":
    root = tk.Tk()
    MovementRequestApp(root)
    root.mainloop()
