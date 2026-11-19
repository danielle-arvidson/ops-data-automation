"""Warehouse Order Processor: checks a warehouse order against 3PL stock."""
import os
import tkinter as tk
from datetime import datetime
from tkinter import filedialog, messagebox, ttk

from _gui import ToolWindow, open_path
from stocktools import config, warehouse_orders
from stocktools.inventory import Inventory


class WarehouseOrderApp:
    def __init__(self, root):
        self.cfg = config.load()
        self.order_path = self.inventory_path = None
        self.win = ToolWindow(root, "Warehouse Order Processor",
                              "1. Load the order  2. Load the 3PL stock CSV  3. Process")
        self.order_label = self._file_row("Warehouse order:", "Browse Excel", self.pick_order)
        self.inv_label = self._file_row("3PL stock CSV:", "Browse CSV", self.pick_inventory)
        self.win.run_button.config(text="Process order", command=self.run, state="disabled")

    def _file_row(self, label, button, command):
        row = ttk.Frame(self.win.controls)
        row.pack(fill=tk.X, pady=4)
        ttk.Label(row, text=label, width=18).pack(side=tk.LEFT)
        status = ttk.Label(row, text="No file selected", foreground="gray")
        status.pack(side=tk.LEFT, padx=10)
        ttk.Button(row, text=button, command=command).pack(side=tk.RIGHT)
        return status

    def _ready(self):
        if self.order_path and self.inventory_path:
            self.win.run_button.config(state="normal")

    def pick_order(self):
        path = filedialog.askopenfilename(filetypes=[("Excel", "*.xlsx")])
        if path:
            self.order_path = path
            self.order_label.config(text=os.path.basename(path), foreground="green")
            self._ready()

    def pick_inventory(self):
        folder = self.cfg["paths"]["inventory_folder"]
        path = Inventory.find_newest(folder)
        if not (path and messagebox.askyesno("Use latest?", f"Use {os.path.basename(path)}?")):
            path = filedialog.askopenfilename(initialdir=folder, filetypes=[("CSV", "*.csv")])
        if path:
            self.inventory_path = path
            self.inv_label.config(text=os.path.basename(path), foreground="green")
            self._ready()

    def run(self):
        name = os.path.splitext(os.path.basename(self.order_path))[0]
        out = filedialog.asksaveasfilename(
            defaultextension=".xlsx", initialfile=f"COMPLETED_{name}_{datetime.now():%Y%m%d_%H%M%S}.xlsx")
        if not out:
            return
        owner = self.cfg["default_owner_id"]
        self.win.run_in_background(
            lambda: warehouse_orders.process(
                self.order_path, Inventory(self.inventory_path, owner), out, log=self.win.log),
            lambda result: open_path(result["output"]))


if __name__ == "__main__":
    root = tk.Tk()
    WarehouseOrderApp(root)
    root.mainloop()
