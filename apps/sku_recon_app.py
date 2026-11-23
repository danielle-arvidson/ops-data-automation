"""Single SKU reconciliation window. Logic is in stocktools/sku_reconciliation.py.

    python apps/sku_recon_app.py
"""
import os
import traceback

from _gui import open_path
from stocktools import sku_reconciliation as sku

KPI_FOLDER = sku.KPI_FOLDER


def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    root = tk.Tk(); root.title('Single SKU Reconciliation: 3PL vs ERP'); root.geometry('860x760')
    f = ttk.Frame(root, padding=8); f.pack(fill='both', expand=True)
    part, epi, out = tk.StringVar(), tk.StringVar(), tk.StringVar()
    traces = []

    ttk.Label(f, text='Part number *').grid(row=0, column=0, sticky='w')
    ttk.Entry(f, textvariable=part, width=24).grid(row=0, column=1, sticky='w', pady=3)
    ttk.Label(f, text='3PL inventory trace(s) *').grid(row=1, column=0, sticky='w')
    lt = ttk.Label(f, text='0 file(s)'); lt.grid(row=1, column=1, sticky='w')

    def pick_traces():
        traces[:] = filedialog.askopenfilenames(filetypes=[('3PL trace', '*.xls *.xlsx *.csv'), ('All', '*.*')])
        lt.config(text=f'{len(traces)} file(s): ' + ', '.join(os.path.basename(p) for p in traces)[:80])
    ttk.Button(f, text='Browse', command=pick_traces).grid(row=1, column=2)
    ttk.Label(f, text='ERP transaction history *').grid(row=2, column=0, sticky='w')
    ttk.Entry(f, textvariable=epi, width=62).grid(row=2, column=1, sticky='w')
    ttk.Button(f, text='Browse', command=lambda: epi.set(filedialog.askopenfilename(
        filetypes=[('ERP export', '*.xlsx *.txt *.csv'), ('All', '*.*')]) or epi.get())).grid(row=2, column=2)
    ttk.Label(f, text='3PL on hand (paste: Lot, Owner, Qty)').grid(row=3, column=0, sticky='nw', pady=4)
    dt_box = tk.Text(f, height=6, width=62); dt_box.grid(row=3, column=1, sticky='w', pady=4)
    ttk.Label(f, text='ERP on hand (paste: Lot, Warehouse, Qty)').grid(row=4, column=0, sticky='nw', pady=4)
    ep_box = tk.Text(f, height=6, width=62); ep_box.grid(row=4, column=1, sticky='w', pady=4)
    kpi_var = tk.StringVar(value=KPI_FOLDER)
    ttk.Label(f, text='KPI log folder (blank = results folder)').grid(row=6, column=0, sticky='w')
    ttk.Entry(f, textvariable=kpi_var, width=62).grid(row=6, column=1, sticky='w')
    ttk.Button(f, text='Browse', command=lambda: kpi_var.set(filedialog.askdirectory() or kpi_var.get())).grid(row=6, column=2)
    ttk.Label(f, text='Save results as *').grid(row=5, column=0, sticky='w')
    ttk.Entry(f, textvariable=out, width=62).grid(row=5, column=1, sticky='w')
    ttk.Button(f, text='Browse', command=lambda: out.set(filedialog.asksaveasfilename(
        defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')]) or out.get())).grid(row=5, column=2)
    logbox = scrolledtext.ScrolledText(f, height=14, width=104)
    logbox.grid(row=8, column=0, columnspan=3, pady=8)

    def log(m):
        logbox.insert('end', m + '\n'); logbox.see('end'); root.update()

    def go():
        if not (part.get().strip() and traces and epi.get() and out.get()):
            messagebox.showwarning('Missing', 'Enter the part, choose the 3PL trace(s), the ERP history and where to save.')
            return
        try:
            sku.run(part.get(), list(traces), epi.get(), out.get(), dt_box.get('1.0', 'end'), ep_box.get('1.0', 'end'), log,
                kpi_var.get() or None)
            if messagebox.askyesno('Done', 'Finished. Open the results?'):
                open_path(out.get())
        except Exception as ex:
            log('ERROR: ' + str(ex)); log(traceback.format_exc())
    ttk.Button(f, text='Run', command=go).grid(row=7, column=1, pady=6)
    root.mainloop()


if __name__ == '__main__':
    gui()
