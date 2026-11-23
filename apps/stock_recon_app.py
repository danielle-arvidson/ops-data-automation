"""Stock reconciliation: ERP stock against a 3PL inventory report or a
physical stock take count. Logic is in stocktools/reconciliation.py.

    python apps/stock_recon_app.py
"""
import datetime
import os
import sys
import traceback

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from _gui import open_path  # noqa: E402
from stocktools import reconciliation  # noqa: E402


def gui():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    root = tk.Tk(); root.title('Stock Reconciliation'); root.geometry('820x640')
    nb = ttk.Notebook(root); nb.pack(fill='both', expand=False, padx=6, pady=6)
    kf = ttk.Frame(root); kf.pack(fill='x', padx=6)
    kpi_var = tk.StringVar(value=reconciliation.KPI_FOLDER)
    ttk.Label(kf, text='KPI log folder (blank = results folder)').pack(side='left', padx=4)
    ttk.Entry(kf, textvariable=kpi_var, width=60).pack(side='left', padx=4)
    ttk.Button(kf, text='Browse', command=lambda: kpi_var.set(filedialog.askdirectory() or kpi_var.get())).pack(side='left')
    out = scrolledtext.ScrolledText(root, height=12); out.pack(fill='both', expand=True, padx=6, pady=6)

    def log(msg):
        out.insert('end', msg + '\n'); out.see('end'); root.update()

    def browse(var, save=False, ft=(('Excel', '*.xlsx'), ('All', '*.*'))):
        p = (filedialog.asksaveasfilename(defaultextension='.xlsx', filetypes=[('Excel', '*.xlsx')]) if save
             else filedialog.askopenfilename(filetypes=list(ft)))
        if p: var.set(p)
        return p

    def row(frame, r, label, var, save=False, ft=(('Excel', '*.xlsx'), ('All', '*.*')), after=None):
        ttk.Label(frame, text=label).grid(row=r, column=0, sticky='w', padx=6, pady=3)
        ttk.Entry(frame, textvariable=var, width=62).grid(row=r, column=1, padx=4)
        ttk.Button(frame, text='Browse', command=lambda: (browse(var, save, ft), after and after())).grid(row=r, column=2)

    def finish(path):
        if messagebox.askyesno('Done', 'Finished. Open the results?'):
            open_path(path)

    # --- tab 1: 3PL vs ERP
    t1 = ttk.Frame(nb); nb.add(t1, text='3PL vs ERP')
    v1 = {k: tk.StringVar() for k in ('erp', 'tpl', 'txn', 'out', 'date')}
    v1['date'].set(datetime.date.today().strftime('%d/%m/%Y')); orders = []
    row(t1, 0, 'ERP stock report (.xlsx) *', v1['erp'])
    row(t1, 1, '3PL full inventory (.xlsx) *', v1['tpl'])
    row(t1, 2, 'ERP transactions (.txt, optional)', v1['txn'], ft=(('Text', '*.txt *.tsv'), ('All', '*.*')))
    ttk.Label(t1, text='3PL order summaries (.csv, optional)').grid(row=3, column=0, sticky='w', padx=6)
    lo = ttk.Label(t1, text='0 file(s)'); lo.grid(row=3, column=1, sticky='w')

    def pick_orders():
        orders[:] = filedialog.askopenfilenames(filetypes=[('CSV', '*.csv')]); lo.config(text=f'{len(orders)} file(s)')
    ttk.Button(t1, text='Browse', command=pick_orders).grid(row=3, column=2)
    ttk.Label(t1, text='Report date (dd/mm/yyyy)').grid(row=4, column=0, sticky='w', padx=6)
    ttk.Entry(t1, textvariable=v1['date'], width=14).grid(row=4, column=1, sticky='w', padx=4)
    row(t1, 5, 'Save results as *', v1['out'], save=True)

    def go1():
        if not (v1['erp'].get() and v1['tpl'].get() and v1['out'].get()):
            messagebox.showwarning('Missing', 'Choose the ERP report, 3PL report and where to save.'); return
        try:
            reconciliation.run_3pl(v1['erp'].get(), v1['tpl'].get(), v1['out'].get(), v1['txn'].get() or None, list(orders),
                reconciliation.to_date(v1['date'].get()), log, kpi_var.get() or None); finish(v1['out'].get())
        except Exception as ex:
            log('ERROR: ' + str(ex)); log(traceback.format_exc())
    ttk.Button(t1, text='Run', command=go1).grid(row=6, column=1, pady=6)

    # --- tab 2: stock take vs ERP
    t2 = ttk.Frame(nb); nb.add(t2, text='Stock take vs ERP')
    v2 = {k: tk.StringVar() for k in ('erp', 'cnt', 'txn', 'out', 'sheet', 'part', 'lot', 'qty', 'warehouse', 'bin')}
    bybin = tk.BooleanVar(value=False)
    whlist = tk.Listbox(t2, selectmode='multiple', height=6, width=62, exportselection=False)
    wh_pairs = []

    def load_wh():
        try:
            wh_pairs[:] = reconciliation.erp_warehouses(v2['erp'].get()); whlist.delete(0, 'end')
            for s, w in wh_pairs: whlist.insert('end', f'{s} / {w}')
        except Exception as ex:
            log('Could not read warehouses: ' + str(ex))
    combos = {}

    def load_cols():
        try:
            p = v2['cnt'].get(); sheets = []
            if not p.lower().endswith('.csv'):
                from openpyxl import load_workbook
                sheets = load_workbook(p, read_only=True).sheetnames
                sh_combo['values'] = sheets
                if not v2['sheet'].get() or v2['sheet'].get() not in sheets: v2['sheet'].set(sheets[0])
            hdr, _ = reconciliation.read_table(p, v2['sheet'].get() or None)
            for k, cb in combos.items():
                cb['values'] = [''] + hdr
                guess = {'part': ('part', 'sku', 'stock code', 'part number'), 'lot': ('lot', 'lot number', 'batch'),
                         'qty': ('qty', 'quantity', 'count', 'final count', 'counted'),
                         'warehouse': ('warehouse',), 'bin': ('bin', 'location')}[k]
                hit = [h for h in hdr if h.strip().lower() in guess]
                v2[k].set(hit[0] if hit else '')
        except Exception as ex:
            log('Could not read count columns: ' + str(ex))
    row(t2, 0, 'ERP stock report (.xlsx) *', v2['erp'], after=load_wh)
    ttk.Label(t2, text='Warehouses counted *').grid(row=1, column=0, sticky='nw', padx=6)
    whlist.grid(row=1, column=1, padx=4, pady=3, sticky='w')
    row(t2, 2, 'Count sheet (.xlsx or .csv) *', v2['cnt'], ft=(('Excel or CSV', '*.xlsx *.csv'), ('All', '*.*')), after=load_cols)
    ttk.Label(t2, text='Count sheet tab').grid(row=3, column=0, sticky='w', padx=6)
    sh_combo = ttk.Combobox(t2, textvariable=v2['sheet'], width=30); sh_combo.grid(row=3, column=1, sticky='w', padx=4)
    sh_combo.bind('<<ComboboxSelected>>', lambda e: load_cols())
    cf = ttk.Frame(t2); cf.grid(row=4, column=0, columnspan=3, sticky='w', padx=6, pady=3)
    for i, (k, lab) in enumerate((('part', 'Part *'), ('lot', 'Lot'), ('qty', 'Qty *'), ('warehouse', 'Warehouse'), ('bin', 'Bin'))):
        ttk.Label(cf, text=lab).grid(row=0, column=i * 2, padx=3)
        combos[k] = ttk.Combobox(cf, textvariable=v2[k], width=14); combos[k].grid(row=0, column=i * 2 + 1, padx=3)
    ttk.Checkbutton(t2, text='Compare by bin (needs a Bin column)', variable=bybin).grid(row=5, column=1, sticky='w')
    row(t2, 6, 'ERP transactions (.txt, optional)', v2['txn'], ft=(('Text', '*.txt *.tsv'), ('All', '*.*')))
    row(t2, 7, 'Save results as *', v2['out'], save=True)

    def go2():
        scope = {wh_pairs[i] for i in whlist.curselection()}
        if not (v2['erp'].get() and v2['cnt'].get() and v2['out'].get() and scope and v2['part'].get() and v2['qty'].get()):
            messagebox.showwarning('Missing', 'Choose the ERP report, the warehouses counted, the count sheet, '
                                              'the Part and Qty columns, and where to save.'); return
        mp = {k: v2[k].get() or None for k in ('part', 'lot', 'qty', 'warehouse', 'bin')}
        try:
            reconciliation.run_stocktake(v2['erp'].get(), v2['cnt'].get(), v2['out'].get(), scope, mp, bybin.get(),
                          v2['txn'].get() or None, v2['sheet'].get() or None, log, kpi_var.get() or None); finish(v2['out'].get())
        except Exception as ex:
            log('ERROR: ' + str(ex)); log(traceback.format_exc())
    ttk.Button(t2, text='Run', command=go2).grid(row=8, column=1, pady=6)
    root.mainloop()


if __name__ == '__main__':
    gui()
