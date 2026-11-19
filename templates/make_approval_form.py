"""Build the stock adjustment and miscellaneous transaction approval form.

    python templates/make_approval_form.py

This is the control document for every manual stock correction. It has
one job beyond collecting signatures: telling the approver, before they
sign, whether a line is already covered by the finance provision. If a
written-off line is also in the provision, the same loss is counted twice
(once as the provision, once as the adjustment). The "In Finance
Provision?" column checks every line against the current provision
snapshot automatically.

The provision rows in the template are synthetic. In use, the hidden
FP Data sheet is replaced each quarter with the live provision list.
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.worksheet.datavalidation import DataValidation

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "adjustment_approval_form.xlsx")

TRANSACTION_TYPES = ["Adjustment (Qty)", "Adjustment (Cost)", "Misc Issue", "Misc Receipt"]
APPROVAL_TYPES = ["Pre-approval", "Retrospective"]
REASON_CODES = [
    ("A-01", "SLOB", "Scrap following the slow-moving and obsolete stock review"),
    ("A-02", "Stock Disposal Form", "Scrap of tail ends, expired or short-dated stock on an approved disposal form"),
    ("A-03", "Cycle Count", "Variance found by a cycle count"),
    ("A-04", "Stock Take", "Variance found by a stock take"),
    ("A-05", "Customer Tail Ends", "Tail ends charged back under a customer agreement"),
    ("A-06", "Correction", "Correction to a previous adjustment"),
    ("A-07", "Material Variance", "Component over or under use found at the end of a batch"),
    ("A-08", "PO Receipt Error", "Receipt error that can't be reversed on the receipt itself"),
    ("A-09", "Job Receipt Error", "Job receipt error that can't be reversed on the job itself"),
    ("A-10", "Attribute Change", "Net-zero adjustment to change a lot attribute"),
    ("A-11", "System Error", "ERP fault on a transfer, receipt or invoice, where stock is needed before the fix"),
    ("A-12", "Quality Issue", "Error in nonconformance or rework processing"),
    ("A-13", "Not Established", "Cause can't be identified"),
    ("M-01", "Development Projects", "Issues to development projects"),
    ("M-02", "Samples", "Sales and purchasing samples"),
    ("M-03", "Test Lab", "Release and complaint testing"),
    ("M-04", "Engineering", "Engineering samples and requests"),
    ("M-05", "Production", "Production samples and requests"),
    ("M-06", "Consumables - Production", "Consumables issued to production"),
    ("M-07", "Consumables - Warehouse", "Consumables used at the third-party warehouse"),
    ("M-08", "Issue Error", "Stock loss where the job can't be identified"),
    ("M-09", "Return Error", "Stock found where the job can't be identified"),
]
SAMPLE_PROVISION = [
    # part, lot, description, category, value
    ("009104", "L004", "Sample pen needle 6mm", "3 - Over 2yr no use", 300.00),
    ("009105", "L005", "Sample swab pack", "2 - SLOB", 400.00),
    ("009120", "", "Sample carry case", "2 - SLOB", 120.00),
]

HEADER_FILL = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
INPUT_FILL = PatternFill(start_color="FFFFFF", end_color="FFFFFF", fill_type="solid")
THIN = Side(style="thin", color="999999")
BOX = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
FIRST_LINE, LAST_LINE = 9, 14
FP_RANGE = "'FP Data'!$A$3:$A$10000"


def lists_sheet(wb):
    ws = wb.create_sheet("Lists")
    for i, v in enumerate(TRANSACTION_TYPES, start=1):
        ws.cell(i, 1, v)
    for i, v in enumerate(APPROVAL_TYPES, start=1):
        ws.cell(i, 2, v)
    ws.cell(1, 3, "Yes")
    ws.cell(1, 4, "Reason Code")
    for i, (code, name, _) in enumerate(REASON_CODES, start=2):
        ws.cell(i, 4, f"{code} - {name}")
    ws.sheet_state = "hidden"
    return f"Lists!$D$2:$D${len(REASON_CODES) + 1}"


def reason_codes_sheet(wb):
    ws = wb.create_sheet("Reason Codes")
    ws.append(["Code", "Name", "When to use"])
    for c in ws[1]:
        c.font = BOLD
        c.fill = HEADER_FILL
    for row in REASON_CODES:
        ws.append(list(row))
    ws.column_dimensions["A"].width = 8
    ws.column_dimensions["B"].width = 26
    ws.column_dimensions["C"].width = 80


def fp_data_sheet(wb):
    ws = wb.create_sheet("FP Data")
    ws["A1"] = "Finance provision snapshot. Replace the rows below each quarter. Column A is the match key (part + lot)."
    ws.append(["Part + Lot Key", "Part", "Lot", "Description", "Provision Category", "Value"])
    for part, lot, desc, cat, value in SAMPLE_PROVISION:
        ws.append([part + lot, part, lot, desc, cat, value])
    for row in ws.iter_rows(min_row=3, max_col=3):
        for c in row:
            c.number_format = "@"
    ws.sheet_state = "hidden"


def instructions_sheet(wb):
    ws = wb.create_sheet("Instructions", 0)
    ws.column_dimensions["A"].width = 4
    ws.column_dimensions["B"].width = 110
    ws["B1"] = "Requester instructions"
    ws["B1"].font = Font(bold=True, size=14)
    steps = [
        "Date Raised (H3): today's date.",
        "Transaction Type (C4): choose from the list.",
        "Approval Type (H4): Pre-approval if not yet posted, Retrospective if already posted.",
        "Job Number (C5): required for a Misc Issue, otherwise leave blank.",
        "Linked NCR / Investigation Ref (E6): if the correction relates to a quality record.",
        "Lines 9 to 14: one row per part and lot. Qty is positive for an increase, negative for a decrease. "
        "Value of Impact and In Finance Provision? fill in automatically. Justification is required on every line.",
        "Requester row: print your name, set Signed to Yes, then set Dated to Yes to stamp today's date.",
        "Pass the form to the Stock and Data Integrity Manager. Don't complete the Administrator or Approver rows.",
    ]
    for i, text in enumerate(steps, start=1):
        ws.cell(i + 2, 1, i)
        ws.cell(i + 2, 2, text).alignment = Alignment(wrap_text=True, vertical="top")


def form_sheet(ws, reason_list):
    widths = {"A": 2.5, "B": 18.5, "C": 9, "D": 12, "E": 22, "F": 12, "G": 15, "H": 16,
              "I": 9.5, "J": 11, "K": 14.5, "L": 32, "M": 14}
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

    ws.merge_cells("B2:M2")
    ws["B2"] = "Stock Adjustment & Miscellaneous Transaction Approval Form"
    ws["B2"].font = Font(bold=True, size=18)

    labels = {"B3": "File Name:", "F3": "Date Raised:", "I3": "Form Ref:", "B4": "Transaction Type:",
              "G4": "Approval Type:", "B5": "Job Number:", "B6": "Linked NCR / Investigation Ref:"}
    for ref, text in labels.items():
        ws[ref] = text
        ws[ref].font = BOLD
    for rng in ("C3:E3", "F3:G3", "I3:J3", "K3:M3", "C4:F4", "H4:M4", "C5:F5", "G5:M5", "B6:D6", "E6:M6"):
        ws.merge_cells(rng)
    ws["G5"] = "Job Number is required for a Misc Issue"
    ws["G5"].font = Font(italic=True, color="808080")
    ws["H3"].number_format = "DD/MM/YYYY"

    headers = ["Part Number", "Lot No", "MFG Batch", "Part Description", "Warehouse", "Location",
               "Reason Code", "Qty (+/-)", "Unit Cost", "Value of Impact", "Justification / Root Cause (required)",
               "In Finance Provision?"]
    for i, h in enumerate(headers):
        c = ws.cell(8, 2 + i, h)
        c.font = BOLD
        c.fill = HEADER_FILL
        c.border = BOX
        c.alignment = Alignment(wrap_text=True, horizontal="center", vertical="center")
    ws.row_dimensions[8].height = 32

    for r in range(FIRST_LINE, LAST_LINE + 1):
        for col in range(2, 14):
            ws.cell(r, col).border = BOX
        ws.cell(r, 2).number_format = "@"
        ws.cell(r, 3).number_format = "@"
        ws.cell(r, 11, f'=IF(AND(I{r}="",J{r}=""),"",N(I{r})*N(J{r}))').number_format = "#,##0.00"
        # The double-count check: is this part + lot already in the finance provision?
        ws.cell(r, 13, f'=IF($B{r}="","",IF(COUNTA({FP_RANGE})=0,"No FP data",'
                       f'IF(COUNTIF({FP_RANGE},$B{r}&$C{r})>0,"Yes","No")))')
        ws.cell(r, 13).alignment = Alignment(horizontal="center")

    ws.merge_cells("B15:J15")
    ws["B15"] = "Total Value of Impact"
    ws["B15"].font = BOLD
    ws["K15"] = f"=SUM(K{FIRST_LINE}:K{LAST_LINE})"
    ws["K15"].number_format = "#,##0.00"
    ws["K15"].font = BOLD

    ws.merge_cells("B17:M17")
    ws["B17"] = "All stock adjustments and miscellaneous transactions need the approvals below before posting in the ERP."
    ws.merge_cells("B18:M18")
    ws["B18"] = ("Sign in order: Requester, then Administrator (Stock and Data Integrity Manager), then Approver "
                 "(line manager or head of supply chain). For a transaction already posted, complete the form and "
                 "get sign-off retrospectively.")
    ws["B18"].alignment = Alignment(wrap_text=True)
    ws.row_dimensions[18].height = 30

    yes_no = DataValidation(type="list", formula1="Lists!$C$1", allow_blank=True)
    ws.add_data_validation(yes_no)
    for r, role in ((19, "Requester:"), (20, "Administrator:"), (21, "Approver:")):
        ws.merge_cells(f"B{r}:C{r}")
        ws.merge_cells(f"E{r}:F{r}")
        ws.merge_cells(f"K{r}:M{r}")
        ws[f"B{r}"] = role
        ws[f"B{r}"].font = BOLD
        ws[f"D{r}"] = "Print Name:"
        ws[f"G{r}"] = "Signed:"
        ws[f"I{r}"] = "Dated:"
        yes_no.add(f"H{r}")
        yes_no.add(f"J{r}")
        # Stamps today's date once, then keeps it. Relies on iterative calculation (set below),
        # so the date doesn't roll forward every time the file is opened.
        ws[f"K{r}"] = f'=IF(J{r}="Yes",IF(AND(ISNUMBER(K{r}),K{r}>1),K{r},TODAY()),"<- set Dated to Yes")'
        ws[f"K{r}"].number_format = "DD/MM/YYYY"

    ws.merge_cells("B23:M23")
    ws["B23"] = ("Once fully approved, return to the Stock and Data Integrity Manager to log and post the "
                 "transaction. Keep a copy for audit.")

    dv_type = DataValidation(type="list", formula1="Lists!$A$1:$A$4", allow_blank=True)
    dv_approval = DataValidation(type="list", formula1="Lists!$B$1:$B$2", allow_blank=True)
    dv_reason = DataValidation(type="list", formula1=reason_list, allow_blank=True)
    for dv, ref in ((dv_type, "C4"), (dv_approval, "H4"), (dv_reason, f"H{FIRST_LINE}:H{LAST_LINE}")):
        ws.add_data_validation(dv)
        dv.add(ref)


def main():
    wb = Workbook()
    form = wb.active
    form.title = "Approval Form"
    reason_list = lists_sheet(wb)
    reason_codes_sheet(wb)
    fp_data_sheet(wb)
    instructions_sheet(wb)
    form_sheet(form, reason_list)
    wb.active = wb.sheetnames.index("Approval Form")
    wb.calculation.iterate = True
    wb.calculation.iterateCount = 1
    wb.save(OUT)
    print("Saved", OUT)


if __name__ == "__main__":
    main()
