import os, openpyxl
ROOT = os.path.dirname(os.path.abspath(__file__))
wb = openpyxl.load_workbook(os.path.join(ROOT, "recipients.xlsx"), data_only=True)
ws = wb.active
for i, row in enumerate(ws.iter_rows(values_only=True)):
    print(i, row)
