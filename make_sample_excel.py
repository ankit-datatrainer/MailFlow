"""Generate a sample recipients_sample.xlsx you can copy/edit as recipients.xlsx."""
import os
import openpyxl

ROOT = os.path.dirname(os.path.abspath(__file__))
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "recipients"
ws.append(["name", "email", "company"])  # 'name' and 'email' are required; add any extra columns
ws.append(["John Doe", "john@example.com", "Acme Inc"])
ws.append(["Jane Smith", "jane@example.com", "Globex"])
ws.append(["Sam Lee", "sam@example.com", "Initech"])
out = os.path.join(ROOT, "recipients_sample.xlsx")
wb.save(out)
print("Saved:", out)
