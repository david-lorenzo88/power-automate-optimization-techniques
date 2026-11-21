"""Builds lab/assets/Uc9Orders.xlsx: one sheet with an empty Excel table `Orders`
(OrderId, Region, Amount). UC9 SLOW adds rows one by one with the Excel connector;
UC9 FAST writes them in one Office Script call.

    python lab/assets/make_workbook.py        (needs openpyxl)
"""
from pathlib import Path

from openpyxl import Workbook
from openpyxl.worksheet.table import Table, TableStyleInfo

wb = Workbook()
ws = wb.active
ws.title = "Orders"
ws.append(["OrderId", "Region", "Amount"])
ws.append(["", "", None])  # a table needs at least one data row; UC9-RESET clears it
table = Table(displayName="Orders", ref="A1:C2")
table.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
ws.add_table(table)
ws.column_dimensions["A"].width = 14
ws.column_dimensions["B"].width = 10
ws.column_dimensions["C"].width = 12
out = Path(__file__).with_name("Uc9Orders.xlsx")
wb.save(out)
print(f"wrote {out}")
