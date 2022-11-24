import string
import openpyxl
from openpyxl import Workbook, load_workbook

wb = load_workbook("deneme.xlsx")
ws = wb.active

print(ws["B7"].value)
print(ws.cell(7,2).value)
print(ws.max_row)

for satir in range(6, ws.max_row):
    for sutun in range(8,9):
        print("")
        if (str(ws.cell(satir,sutun).value) == "X"):
            a = "=+C"
            c = str(satir)
            b = a + c
            ws.cell(satir, sutun).value = b  # type: ignore

        print(" | " + str(ws.cell(satir,sutun).value) + " | ",end="")
        print()

wb.save("denemenindenemesi.xlsx")



