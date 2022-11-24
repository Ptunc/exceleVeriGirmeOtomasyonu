from openpyxl import Workbook, load_workbook

wb = load_workbook("deneme.xlsx")
ws = wb.active


for satir in range(5, ws.max_row):
    for sutun in range(6,18):
        if(str(ws.cell(satir,sutun).value) == "X"):
            a = "=+C"
            c = str(satir)
            b = a + c
            ws.cell(satir, sutun).value = b

for satir in range(5, ws.max_row):
    for sutun in range(29,70):
        if(str(ws.cell(satir,sutun).value) == "X"):
            a = "=+C"
            c = str(satir)
            b = a + c
            ws.cell(satir, sutun).value = b

print(ws.cell(5,7).value)
wb.save("denemenindenemesi.xlsx")



