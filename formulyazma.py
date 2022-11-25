from openpyxl import Workbook, load_workbook

def mainFonk(x,y):
        
    for satir in range(5, ws.max_row):
        for sutun in range(x,y):
            if(str(ws.cell(satir,sutun).value) == "X"):
                a = "=+C"
                c = str(satir)
                b = a + c
                ws.cell(satir, sutun).value = b

wb = load_workbook("deneme.xlsx")
ws = wb.active

mainFonk(6,18)
mainFonk(29,70)

wb.save("denemenindenemesi.xlsx")



