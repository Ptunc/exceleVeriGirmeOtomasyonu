from openpyxl import Workbook, load_workbook

#hangi sütun aralığını değiştireceğinin verisini alan ve o sütun aralığında X gördüğü yeri miniFonk'taki fonksiyonla değiştiren kod
def anaFonk(x,y):

    for satir in range(5, ws.max_row):
        for sutun in range(x,y):
            if(str(ws.cell(satir,sutun).value) == "X"):
                ws.cell(satir, sutun).value = miniFonk("=+C", str(satir))
                print(ws.cell(satir, sutun).value) 

def miniFonk(a,b):
    return a + b

#belirli bir satır-sütun aralığında X gördüğü yeri miniFonk'taki fonksiyonları birbirine ekleyerek değiştiren kod
def ozelCarpimFonk(a,b,c,d,e):

    for satir in range(a, b):
        for sutun in range(c, d):
            if(str(ws.cell(satir,sutun).value) == "X"):
                ws.cell(satir, sutun).value = miniFonk("=+D", str(satir)) + "*" + e + miniFonk("-E", str(satir))
                print(ws.cell(satir, sutun).value) 



wb = load_workbook("deneme.xlsx")
ws = wb.active

#sırayla fonksiyonları çağırma
anaFonk(6,18)
anaFonk(29,70)
ozelCarpimFonk(5, 120, 19, 22, "4.03")
ozelCarpimFonk(120, 160, 19, 22, "5.53")
ozelCarpimFonk(5, 120, 22, 23, "3.6")
ozelCarpimFonk(120, 160, 22, 23, "5.03")
ozelCarpimFonk(5, 120, 23, 29, "4.03")
ozelCarpimFonk(120, 160, 23, 29, "5.53")


wb.save("denemenindenemesi.xlsx")



