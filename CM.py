
import pandas as pd

file_name = "./Mov y CM anterior.XLSX"

xls = pd.ExcelFile(file_name)

sheet_mov = xls.parse(sheet_name="Movimientos")
sheet_stock = xls.parse(sheet_name="Stock")



i = 23
j = 0
new_sheet = 'Material;Costo medio'+'\n'

while i < 35:
    material = str(sheet_stock['Material'][i])
    stock = sheet_stock['Cantidad'][i]
    CM_dolar = sheet_stock['CM Dolares'][i]
    CM_pesos = sheet_stock['CM ARS'][i]
    stock_actual = stock
    pxq = 0
    wac = 0
    while j < len(sheet_mov):
        mat = str(sheet_mov['Material'][j])
        cantidad = sheet_mov['Cantidad'][j]
        precio = sheet_mov['Precio unitario'][j]
        if mat == material:
            if cantidad <= stock:
                pxq += precio*cantidad
                stock = stock - cantidad
                j += 1
                if stock == 0:
                    wac = pxq/stock_actual
                    new_sheet += (str(material)+';'+str(wac)+'\n')
                    print(str(material)+';'+str(wac))
                    j = 0
                    break
                if stock > 0 and sheet_mov['Material'][j+1] != mat:
                    pxq += stock*CM_pesos
                    wac = pxq/stock_actual
                    new_sheet += (str(material)+';'+str(wac)+'\n')
                    print(str(material)+';'+str(wac))
                    j=0
                    break
            else:
                pxq += precio*stock
                wac = pxq/stock_actual
                new_sheet += (str(material)+';'+str(wac)+'\n')
                print(str(material)+';'+str(wac))
                j=0
                break
        else:
            j+=1
       
    i += 1
    j = 0
    