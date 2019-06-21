# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 14:40:23 2019

@author: luca_
"""

import pandas as pd
from io import StringIO


file_name = "./MOPAR - Movimientos.xlsx"

xls = pd.ExcelFile(file_name)

sheet_mov = xls.parse(sheet_name="Movimientos")
sheet_stock = xls.parse(sheet_name="Stock con movimientos")


i = 0
j = 0
new_sheet = 'Material;Costo medio pesos;'+'Costo medio dolares'+'\n'

#Obtenemos cada uno de los materiales con su stock
while i < len(sheet_stock):
    material = str(sheet_stock['Material'][i])
    stock = sheet_stock['Stock'][i]
    stock_actual = stock
    pxq_pesos = 0
    pxq_dolares = 0
    wac_pesos = 0
    wac_dolares = 0
    while j < len(sheet_mov):
        mat = str(sheet_mov['Material'][j])
        cantidad = sheet_mov['Cantidad'][j]
        precio = sheet_mov['Precio unitario'][j]
        dolar = sheet_mov['Importe Unitario ML3'][j]
        if mat == material:
            if cantidad <= stock:
                pxq_pesos += precio*cantidad
                pxq_dolares += dolar*cantidad
                stock = stock - cantidad
                j += 1
                if stock == 0:
                    wac_pesos = pxq_pesos/stock_actual
                    wac_dolares = pxq_dolares/stock_actual
                    new_sheet += (str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                    print(str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                    j = 0
                    break
            else:
                pxq_pesos += precio*stock
                pxq_dolares += dolar*stock
                wac_pesos = pxq_pesos/stock_actual
                wac_dolares = pxq_dolares/stock_actual
                new_sheet += (str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                print(str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                j=0
                break
        else:
            j+=1
    i += 1
    j = 0
    
new_sheet = StringIO(new_sheet)

df = pd.read_csv(new_sheet, sep=";")
with pd.ExcelWriter('Costo medio.xlsx') as newExcel:
    df.to_excel(newExcel, sheet_name='Sheet1',index=False)
    