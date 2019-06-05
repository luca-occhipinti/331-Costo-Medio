# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 14:40:23 2019

@author: luca_
"""

import pandas as pd


file_name = "./MOPAR dolar.xlsx"

xls = pd.ExcelFile(file_name)

sheet_mov = xls.parse(sheet_name="Movimientos")
sheet_stock = xls.parse(sheet_name="Stock con movimientos")


i = 0
j = 0
new_sheet = 'Material;Costo medio'+'\n'

#Obtenemos cada uno de los materiales con su stock
while i < len(sheet_stock):
    material = str(sheet_stock['Material'][i])
    stock = sheet_stock['Stock'][i]
    stock_actual = stock
    pxq = 0
    wac = 0
    while j < len(sheet_mov):
        mat = str(sheet_mov['Material'][j])
        cantidad = sheet_mov['Cantidad'][j]
        precio = sheet_mov['Precio unitario'][j]
        dolar = sheet_mov['TIPO DE CAMBIO'][j]
        precio = precio/dolar
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
    
    
    
df = pd.read_csv(new_sheet, sep=";")
with pd.ExcelWriter('Costo medio dolares.xlsx') as newExcel:
    df.to_excel(newExcel, sheet_name='Sheet1',index=False)
    