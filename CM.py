# -*- coding: utf-8 -*-
"""
Created on Mon Apr 22 14:40:23 2019

@author: Luca Occhipinti
"""

import pandas as pd
from io import StringIO

#file_name es lo nombre del archivo Excel a procesar que se va a abrir.
#Siempre lleva la extension al final
file_name = "./MOPAR - Movimientos.xlsx"

#Se lee el archivo y se guarda en la variable xls.
xls = pd.ExcelFile(file_name)

#En la variable sheet_mov se guardan los datos de la hoja de movimiento de piezas.
sheet_mov = xls.parse(sheet_name="Movimientos")
#En la variable sheet_stock se guardan los datos de la hoja de stock.
sheet_stock = xls.parse(sheet_name="Stock con movimientos")

#Contador que se usa para recorrer la hoja de stock.
i = 0
#Contador que se usa para recorrer la hoja de movimientos.
j = 0

#La variable new_sheet es un string que va a ir concatenando todos los resultados de los costos medios
#En esta linea se colocan los titulos de las columnas usando el ; como separador de columna.
new_sheet = 'Material;Costo medio pesos;'+'Costo medio dolares'+'\n'

#Primer while. Recorre la hoja de stock
while i < len(sheet_stock):
    #Se obtiene el PartNumber de la hoja de stock
    material = str(sheet_stock['Material'][i])
    #Se obtiene la cantidad que hay en stock del Part Number
    stock = sheet_stock['Stock'][i]
    #El stock se guarda en la variable stock_actual
    stock_actual = stock
    #pxq_pesos: Variable que se utiliza para guardar los precios*cantidad de cada movimiento en PESOS
    pxq_pesos = 0
    #pxq_dolares: Variable que se utiliza para guardar los precios*cantidad de cada movimiento en DOLARES
    pxq_dolares = 0
    #wac_pesos: variable que se usa para guardar el costo medio calculado del partnumber en PESOS
    wac_pesos = 0
    #wac_dolares: variable que se usa para guardar el costo medio calculado del partnumber en DOLARES
    wac_dolares = 0
    #Segundo while. Recorre los movimientos para compararlos con el PartNumber obtenido en el while anterior
    while j < len(sheet_mov):
        #mat: PartNumber de la hoja de movimientos.
        mat = str(sheet_mov['Material'][j])
        #cantidad: La cantidad que ingresa el partnumber en un registro de movimientos.
        cantidad = sheet_mov['Cantidad'][j]
        #precio: Precio en PESOS del material en un registro de movimientos.
        precio = sheet_mov['Precio unitario'][j]
        #dolar: Precio en DOLARES del material en un registro de movimientos.
        dolar = sheet_mov['Importe Unitario ML3'][j]
        #Compara si el material obtenido en la hoja de stock es igual al que esta en la hoja de movimientos
        if mat == material:
            #Verifica qur la cantidad ingresada sea menor o igual que el stock actual
            if cantidad <= stock:
                #Calcula lo pxq tanto en pesos como en dolares
                pxq_pesos += precio*cantidad
                pxq_dolares += dolar*cantidad
                #Resta lo calculado del stock, para marcarlo como calculado.
                stock = stock - cantidad
                j += 1
                #Verifica si luego del calculo de PxQ el stock llego a 0
                if stock == 0:
                    #Si llegó a 0, calcula los costos medios y los guarda en la cadena new_sheet y lo muestra por pantalla
                    wac_pesos = pxq_pesos/stock_actual
                    wac_dolares = pxq_dolares/stock_actual
                    new_sheet += (str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                    print(str(material)+';'+str(wac_pesos)+';'+str(wac_dolares)+'\n')
                    j = 0
                    break
            else:
                #Si la cantidad ingresada es superior al stock, calcula el pxq por lo que hay en stock y lo guarda
                #en la cadena new_sheet y lo muestra por pantalla.
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

#Convierte el string new_seet a un formato de csv.
new_sheet = StringIO(new_sheet)

#Exporta todos los datos contenidos en new_sheet a un Excel ubicado en la carpeta donde se encuentra el scrit corriendo
#El excel será con nombre "Costo medio.xlsx"
df = pd.read_csv(new_sheet, sep=";")
with pd.ExcelWriter('Costo medio.xlsx') as newExcel:
    df.to_excel(newExcel, sheet_name='Sheet1',index=False)
    
