# WCF - Ideas para la mejora

# Idea 331 - Cálculo Costo Medio - MOPAR

# Cuál es el problema?
Impacto de log de remitos en calculo de costo medio de SAPiens, por lo que precisamos una base de Contratos desde NSC histórica para el período 09/2017 a 03/2019 para el cálculo gestional de Valuación del Stock.

# Impacto del problema
Impacto en la confibilidad de la informacion contable

# Referente de Propuesta
Daniel Mandrille

# Propuesta de solucion

Procesar dos planillas en formato Excel con un script en lenguaje Python.
El contenido de las planillas será:

    -Una hoja con los movimientos de materiales desde Octubre 2017 a Marzo 2019.
    -Una hoja con el stock de piezas y sus cantidades.


# Pre-requisitos

# Requisitos para ejecucion
Recomendados:
    -Anaconda Python (link de desarga: https://repo.anaconda.com/archive/Anaconda3-2019.03-Windows-x86_64.exe)
La instalacion de Anaconda ya viene con todas las librerias incluidas. Con esta instalacion ya incluye todo lo necesario para ejecutar el script.

Alternativa
    -Interprete de Python 3.6.* o superior (link: https://www.python.org/ftp/python/3.7.3/python-3.7.3-amd64.exe)
    -Libreria pandas
    -Libreria io
Con esta alternativa hay que instalar todo por separado para poder ejecutar el script.
Link con instructivo para la instalacion de librerias: https://programminghistorian.org/es/lecciones/instalar-modulos-python-pip
    -Los comandos a instalar luego de hecho este instructivo son:
    pip install pandas
    pip install StringIO

# Hoja de movimientos:
    -La hoja de movimientos de materiales den el periodo detallado debe ocntener solo las clases de movimiento 101 y 561.
    -La hoja debe llamarse "Movimientos"
    -Las columnas que debe tener obligatoriamente y con los nombres exactos se detallan a continuacion:
        -Material: PartNumber de pieza o material
        -Cantidad: Con la cantidad en numero del registro de ingreso o migración (Depende de la clase de movimiento)
        -Precio unitario: El precio unitario del material en PESOS. Calculo de Importe ML3/Cantidad
        -Importe Unitario ML3: El precio unitario del material en DOALRES. Calculo de Importe ML3/Cantidad/Tipo de cambio de la fecha de         documento de material.



# Hoja de Stock:
    -Identificar cuales son los PartNumber que tuvieron movimientos, haciendo un cruce de datos entre el Stock total y los movimientos. Aquellos que tuvieron movimientos colocarlos en una hoja llamada "Stock con movimientos" (sin comillas).
    -La hoja de "Stock con movimientos" debe contener dos columnas con los siguientes nombres:
    -Materiales: con los PartNumber de materiales
    -Stock: con las cantidades en numero de cada PartNumber

# Resultado del script:
El resultado del scrips va a ser una planilla de Excel llamada "Costo medio.xlsx" que va a contener las siguientes columnas:

    -Material: PartNumber de la pieza o material
    -Costo medio pesos: El costo medio en PESOS del PartNumber calculado.
    -Costo medio dolares: El costo medio en DOLARES del PartNumber calculado.

Estos pre-requisitos deben ser entregados por el negocio. La ejecucion del script será por parte de ICT junto con la entrega de resultados.


El codigo de el sricpt se encuentra todo comentado operacion por operacion en caso de tener que modificar algo.
Para ejecutar el script es necesario instalar el intérprete de Python version 3.6.* o posterior.



