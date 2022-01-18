import excepciones
import ui

import openpyxl
from openpyxl.styles import PatternFill
import re
import pandas as pd

pathEnvases = ''
pathMercadonaCsv = ''


newpath = 'Envases.xlsx'


def compruebaArchivos():
    archivosLeidos=True
    try:
        envasesFile = openpyxl.load_workbook(filename=pathEnvases)
        envases = envasesFile[
            "SALIDAS DE ENVASES MERCADONA"]
    except:
        excepciones.error(0, "Error al leer " + pathEnvases)
        archivosLeidos=False

    try:
        mercadonaCsv = pd.read_csv(pathMercadonaCsv, encoding='latin-1', sep=';', header=1)
    except:
        excepciones.error(0, "Error al leer " + pathMercadonaCsv)
        archivosLeidos=False

    if(archivosLeidos):
        envasesCorrecto = False
        mercadonaCorrecto = False

        if (envases['A1'].value == 'Tipo mov producto') and (envases['A2'].value == 'Grupocontableproducto'):
            envasesCorrecto = True
        if (mercadonaCsv.columns.values[0] == 'Centro') and (mercadonaCsv.columns.values[1] == 'Tipo Movimiento'):
            mercadonaCorrecto = True

        if (mercadonaCorrecto and envasesCorrecto):
            recorreEnvases(envases,mercadonaCsv,envasesFile)

        elif (mercadonaCorrecto==False and envasesCorrecto==False):
            excepciones.excepciones(0)
        elif envasesCorrecto==False:
            excepciones.excepciones(1)
        elif mercadonaCorrecto==False:
            excepciones.excepciones(2)


def recorreEnvases(envases, mercadonaCsv, envasesFile):
    lista = ['B', 'C', 'D', 'E']
    lastfile = len(envases['A'])
    for i in range(11, lastfile):
        cellAlbaran = 'A' + str(i)
        if re.search('[0-9]{5}', envases[cellAlbaran].value):

            for columna in lista:
                currentCell = columna + str(i)
                if (envases[currentCell].value):
                    encontrado= buscaAlbaran(mercadonaCsv,envases[cellAlbaran].value, columna, envases[currentCell].value)
                    if encontrado:
                        coloreaCelda(envases,columna + str(i))

    envasesFile.save(filename=pathEnvases)
    ui.final()
    print('done')


def buscaAlbaran(mercadonaCsv,numAlbaran, columna, valor):
    articulo = ''
    if columna == 'B':
        articulo = '(612)'
    elif columna == 'C':
        articulo = '(618)'
    elif columna == 'D':
        articulo = '(624)'
    elif columna == 'E':
        articulo = '(316)'

    tabla=pd.DataFrame(mercadonaCsv, columns=['Centro','Tipo Movimiento','Articulo','Almacén Origen/Destino','Fecha','Albaran','Referencia Destino','Cantidad'])
    filtro = tabla[tabla['Albaran'].str.contains(numAlbaran)]
    if(filtro.empty==False):
        filtro2=filtro[filtro['Articulo'].str.contains(articulo)]
        if (filtro2.empty==False):
            if (valor+filtro2['Cantidad'].values[0]==0):
                return True




def coloreaCelda(envases,celda):
    cell = envases[celda]
    cell.fill = PatternFill(fill_type='solid', start_color='FF33FFCE', end_color='FF33FFCE')
