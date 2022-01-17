import excepciones
import ui

import openpyxl
from openpyxl.styles import PatternFill
import re

pathEnvases = ''
pathMercadonaCsv = ''
# envasesFile = None
# envases = None
# mercadonaCsv = None

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
        mercadonaCsv = openpyxl.load_workbook(filename=pathMercadonaCsv).active
    except:
        excepciones.error(0, "Error al leer " + pathMercadonaCsv)
        archivosLeidos=False

    if(archivosLeidos):
        envasesCorrecto = False
        mercadonaCorrecto = False

        if (envases['A1'].value == 'Tipo mov producto') and (envases['A2'].value == 'Grupocontableproducto'):
            envasesCorrecto = True
        if (mercadonaCsv['A1'].value == 'Factura') and (mercadonaCsv['A2'].value == 'Centro'):
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
                    encontrado = buscaAlbaran(mercadonaCsv,envases[cellAlbaran].value, columna, envases[currentCell].value)
                    if encontrado:
                        coloreaCelda(envases,columna + str(i))

    envasesFile.save(filename=pathEnvases)
    ui.final()

    # print('Error de escritura')
    print('done')


def buscaAlbaran(mercadonaCsv,numAlbaran, columna, valor):
    articulo = ''
    if columna == 'B':
        articulo = '612 - Envase Plegable AECOC 60-40-12 (612)'
    elif columna == 'C':
        articulo = '618 - Envase Plegable AECOC 60-40-18 (618)'
    elif columna == 'D':
        articulo = '624 - Envase Plegable AECOC 60-40-24 (624)'
    elif columna == 'E':
        articulo = '316 - Envase Plegable AECOC 30-40-16 (316)'

    lastfile = len(mercadonaCsv['A'])

    for i in range(3, lastfile):
        cellAlbaran = 'F' + str(i)
        if ((numAlbaran in mercadonaCsv[cellAlbaran].value) and ((mercadonaCsv)['C' + str(i)].value == articulo) and (
                (mercadonaCsv)['H' + str(i)].value + valor == 0)):
            return (True)


def coloreaCelda(envases,celda):
    cell = envases[celda]
    cell.fill = PatternFill(fill_type='solid', start_color='FF33FFCE', end_color='FF33FFCE')
