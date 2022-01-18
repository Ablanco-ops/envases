import tkinter as tk
from tkinter import ttk
from tkinter.ttk import *
from tkinter import filedialog as fd
from tkinter import StringVar
from tkinter.messagebox import *
import threading
import os


import comparador

root = tk.Tk()
progressBar = Progressbar()

def final():
    fin = askyesno(title='Proceso terminado', message='Archivo modificado con éxito, ¿desea salir del programa?')
    if fin:
        os._exit(0)
    else:
        pass


def uiMain():

    pathEnvases = StringVar()
    pathMercadona = StringVar()

    root.title('Comparador de archivos de envases')
    root.geometry('600x400')
    root.resizable(False, False)

    def elegir_archivo(archivo):
        filetypes = None
        if archivo=='envases':
            filetypes = (
                ('Archivos Excel', ('*.xlsx')),
            )
        else:
            filetypes = (
                ('Archivos Excel', ('*.csv')),
            )

        filename = fd.askopenfilename(
            title='Elegir archivo',
            initialdir='/',
            filetypes=filetypes
        )
        if (archivo == 'envases'):
            pathEnvases.set(filename)
            comparador.pathEnvases = filename
        else:
            pathMercadona.set(filename)
            comparador.pathClienteCsv = filename

    def uiProcesar():
        uiProgressBar()
        threading.Thread(target=comparador.compruebaArchivos).start()

    def uiProgressBar():
        progressBar = ttk.Progressbar(root, orient='horizontal', mode='indeterminate', length=300, )
        progressBar.grid(columnspan=2, row=6, column=0, pady=10)
        progressBar.start()
        progressBar.update_idletasks()





    root.rowconfigure(0, weight=3)
    root.rowconfigure(1, weight=1)
    root.rowconfigure(2, weight=1)
    root.rowconfigure(3, weight=1)
    root.rowconfigure(4, weight=1)
    root.rowconfigure(5, weight=3)
    root.rowconfigure(6, weight=1)
    root.columnconfigure(0, weight=2)
    root.columnconfigure(1, weight=2)

    etiquetaInstrucciones = ttk.Label(root, text='Selecciona los archivos a utilizar y luego pulsa procesar').grid(
        column=0, row=0, columnspan=2)

    etiquetaArchivo1 = ttk.Label(root, text='Selecciona el archivo de envases').grid(column=0, row=1)
    botonArchivo1 = ttk.Button(root, text='Selecciona el archivo de envases',
                               command=lambda: elegir_archivo('envases')).grid(column=1, row=1,
                                                                               sticky=tk.W)
    campoPathEnvases = ttk.Label(root, textvariable=pathEnvases).grid(column=0, row=2)
    etiquetaArchivo2 = ttk.Label(root, text='Selecciona el archivo de cliente').grid(column=0, row=3)
    botonArchivo2 = ttk.Button(root, text='Selecciona el archivo de cliente',
                               command=lambda: elegir_archivo('cliente')).grid(column=1, row=3,
                                                                                 sticky=tk.W)
    campoPathEnvases = ttk.Label(root, textvariable=pathMercadona).grid(column=0, row=4)

    botonProcesar = ttk.Button(root, text='Procesar', command=lambda: uiProcesar()).grid(column=0, row=5,
                                                                                                        columnspan=2)





    root.mainloop()



