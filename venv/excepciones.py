import ui


def excepciones(numero):
    mensaje = ''
    if (numero == 0):
        mensaje = 'Archivos erróneos'
    elif (numero == 1):
        mensaje = 'Archivo de envases erroneo'
    elif (numero == 2):
        mensaje = 'Archivo de Mercadona erroneo'

    ui.showwarning(title='Error', message=mensaje)

def error(num, mensaje):
    if num==0:
        ui.showerror(title='Error de lectura', message=mensaje)
    if num==1:
        ui.showerror(title='Error de escritura', message=mensaje)