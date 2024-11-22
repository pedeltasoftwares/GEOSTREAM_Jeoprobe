"""
VERIFICA LA MAC PARA DAR ACCESO
"""
from _lib.get_files_path import get_file_paths
import os
from getmac import get_mac_address
import gspread
import tkinter
import tkinter.messagebox

# Función para verificar la dirección MAC y mostrar la interfaz si está autorizada
def verificar_mac_y_ejecutar_programa():

    #Path de la librería
    lib_path = get_file_paths("_lib")
    #conecta con la hoja
    file_key = "1eqoRsAenQj5azn99eG0fPyj8kRFHaySvm6jBldVBmQw"
    sheet_name = "Equipos autorizados"
    token_path = os.path.join(lib_path, "service_account.json")
    sheet = googleSpreadSheetConnect(file_key,sheet_name,token_path)
    mac = read_mac(sheet)
    print(mac)
    mac_address_actual = get_mac_address()
    print(mac_address_actual)
    if mac_address_actual in mac:
        # Si la dirección MAC coincide, mostrar la interfaz
        return True
    else:
        # Si la dirección MAC no coincide, mostrar un mensaje de error y salir del programa
        tkinter.messagebox.showerror("Error", "Licencia no activa.")
        return False

def googleSpreadSheetConnect(file_key:str,sheet_name:str,token_path:str):
    """
    Input args:
        file_key: Código genérico del archivo (se encuentra en la url)
        sheet_name: Nombre de la hoja
        token_path: Ruta del token para conectarse a la API
    Output args:
        sheet: Hoja como objeto
    """

    #Conecta con google spreadsheet
    service = gspread.service_account(token_path)

    #Obtiene la hoja
    workbook = service.open_by_key(file_key)
    sheet = workbook.worksheet(sheet_name)

    return sheet

def read_mac(sheet:object):
    """
    Input args:
        sheet: Hoja como objeto
    Output args:
        dataframe: DataFrame con los registros de los usuarios
    """

    mac  = sheet.col_values(1)[1:]

    return mac