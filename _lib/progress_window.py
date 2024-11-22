"""
CREA VENTANAS DE PROGRESO
"""
import tkinter
import os
from _lib.get_files_path import get_file_paths
from customtkinter import *

def create_progress_window(title_window,ico_name_fime,text_progress):

    images_path = get_file_paths("_images")

    # Crear una nueva ventana de progreso
    ventana_progreso = tkinter.Toplevel()
    ventana_progreso.title(title_window)
    ventana_progreso.resizable(False, False)
    width = 400
    height = 100
    screen_width = ventana_progreso.winfo_screenwidth()
    screen_height = ventana_progreso.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    ventana_progreso.geometry(f"{width}x{height}+{x}+{y}")
    ventana_progreso.iconbitmap(os.path.join(images_path, ico_name_fime))  
    
    # Ventana esté encima de la ventana principal
    ventana_progreso.lift()
    ventana_progreso.attributes('-topmost', True)  
    ventana_progreso.after(100, lambda: ventana_progreso.attributes('-topmost', False))  

    #Crear la barra de progreso en la nueva ventana
    barra_progreso = CTkProgressBar(master=ventana_progreso, 
                                    width=250,
                                    progress_color="green")
    barra_progreso.set(0)
    barra_progreso.pack(pady=20)

    # Crear texto de progreso
    texto_progreso = tkinter.StringVar(value=text_progress)
    label_progreso = CTkLabel(master=ventana_progreso, 
                              textvariable=texto_progreso,
                              text_color="black", 
                              bg_color='transparent',
                              font=('Gothic A1',15))
    label_progreso.pack(pady=2)

    #Actualiza la ventana
    ventana_progreso.update_idletasks()

    # Retornar los elementos necesarios para actualizar el progreso
    return ventana_progreso, barra_progreso, texto_progreso