"""
INTERFAZ GRÁFICA
"""

import os
from _lib.mac_verification import verificar_mac_y_ejecutar_programa
from customtkinter import *
from _lib.get_files_path import get_file_paths
from PIL import Image, ImageTk
from _lib.f01_masw_function import open_masw_window

#Paths
images_path = get_file_paths("_images")

#Ejecuta la interfaz gráfica
if verificar_mac_y_ejecutar_programa():

    def ejecutar_masw():
        open_masw_window(menu_window,images_path)

    #Inicializa la ventana
    menu_window = CTk()
    #Geometría
    width = 300
    height = 100
    screen_width = menu_window.winfo_screenwidth()
    screen_height = menu_window.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    menu_window.geometry(f"{width}x{height}+{x}+{y}")
    #Nombre de la ventana
    menu_window.title("GEOSTREAM")
    #Resizable
    menu_window.resizable(False,False)
    #Tema de la ventana
    set_appearance_mode("dark")
    #Ícono ventana
    menu_window.after(201, lambda :menu_window.iconbitmap(os.path.join(images_path, "logo.ico")))
    #Botón principal para escalar señales
    masw_image = Image.open(os.path.join(images_path, "formulario-de-llenado.png"))
    masw_image = masw_image.resize((30, 30), Image.LANCZOS)
    masw_image_tk = ImageTk.PhotoImage(masw_image)
    masw = CTkButton(master=menu_window, 
                    text="MASW", 
                    image = masw_image_tk,
                    width=180, 
                    height=40, 
                    compound="left",
                    font=('Gothic A1',15),
                    fg_color="#3A3A3A",
                    hover_color="#4C4C4C",
                    text_color="#E0E0E0",
                    corner_radius=5,
                    border_width=2,
                    border_color="#606060",
                    command=ejecutar_masw)

    masw.place(x=60, y=25)
    #Ejecuta la ventana
    menu_window.mainloop()