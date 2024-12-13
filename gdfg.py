import os
from _lib.f01_masw_function import save_to_pdf
import xlwings as xw
import PyPDF2

#Ruta de documentos
documents_path = os.path.expanduser("~\\Documents")

for folder in os.listdir("C:/Users/Katerine Arias/Documents/GEOSTREAM/MASW"):
    print(folder)
    #Imprimir en PDF la memoria
    app = xw.App(visible=False)  
    app.display_alerts = False 
    app.screen_updating = False
    libro = xw.Book(f'{documents_path}//GEOSTREAM//MASW//{folder}//template lineas.xlsx', update_links=True)
    save_to_pdf(libro, folder, documents_path )
    # Guardar los cambios
    libro.save()
    libro.close()
    app.quit()

#Compilar versión final
merger = PyPDF2.PdfMerger()
for folder in os.listdir("C:/Users/Katerine Arias/Documents/GEOSTREAM/MASW"):
        merger.append(f'{documents_path}//GEOSTREAM//MASW//{folder}//{folder}_combinado.pdf')

# Guardar el PDF combinado en un archivo
merger.write(f'{documents_path}//GEOSTREAM//MASW//combinado.pdf')
merger.close()