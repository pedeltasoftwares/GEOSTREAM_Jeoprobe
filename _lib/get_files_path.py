"""
Obtiene el path del archivo a buscar. Sin importar que sea ejecutable o script
"""
import os
import sys

def get_file_paths(file_name):

    if hasattr(sys, '_MEIPASS'):  # Cuando está ejecutado desde un archivo compilado
        base_path = sys._MEIPASS
        if file_name == "_lib":
            return os.path.join(base_path, "_lib\\")
        elif file_name == "_images":
            return os.path.join(base_path, "_images\\")
        elif file_name == "_templates":
            return os.path.join(base_path, "_templates\\")
    else:  # Cuando está ejecutado como un script
        base_path = os.path.dirname(os.path.realpath(__file__))
        if file_name == "_lib":
            return base_path
        elif file_name == "_images":
            return os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "_images")
        elif file_name == "_templates":
            return os.path.join(os.path.dirname(os.path.realpath(sys.argv[0])), "_templates")
