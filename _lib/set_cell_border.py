"""
BORDE DE TABLAS
"""

def set_border(hoja, columna,fila):

    for border_id in range(1, 5):
        hoja.range(f"{columna}{fila}").api.Borders(border_id).LineStyle = 1  