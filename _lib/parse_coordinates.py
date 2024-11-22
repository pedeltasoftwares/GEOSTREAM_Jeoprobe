"""
FORMATEA LAS COORDENADAS    
"""
def parse_coordinates(coordenadas):

    latitud = coordenadas.split(' ')[0:3]
    longitud = coordenadas.split(' ')[3:]

    # Procesar latitud
    lat_sign = latitud[0][0]  # N o S
    lat_grados = latitud[0][1:]
    lat_minutos = latitud[1]
    lat_segundos = latitud[2]
    latitud_final = f"{lat_grados}° {lat_minutos}'{lat_segundos}''{lat_sign}"

    # Procesar longitud
    lon_sign = longitud[0][0]  # W o E
    lon_grados = longitud[0][1:]
    lon_minutos = longitud[1]
    lon_segundos = longitud[2]
    lon_sign_final = 'O' if lon_sign == 'W' else 'E'  # Convertir W a O y mantener E
    longitud_final = f"{lon_grados}° {lon_minutos}'{lon_segundos}''{lon_sign_final}"

    return latitud_final, longitud_final
