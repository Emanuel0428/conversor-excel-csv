import pandas as pd
import os

# Ruta donde están guardados los archivos de Excel
ruta_archivos = "C:/Users/cod28/OneDrive - UPB/Desktop/RetoTalentoB/old"

# Listas para almacenar los DataFrames procesados
datos_capitacion = []  # Para "GIRO DIRECTO CAPITACION"
datos_evento = []  # Para "GIRO DIRECTO EVENTO"

# Diccionario para traducir los nombres de las hojas
hojas = {
    "GIRO DIRECTO CAPITACION": "Capitacion",
    "GIRO DIRECTO EVENTO": "Evento"
}

# Nombres de meses en español para manejar los archivos y agregar al DataFrame
nombres_meses = [
    "enero", "febrero", "marzo", "abril", "mayo", "junio",
    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"
]

# Función para leer archivos Excel con soporte para múltiples motores
def leer_excel(archivo, hoja):
    try:
        return pd.read_excel(archivo, sheet_name=hoja, engine="xlrd")  # Para .xls
    except Exception as e_xlrd:
        try:
            return pd.read_excel(archivo, sheet_name=hoja, engine="openpyxl")  # Para .xlsx
        except Exception as e_openpyxl:
            print(f"Error leyendo archivo {archivo}, hoja {hoja}: {e_xlrd} | {e_openpyxl}")
            return pd.DataFrame()  # Retorna un DataFrame vacío en caso de error

# Función para procesar archivos de Excel
def procesar_archivo(archivo, anio, nombre_mes):
    for hoja, tipo in hojas.items():
        datos = leer_excel(archivo, hoja)
        if datos.empty:
            print(f"Advertencia: La hoja {hoja} en el archivo {archivo} está vacía o no se pudo leer.")
            continue
        try:
            # Convertir TotalGiroMes eliminando las comas y puntos decimales
            datos['TotalGiroMes'] = datos['TotalGiroMes'].replace('[,]', '', regex=True).astype(float)

            # Añadir columnas para identificar el tipo de hoja y el mes
            datos['Tipo'] = tipo
            datos['Mes'] = f"{nombre_mes.capitalize()} {anio}"

            # Separar los datos según el tipo de hoja
            if hoja == "GIRO DIRECTO CAPITACION":
                datos_capitacion.append(datos)
            elif hoja == "GIRO DIRECTO EVENTO":
                datos_evento.append(datos)

        except Exception as e:
            print(f"Error procesando {archivo}, hoja {hoja}: {e}")

# Procesar los archivos desde noviembre 2023 hasta noviembre 2024
for anio in [2023, 2024]:
    if anio == 2023:
        rango_meses = range(10, 12)  # Noviembre y Diciembre
    else:
        rango_meses = range(11)  # Hasta Noviembre

    for mes in rango_meses:
        nombre_mes = nombres_meses[mes]
        archivo = os.path.join(
            ruta_archivos,
            f"giro-directo-discriminado-capita-y-evento-{nombre_mes}-{anio}.xls"
        )
        if os.path.exists(archivo):
            procesar_archivo(archivo, anio, nombre_mes)
        else:
            print(f"Archivo no encontrado: {archivo}")

# Combinar los DataFrames de "GIRO DIRECTO CAPITACION" y "GIRO DIRECTO EVENTO"
if datos_capitacion:
    datos_capitacion_final = pd.concat(datos_capitacion, ignore_index=True)
    archivo_salida_capitacion = "C:/Users/cod28/OneDrive - UPB/Desktop/RetoTalentoB/datos_capitacion.csv"
    datos_capitacion_final.to_csv(archivo_salida_capitacion, index=False)
    print(f"Archivo consolidado de CAPITACION guardado en: {archivo_salida_capitacion}")
else:
    print("No se procesaron datos de CAPITACION. Verifica los archivos y su formato.")

if datos_evento:
    datos_evento_final = pd.concat(datos_evento, ignore_index=True)
    archivo_salida_evento = "C:/Users/cod28/OneDrive - UPB/Desktop/RetoTalentoB/datos_evento.csv"
    datos_evento_final.to_csv(archivo_salida_evento, index=False)
    print(f"Archivo consolidado de EVENTO guardado en: {archivo_salida_evento}")
else:
    print("No se procesaron datos de EVENTO. Verifica los archivos y su formato.")
