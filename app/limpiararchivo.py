def insertar_evento(input_file, output_file):
    try:
        with open(input_file, 'r', encoding='utf-8') as archivo_entrada, \
                open(output_file, 'w', encoding='utf-8') as archivo_salida:

            for linea in archivo_entrada:
                partes = linea.strip().split(',')  # Divide la línea en partes usando coma

                # Inserta ", Evento" antes del número
                if len(partes) > 4:  # Asegúrate de que hay suficientes columnas
                    partes.insert(4, "Evento")

                # Une las partes nuevamente y escribe al archivo de salida
                nueva_linea = ','.join(partes) + '\n'
                archivo_salida.write(nueva_linea)

        print(f"Archivo procesado y guardado en: {output_file}")
    except FileNotFoundError:
        print(f"El archivo '{input_file}' no existe.")
    except Exception as e:
        print(f"Error procesando el archivo: {e}")


input_file = "C:/Users/cod28/OneDrive - UPB/Desktop/RetoTalentoB/pacomas.csv"  # Ruta del archivo original
output_file = "C:/Users/cod28/OneDrive - UPB/Desktop/RetoTalentoB/concomas.csv"  # Ruta del archivo de salida
insertar_evento(input_file, output_file)
