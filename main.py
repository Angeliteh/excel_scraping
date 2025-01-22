from utils.excel_utils import procesar_archivo, consolidar_archivos, inyectar_datos_en_plantilla
import pandas as pd

# Configuración
archivos = [
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_1.xlsx",
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_2.xlsx",
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_3.xlsx"
]  # Lista de rutas de los archivos
rango_tabla = (5, 1, 16, 26) # Define el rango de la tabla
rango_sumatoria = (3, 8, 11, 19)  # Rango para sumar valores
rango_inyeccion= (8, 8, 16, 19)
hoja_nombre = "ESC2"  # Nombre de la hoja donde se encuentra la tabla
archivo_plantilla = r"C:\Users\Angel\Desktop\plantilla base.xlsx" # Ruta del archivo plantilla


# Procesar y depurar un solo archivo (una tabla)
print("Procesando archivo individual para depuración...")
df_individual = procesar_archivo(archivos[0], rango_tabla, hoja_nombre)
print("\nTabla procesada del archivo individual:")
print(df_individual)

 # Generar archivo Excel de depuración
# archivo_salida = r"C:\Users\Angel\Desktop\archivo_individual_depurado_prueba1.xlsx"
# with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
#     # Exportar el DataFrame al archivo Excel
#     df_individual.to_excel(writer, index=False,header=True, sheet_name='Datos Procesados')

# print(f"Archivo de depuración guardado en: {archivo_salida}")

# Consolidar los datos
print("\nConsolidando datos de los archivos...")
df_consolidado = consolidar_archivos(archivos, rango_sumatoria, rango_tabla, hoja_nombre)

# Guardar archivo de depuración para el consolidado
# archivo_consolidado_depuracion = r"C:\Users\Angel\Desktop\archivo_consolidado_depurado_prueba.xlsx"
# with pd.ExcelWriter(archivo_consolidado_depuracion, engine='openpyxl') as writer:
#     df_consolidado.to_excel(writer, index=False, header=False, sheet_name='Consolidado')
# print(f"Archivo de depuración consolidado guardado en: {archivo_consolidado_depuracion}")

# Cargar la plantilla base (sin procesar, solo cargarla tal como está)
print("\nCargando plantilla base para la inyección de datos...")
from openpyxl import load_workbook
wb_plantilla = load_workbook(archivo_plantilla)
hoja_plantilla = wb_plantilla["ESC2"]

# Reinyectar los datos consolidados en el archivo base (plantilla)
print("\nInyectando los datos consolidados en la plantilla...")
inyectar_datos_en_plantilla(df_consolidado, hoja_plantilla, rango_inyeccion)  # Rango H8:S16

# Guardar archivo final después de la reinyección
archivo_final_depuracion = r"C:\Users\Angel\Desktop\archivo_final_consolidado_PRUEBA.xlsx"
wb_plantilla.save(archivo_final_depuracion)
print(f"Archivo final después de la reinyección guardado en: {archivo_final_depuracion}")