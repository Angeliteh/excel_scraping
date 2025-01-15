from utils.excel_utils import procesar_archivo, consolidar_archivos, calcular_totales
import pandas as pd

# Configuración
archivos = [
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_1.xlsx",
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_2.xlsx",
    r"C:\Users\Angel\Desktop\excel_scraping\excel_scraping\data\FORMATO_ESTADISTICA_ESCUELA_3.xlsx"
]  # Lista de rutas de los archivos
rango_tabla = (5, 1, 16, 26) # Define el rango de la tabla
hoja_nombre = "ESC2"  # Nombre de la hoja donde se encuentra la tabla


# Procesar y depurar un solo archivo
print("Procesando archivo individual para depuración...")
df_individual = procesar_archivo(archivos[0], rango_tabla, hoja_nombre)
print("\nTabla procesada del archivo individual:")
print(df_individual)

# # Generar archivo Excel de depuración
# archivo_salida = r"C:\Users\Angel\Desktop\archivo_individual_depurado.xlsx"
# with pd.ExcelWriter(archivo_salida, engine='openpyxl') as writer:
#     # Exportar el DataFrame al archivo Excel
#     df_individual.to_excel(writer, index=False, sheet_name='Datos Procesados')

# print(f"Archivo de depuración guardado en: {archivo_salida}")

# Consolidar múltiples archivos
print("\nConsolidando archivos...")
df_consolidado = consolidar_archivos(archivos, rango_tabla)
print("\nTabla consolidada (sin totales):")
print(df_consolidado)

# # Generar archivo Excel de consolidación
# archivo_salida_consolidado = r"C:\Users\Angel\Desktop\archivo_concentrado_prueba.xlsx"
# with pd.ExcelWriter(archivo_salida_consolidado, engine='openpyxl') as writer:
#     # Exportar el DataFrame consolidado al archivo Excel
#     df_consolidado.to_excel(writer, index=False, sheet_name='Datos Consolidado')

# print(f"Archivo consolidado guardado en: {archivo_salida_consolidado}")

# # Calcular subtotales y totales
# print("\nCalculando subtotales y totales...")
# df_consolidado_con_totales = calcular_totales(df_consolidado)
# print("\nTabla consolidada con subtotales y totales:")
# print(df_consolidado_con_totales)