import PySimpleGUI as sg
from utils.data_processing import consolidar_archivos, inyectar_datos_en_plantilla, inyectar_formulas_totales_y_subtotales, convertir_formulas_a_valores
from openpyxl import load_workbook

# Ruta fija de la plantilla base
ARCHIVO_PLANTILLA = r"C:\Users\Angel\Desktop\plantilla base.xlsx"

def ejecutar_proceso(archivos, archivo_final):
    try:
        # Verificar si la lista de archivos es válida
        if not archivos:
            sg.popup_error("No se seleccionaron archivos Excel.")
            return

        # Mostrar los archivos seleccionados (para depuración)
        sg.popup(f"Archivos seleccionados:\n{', '.join(archivos)}", title="Archivos Excel")

        rango_tabla = (5, 1, 16, 26)  # Rango de tabla
        rango_sumatoria = (3, 8, 11, 19)  # Rango de sumatorias
        rango_inyeccion = (6, 8, 14, 19)  # Rango de reinyección
        hoja_entrada = "ESC2"  # Nombre de hoja de entrada
        hoja_salida = "ZONA3"  # Nombre de hoja de salida

        # Consolidar los datos
        sg.popup("Consolidando datos...", title="Procesando")
        df_consolidado = consolidar_archivos(archivos, rango_sumatoria, rango_tabla, hoja_entrada)

        # Cargar la plantilla base
        wb_plantilla = load_workbook(ARCHIVO_PLANTILLA)
        hoja_plantilla = wb_plantilla[hoja_salida]

        # Inyectar datos consolidados en la plantilla
        inyectar_datos_en_plantilla(df_consolidado, hoja_plantilla, rango_inyeccion)

        # Guardar archivo consolidado
        wb_plantilla.save(archivo_final)
        inyectar_formulas_totales_y_subtotales(archivo_final)
        convertir_formulas_a_valores(archivo_final)

        sg.popup("Proceso completado con éxito.", title="Éxito")
        sg.popup(f"Archivo consolidado guardado en: {archivo_final}", title="Archivo generado")
    except Exception as e:
        sg.popup_error(f"Se produjo un error durante el proceso:\n{e}")

# Configuración de la interfaz
layout = [
    [sg.Text("Selecciona los archivos Excel:")],
    [sg.Input(key="-ARCHIVOS-", enable_events=True), sg.FilesBrowse(button_text="Buscar", file_types=(("Archivos Excel", "*.xlsx"),))],
    [sg.Text("Archivo final consolidado se guardará en:")],
    [sg.Input(default_text=r"C:\Users\Angel\Desktop\archivo_final_consolidado.xlsx", key="-ARCHIVO_FINAL-")],
    [sg.Button("Ejecutar"), sg.Button("Salir")],
]

# Crear ventana
window = sg.Window("Consolidación de Archivos Excel", layout)

while True:
    event, values = window.read()

    if event == sg.WINDOW_CLOSED or event == "Salir":
        break

    if event == "Ejecutar":
        archivos = values["-ARCHIVOS-"].split(";")  # Los archivos se devuelven como una cadena separada por puntos y coma
        archivo_final = values["-ARCHIVO_FINAL-"]

        if not archivos or not archivo_final:
            sg.popup_error("Por favor, completa todos los campos antes de continuar.")
        else:
            ejecutar_proceso(archivos, archivo_final)

window.close()
