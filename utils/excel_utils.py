from openpyxl import load_workbook
import pandas as pd

def extraer_tabla_y_limpiar(hoja, rango):
    """
    Extrae los datos de un rango específico en una hoja de Excel y descombina las celdas combinadas.

    Args:
        hoja: Objeto Worksheet de openpyxl.
        rango: Tupla con (min_row, min_col, max_row, max_col).

    Returns:
        Una lista de listas representando la tabla extraída.
    """
    min_row, min_col, max_row, max_col = rango

    # Hacer una copia de los rangos de celdas combinadas
    rangos_combinados = list(hoja.merged_cells.ranges)
    for rango_combinado in rangos_combinados:
        if (rango_combinado.min_row >= min_row and rango_combinado.max_row <= max_row and
            rango_combinado.min_col >= min_col and rango_combinado.max_col <= max_col):
            
            valor = hoja.cell(rango_combinado.min_row, rango_combinado.min_col).value
            hoja.unmerge_cells(rango_combinado.coord)

            for fila in hoja.iter_rows(min_row=rango_combinado.min_row, max_row=rango_combinado.max_row,
                                       min_col=rango_combinado.min_col, max_col=rango_combinado.max_col):
                for celda in fila:
                    celda.value = valor

    tabla = [
        [hoja.cell(row=fila, column=col).value for col in range(min_col, max_col + 1)]
        for fila in range(min_row, max_row + 1)
    ]
    return tabla

def procesar_archivo(archivo, rango_tabla, hoja_nombre="ESC2"):
    """
    Extrae y procesa la tabla de un archivo Excel.
    Retorna un DataFrame con los valores de la tabla.
    """
    wb = load_workbook(archivo)
    hoja = wb[hoja_nombre]
    tabla = extraer_tabla_y_limpiar(hoja, rango_tabla)
    df = pd.DataFrame(tabla)
    
    # Renombrar columnas usando la fila de encabezados
    encabezados = df.iloc[1]
    df.columns = encabezados
    df = df.iloc[2:].reset_index(drop=True)
    
    return df

def consolidar_archivos(archivos, rango_tabla):
    """
    Consolida los datos de múltiples archivos Excel.
    Retorna un DataFrame con los valores sumados.
    """
    consolidado = None

    for archivo in archivos:
        df = procesar_archivo(archivo, rango_tabla)
        
        columnas_h = [col for col in df.columns if col.endswith("H")]
        columnas_m = [col for col in df.columns if col.endswith("M")]
        df[columnas_h] = df[columnas_h].apply(pd.to_numeric, errors='coerce').fillna(0)
        df[columnas_m] = df[columnas_m].apply(pd.to_numeric, errors='coerce').fillna(0)

        if consolidado is None:
            consolidado = df
        else:
            consolidado[columnas_h] += df[columnas_h]
            consolidado[columnas_m] += df[columnas_m]
            print(f"Columnas H: {columnas_h}")
            print(f"Columnas M: {columnas_m}")
            print(f"Primer archivo procesado:\n{df[columnas_h].head()}")


    return consolidado

def calcular_totales(df):
    """
    Calcula subtotales y totales de un DataFrame consolidado.
    Retorna un DataFrame con los totales añadidos.
    """
    # Asegurar que el índice sea único y resetearlo si es necesario
    df = df.reset_index(drop=True)

    # Identificar columnas H y M
    columnas_h = [col for col in df.columns if col.endswith("H")]
    columnas_m = [col for col in df.columns if col.endswith("M")]

    # Calcular subtotales por fila
    df["Subtotal H"] = df[columnas_h].sum(axis=1)
    df["Subtotal M"] = df[columnas_m].sum(axis=1)

    # Calcular totales generales
    total_h = df["Subtotal H"].sum()
    total_m = df["Subtotal M"].sum()

    # Crear un DataFrame para los totales generales
    df_totales = pd.DataFrame([["TOTAL", total_h, total_m]], columns=["Concepto", "Subtotal H", "Subtotal M"])

    # Concatenar el DataFrame original con los totales
    return pd.concat([df, df_totales], ignore_index=True)
