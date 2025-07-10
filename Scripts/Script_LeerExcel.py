import openpyxl

def leer_excel(ruta_archivo):
    try:
        workbook = openpyxl.load_workbook(ruta_archivo)

        # Seleccionar la primera hoja
        sheet = workbook.active
        
        for row in sheet.iter_rows():
            # Crear una lista para almacenar los valores de cada celda
            fila_datos = []
            for cell in row:
                # Agregar el valor de la celda a la lista
                fila_datos.append(cell.value)
            # Imprimir la fila de datos
            print(fila_datos)

    except FileNotFoundError:
        print(f"Error: Archivo no encontrado en {ruta_archivo}")
    except Exception as e:
        print(f"Error al leer el archivo: {e}")

# Ejemplo de uso
ruta_del_archivo = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
leer_excel(ruta_del_archivo)
