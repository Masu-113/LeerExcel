from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
import os
import numpy as np

def sobrescribir_imagen_con_excel(imagen_path, excel_path, hoja_excel, rango_celdas, fuente_path=None, tamaño_fuente=12, anchos_columnas_definidos=None):
    try:
        # Cargar la imagen
        imagen = Image.open(imagen_path)
        dibujo = ImageDraw.Draw(imagen)

        # Cargar los datos de Excel
        workbook = load_workbook(excel_path)
        sheet = workbook[hoja_excel]

        # Extraer datos del rango especificado
        datos = []
        for row in sheet[rango_celdas]:
            fila_datos = [cell.value for cell in row]
            datos.append(fila_datos)

        # Convertir datos a un array de NumPy para un mejor manejo
        datos_array = np.array(datos)

        # Cargar la fuente si se proporciona
        if fuente_path and os.path.exists(fuente_path):
            try:
                fuente = ImageFont.truetype(fuente_path, tamaño_fuente)
                print(f"Fuente cargada: {fuente_path} con tamaño {tamaño_fuente}")
            except Exception as e:
                print(f"Error al cargar la fuente: {e}. Usando fuente predeterminada.")
                fuente = ImageFont.load_default()
        else:
            fuente = ImageFont.load_default()

        # Definir posición inicial
        posicion_x = 147
        posicion_y = 153
        separacion_lineas = tamaño_fuente + 13  # Espaciado vertical entre filas

        # Medir los anchos maximos por columna para alinear
        num_columnas = datos_array.shape[1]
        anchos_columnas = [0] * num_columnas

        if anchos_columnas_definidos:
            # Usar anchos definidos
            anchos_columnas = anchos_columnas_definidos
        else:
            # Calcular anchos automáticamente
            for fila in datos_array:
                for idx, dato in enumerate(fila):
                    texto_dato = str(dato) if dato is not None else ""
                    bbox = dibujo.textbbox((0, 0), texto_dato, font=fuente)
                    ancho_texto = bbox[2] - bbox[0]
                    if ancho_texto > anchos_columnas[idx]:
                        anchos_columnas[idx] = ancho_texto

        # Dibujar alineando las columnas
        for fila in datos_array:
            posicion_x_columna = posicion_x
            for idx, dato in enumerate(fila):
                texto_dato = str(dato) if dato is not None else ""
                dibujo.text((posicion_x_columna, posicion_y), texto_dato, font=fuente, fill=(0, 0, 0))
                posicion_x_columna += anchos_columnas[idx] + 42  # espacios entre columnas
            posicion_y += separacion_lineas

        # Guardar la imagen modificada
        nombre_imagen, extension = os.path.splitext(imagen_path)
        imagen_modificada = f"{nombre_imagen}_modificada{extension}"
        imagen.save(imagen_modificada)
        print(f"Imagen guardada como {imagen_modificada}")

    except FileNotFoundError:
        print("Error: Archivo no encontrado.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Parámetros de entrada
imagen_a_modificar = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1.jpg'
archivo_excel = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
hoja_a_usar = "Sheet1"
rango_a_leer = "A1:F10"
fuente_personalizada = r"C:\Windows\Fonts\Arial.ttf"
tamaño_fuente = 30

# Definir anchos de columnas
anchos_definidos = [124, 125, 126, 124, 125, 124]

sobrescribir_imagen_con_excel(imagen_a_modificar, archivo_excel, hoja_a_usar, rango_a_leer, fuente_personalizada, tamaño_fuente, anchos_definidos)