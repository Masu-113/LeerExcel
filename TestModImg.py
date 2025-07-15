from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
import os
import numpy as np
import pytesseract

def obtener_textos_originales(imagen, filas, columnas, pos_x, pos_y, sep_lineas, anchos_col):

    imagen_gris = imagen.convert("L")
    imagen_binarizada = imagen_gris.point(lambda x: 0 if x < 160 else 255, '1')  # Binariza para mejorar OCR
    #imagen_binarizada = imagen_gris.point(lambda x: 255 if x > 90 else 0, mode = '1')
    datos_originales = []
    margen = 12

    for fila in range(filas):
        fila_textos = []
        x = pos_x
        y = pos_y + fila * sep_lineas
        for col in range(columnas):
            ancho = anchos_col[col]
            recorte = imagen_binarizada.crop((x - margen, y, x + ancho + margen, y + sep_lineas))
            texto = pytesseract.image_to_string(recorte, config='--psm 6').strip()
            fila_textos.append(texto)
            x += ancho + 42
        datos_originales.append(fila_textos)

    return np.array(datos_originales)

def sobrescribir_imagen_con_excel(imagen_path, excel_path, hoja_excel, rango_celdas, fuente_path=None, tamaño_fuente=11, anchos_columnas_definidos=None):
    try:
        # Cargar imagen
        imagen = Image.open(imagen_path)
        dibujo = ImageDraw.Draw(imagen)

        # Excel
        workbook = load_workbook(excel_path)
        sheet = workbook[hoja_excel]
        datos = [[cell.value for cell in row] for row in sheet[rango_celdas]]
        datos_array = np.array(datos)

        # Fuente
        if fuente_path and os.path.exists(fuente_path):
            try:
                fuente = ImageFont.truetype(fuente_path, tamaño_fuente)
            except:
                fuente = ImageFont.load_default()
        else:
            fuente = ImageFont.load_default()

        # Posiciones
        posicion_x = 147
        posicion_y = 153
        separacion_lineas = tamaño_fuente + 13
        num_columnas = datos_array.shape[1]
        num_filas = datos_array.shape[0]

        # Anchos
        anchos_columnas = [0] * num_columnas
        if anchos_columnas_definidos:
            anchos_columnas = anchos_columnas_definidos
        else:
            for fila in datos_array:
                for idx, dato in enumerate(fila):
                    texto_dato = str(dato) if dato is not None else ""
                    bbox = dibujo.textbbox((0, 0), texto_dato, font=fuente)
                    anchos_columnas[idx] = max(anchos_columnas[idx], bbox[2] - bbox[0])

        # Extraer texto original
        originales_array = obtener_textos_originales(imagen, num_filas, num_columnas, posicion_x, posicion_y, separacion_lineas, anchos_columnas)

        # Dibujar y comparar
        y_actual = posicion_y
        tolerancia_pixeles = 15  # margen de comparación

        for f_idx, fila in enumerate(datos_array):
            x_actual = posicion_x
            for c_idx, dato in enumerate(fila):
                texto_nuevo = str(dato) if dato is not None else ""
                bbox_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)
                ancho_nuevo = bbox_nuevo[2] - bbox_nuevo[0]

                texto_original = originales_array[f_idx][c_idx] if f_idx < originales_array.shape[0] and c_idx < originales_array.shape[1] else ""
                bbox_original = dibujo.textbbox((0, 0), texto_original, font=fuente)
                ancho_original = bbox_original[2] - bbox_original[0]

                if ancho_nuevo > ancho_original + tolerancia_pixeles:
                    print(f"Advertencia: '{texto_nuevo}' no cabe en el espacio de '{texto_original}' (columna {c_idx + 1}, fila {f_idx + 1})")

                dibujo.text((x_actual, y_actual), texto_nuevo, font=fuente, fill=(0, 0, 0))
                x_actual += anchos_columnas[c_idx] + 42
            y_actual += separacion_lineas

        # Guardar imagen
        nombre_imagen, extension = os.path.splitext(imagen_path)
        imagen_modificada = f"{nombre_imagen}_modificada{extension}"
        imagen.save(imagen_modificada)
        print(f"Imagen guardada como {imagen_modificada}")

    except FileNotFoundError:
        print("Error: Archivo no encontrado.")
    except Exception as e:
        print(f"Ocurrió un error: {e}")

# Parámetros
imagen_a_modificar = r'C:\Users\Marlon Jose\source\repos\LeerExcel\prueba\page_1.jpg'
archivo_excel = r'C:\Users\Marlon Jose\Documents\PruebaExcel.xlsx'
hoja_a_usar = "Hoja1"
rango_a_leer = "A1:M10"
fuente_personalizada = r'C:\Windows\Fonts\Arial.ttf'
tamaño_fuente = 30
anchos_definidos = [124, 125, 126, 124, 125, 124, 125, 124, 125, 124, 125, 124, 125, 124]

# Ejecutar
sobrescribir_imagen_con_excel(imagen_a_modificar, archivo_excel, hoja_a_usar, rango_a_leer, fuente_personalizada, tamaño_fuente, anchos_definidos)
