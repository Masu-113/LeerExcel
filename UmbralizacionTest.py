from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
import os
import numpy as np
import pytesseract
import xml.etree.ElementTree as ET

# Función para obtener los cuadros delimitadores de las columnas desde un XML
def get_column_bounding_box(column_xml_path, old_image_shape, new_image_shape, table_bounding_box, threshhold=3):
    root = ET.parse(column_xml_path).getroot()
    column_bounding_box = []

    for bndbox in root.findall('.//bndbox'):
        xmin = int(bndbox.find('xmin').text)
        ymin = int(bndbox.find('ymin').text)
        xmax = int(bndbox.find('xmax').text)
        ymax = int(bndbox.find('ymax').text)

        # Escala las coordenadas
        scale_x = new_image_shape[1] / old_image_shape[1]
        scale_y = new_image_shape[0] / old_image_shape[0]

        # Redondea las coordenadas escaladas
        scaled_xmin = round(xmin * scale_x)
        scaled_ymin = round(ymin * scale_y)
        scaled_xmax = round(xmax * scale_x)
        scaled_ymax = round(ymax * scale_y)

        column_bounding_box.append((scaled_xmin, scaled_ymin, scaled_xmax, scaled_ymax))

    # Aproximar el cuadro delimitador de la tabla si no se proporciona
    if len(table_bounding_box) == 0:
        min_x = min(bbox[0] for bbox in column_bounding_box) - threshhold
        min_y = min(bbox[1] for bbox in column_bounding_box) - threshhold
        max_x = max(bbox[2] for bbox in column_bounding_box) + threshhold
        max_y = max(bbox[3] for bbox in column_bounding_box) + threshhold
        table_bounding_box = [(min_x, min_y, max_x, max_y)]

    return column_bounding_box, table_bounding_box

def obtener_textos_originales(imagen, filas, columnas, pos_x, pos_y, sep_lineas, anchos_col):
    imagen_gris = imagen.convert("L")
    imagen_binarizada = imagen_gris.point(lambda x: 0 if x < 150 else 255, '1')  # Binariza para mejorar OCR
    datos_originales = []
    margen = 20

    for fila in range(filas):
        fila_textos = []
        x = pos_x
        y = pos_y + fila * sep_lineas
        for col in range(columnas):
            ancho = anchos_col[col]
            recorte = imagen_binarizada.crop((x - margen, y, x + ancho + margen, y + sep_lineas))
            # Escalar recorte (doble tamaño)
            recorte = recorte.resize((recorte.width * 2, recorte.height * 2))
            # Guardar cada recorte como imagen temporal para inspección
            recorte.save(f"recorte_f{fila}_c{col}.png")
            # Realizar OCR
            texto = pytesseract.image_to_string(recorte, config='--psm 7').strip()

            fila_textos.append(texto)
            x += ancho + 42
        datos_originales.append(fila_textos)

    return np.array(datos_originales)

def sobrescribir_imagen_con_excel(imagen_path, excel_path, hoja_excel, rango_celdas, fuente_path=None, tamaño_fuente=11, anchos_columnas_definidos=None, column_xml_path=None):
    try:
        # Cargar imagen
        imagen = Image.open(imagen_path)
        dibujo = ImageDraw.Draw(imagen)

        # Excel
        workbook = load_workbook(excel_path)
        sheet = workbook[hoja_excel]
        datos = [[cell.value for cell in row] for row in sheet[rango_celdas]]
        datos_array = np.array(datos)

        # Imprimir los datos leídos
        print("Datos leídos del Excel:", datos_array)

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

        # Obtener la forma antigua y nueva de la imagen
        old_image_shape = imagen.size  # (width, height)
        new_image_shape = (imagen.height, imagen.width)

        # Obtener los cuadros delimitadores de columnas y la tabla
        column_bounding_box, table_bounding_box = get_column_bounding_box(column_xml_path, old_image_shape, new_image_shape, [])

        # Imprimir los cuadros delimitadores
        print("Cuadros delimitadores de columnas:", column_bounding_box)

        # Definir posiciones y tamaños
        pos_x = table_bounding_box[0][0]  # x_min de la tabla
        pos_y = table_bounding_box[0][1]  # y_min de la tabla
        sep_lineas = table_bounding_box[0][3] - table_bounding_box[0][1]  # altura de la tabla
        anchos_col = [column[2] - column[0] for column in column_bounding_box]  # ancho de cada columna

        # Extraer texto original
        originales_array = obtener_textos_originales(imagen, num_filas, num_columnas, pos_x, pos_y, sep_lineas, anchos_col)

        # Imprimir los textos originales
        print("Textos originales extraídos:", originales_array)

        # Dibujar y comparar
        y_actual = posicion_y
        tolerancia_pixeles = 15  # margen de comparación

        for f_idx, fila in enumerate(datos_array):
            if all(dato is None for dato in fila):  # Saltar filas completamente vacías
                continue
            x_actual = posicion_x
            for c_idx, dato in enumerate(fila):
                texto_nuevo = str(dato) if dato is not None else ""
                bbox_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)
                ancho_nuevo = bbox_nuevo[2] - bbox_nuevo[0]

                # Verificar que el índice está dentro de los límites
                if f_idx < originales_array.shape[0] and c_idx < originales_array.shape[1]:
                    texto_original = originales_array[f_idx][c_idx]
                else:
                    texto_original = ""  # O manejar el caso de otra manera

                bbox_original = dibujo.textbbox((0, 0), texto_original, font=fuente)
                ancho_original = bbox_original[2] - bbox_original[0]

                if texto_original.strip() == "":
                    # Si el OCR no detectó texto, asumir que la celda esta vacía
                    if ancho_nuevo > anchos_col[c_idx] + tolerancia_pixeles:
                        print(f"Advertencia: '{texto_nuevo}' podría no caber en columna {c_idx + 1}, fila {f_idx + 1} (celda vacía en imagen)")
                else:
                    # Si sí hay texto original, comparar con el nuevo
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
rango_a_leer = "A1:F10"
fuente_personalizada = r'C:\Windows\Fonts\Arial.ttf'
tamaño_fuente = 30
anchos_definidos = [124, 125, 126, 124, 125, 124, 125, 124, 125, 124, 125, 124, 125, 124]
column_xml_path = r'C:\Users\Marlon Jose\source\repos\LeerExcel\column_bounding_boxes.xml'  # Ruta del XML de columnas

# Ejecutar
sobrescribir_imagen_con_excel(imagen_a_modificar, archivo_excel, hoja_a_usar, rango_a_leer, fuente_personalizada, tamaño_fuente, anchos_definidos, column_xml_path)
