from openpyxl import load_workbook
from PIL import Image, ImageDraw, ImageFont
import os
import numpy as np
import pytesseract
import xml.etree.ElementTree as ET
import re

def get_column_bounding_box(column_xml_path, old_image_shape, new_image_shape, table_bounding_box, threshhold=3):
    root = ET.parse(column_xml_path).getroot()
    column_bounding_box = []

    for bndbox in root.findall('.//bndbox'):
        xmin = int(bndbox.find('xmin').text)
        ymin = int(bndbox.find('ymin').text)
        xmax = int(bndbox.find('xmax').text)
        ymax = int(bndbox.find('ymax').text)

        scale_x = new_image_shape[1] / old_image_shape[1]
        scale_y = new_image_shape[0] / old_image_shape[0]

        scaled_xmin = round(xmin * scale_x)
        scaled_ymin = round(ymin * scale_y)
        scaled_xmax = round(xmax * scale_x)
        scaled_ymax = round(ymax * scale_y)

        # Agregar solo cajas únicas
        if (scaled_xmin, scaled_ymin, scaled_xmax, scaled_ymax) not in column_bounding_box:
            column_bounding_box.append((scaled_xmin, scaled_ymin, scaled_xmax, scaled_ymax))

    if len(table_bounding_box) == 0 and column_bounding_box:
        min_x = min(bbox[0] for bbox in column_bounding_box) - threshhold
        min_y = min(bbox[1] for bbox in column_bounding_box) - threshhold
        max_x = max(bbox[2] for bbox in column_bounding_box) + threshhold
        max_y = max(bbox[3] for bbox in column_bounding_box) + threshhold
        table_bounding_box = [(min_x, min_y, max_x, max_y)]

    return column_bounding_box, table_bounding_box

def obtener_textos_originales(imagen, filas_xml, columnas, pos_x, pos_y, sep_lineas, anchos_col, offset_x=0, margen_lateral=20, espacios_adicionales=None):
    imagen_gris = imagen.convert("L")
    imagen_binarizada = imagen_gris.point(lambda x: 0 if x < 150 else 255, '1')
    datos_originales = []

    # Procesar cada fila definida en el XML
    for fila_idx in range(len(filas_xml)):  # Cambiado para asegurar que iteramos sobre el rango
        fila = filas_xml[fila_idx]
        fila_textos = []
        # Obtener las coordenadas de la fila actual
        y = pos_y + fila_idx * sep_lineas
        x = pos_x + offset_x
        
        for col_idx in range(columnas):
            if col_idx >= len(fila):  # Asegúrate de no exceder el número de columnas
                break
            
            # Obtener el ancho de la columna desde el XML
            ancho = anchos_col[col_idx] if col_idx < len(anchos_col) else 0
            recorte = imagen_binarizada.crop((x - margen_lateral, y, x + ancho + margen_lateral, y + sep_lineas))
            recorte = recorte.resize((recorte.width * 3, recorte.height * 2))
            recorte.save(f"Recortes/recorte_f{fila_idx}_c{col_idx}.png")
            texto = pytesseract.image_to_string(recorte, config='--psm 6').strip()
            fila_textos.append(texto)
            
            # Usar el espacio correspondiente de la lista
            espacio_variable = espacios_adicionales[col_idx] if espacios_adicionales and col_idx < len(espacios_adicionales) else 0
            x += ancho + (2 * margen_lateral) + espacio_variable
        
        datos_originales.append(fila_textos)

    return np.array(datos_originales)

def es_texto_ruido(texto):
    texto = texto.strip()
    if texto == "":
        return True
    if re.fullmatch(r"[|/\\\-_.,;:]+", texto):
        return True
    if re.fullmatch(r"[A-Za-z]{1,2}", texto):
        return True
    if len(set(texto.lower())) == 1 and texto.isalpha():
        return True
    if not re.search(r"[A-Za-z0-9]", texto):
        return True
    if re.fullmatch(r"ee \|", texto):
        return True
    if len(texto.replace(" ", "")) <= 4 and texto.replace(" ", "").isalpha():
        return True

    return False

def sobrescribir_imagen_con_excel(imagen_path, excel_path, hoja_excel, rango_celdas,
                                  fuente_path=None, tamaño_fuente=11,
                                  anchos_columnas_definidos=None, column_xml_path=None):
    try:
        imagen = Image.open(imagen_path)
        dibujo = ImageDraw.Draw(imagen)

        # Obtener dimensiones de la imagen original
        old_image_shape = imagen.size  # (width, height)
        
        # Definir una nueva forma de imagen si es necesario
        new_image_shape = old_image_shape

        # Cargar las bounding boxes de las columnas
        column_bounding_box, table_bounding_box = get_column_bounding_box(column_xml_path, old_image_shape, new_image_shape, [])
        
        workbook = load_workbook(excel_path)
        sheet = workbook[hoja_excel]
        datos = [[cell.value for cell in row] for row in sheet[rango_celdas]]
        datos_array = np.array(datos)
        print(f"Datos leídos del Excel: {datos_array}")

        if fuente_path and os.path.exists(fuente_path):
            try:
                fuente = ImageFont.truetype(fuente_path, tamaño_fuente)
            except Exception as e:
                print(f"Error al cargar la fuente: {e}")
                fuente = ImageFont.load_default()
        else:
            fuente = ImageFont.load_default()

        posicion_x = 140
        posicion_y = 948
        separacion_lineas = tamaño_fuente + 13
        num_columnas = datos_array.shape[1]
        num_filas = datos_array.shape[0]

        # Calcular los anchos de las columnas desde las bounding boxes
        anchos_col = [bbox[2] - bbox[0] for bbox in column_bounding_box]
        if not anchos_col:
            print("Error: No se encontraron cuadros delimitadores de columnas.")
            return

        # Definir los espacios adicionales entre las capturas
        espacios_adicionales = [20, 5, 15, 8, 12, 6, 10, 7, 5, 10]

        originales_array = obtener_textos_originales(imagen, datos_array.tolist(), num_columnas, posicion_x, posicion_y, separacion_lineas, anchos_col, offset_x=6, margen_lateral=14, espacios_adicionales=espacios_adicionales)
        print("Textos originales extraídos:", originales_array)

        y_actual = posicion_y
        tolerancia_pixeles = 35
        margen_seguridad = 12

        # Dentro del bucle que itera sobre las filas
        for f_idx, fila in enumerate(datos_array):
            if all(dato is None for dato in fila):
                continue
            x_actual = posicion_x
            for c_idx, dato in enumerate(fila):
                if c_idx >= len(originales_array[0]):
                    print(f"Advertencia: Índice de columna {c_idx} fuera de rango para fila {f_idx}.")
                    continue

                texto_nuevo = str(dato) if dato is not None else ""
                bbox_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)
                ancho_nuevo = bbox_nuevo[2] - bbox_nuevo[0]

                # Usar anchos_definidos para el ancho de la casilla
                ancho_casilla = anchos_col[c_idx] if c_idx < len(anchos_col) else 0

                # Verifica si el texto cabe en la casilla
                if ancho_nuevo > ancho_casilla and (ancho_nuevo - ancho_casilla) > tolerancia_pixeles:
                    print(f"No se sobrescribió: [{texto_nuevo}] excede el ancho de la casilla en columna {c_idx + 1}, fila {f_idx + 1}")
                    continue

                # --- fondo del texto ---
                ancho_texto_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)[2]

                if c_idx == 5:
                    x_pos_dibujo = x_actual + anchos_col[c_idx] - ancho_texto_nuevo
                else:
                    x_pos_dibujo = x_actual

                left, top, right, bottom = dibujo.textbbox((x_pos_dibujo, y_actual), texto_nuevo, font=fuente)
                dibujo.rectangle((left-4, top-2, right+2, bottom+2), fill="blue")

                # Dibuja el texto
                dibujo.text((x_actual, y_actual), texto_nuevo, font=fuente, fill=(0, 0, 0))
                
                espacio_variable = espacios_adicionales[c_idx] if c_idx < len(espacios_adicionales) else 0
                x_actual += ancho_casilla + espacio_variable
            y_actual += separacion_lineas
        
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
rango_a_leer = "A2:J2"
fuente_personalizada = r'C:\Windows\Fonts\Arial.ttf'
tamaño_fuente = 20
anchos_definidos = [406, 56, 56, 80, 56, 30, 80, 60, 56, 56]
column_xml_path = r'Test\fila_boundign_boxes.xml'

sobrescribir_imagen_con_excel(imagen_a_modificar, archivo_excel, hoja_a_usar, rango_a_leer, fuente_personalizada, tamaño_fuente, anchos_definidos, column_xml_path)
