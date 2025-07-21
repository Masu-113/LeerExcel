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

        column_bounding_box.append((scaled_xmin, scaled_ymin, scaled_xmax, scaled_ymax))

    if len(table_bounding_box) == 0:
        min_x = min(bbox[0] for bbox in column_bounding_box) - threshhold
        min_y = min(bbox[1] for bbox in column_bounding_box) - threshhold
        max_x = max(bbox[2] for bbox in column_bounding_box) + threshhold
        max_y = max(bbox[3] for bbox in column_bounding_box) + threshhold
        table_bounding_box = [(min_x, min_y, max_x, max_y)]

    return column_bounding_box, table_bounding_box

def obtener_textos_originales(imagen, filas, columnas, pos_x, pos_y, sep_lineas, anchos_col, offset_x=0, margen_lateral=20):
    imagen_gris = imagen.convert("L")
    imagen_binarizada = imagen_gris.point(lambda x: 0 if x < 150 else 255, '1')
    datos_originales = []

    for fila in range(filas):
        fila_textos = []
        x = pos_x + offset_x
        y = pos_y + fila * sep_lineas
        for col in range(columnas):
            ancho = anchos_col[col]
            recorte = imagen_binarizada.crop((x - margen_lateral, y, x + ancho + margen_lateral, y + sep_lineas))
            recorte = recorte.resize((recorte.width * 3, recorte.height * 2))
            # --- esta parte de aqui guarda los recortes temporales ---
            recorte.save(f"Recortes/recorte_f{fila}_c{col}.png")
            # --------
            texto = pytesseract.image_to_string(recorte, config='--psm 6').strip()
            fila_textos.append(texto)
            x += ancho + (2 * margen_lateral)
        datos_originales.append(fila_textos)

    return np.array(datos_originales)
#esto es para limpiar el ruido en  la imagen
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
<<<<<<< HEAD:UmbralizacionTest.py
    #menos de 5 caracteres y solo letras y espacios
    if len(texto.replace(" ", "")) <= 4 and texto.replace(" ", "    ").isalpha():
=======
    if re.fullmatch(r"ee \|", texto):  # Considerar 'ee |' como ruido
        return True
    # Añadir condición: menos de 5 caracteres y solo letras y espacios
    if len(texto.replace(" ", "")) <= 4 and texto.replace(" ", "").isalpha():
>>>>>>> 96118c7bc00b0f9081fff806aca7192885882925:Test/UmbralizacionTest.py
        return True

    return False

def sobrescribir_imagen_con_excel(imagen_path, excel_path, hoja_excel, rango_celdas, fuente_path=None, tamaño_fuente=11, anchos_columnas_definidos=None, column_xml_path=None):
    try:
        imagen = Image.open(imagen_path)
        dibujo = ImageDraw.Draw(imagen)

        workbook = load_workbook(excel_path)
        sheet = workbook[hoja_excel]
        datos = [[cell.value for cell in row] for row in sheet[rango_celdas]]
        datos_array = np.array(datos)
        print(f"Datos leidos del Excel: {datos_array}")

        if fuente_path and os.path.exists(fuente_path):
            try:
                fuente = ImageFont.truetype(fuente_path, tamaño_fuente)
            except:
                fuente = ImageFont.load_default()
        else:
            fuente = ImageFont.load_default()

        posicion_x = 147
        posicion_y = 153
        separacion_lineas = tamaño_fuente + 13
        num_columnas = datos_array.shape[1]
        num_filas = datos_array.shape[0]

        anchos_columnas = [0] * num_columnas
        if anchos_columnas_definidos:
            anchos_columnas = anchos_columnas_definidos
        else:
            for fila in datos_array:
                for idx, dato in enumerate(fila):
                    texto_dato = str(dato) if dato is not None else ""
                    bbox = dibujo.textbbox((0, 0), texto_dato, font=fuente)
                    anchos_columnas[idx] = max(anchos_columnas[idx], bbox[2] - bbox[0])

        old_image_shape = imagen.size
        new_image_shape = (imagen.height, imagen.width)

        column_bounding_box, table_bounding_box = get_column_bounding_box(column_xml_path, old_image_shape, new_image_shape, [])
        print("Cuadros delimitadores de columnas: ", column_bounding_box)

        anchos_col = [column[2] - column[0] for column in column_bounding_box]
<<<<<<< HEAD:UmbralizacionTest.py
        #posiciones para marcar donde va a dibujar y el espacio entre cada cap
        originales_array = obtener_textos_originales(imagen, num_filas, num_columnas, posicion_x, posicion_y, separacion_lineas, anchos_col, offset_x=39, margen_lateral=46)
        print("Textos originales extraidos: ", originales_array)
=======
        #posiciones para marcar donde va a dibujar y el espacio entre cada captura
        originales_array = obtener_textos_originales(imagen, num_filas, num_columnas, posicion_x, posicion_y, separacion_lineas, anchos_col, offset_x=6, margen_lateral=15)
        print("Textos originales extraidos:", originales_array)
>>>>>>> 96118c7bc00b0f9081fff806aca7192885882925:Test/UmbralizacionTest.py

        y_actual = posicion_y
        tolerancia_pixeles = 35
        margen_seguridad = 12  # o el valor que te funcione mejor
        posiciones_columnas = [bbox[0] for bbox in column_bounding_box]

        for f_idx, fila in enumerate(datos_array):
            if all(dato is None for dato in fila):
                continue
            x_actual = posicion_x
            for c_idx, dato in enumerate(fila):
                texto_nuevo = str(dato) if dato is not None else ""
                bbox_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)
                ancho_nuevo = bbox_nuevo[2] - bbox_nuevo[0]

                texto_original = ""
                if f_idx < originales_array.shape[0] and c_idx < originales_array.shape[1]:
                    texto_original = originales_array[f_idx][c_idx]
                ancho_original = dibujo.textbbox((0, 0), texto_original, font=fuente)[2] - dibujo.textbbox((0, 0), texto_original, font=fuente)[0]
                print(f"Dato: [{texto_nuevo}] hubicado en (columna {c_idx + 1}, fila {f_idx + 1})")

                # Calcular el ancho de la casilla en la imagen (basado en coordenadas)
                if c_idx < len(posiciones_columnas) - 1:
                    ancho_casilla = posiciones_columnas[c_idx + 1] - posiciones_columnas[c_idx] - margen_seguridad
                else:
<<<<<<< HEAD:UmbralizacionTest.py
                    if ancho_nuevo > ancho_original + tolerancia_pixeles:
                        print(f"No se sobrescribio: [{texto_nuevo}] no cabe en el espacio de '{texto_original}' (columna {c_idx + 1}, fila {f_idx + 1})")
                        x_actual += anchos_columnas[c_idx] +  42
                        continue # si no alcanza el nuevo dato en la casilla no lo sobrescribe y pasa al siguiente
=======
                    ancho_casilla = anchos_col[c_idx]  # última columna, usar valor estimado
                print(f"Texto: '{texto_nuevo}' | Ancho texto: {ancho_nuevo} | Ancho casilla: {ancho_casilla}")

                if ancho_nuevo > ancho_casilla and (ancho_nuevo - ancho_casilla) > tolerancia_pixeles:
                    print(f"No se sobrescribio: [{texto_nuevo}] excede el ancho de la casilla en columna {c_idx + 1}, fila {f_idx + 1}")
                    x_actual += anchos_columnas[c_idx] + 42
                    continue

>>>>>>> 96118c7bc00b0f9081fff806aca7192885882925:Test/UmbralizacionTest.py
                # --- fondo del texto ---
                # Calcular el ancho del texto a dibujar
                ancho_texto_nuevo = dibujo.textbbox((0, 0), texto_nuevo, font=fuente)[2]

                # esto es para modificar la alineacion de todos los datos de columnas en especifico
                if c_idx == 5:
                    x_pos_dibujo = x_actual + anchos_columnas[c_idx] - ancho_texto_nuevo
                else:
                    x_pos_dibujo = x_actual

                # Fondo del texto
                left, top, right, bottom = dibujo.textbbox((x_pos_dibujo, y_actual), texto_nuevo, font=fuente)
                dibujo.rectangle((left-4, top-2, right+2, bottom+2), fill="white")

                # Dibujo del texto
                dibujo.text((x_pos_dibujo, y_actual), texto_nuevo, font=fuente, fill=(0, 0, 0))

                x_actual += anchos_columnas[c_idx] +  42
            y_actual += separacion_lineas
        
        nombre_imagen, extension = os.path.splitext(imagen_path)
        imagen_modificada = f"{nombre_imagen}_modificada{extension}"
        imagen.save(imagen_modificada)
        print(f"Imagen guardada como {imagen_modificada}")

    except FileNotFoundError:
        print("Error: Archivo no encontrado.")
    except Exception as e:
        print(f"Ocurrio un error: {e}")

<<<<<<< HEAD:UmbralizacionTest.py
imagen_a_modificar = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1.jpg'
archivo_excel = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
hoja_a_usar = "Sheet1"
rango_a_leer = "A1:G10"
fuente_personalizada = r'C:\Windows\Fonts\Arial.ttf'
tamaño_fuente = 30
anchos_definidos = [125, 126, 126, 124, 125, 150, 125] #espacio entre las columnas
column_xml_path = r'C:\Users\msuarez\source\repos\LeerExcel\column_bounding_boxes.xml'
=======
# Parámetros
imagen_a_modificar = r'C:\Users\Marlon Jose\source\repos\LeerExcel\prueba\page_1.jpg'
archivo_excel = r'C:\Users\Marlon Jose\Documents\PruebaExcel.xlsx'
hoja_a_usar = "Hoja1"
rango_a_leer = "A1:G10"
fuente_personalizada = r'C:\Windows\Fonts\Arial.ttf'
tamaño_fuente = 30
#esto es para definir la distancia en pixeles de las columnas q se imprimen
anchos_definidos = [125, 126, 126, 124, 125, 173, 125]
column_xml_path = r'C:\Users\Marlon Jose\source\repos\LeerExcel\Scripts\column_bounding_boxes.xml'
>>>>>>> 96118c7bc00b0f9081fff806aca7192885882925:Test/UmbralizacionTest.py

sobrescribir_imagen_con_excel(imagen_a_modificar, archivo_excel, hoja_a_usar, rango_a_leer, fuente_personalizada, tamaño_fuente, anchos_definidos, column_xml_path)
