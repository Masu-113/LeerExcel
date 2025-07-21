import os
from Scripts.PDF_A_IMG import pdf_to_images
from Scripts.IMG_A_PDF import convert_images_to_pdf
from Scripts.Script_Mod_img import sobrescribir_imagen_con_excel

# ====== PARÁMETROS GENERALES ======
ruta_pdf_entrada = r"C:\Users\Marlon Jose\source\repos\LeerExcel\documento_modificado.pdf"
carpeta_imagenes = r"C:\Users\Marlon Jose\source\repos\LeerExcel\prueba"
archivo_excel = r"C:\Users\Marlon Jose\Documents\PruebaExcel.xlsx"
hoja_excel = "Hoja1"
rango_celdas = "A1:G10"
fuente_path = r"C:\Windows\Fonts\Arial.ttf"
tamaño_fuente = 30
anchos_definidos = [125, 126, 126, 124, 125, 173, 125]
column_xml_path = r'C:\Users\Marlon Jose\source\repos\LeerExcel\Scripts\column_bounding_boxes.xml'
salida_pdf_final = r"C:\Users\Marlon Jose\Documents\ResultadoFinal.pdf"

# ====== PASO 1: PDF a IMÁGENES ======
print("Convirtiendo PDF a imágenes...")
pdf_to_images(ruta_pdf_entrada, carpeta_imagenes)

# ====== PASO 2: SOBRESCRIBIR TEXTO EN CADA IMAGEN ======
print("Procesando imágenes...")
for nombre_archivo in os.listdir(carpeta_imagenes):
    if nombre_archivo.lower().endswith(('.jpg', '.jpeg', '.png')) and "modificada" not in nombre_archivo:
        ruta_imagen = os.path.join(carpeta_imagenes, nombre_archivo)
        print(f"Modificando imagen: {ruta_imagen}")
        sobrescribir_imagen_con_excel(
            imagen_path=ruta_imagen,
            excel_path=archivo_excel,
            hoja_excel=hoja_excel,
            rango_celdas=rango_celdas,
            fuente_path=fuente_path,
            tamaño_fuente=tamaño_fuente,
            anchos_columnas_definidos=anchos_definidos,
            column_xml_path=column_xml_path
        )

# ====== PASO 3: IMÁGENES A PDF ======
print("Generando PDF final...")
imagenes_modificadas = sorted([
    os.path.join(carpeta_imagenes, f)
    for f in os.listdir(carpeta_imagenes)
    if f.lower().endswith(('.jpg', '.jpeg', '.png')) and "modificada" in f
])

convert_images_to_pdf(imagenes_modificadas, salida_pdf_final)
print(f"PDF final generado: {salida_pdf_final}")
