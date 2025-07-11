from pdf2image import convert_from_path
from PIL import Image, ImageDraw, ImageFont
import openpyxl

def modificar_pdf_con_excel(pdf_path, excel_path, output_image_path):
    # 1. Convertir PDF a imagen
    try:
        images = convert_from_path(pdf_path)
    except Exception as e:
        print(f"Error al convertir PDF a imagen: {e}")
        return

    if not images:
        print("No se pudo generar ninguna imagen del PDF.")
        return

    # Tomar la primera página (puedes iterar si hay múltiples páginas)
    image = images[0]
    image = image.convert("RGB")  # Asegurarse de que la imagen sea RGB

    # 2. Leer datos de Excel
    try:
        workbook = openpyxl.load_workbook(excel_path)
        sheet = workbook.active  # Asume que estás trabajando con la hoja activa
        # Ejemplo: Leer datos de celdas específicas
        dato1 = sheet['A1'].value
        dato2 = sheet['B2'].value
    except Exception as e:
        print(f"Error al leer el archivo Excel: {e}")
        return


    # 3. Modificar la imagen
    draw = ImageDraw.Draw(image)
    # Define una fuente (asegúrate de que la fuente exista en tu sistema)
    try:
        font = ImageFont.truetype("arial.ttf", size=20)  # Ajusta el tamaño y la ruta de la fuente
    except IOError:
        font = ImageFont.load_default()  # Usa la fuente predeterminada si la especificada no existe
        print("Advertencia: La fuente especificada no se encontró. Se usará la fuente predeterminada.")

    # Dibuja texto en la imagen
    draw.text((10, 10), f"Dato 1: {dato1}", fill="black", font=font)
    draw.text((10, 40), f"Dato 2: {dato2}", fill="black", font=font)

    # 4. Guardar la imagen modificada
    try:
        image.save(output_image_path, "JPEG")  # Guarda como JPEG
        print(f"Imagen modificada guardada en: {output_image_path}")
    except Exception as e:
        print(f"Error al guardar la imagen: {e}")

# Ejemplo de uso
pdf_path = r'C:\Users\msuarez\Documents\TestExcel.pdf'
excel_path = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
output_image_path = "imagen_modificada.jpg"

modificar_pdf_con_excel(pdf_path, excel_path, output_image_path)