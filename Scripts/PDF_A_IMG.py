from pdf2image import convert_from_path
from PIL import Image
import os

def pdf_to_images(pdf_path, output_dir):

    try:
        images = convert_from_path(pdf_path)

        if not os.path.exists(output_dir):
            os.makedirs(output_dir)

        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, f"page_{i+1}.jpg")
            image.save(image_path, "PNG")
            print(f"Página {i+1} guardada como {image_path}")

    except Exception as e:
        print(f"Error durante la conversión: {e}")

if __name__ == "__main__":
    pdf_to_images(pdf_path = r'C:\Users\Marlon Jose\source\repos\LeerExcel\documento_modificado.pdf', output_dir="prueba")