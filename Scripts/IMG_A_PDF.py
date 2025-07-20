import img2pdf
import os

def convert_images_to_pdf(image_list, pdf_path):
    with open(pdf_path, "wb") as f:
        f.write(img2pdf.convert(image_list))

image_folder = r"C:\Users\Marlon Jose\source\repos\LeerExcel\prueba"
image_files = [os.path.join(image_folder, f) for f in os.listdir(image_folder) if f.endswith(('.jpg', '.jpeg', '.png'))]
output_pdf = "mi_documento.pdf"
convert_images_to_pdf(image_files, output_pdf)
print(f"PDF generado: {output_pdf}")