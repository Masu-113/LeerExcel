import pandas as pd
from PyPDF2 import  PdfReader, PdfWriter

excel_file = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
df = pd.read_excel(excel_file)

pdf_file = r'C:\Users\msuarez\Documents\TestExcel.pdf'
reader = PdfReader(pdf_file)
writer = PdfWriter()

for page_num in range(len(reader.pages)):
    page = reader.pages[page_num]
    page_content = page.extract_text() # Extraer texto de la página
    
    for index, row in df.iterrows():
        # Buscar y reemplazar texto en el contenido de la página
        if "MARCADOR_NOMBRE" in page_content:
            page_content = page_content.replace("MARCADOR_NOMBRE", row['Nombre'])
        if "MARCADOR_FECHA" in page_content:
            page_content = page_content.replace("MARCADOR_FECHA", str(row['Fecha']))
        # ... agregar más reemplazos según sea necesario ...
    
    # Actualizar el contenido de la página
    page.merge_page(page) # Merge para que se aplique el contenido modificado
    writer.add_page(page)

    with open("documento_modificado.pdf", "wb") as output_pdf:
        writer.write(output_pdf)