import pandas as pd
from PIL import Image, ImageDraw, ImageFont

# Reemplaza con la ruta de tu archivo Excel
excel_file = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
df = pd.read_excel(excel_file)

# Imprime el DataFrame para verificar
print(df)

# Reemplaza con la ruta de tu imagen
image_path = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1.jpg'
image = Image.open(image_path)
draw = ImageDraw.Draw(image)

# Reemplaza con la ruta a tu archivo de fuente y el tamaño deseado
font_path =  r"C:\Windows\Fonts\Arial.ttf"
font_size = 20
font = ImageFont.truetype(font_path, font_size)

# Coordenadas iniciales del pixel
x_start = 100
y_start = 100

# Espaciado entre líneas
line_height = font_size + 5  # Ajusta según sea necesario

# Itera sobre las filas del DataFrame y dibuja cada valor
for index, row in df.iterrows():
    text = str(row['Nombres'])  
    text_id = str(row['id']) # Reemplaza 'columna_a_dibujar'
    draw.text((x_start, y_start), text, fill=(0, 0, 0), font=font)
    draw.text((x_start, y_start), text_id, fill=(0, 0, 0), font=font)  # Dibuja el texto
    y_start += line_height  # Incrementa la posición vertical

# Guarda la imagen modificada
image.save('imagen_con_texto.png')