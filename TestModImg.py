from openpyxl import load_workbook
from PIL  import Image, ImageDraw, ImageFont
import tkinter as tk
from tkinter import ttk, filedialog
from PIL import ImageTk, Image

def leer_excel (ruta_archivo, hoja, rango_celdas):
    ruta_archivo = r'C:\Users\msuarez\Documents\TestExcel.xlsx'
    workbook = load_workbook(ruta_archivo)
    hoja = workbook[hoja]
    datos = []
    for fila in hoja[rango_celdas]:
        fila_datos = [celda.value for celda in fila]
        datos.append(fila_datos)
    return datos

def agregar_texto_a_imagen(ruta_imagen, texto, posicion, fuente_nombre, tamaño_fuente, color):
    ruta_imagen = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1.jpg'
    imagen = Image.open(ruta_imagen)
    dibujo = ImageDraw.Draw(imagen)
    fuente = ImageFont.truetype(fuente_nombre, tamaño_fuente)
    dibujo.text(posicion, texto, font=fuente, fill=color)
    return imagen


def mostrar_interfaz(imagen, datos):
    ventana = tk.Tk()
    ventana.title("Mostrar Datos de Excel en Imagen")

    # Marco para la imagen
    marco_imagen = ttk.Frame(ventana, padding=10)
    marco_imagen.pack()

    # Mostrar la imagen
    imagen_tk = ImageTk.PhotoImage(imagen)
    etiqueta_imagen = tk.Label(marco_imagen, image=imagen_tk)
    etiqueta_imagen.image = imagen_tk  # Mantener referencia
    etiqueta_imagen.pack()

    # Marco para el textbox
    marco_textbox = ttk.Frame(ventana, padding=10)
    marco_textbox.pack()

    # Textbox
    textbox = tk.Text(marco_textbox, height=10, width=50)
    textbox.pack()

    # Insertar datos en el textbox
    for fila in datos:
        textbox.insert(tk.END, ", ".join(map(str, fila)) + "\n")

    ventana.mainloop()