import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from Scripts.PDF_A_IMG import pdf_to_images
from Scripts.IMG_A_PDF import convert_images_to_pdf
from Test.Script_PDF_EXCEL import sobrescribir_imagen_con_excel

def mostrar_gui():
    def ejecutar_proceso():
        try:
            pdf = entry_pdf.get()
            excel = entry_excel.get()
            hoja = entry_hoja.get()
            rango = entry_rango.get()

            if not all([pdf, excel, hoja, rango]):
                messagebox.showwarning("Campos incompletos", "Por favor, completa todos los campos obligatorios.")
                return

            # Rutas por defecto
            fuente = r"C:\Windows\Fonts\Arial.ttf"
            tamaño_fuente = 30
            xml = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "Scripts", "column_bounding_boxes.xml"))
            anchos_definidos = [125, 126, 126, 124, 125, 173, 125]

            output_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "output"))
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)

            image_path = os.path.join(output_dir, "page_1.jpg")
            image_modificada = os.path.join(output_dir, "page_1_modificada.jpg")
            final_pdf_path = os.path.join(output_dir, "final.pdf")

            # Paso 1: PDF → Imagen
            pdf_to_images(pdf, output_dir)

            # Paso 2: Sobrescribir con datos del Excel
            sobrescribir_imagen_con_excel(image_path, excel, hoja, rango, fuente, tamaño_fuente, anchos_definidos, xml)

            # Paso 3: Convertir a PDF
            convert_images_to_pdf([image_modificada], final_pdf_path)

            messagebox.showinfo("Éxito", f"Proceso completado.\nPDF generado: {final_pdf_path}")

        except Exception as e:
            messagebox.showerror("Error", f"Ocurrió un error: {e}")

    def seleccionar_pdf():
        ruta = filedialog.askopenfilename(filetypes=[("Archivos PDF", "*.pdf")])
        if ruta:
            entry_pdf.delete(0, tk.END)
            entry_pdf.insert(0, ruta)

    def seleccionar_excel():
        ruta = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
        if ruta:
            entry_excel.delete(0, tk.END)
            entry_excel.insert(0, ruta)

    ventana = tk.Tk()
    ventana.title("Generar PDF desde Excel y PDF Base")
    ventana.geometry("600x300")

    def crear_entrada(label_text, row, browse_func=None):
        label = tk.Label(ventana, text=label_text)
        label.grid(row=row, column=0, sticky='e', padx=5, pady=5)
        entry = tk.Entry(ventana, width=50)
        entry.grid(row=row, column=1, padx=5, pady=5)
        if browse_func:
            boton = tk.Button(ventana, text="Buscar", command=browse_func)
            boton.grid(row=row, column=2, padx=5, pady=5)
        return entry

    entry_pdf = crear_entrada("Archivo PDF base:", 0, seleccionar_pdf)
    entry_excel = crear_entrada("Archivo Excel:", 1, seleccionar_excel)
    entry_hoja = crear_entrada("Nombre de la hoja:", 2)
    entry_hoja.insert(0, "Hoja1")
    entry_rango = crear_entrada("Rango de celdas:", 3)
    entry_rango.insert(0, "A1:G10")

    boton_ejecutar = tk.Button(ventana, text="Ejecutar", bg="green", fg="white", command=ejecutar_proceso)
    boton_ejecutar.grid(row=4, column=1, pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    mostrar_gui()
