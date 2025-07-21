# interfaz.py dentro de /view
import sys
import os
import tkinter as tk
from tkinter import filedialog, messagebox
sys.path.append(os.path.abspath(os.path.join(os.path.dirname(__file__), "..")))

from Test.Script_PDF_EXCEL import sobrescribir_imagen_con_excel  # Asegúrate de que esta función haga todo (PDF→IMG, modificar, IMG→PDF)

def mostrar_gui():
    def ejecutar_proceso():
        try:
            pdf = entry_pdf.get()
            excel = entry_excel.get()
            hoja = entry_hoja.get()
            rango = entry_rango.get()
            fuente = entry_fuente.get()
            tamaño_fuente = int(entry_tamaño_fuente.get())
            xml = entry_xml.get()

            if not all([pdf, excel, hoja, rango, xml]):
                messagebox.showwarning("Campos incompletos", "Por favor, completa todos los campos obligatorios.")
                return

            anchos_definidos = [125, 126, 126, 124, 125, 173, 125]

            # Este método debe encargarse de TODO:
            # PDF → imagen → modificación → nuevo PDF
            sobrescribir_imagen_con_excel(pdf, excel, hoja, rango, fuente, tamaño_fuente, anchos_definidos, xml)

            messagebox.showinfo("Éxito", "¡Proceso completado con éxito!")

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

    def seleccionar_fuente():
        ruta = filedialog.askopenfilename(filetypes=[("Fuentes", "*.ttf")])
        if ruta:
            entry_fuente.delete(0, tk.END)
            entry_fuente.insert(0, ruta)

    def seleccionar_xml():
        ruta = filedialog.askopenfilename(filetypes=[("Archivos XML", "*.xml")])
        if ruta:
            entry_xml.delete(0, tk.END)
            entry_xml.insert(0, ruta)

    ventana = tk.Tk()
    ventana.title("Generar PDF desde Excel y PDF Base")
    ventana.geometry("600x400")

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
    entry_fuente = crear_entrada("Fuente (opcional):", 4, seleccionar_fuente)
    entry_tamaño_fuente = crear_entrada("Tamaño de fuente:", 5)
    entry_tamaño_fuente.insert(0, "30")
    entry_xml = crear_entrada("XML de columnas:", 6, seleccionar_xml)

    boton_ejecutar = tk.Button(ventana, text="Ejecutar", bg="green", fg="white", command=ejecutar_proceso)
    boton_ejecutar.grid(row=7, column=1, pady=20)

    ventana.mainloop()

if __name__ == "__main__":
    mostrar_gui()
