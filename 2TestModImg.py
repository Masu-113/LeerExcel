from PIL import Image

def modificar_imagen(ruta_imagen, nuevo_ancho, nuevo_alto):
    ruta_imagen = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1_modificada.jpg'
    try:
        img = Image.open(ruta_imagen)
        ancho, alto = img.size

        # Ejemplo: Redimensionar la imagen
        img_redimensionada = img.resize((nuevo_ancho, nuevo_alto))

        # Ejemplo: Cambiar el color de ciertos píxeles (fila 10, columna 20)
        if ancho > 100 and alto > 100:
            pix = img_redimensionada.load()
            pix[10, 20] = (255, 0, 0)  # Rojo

        # Guarda la imagen modificada
        nombre_base, extension = ruta_imagen.split('.')
        nueva_ruta = f"{nombre_base}_modificada.{extension}"
        img_redimensionada.save(nueva_ruta)
        print(f"Imagen modificada guardada en: {nueva_ruta}")

    except FileNotFoundError:
        print(f"Error: Archivo no encontrado en {ruta_imagen}")
    except Exception as e:
        print(f"Ocurrió un error: {e}")


# Ejemplo de uso
modificar_imagen("imagen.jpg", 2700, 4200)