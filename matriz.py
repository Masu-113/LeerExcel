def imprimir_matriz(matriz):
    for fila in matriz:
        for elemento in fila:
            print(f"{elemento:<30}", end=" ")  # Ajusta el ancho según tus necesidades
        print()

# Ejemplo de uso:
matriz_ejemplo = [
    ["nombre", "id", "cantidad"],
    ["dog", "1", "30"],
    ["murcielago", "44", "100"]
]

imprimir_matriz(matriz_ejemplo)