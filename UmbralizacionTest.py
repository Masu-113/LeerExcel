import numpy as np  
import pandas as pd 
import matplotlib.pyplot as plt 
from PIL import Image
import pytesseract 
import os

# Asegúrate de especificar la ruta de la imagen
image_path = r'C:\Users\msuarez\source\repos\LeerExcel\prueba\page_1.jpg'  # Especifica la ruta de la imagen aquí
imagen = Image.open(image_path)

# Cambié 'tipo_de_salida' a 'output_type'
datos = pytesseract.image_to_data(imagen, output_type=pytesseract.Output.DATAFRAME)

def optimizarDf(old_df: pd.DataFrame) -> pd.DataFrame: 
    df = old_df[["left", "top", "width", "text"]].copy()  # Haciendo una copia para evitar SettingWithCopyWarning
    df['left+width'] = df['left'] + df['width']  # Usando .loc para evitar SettingWithCopyWarning
    df = df.sort_values(by=['top'], ascending=True)  # Cambié 'ascendente' a 'ascending'
    df = df.groupby(['top', 'left+width'], sort=False)['text'].sum().unstack('left+width') 
    df = df.reindex(sorted(df.columns), axis=1).dropna(how='all').dropna(axis='columns', how='all') 
    df = df.fillna('') 
    return df 

def mergeDfColumns(old_df: pd.DataFrame, threshold: int = 10, rotations: int = 5) -> pd.DataFrame: 
    df = old_df.copy() 
    for j in range(rotations): 
        new_columns = {} 
        old_columns = df.columns.tolist()  # Convertir a lista para iterar
        i = 0 
        while i < len(old_columns):  # Cambié 'if' por 'while' para iterar sobre todas las columnas
            if i < len(old_columns) - 1: 
                # Si la diferencia entre nombres de columnas consecutivos es menor que el umbral
                if abs(old_columns[i + 1] - old_columns[i]) < threshold:  # Comparar directamente los nombres de las columnas
                    new_col = df[old_columns[i]].astype(str) + df[old_columns[i + 1]].astype(str) 
                    new_columns[old_columns[i]] = new_col  # Cambié 'new_columns[old_columns[i + 1]]' a 'new_columns[old_columns[i]]'
                    i += 1 
                else:  # Si la diferencia entre nombres de columnas consecutivos es mayor o igual que el umbral
                    new_columns[old_columns[i]] = df[old_columns[i]] 
            else:  # Si la columna actual es la última columna
                new_columns[old_columns[i]] = df[old_columns[i]] 
            i += 1  # Mover el incremento aquí para evitar bucles infinitos
        df = pd.DataFrame.from_dict(new_columns).replace('', np.nan).infer_objects().dropna(axis='columns', how='all') 
    return df.replace(np.nan, '') 

def mergeDfRows(old_df: pd.DataFrame, threshold: int = 10) -> pd.DataFrame: 
    new_df = old_df.iloc[:1].copy()  # Haciendo una copia de la primera fila
    for i in range(1, len(old_df)): 
        # Si la diferencia entre valores de índice consecutivos es menor que el umbral 
        if abs(old_df.index[i] - old_df.index[i - 1]) < threshold:  # Cambié 'umbral' a 'threshold'
            new_df.iloc[-1] = new_df.iloc[-1].astype(str) + old_df.iloc[i].astype(str) 
        else:  # Si la diferencia es mayor que el umbral, anexar la fila actual
            new_df = pd.concat([new_df, old_df.iloc[[i]]], ignore_index=True)  # Usando pd.concat en lugar de append
    return new_df.reset_index(drop=True)  # Cambié 'devolver' a 'return'

# Cambié 'data_imp_sort' a 'datos' para asegurar que se utilice el DataFrame correcto
data_imp_sort = optimizarDf(datos) 
df_nueva_col = mergeDfColumns(data_imp_sort) 
fila_combinada_df = mergeDfRows(df_nueva_col)

def clean_df(df): 
    # Eliminar columnas con todas las celdas que tienen el mismo valor y su longitud es 0 o 1
    df = df.loc[:, (df != df.iloc[0]).any()] 
    # Eliminar filas con celdas vacías o celdas con solo el símbolo '|'
    df = df[(df != '|') & (df != '') & (pd.notnull(df))] 
    # Eliminar columnas con solo celdas vacías
    df = df.dropna(axis=1, how='all') 
    return df.fillna('') 

# Cambié 'merged_row_df' a 'fila_combinada_df' para asegurar que se utilice el DataFrame correcto
installed_df = clean_df(fila_combinada_df.copy())
