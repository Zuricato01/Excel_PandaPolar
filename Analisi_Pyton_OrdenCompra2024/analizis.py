import pandas as pd
from time import time
def main():
    print('Inicio del programa')
    #==variables
    archivo = 'Base servicios KOF 2024 con telemetria.xlsx'
    columnas = ['Servicio', 'Facturacion', 'Orden Servicio']
    #==codigo 
    #leectura_excel('Base servicios KOF 2024 con telemetria.xlsx')

    file_path = archivo
    chunk_size = 1000  # Número de filas por carga
    start_row = 0

    while True:
        try:
            df = pd.read_excel(file_path, 
                               engine="openpyxl", 
                               skiprows=start_row, 
                               nrows=chunk_size, 
                               sheet_name='Base',
                               usecols=columnas,
                               )
            if df.empty:
                break  # Termina cuando no hay más filas
            
            print(df.head())  # Aquí puedes procesar cada chunk
            start_row += chunk_size
        except Exception as e:
            print(f"Error: {e}")
            break
      
if __name__ == "__main__":
    main()

def leectura_excel(archivo):
    with pd.ExcelFile(archivo) as f_xls:
        df = pd.read_excel(f_xls, 
                        sheet_name='Base', 
                        usecols=['Servicio', 'Facturacion', 'Orden Servicio'], 
                        #iterator=True,
                        engine='openpyxl',
                        )
    
