import pandas as Pd

def main():
    print("data analysis of excel")

    fileExcel = 'FileExcel.xlsx'
    
    extract_create_excel(fileExcel)
    data_table = read_csv_dt('origin.csv')
    proccess_dt(data_table)

def extract_create_excel(path_file):
    """
    This function takes an Excel file, 
    extracts the columns of interest, 
    and creates a CSV file to speed up 
    analysis operations.
    """
    with Pd.ExcelFile(path_file) as file:
        dt = Pd.read_excel(
            file,
            sheet_name='Base',#pointed sheet
            usecols=['Servicio', 'Facturacion', 'Orden Servicio'],#columns to extract
            dtype={'Servicio': str, 'Facturacion': float, 'Orden Servicio': str},#type of data to extract
            engine='openpyxl'#reading engine
        )
        dt.to_csv('origen.csv', index=False)

def read_csv_dt(path_file: str) -> Pd.DataFrame:
    dt = Pd.read_csv(path_file)
    return dt

def proccess_dt(dt):
    dt['Servicio'] = dt['Servicio'].apply(lambda x: 'GARANTIA' if x.startswith('redactado.') else x)
    grouped = dt.groupby(['Servicio']).agg({
        'Orden Servicio': 'count',
        'Facturacion': 'sum'
    })

    #grouped["Facturacion"] = grouped["Facturacion"].astype(int)
    grouped["Facturacion"] = grouped["Facturacion"].round(2)

    grouped['Media'] = grouped['Facturacion'] / grouped['Orden Servicio'].astype(float)
    grouped['Media'] = grouped['Media'].round(2)
    
    grouped.to_excel('proccessed_data.xlsx', sheet_name='Resumen', engine='openpyxl')
    

    #grouped.reset_index().to_csv('proccess.csv', index=False)
    #grouped_garant = grouped[grouped.index.str.startswith('GTIA.')]
    #print(grouped_garant)
    
    #grouped_gt.reset_index().to_csv('proccess_gt.csv', index=False)




if __name__ =='__main__':
    main()
