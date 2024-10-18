import pandas as pd
# Cargar el archivo Excel proporcionado
file_path = 'C:/Users/julio/Desktop/VISA/data.xlsm'
data = pd.read_excel(file_path)

# Filtrar las filas donde la columna 'FECHA DEVOLUCIÓN' está vacía
filtered_data = data[data['FECHA DEVOLUCIÓN'].isna()]

# Extraer las columnas 'NRO DE CUENTA \n(DEBITO/CREDITO)' y 'AUTHORIZATION \nCODE'
data_to_use = filtered_data[['NRO DE CUENTA \n(DEBITO/CREDITO)', 'AUTHORIZATION \nCODE']]

# Ver los primeros datos para comprobar la lectura
print("Datos extraídos del Excel:")
print(data_to_use.head())
