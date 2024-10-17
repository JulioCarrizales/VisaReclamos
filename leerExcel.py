import openpyxl
from tabulate import tabulate

excel_dataframe=openpyxl.load_workbook("C:/Users/Administrador/Desktop/PROYECTOS_JULIO/Canales_virtuales/data.xlsm")

dataframe=excel_dataframe.active

print(dataframe)