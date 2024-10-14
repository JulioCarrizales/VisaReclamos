from playwright.sync_api import sync_playwright
import pandas as pd
from datetime import datetime, timedelta

# Cargar los datos desde el archivo Excel
file_path = 'C:/Users/Administrador/Desktop/PROYECTOS_JULIO/VISA/data.xlsm'
df = pd.read_excel(file_path, sheet_name=0)

# Filtrar los datos donde no hay VROL
filtered_df = df[df['VROL'].isna()]

# Iniciar Playwright
def run(playwright):
    # Conectar al navegador ya abierto (asegúrate de tener el navegador conectado a Playwright)
    browser = playwright.chromium.connect_over_cdp("http://localhost:9222")  # Asegúrate de abrir tu navegador con este puerto
    page = browser.pages[0]  # Usar la primera página abierta

    # Procesar cada fila del dataframe filtrado
    for index, row in filtered_df.iterrows():
        card_number = str(row['NRO DE TARJETA \nCOMPROMETIDA']).strip()  # Tarjeta
        transaction_date = pd.to_datetime(row['TRANSACTION \nDATE'])  # Fecha de transacción
        authorization_code = str(row['AUTHORIZATION \nCODE']).strip()  # Código de autorización

        # Restar un mes a la fecha de transacción para el campo Start Date
        start_date = transaction_date - timedelta(days=30)
        start_date_formatted = start_date.strftime("%m/%d/%y")

        # Obtener la fecha actual para el campo End Date
        end_date_formatted = datetime.now().strftime("%m/%d/%y")

        # Ingresar el "Card/Account Number"
        page.fill("input[name='CardAccountNumber']", card_number)

        # Ingresar el "Start Date"
        page.fill("input[name='StartDate']", start_date_formatted)

        # Ingresar el "End Date"
        page.fill("input[name='EndDate']", end_date_formatted)

        # Ingresar el "Authorization Code"
        page.fill("input[name='AuthorizationCode']", authorization_code)

        # Hacer clic en el botón "Submit"
        page.click("button[type='submit']")

        # Esperar unos segundos para verificar el envío
        page.wait_for_timeout(3000)

# Ejecutar Playwright
with sync_playwright() as playwright:
    run(playwright)
