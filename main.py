from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from webdriver_manager.chrome import ChromeDriverManager
import pandas as pd
from datetime import datetime, timedelta

#primer  cambio
#Dany
# Cargar el archivo Excel
file_path = 'C:/Users/Administrador/Desktop/PROYECTOS_JULIO/VISA/data.xlsm'
df = pd.read_excel(file_path, sheet_name=0)

# Filtrar los datos donde no hay VROL
filtered_df = df[df['VROL'].isna()]

# Configuración del controlador de Chrome con WebDriver Manager
options = webdriver.ChromeOptions()  # Inicializar ChromeOptions
service = Service(ChromeDriverManager().install())  # Usar WebDriver Manager para instalar y configurar ChromeDriver

# Inicializar el controlador de Chrome con las opciones y el servicio
driver = webdriver.Chrome(service=service, options=options)

# Abrir la página web
driver.get("https://www.vrol.visaonline.com/rolol/AdminDispatcher?action=RouteToLandingPage&HTTP_ENTRY_POINT=&V=13")

# Procesar cada fila del DataFrame filtrado
for index, row in filtered_df.iterrows():
    card_number = str(row['NRO DE TARJETA \nCOMPROMETIDA']).strip()
    transaction_date = pd.to_datetime(row['TRANSACTION \nDATE'])
    authorization_code = str(row['AUTHORIZATION \nCODE']).strip()

    # Calcular Start Date restando un mes
    start_date = transaction_date - timedelta(days=30)
    start_date_formatted = start_date.strftime("%m/%d/%y")

    # Formatear la fecha actual para End Date
    end_date_formatted = datetime.now().strftime("%m/%d/%y")

    # Rellenar los campos del formulario
    driver.find_element(By.NAME, "CardAccountNumber").send_keys(card_number)
    driver.find_element(By.NAME, "StartDate").send_keys(start_date_formatted)
    driver.find_element(By.NAME, "EndDate").send_keys(end_date_formatted)
    driver.find_element(By.NAME, "AuthorizationCode").send_keys(authorization_code)

    # Hacer clic en el botón "Submit"
    driver.find_element(By.XPATH, '//button[@type="submit"]').click()

    # Esperar unos segundos para verificar el envío antes de pasar al siguiente
    driver.implicitly_wait(3)

# Mantener el navegador abierto para inspección manual
input("Presiona Enter para cerrar el navegador...")

# Cerrar el navegador cuando se complete el proceso
driver.quit()
