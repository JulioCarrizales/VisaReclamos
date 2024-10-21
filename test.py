import pyperclip
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.chrome.service import Service
import pandas as pd
from datetime import datetime, timedelta

# Cargar el archivo Excel proporcionado
file_path = 'C:/Users/Administrador/Desktop/PROYECTOS_JULIO/Canales_virtuales/data.xlsm'
data = pd.read_excel(file_path)

# Filtrar las filas donde la columna 'FECHA DEVOLUCIÓN' está vacía
filtered_data = data[data['FECHA DEVOLUCIÓN'].isna()]

# Extraer las columnas 'NRO DE TARJETA \nCOMPROMETIDA' y 'AUTHORIZATION \nCODE'
data_to_use = filtered_data[['NRO DE TARJETA \nCOMPROMETIDA', 'AUTHORIZATION \nCODE']]

# Ver los primeros datos para comprobar la lectura
print("Datos extraídos del Excel:")
print(data_to_use.head())

class TestTest():
    def setup_method(self):
        # Configurar las opciones de Chrome
        options = webdriver.ChromeOptions()
        options.add_argument("--start-maximized")
        
        # Usar el Service para manejar la instalación de ChromeDriver
        service = Service(ChromeDriverManager().install())
        
        # Inicializar el driver de Chrome con el Service y las opciones
        self.driver = webdriver.Chrome(service=service, options=options)
        self.driver.implicitly_wait(10)  # Espera implícita para que los elementos carguen
        self.vars = {}

    def wait_for_window(self, timeout=10):
        WebDriverWait(self.driver, timeout).until(
            lambda driver: len(self.driver.window_handles) > 1
        )
        wh_now = self.driver.window_handles
        wh_then = self.vars["window_handles"]
        return list(set(wh_now).difference(set(wh_then)))[0]

    def test_fill_form(self, data_to_use):
        try:
            # Abrir la página de login
            self.driver.get("https://www.visaonline.com/login/?realm=vol&goto=https:%2F%2Fsecure.visaonline.com:443%2Fagent%2Fcustom-login-response%3Fstate%3DruSFimU-nVV4ft9utXp5rzLOuQ0%26realm%3Dvol&original_request_url=https:%2F%2Fsecure.visaonline.com:443%2F")
            
            # Ajustar el tamaño de la ventana
            self.driver.set_window_size(1296, 688)

            # Esperar a que el campo de usuario esté presente y luego hacer clic y escribir en él
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "username"))
            )
            username_field = self.driver.find_element(By.ID, "username")
            username_field.click()
            username_field.send_keys("ggutierrezag@bn.com.pe")

            # Esperar a que el campo de contraseña esté presente y luego escribir en él
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "password"))
            )
            password_field = self.driver.find_element(By.ID, "password")
            password_field.send_keys("Estrellaviviana24.")
            
            # Enviar la contraseña con la tecla Enter
            password_field.send_keys(Keys.ENTER)

            # Guardar el handle de la ventana principal
            self.vars["root"] = self.driver.current_window_handle

            # Esperar a que el elemento con id 'ROL' esté presente
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "ROL"))
            )

            # Moverse al elemento 'ROL' y hacer clic en él
            element = self.driver.find_element(By.ID, "ROL")
            actions = ActionChains(self.driver)
            actions.move_to_element(element).perform()

            # Almacenar el handle de las ventanas actuales
            self.vars["window_handles"] = self.driver.window_handles

            # Hacer clic en el elemento 'ROL' y esperar a que se abra una nueva ventana
            element.click()

            # Esperar que la nueva ventana esté disponible (hasta 10 segundos)
            self.vars["win4167"] = self.wait_for_window(10)

            # Cambiar a la nueva ventana
            self.driver.switch_to.window(self.vars["win4167"])

            # Aquí es donde hacemos hover sobre el menú "Inquiry"
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Inquiry"))
            )
            inquiry_element = self.driver.find_element(By.LINK_TEXT, "Inquiry")
            
            # Mover el mouse sobre el elemento "Inquiry" para desplegar el menú
            actions.move_to_element(inquiry_element).perform()

            # Esperar a que aparezca el menú desplegable con "Transaction Inquiry"
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.LINK_TEXT, "Transaction Inquiry"))
            )
            
            # Hacer clic en "Transaction Inquiry" una vez que aparece en el menú desplegable
            self.driver.find_element(By.LINK_TEXT, "Transaction Inquiry").click()

            # COPIAR Y PEGAR DATOS DESDE EXCEL
            nro_de_cuenta = data_to_use.iloc[0]['NRO DE TARJETA \nCOMPROMETIDA']
            authorization_code = data_to_use.iloc[0]['AUTHORIZATION \nCODE']

            # Copiar el dato de la tarjeta al portapapeles
            pyperclip.copy(nro_de_cuenta)

            # Esperar a que el campo "Card/Account Number" esté visible y seleccionarlo
            WebDriverWait(self.driver, 20).until(
                EC.presence_of_element_located((By.ID, "CardNumber"))
            )
            card_number_field = self.driver.find_element(By.ID, "CardNumber")
            card_number_field.click()

            # Pegar con Ctrl+V
            card_number_field.send_keys(Keys.CONTROL, 'v')

            # Copiar el Authorization Code al portapapeles
            pyperclip.copy(authorization_code)

            # Hacer scroll para que el campo "Authorization Code" esté visible
            auth_code_field = self.driver.find_element(By.ID, "AuthCode")
            self.driver.execute_script("arguments[0].scrollIntoView();", auth_code_field)
            
            # Pegar el Authorization Code
            auth_code_field.click()
            auth_code_field.send_keys(Keys.CONTROL, 'v')  # Pegar con Ctrl + V

            # Cambiar la fecha del campo "Start Date" a tres meses antes de la fecha actual
            start_date_field = self.driver.find_element(By.ID, "StartDate")
            three_months_ago = (datetime.now() - timedelta(days=90)).strftime('%m/%d/%Y')
            start_date_field.clear()
            start_date_field.send_keys(three_months_ago)

            # Hacer clic en el botón "Submit"
            submit_button = self.driver.find_element(By.NAME, "Submit")
            submit_button.click()

        except Exception as e:
            print(f"Error encontrado: {e}")

# Ejecutar el script
test = TestTest()
test.setup_method()
test.test_fill_form(data_to_use)