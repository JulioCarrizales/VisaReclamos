from playwright.sync_api import sync_playwright

def run(playwright):
    # Lanzar el navegador
    browser = playwright.chromium.launch(headless=False)
    page = browser.new_page()

    # Navegar a la página de Visa
    page.goto("https://secure.visaonline.com/")

    # Esperar que la página cargue completamente
    page.wait_for_load_state('networkidle')

    # Completar los campos de login
    page.fill("input[name='username']", "ggutierrezag@bn.com.pe")
    page.fill("input[name='password']", "Estrellaviviana24.")

    # Esperar que el botón de "LOG IN" esté visible y habilitado
    page.wait_for_selector("button[type='submit']", state='visible')

    # Hacer clic en el botón de "LOG IN"
    page.click("button[type='submit']", force=True)  # Forzar clic

    # Esperar para ver los resultados
    page.wait_for_timeout(5000)  # 5 segundos de espera

    # Mantener el navegador abierto hasta que presiones Enter
    input("Presiona Enter para cerrar el navegador...")

with sync_playwright() as playwright:
    run(playwright)
