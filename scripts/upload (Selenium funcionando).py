import os
import time
import sys
import shutil
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
import undetected_chromedriver as uc
from selenium.webdriver import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import json

# SIEMPRE ACTUALIZAR COOKIES DEL ARCHIVO JSON

# Cargar credenciales
dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path)
EMAIL = os.getenv("LIGHTSPEED_EMAIL")
PASSWORD = os.getenv("LIGHTSPEED_PASSWORD")

if len(sys.argv) != 2:
    print("Uso: python upload.py [alm|coke|cub|lion]")
    sys.exit(1)

supplier = sys.argv[1].lower()
supplier_lookup = {
    "alm": "Australian Liqueur Marketers",
    "cub": "Carlton & United Breweries",
    "coke": "Coca-Cola Australia",
    "lion": "Lion - Beer, Spirits & Wine Pty Ltd"
}

input_folder = os.path.join(os.path.dirname(__file__), "../Excel_invoices", supplier)

if not os.path.exists(input_folder):
    print(f"⛔ La carpeta '{input_folder}' no existe.")
    sys.exit(1)

excel_files = [f for f in os.listdir(input_folder) if f.endswith(".xlsx") and not f.startswith("~$")]

if not excel_files:
    print(f"⚠️ No hay archivos .xlsx en '{input_folder}'.")
    sys.exit(1)

# Cargar diccionario desde products.xlsx
lookup_df = pd.read_excel("products.xlsx", sheet_name=supplier.upper())
# product_lookup = dict(zip(lookup_df["Product Code"], lookup_df["Product Name"]))

product_lookup = {}

for _, row in lookup_df.iterrows():
    raw_codes = str(row["Product Code"])
    product_name = row["Product Name"]
    for code in raw_codes.split("/"):
        clean_code = code.strip().lstrip("0")  # elimina espacios y ceros a la izquierda
        if clean_code:
            product_lookup[clean_code] = product_name



# CHEQUEA QUE ESTE LA CARPETA "PROCESADOS" SINO LA CREA
processed_folder = os.path.join(input_folder, "procesados")
os.makedirs(processed_folder, exist_ok=True)

# Chrome setup
chrome_options = Options()
chrome_options.add_argument("--start-maximized")
chrome_options.add_argument("--disable-blink-features=AutomationControlled")
chrome_options.add_argument("--disable-infobars")
chrome_options.add_argument("--disable-extensions")
chrome_options.add_experimental_option("excludeSwitches", ["enable-automation"])
chrome_options.add_experimental_option('useAutomationExtension', False)

service = Service()
driver = webdriver.Chrome(service=service, options=chrome_options)
wait = WebDriverWait(driver, 20)


# Paso 1: Login antidetención
driver.execute_cdp_cmd("Page.addScriptToEvaluateOnNewDocument", {
    "source": """
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined
        });
    """
})

driver.delete_all_cookies()
driver.get("https://purchase.kounta.com/")
time.sleep(3)  # 👈 Pausa para evitar carga agresiva

# Cargar cookies desde archivo exportado
with open("lightspeed_cookies.json", "r") as f:
    cookies = json.load(f)

for cookie in cookies:
    for k in ["sameSite", "storeId", "hostOnly", "session", "id"]:
        cookie.pop(k, None)
    if "expirationDate" in cookie:
        cookie["expiry"] = int(cookie.pop("expirationDate"))

    try:
        driver.add_cookie(cookie)
    except Exception as e:
        print(f"⚠️ Cookie inválida ignorada: {cookie.get('name')}")

# Refrescar con sesión activa
driver.get("https://purchase.kounta.com/")
wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Purchase orders")))
print("✅ Sesión restaurada con cookies.")

# Paso 2: Ir a sección "Purchase Orders"
wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Purchase orders"))).click()

# Procesar todos los archivos Excel de la carpeta
for excel_file in excel_files:
    print(f"\n📄 Procesando: {excel_file}")
    excel_path = os.path.join(input_folder, excel_file)
    df = pd.read_excel(os.path.join(input_folder, excel_file), engine="openpyxl")

    # ADMINN FEE
    admin_fee_value = None
    if supplier == "alm" and "Admin fee" in df.columns:
        try:
            value = float(df.at[0, "Admin fee"])
            if value > 0:
                admin_fee_value = round(value, 2)
        except Exception as e:
            print(f"⚠️ No se pudo leer Admin Fee: {e}")

    # Validar que todos los códigos del Excel existan en el diccionario
    df["Product Code"] = df["Product Code"].astype(str).str.strip().str.lstrip("0")
    missing = set(df["Product Code"]) - set(product_lookup.keys())
    if missing:
        print("⛔ Faltan los siguientes códigos de producto en products.xlsx:", missing)
        continue

    # Volver a "Purchase Orders"
    wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Purchase orders"))).click()
    time.sleep(0.5)

    po_number_raw = str(df.iloc[0].get("PO Number", "")).strip()
    is_new_invoice = not po_number_raw or po_number_raw == "PO00000000"

    if is_new_invoice:
        print("🆕 No se detectó PO Number. Creando nueva orden.")
        # Click en "New order"
        new_order_button = wait.until(EC.element_to_be_clickable((
            By.XPATH, "//button[contains(text(), 'New order')]"
        )))
        driver.execute_script("arguments[0].click();", new_order_button)
        time.sleep(1)

        supplier_name = supplier_lookup[supplier]
        supplier_xpath = f"//tr[.//div[text()='{supplier_name}']]"
        supplier_row = wait.until(EC.element_to_be_clickable((By.XPATH, supplier_xpath)))
        driver.execute_script("arguments[0].click();", supplier_row)
        print(f"📦 Seleccionado proveedor: {supplier_name}")
        time.sleep(1)

        # Esperar a que aparezca el nuevo PO Number y mostrarlo
        try:
            po_input = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input.readonly[data-chaminputid='textInput']"
            )))
            new_po_number = po_input.get_attribute("value")
            print(f"🆕 Nueva invoice creada con PO: PO{new_po_number}")
        except Exception as e:
            print("⚠️ No se pudo obtener el nuevo PO Number:", e)

        

    else:
        clean_po_number = po_number_raw.replace("PO", "")
        po_number = str(int(clean_po_number))
        wait.until(EC.presence_of_element_located((By.LINK_TEXT, "Purchase orders"))).click()
        time.sleep(0.5)
        order_element = wait.until(
            EC.presence_of_element_located((By.XPATH, f"//*[contains(text(), '{po_number}')]"))
        )
        order_element.click()

        # Eliminar producto pre-cargado
        try:
            existing_product = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//div[contains(@class, 'lineInfo')]"
            )))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", existing_product)
            time.sleep(0.5)
            driver.execute_script("arguments[0].click();", existing_product)
            time.sleep(1)

            remove_button = wait.until(EC.element_to_be_clickable((
                By.XPATH, "//span[text()='Remove']/ancestor::button"
            )))
            driver.execute_script("arguments[0].click();", remove_button)
            print("🗑️ Producto precargado eliminado.")

        except Exception as e:
            print(f"ℹ️ No se eliminó ningún producto precargado: {e}")

    # Clic en lupa
    search_button = wait.until(EC.element_to_be_clickable((
        By.CSS_SELECTOR, "button.IconButton__IconButtonStyle-sc-17z8q7c-0"
    )))
    driver.execute_script("arguments[0].click();", search_button)

    # Agregar productos
    for _, row in df.iterrows():
        product_code = row["Product Code"]
        quantity = row["Order Qty"]
        total_cost = row["Total Cost"]
        product_name = product_lookup.get(product_code)

        if not product_name:
            print(f"⚠️ Producto no encontrado para código: {product_code}")
            continue

        print(f"🛒 Agregando: {product_name}")
        try:
            search_input = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[placeholder='Search for products']"
            )))
            search_input.clear()
            search_input.send_keys(product_name)
            search_input.send_keys(Keys.RETURN)

            product_elements = driver.find_elements(By.XPATH, "//span[@class='TextClamp__InnerSpan-phjbkm-1 bFzZsL']")
            time.sleep(0.5)

            found_product = False
            for product_element in product_elements:
                if product_element.text.strip() == product_name.strip():
                    for _ in range(int(quantity)):
                        product_element.click()
                    found_product = True
                    print(f"✅ Producto agregado {int(quantity)}x: {product_name}")

                    product_line_xpath = f"//div[contains(@class, 'lineInfo')]//span[normalize-space(text()) = \"{product_name}\"]"

                    try:
                        product_line = wait.until(EC.element_to_be_clickable((By.XPATH, product_line_xpath)))
                        driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_line)
                        time.sleep(0.5)
                        driver.execute_script("arguments[0].click();", product_line)

                        total_input = wait.until(EC.presence_of_element_located((
                            By.CSS_SELECTOR, "input[placeholder='Enter total price']"
                        )))
                        ActionChains(driver).move_to_element(total_input).click().perform()

                        # Leer precio actual desde el input
                        try:
                            current_cost = float(total_input.get_attribute("value"))
                        except:
                            current_cost = None

                        # Comparar precios y alertar si la diferencia es significativa
                        if current_cost is not None and (total_cost - current_cost) > 1.00:
                            print(f"🚨 Atención: '{product_name}' tiene un costo en invoice de ${total_cost}, mayor que el actual (${current_cost})")

                        # Escribir el nuevo costo
                        actions = ActionChains(driver)
                        # actions.send_keys(str(total_cost)).perform()

                        # Ajustar total_cost con 10% adicional si es CUB, LION o COKE
                        adjusted_cost = total_cost
                        if supplier in ["cub", "lion", "coke"]:
                            adjusted_cost = round(total_cost * 1.10, 2)

                        actions.send_keys(str(adjusted_cost)).perform()



                        ok_button = wait.until(EC.element_to_be_clickable((By.ID, "enter")))
                        driver.execute_script("arguments[0].click();", ok_button)

                        apply_button = wait.until(EC.element_to_be_clickable((
                            By.XPATH, "//button[contains(text(), 'Apply Changes')]"
                        )))
                        driver.execute_script("arguments[0].click();", apply_button)

                        time.sleep(0.5)
                    except Exception as e:
                        print(f"⚠️ No se pudo editar el Total Cost de '{product_name}': {e}")

                    break

            if not found_product:
                print(f"❌ Producto no encontrado exacto: {product_name}")

        except Exception as e:
            print(f"❌ Error con '{product_name}': {e}")

    # Agregar Admin Fee si corresponde
    if supplier == "alm" and admin_fee_value:
        product_name = "Administration Fee"
        quantity = 1
        total_cost = admin_fee_value

        print(f"➕ Agregando Admin Fee: {product_name} ${total_cost}")
        try:
            search_input = wait.until(EC.presence_of_element_located((
                By.CSS_SELECTOR, "input[placeholder='Search for products']"
            )))
            search_input.clear()
            search_input.send_keys(product_name)
            search_input.send_keys(Keys.RETURN)

            product_elements = driver.find_elements(By.XPATH, "//span[@class='TextClamp__InnerSpan-phjbkm-1 bFzZsL']")
            time.sleep(0.5)

            for product_element in product_elements:
                if product_element.text.strip() == product_name.strip():
                    product_element.click()
                    print(f"✅ Admin Fee agregado.")

                    product_line_xpath = f"//div[contains(@class, 'lineInfo')]//span[normalize-space(text()) = \"{product_name}\"]"
                    product_line = wait.until(EC.element_to_be_clickable((By.XPATH, product_line_xpath)))
                    driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", product_line)
                    time.sleep(0.5)
                    driver.execute_script("arguments[0].click();", product_line)

                    total_input = wait.until(EC.presence_of_element_located((
                        By.CSS_SELECTOR, "input[placeholder='Enter total price']"
                    )))
                    ActionChains(driver).move_to_element(total_input).click().perform()

                    actions = ActionChains(driver)
                    actions.send_keys(str(total_cost)).perform()

                    ok_button = wait.until(EC.element_to_be_clickable((By.ID, "enter")))
                    driver.execute_script("arguments[0].click();", ok_button)

                    apply_button = wait.until(EC.element_to_be_clickable((
                        By.XPATH, "//button[contains(text(), 'Apply Changes')]"
                    )))
                    driver.execute_script("arguments[0].click();", apply_button)

                    break
        except Exception as e:
            print(f"❌ Error al agregar Admin Fee: {e}")

    # TOCAR EL BOTON REVIEW ORDER PARA GUARDAR LA INVOICE
    review_order_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[contains(text(), 'Review Order')]")))
    driver.execute_script("arguments[0].click();", review_order_button)
    
    # VOLVER A "https://purchase.kounta.com/purchase#orders"
    driver.get("https://purchase.kounta.com/purchase#orders")

    # BUSCAR PO Number DEVUELTA
    print("✅ Factura procesada con éxito.")

    # Mover a carpeta "Procesados"
    shutil.move(excel_path, os.path.join(processed_folder, excel_file))
    print("📁 Archivo movido a carpeta 'procesados'.")

    time.sleep(5)

print("✔️ Todas las facturas del proveedor fueron procesadas.")
driver.quit()
