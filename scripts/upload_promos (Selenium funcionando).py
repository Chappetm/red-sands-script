import os
import time
import re
import pandas as pd
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.keys import Keys
import json


# Cargar credenciales
dotenv_path = os.path.join(os.path.dirname(__file__), ".env")
load_dotenv(dotenv_path)
EMAIL = os.getenv("LIGHTSPEED_EMAIL")
PASSWORD = os.getenv("LIGHTSPEED_PASSWORD")

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
driver.get("https://my.kounta.com/pricelist")

# 1. Hacer clic en "Create price list"
create_button = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'button[data-analytics="btnPriceLists_createList"]')
))
create_button.click()

# 2. Esperar y escribir "prueba" en el input del modal
input_field = wait.until(EC.visibility_of_element_located(
    (By.CSS_SELECTOR, 'input[placeholder="Enter price list name"]')
))
input_field.send_keys("prueba")

# 3. Esperar y hacer clic en el botón "Create"
create_modal_button = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'button[type="submit"].btnPrimary')
))
create_modal_button.click()

# 4. Leer todos los productos del sistema (products.xlsx)
product_lookup = {}
xls = pd.ExcelFile("assets/products.xlsx")
for sheet in xls.sheet_names:
    df = xls.parse(sheet)
    for _, row in df.iterrows():
        raw_codes = str(row["Product Code"])
        product_name = row["Product Name"]
        for code in raw_codes.split("/"):
            clean_code = code.strip().lstrip("0")
            if clean_code:
                product_lookup[clean_code] = product_name

                

# 5. Leer los productos promocionados
promo_df = pd.read_excel("bottlemart_promos/promo_products.xlsx", sheet_name="Promocionados")

matched = []
unmatched = []

for _, row in promo_df.iterrows():
    try:
        code = str(int(float(row["Brewer code"]))).strip()
    except:
        code = str(row["Brewer code"]).strip()

    price = str(row["Retail Price"]).replace("$", "").strip()
    name = product_lookup.get(code)

    if name:
        promo_name = str(row["Promoted product"])
        match = re.search(r"(\d+)pk$", promo_name.lower())

        if match:
            pack_number = match.group(1)
            # Reemplazar solo el número al final del nombre
            updated_name = re.sub(r"(S|C)\d+$", lambda m: m.group(1) + pack_number, name)
        else:
            updated_name = name  # no se modifica si no termina en 'pk'

        matched.append({
            "Product Code": code,
            "Product Name": updated_name,
            "Retail Price": price
        })

    else:
        unmatched.append({"Product Code": code, "Retail Price": price, "Original Name": row["Promoted product"]})

# 🔍 Mostrar los no encontrados
print("\n❌ Productos NO encontrados en products.xlsx:")
for item in unmatched:
    print(f"- Code: {item['Product Code']} | Name in promo: {item['Original Name']} | Price: {item['Retail Price']}")

# 6. Cargar productos a la Price List
for product in matched:
    name = product["Product Name"]
    price = product["Retail Price"]

    try:
        print(f"🖊️ Escribiendo: {name}")
        # Buscar producto
        search_box = wait.until(EC.element_to_be_clickable(
            (By.CSS_SELECTOR, 'input[data-chaminputid="searchInput"]')
        ))
        search_box.click() 
        search_box.clear()
        search_box.send_keys(name)
        search_box.send_keys(Keys.ENTER)
        time.sleep(1)

        # Esperar que aparezca un div con el texto exacto que se buscó
        result = wait.until(EC.presence_of_element_located(
            (By.XPATH, f'//div[contains(@class, "Alignment__AlignmentContainer") and text()="{name}"]')
        ))
        time.sleep(1)

        # Ingresar precio
        price_input = wait.until(EC.presence_of_element_located(
            (By.CSS_SELECTOR, 'input[data-chaminputid="textInput"][inputmode="decimal"]')
        ))
        price_input.clear()
        price_input.send_keys(price)

        print(f"✅ {name} cargado con precio ${price}")
        time.sleep(1)

    except Exception as e:
        print(f"❌ Error con {name}: {e}")
        continue

# Guardar la lista
save_btn = wait.until(EC.element_to_be_clickable(
    (By.CSS_SELECTOR, 'button[data-analytics="btnPriceLists_saveList"]')
))
save_btn.click()

time.sleep(2)  # 👈 Pausa para evitar carga agresiva
print("✅ Promos guardadas")