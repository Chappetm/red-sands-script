# Este script va a leer el archivo de excel order_ready.xlsx y en base a que proveedor se le pase al ejecutarlo va a leer la pestania y llenar el carrito en la pagina del proveedor

import sys
import re
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import Select
from selenium.common.exceptions import StaleElementReferenceException
from selenium.common.exceptions import ElementNotInteractableException
from selenium.common.exceptions import ElementClickInterceptedException
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import os
import time



####### COKE HELPERS #########
def _extract_int(s: str):
    if not s:
        return None
    m = re.search(r"\d+", s)
    return int(m.group()) if m else None

def find_cart_count_el(driver):
    """
    Usa la estructura que me pasaste: button.toggle-minicart > .badge
    """
    try:
        el = driver.find_element(By.CSS_SELECTOR, "button.toggle-minicart .badge")
        if el.is_displayed():
            return el
    except Exception:
        pass
    # Fallbacks por si el sitio cambia
    for sel in [
        ".toggle-minicart .badge",
        "button.toggle-minicart div.badge",
        "button.toggle-minicart span.badge",
    ]:
        try:
            el = driver.find_element(By.CSS_SELECTOR, sel)
            if el.is_displayed():
                return el
        except Exception:
            continue
    return None

def get_cart_count(driver):
    """
    Lee el número del badge del carrito en el header.
    """
    el = find_cart_count_el(driver)
    if not el:
        return None
    try:
        txt = (el.text or el.get_attribute("innerText") or "").strip()
        return _extract_int(txt)
    except Exception:
        return None

def js_click(driver, el):
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", el)
    driver.execute_script("arguments[0].click();", el)

def first_visible(elements):
    for el in elements:
        try:
            if el.is_displayed() and el.is_enabled():
                return el
        except StaleElementReferenceException:
            continue
    return None

def find_coke_tile_for_code(driver, product_code, wait_seconds=12):
    """
    Devuelve el tile/card que contiene el código como texto visible.
    1) Intenta por CSS (varias clases comunes)
    2) Fallback por XPath: cualquier contenedor con 'product'/'card' que contenga el código
    """
    code = str(product_code).strip()

    # 1) Intento por CSS
    try:
        WebDriverWait(driver, wait_seconds).until(
            EC.presence_of_element_located((By.CSS_SELECTOR,
                "[data-testid='product-card'], .product-tile, .productListItem, .product, .productCard, .search-result-item"
            ))
        )
        tiles = driver.find_elements(By.CSS_SELECTOR,
            "[data-testid='product-card'], .product-tile, .productListItem, .product, .productCard, .search-result-item"
        )
        for tile in tiles:
            try:
                if not tile.is_displayed():
                    continue
                text = (tile.text or "").strip()
                if code in text:
                    return tile
            except StaleElementReferenceException:
                continue
    except TimeoutException:
        pass

    # 2) Fallback por XPath (busca cualquier contenedor relevante que contenga el código)
    try:
        xpath = (
            f"//div[contains(@class,'product') or contains(@class,'card') or contains(@class,'tile') or @data-testid='product-card']"
            f"[.//text()[contains(normalize-space(.), '{code}')]]"
        )
        tile = WebDriverWait(driver, wait_seconds).until(
            EC.presence_of_element_located((By.XPATH, xpath))
        )
        return tile
    except TimeoutException:
        return None

def wait_quantity_ui_in_tile(driver, tile, timeout=10):
    """
    Espera a que el MISMO tile muestre la UI de cantidad (estado 2):
    .number-input-container o input.updateCart
    """
    def _has_qty_ui(_):
        try:
            return (
                len(tile.find_elements(By.CSS_SELECTOR, "div.number-input-container")) > 0 or
                len(tile.find_elements(By.CSS_SELECTOR, "input.updateCart")) > 0
            )
        except StaleElementReferenceException:
            return False
    WebDriverWait(driver, timeout).until(_has_qty_ui)

def set_quantity_in_tile(driver, tile, qty_value):
    """
    Estrategia FIABLE para MyCCA (según tu DOM):
      - NO escribir en input.updateCart (no actualiza backend).
      - Usar SIEMPRE el botón '+' dentro de number-input-container.
      - Verificar incrementos contra el badge del carrito en el header (button.toggle-minicart .badge).
    Devuelve (ok, err|None).
    """
    if qty_value <= 1:
        return True, None

    # Ubicar el botón '+' DENTRO del number-input-container del tile actual
    plus_buttons = tile.find_elements(By.CSS_SELECTOR, "div.number-input-container button.cca-button.secondary.addToCart")
    plus_btn = first_visible(plus_buttons)
    if not plus_btn:
        return False, "No hay botón '+' visible en el tile."

    # Leemos el contador actual del carrito (puede ser None si aún no existe)
    start_count = get_cart_count(driver)
    increments_needed = qty_value - 1  # ya hay 1 por el primer "Add to cart"

    # Si no hay contador visible, igual procedemos con clicks + pequeñas esperas
    if start_count is None:
        for _ in range(increments_needed):
            js_click(driver, plus_btn)
            time.sleep(0.25)  # colchón para backend
        return True, None

    # Con contador visible: esperar que cada click sume +1
    current_target = start_count
    for i in range(increments_needed):
        target = current_target + 1
        js_click(driver, plus_btn)

        # Esperar que el badge aumente a 'target'
        try:
            WebDriverWait(driver, 8).until(lambda d: (get_cart_count(d) or current_target) >= target)
            # Actualizamos base con lo que realmente vea ahora el badge
            current_read = get_cart_count(driver)
            current_target = current_read if current_read is not None else target
        except TimeoutException:
            # Reintento único por si el primer click no pegó
            js_click(driver, plus_btn)
            try:
                WebDriverWait(driver, 6).until(lambda d: (get_cart_count(d) or current_target) >= target)
                current_read = get_cart_count(driver)
                current_target = current_read if current_read is not None else target
            except TimeoutException:
                return False, f"El contador de carrito no incrementó (esperado {target}) en el intento {i+1}/{increments_needed}"

    return True, None

def close_common_overlays(driver):
    for sel in [
        "button[aria-label='Close']",
        "button.close",
        ".modal button.close",
        ".banner button[aria-label='Close']",
        ".toast button[aria-label='Close']",
    ]:
        try:
            for b in driver.find_elements(By.CSS_SELECTOR, sel):
                if b.is_displayed():
                    js_click(driver, b)
                    time.sleep(0.2)
        except Exception:
            pass

def get_qty_from_tile(tile):
    """
    Devuelve la cantidad actual del producto en ese tile.
    Intenta:
      - input.updateCart.value (puede traer '1 in cart')
      - atributos aria-valuenow / data-value / aria-label / placeholder
      - texto del contenedor number-input-container
    """
    # 1) Intentar por input.updateCart
    try:
        inputs = tile.find_elements(By.CSS_SELECTOR, "input.updateCart")
        qty_input = first_visible(inputs)
        if qty_input:
            # value puede ser "1 in cart" o vacío; probar varios atributos
            candidates = [
                qty_input.get_attribute("value") or "",
                qty_input.get_attribute("aria-valuenow") or "",
                qty_input.get_attribute("data-value") or "",
                qty_input.get_attribute("aria-label") or "",
                qty_input.get_attribute("placeholder") or "",
            ]
            for raw in candidates:
                m = re.search(r"\d+", raw)
                if m:
                    return int(m.group())
    except Exception:
        pass

    # 2) Intentar por texto del contenedor de cantidad
    try:
        conts = tile.find_elements(By.CSS_SELECTOR, "div.number-input-container")
        cont = first_visible(conts)
        if cont:
            raw = cont.text or ""
            m = re.search(r"\d+", raw)
            if m:
                return int(m.group())
    except Exception:
        pass

    # 3) Por defecto, si no pudimos leer, asumimos 1
    return 1

##################

load_dotenv()

# Verificar argumento del proveedor
if len(sys.argv) < 2:
    print("❌ Debes indicar el proveedor. Ejemplo: python order.py cub")
    sys.exit()

supplier = sys.argv[1].upper()

# Leer la hoja correspondiente del archivo
order_file = "order_ready.xlsx"

if not os.path.exists(order_file):
    print("❌ No se encontró el archivo 'order_ready.xlsx'. Asegurate de haber generado el reporte antes de ejecutar este script.")
    sys.exit(1)

try:
    df = pd.read_excel(order_file, sheet_name=supplier)
except ValueError:
    print(f"❌ No se encontró la hoja '{supplier}' dentro de '{order_file}'. Verifica el nombre exacto de la pestaña.")
    sys.exit(1)

if supplier == "LION" or supplier == "LION KEGS":
    email = os.getenv("LION_EMAIL")
    password = os.getenv("LION_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de LION en el archivo .env.")
        sys.exit()

    # Iniciar Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # Ir a la página
    driver.get("https://my.lionco.com/login")

    # Ingresar email y password
    wait.until(EC.presence_of_element_located((By.ID, "username")))
    driver.find_element(By.ID, "username").send_keys(email)
    driver.find_element(By.ID, "password").send_keys(password)

    # Hacer clic en el botón de login
    login_button = driver.find_element(By.XPATH, "//button[@type='submit' and contains(text(), 'Login')]")
    login_button.click()
    time.sleep(5)

    for index, row in df.iterrows():
        product_code = str(row["Product Code"]).strip()
        qty_raw = str(row["Quantity"]).strip().upper()

        print(f"🔍 Buscando producto: {product_code} | Cantidad: {qty_raw}")

        # Identificar si es Carton, Layer o Pallet
        if qty_raw.endswith("L"):
            qty_type = "Layer"
            qty_value = int(qty_raw[:-1])
        elif qty_raw.endswith("P"):
            qty_type = "Pallet"
            qty_value = int(qty_raw[:-1])
        else:
            qty_type = "Case"
            qty_value = int(qty_raw)


        # Buscar producto
        search_input = wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input[placeholder='Search products']")))
        time.sleep(0.5)
        search_input.clear()
        search_input.send_keys(product_code)
        search_input.send_keys(Keys.ENTER)

        # Esperar a que aparezca la tarjeta del producto con el select visible
        product_cards = wait.until(EC.presence_of_all_elements_located((By.CSS_SELECTOR, "div.chakra-stack.css-1pv3wlg")))
        product_card = product_cards[0]  # Usamos la primera tarjeta visible
        time.sleep(1)

        # Buscar select dentro de esa tarjeta
        try:
            unit_selector = product_card.find_element(By.CSS_SELECTOR, "select.chakra-select.css-1j263bk")
            select_element = Select(unit_selector)

            qty_type_normalized = qty_type.strip().lower()
            unit_text = None
            for option in select_element.options:
                if qty_type_normalized in option.text.strip().lower():
                    unit_text = option.text
                    break

            if unit_text:
                select_element.select_by_visible_text(unit_text)
                print(f"✅ Unidad seleccionada: {unit_text}")

        except Exception as e:
            print("❌ Error al seleccionar la unidad:", e)

        # Esperar campo de cantidad
        qty_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='number']")))
        qty_input.clear()
        qty_input.send_keys(str(qty_value))

        # Clic en "Add to cart"
        add_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[normalize-space()='Add to cart']")))
        add_button.click()


    time.sleep(5)

    print("✅ Login realizado en LION.")

if supplier == "CUB":
    email = os.getenv("CUB_EMAIL")
    password = os.getenv("CUB_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de CUB en el archivo .env.")
        sys.exit()

    # Listas para resumen
    cub_oos = []        # out of stock
    cub_not_found = []  # no se encontró card del producto
    cub_qty_failed = [] # falló setear cantidad

    # Iniciar Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # Ir a la página
    driver.get("https://online.cub.com.au/sabmStore/en/login")

    # Ingresar email y password
    wait.until(EC.presence_of_element_located((By.ID, "j_username")))
    driver.find_element(By.ID, "j_username").send_keys(email)
    driver.find_element(By.ID, "j_password").send_keys(password)
    time.sleep(1)

    # Hacer clic en el botón de login
    login_button = wait.until(EC.element_to_be_clickable((By.XPATH, "//button[@type='button' and text()='Login']")))
    login_button.click()

    for index, row in df.iterrows():
        product_code = str(row["Product Code"]).strip()
        qty_raw = str(row["Quantity"]).strip().upper()

        print(f"🔍 Searching product: {product_code} | Quantity: {qty_raw}")

        # Identificar si es Carton, Layer o Pallet
        if qty_raw.endswith("L"):
            qty_type = "Layer"
            qty_value = int(qty_raw[:-1])
        elif qty_raw.endswith("P"):
            qty_type = "Pallet"
            qty_value = int(qty_raw[:-1])
        else:
            qty_type = "Case"
            qty_value = int(qty_raw)

        # Buscar producto
        try:
            wait.until(EC.presence_of_element_located((By.ID, "input_SearchBox")))
            search_input = driver.find_element(By.ID, "input_SearchBox")
            search_input.click()
            search_input.clear()
            search_input.send_keys(product_code)
            search_input.send_keys(Keys.RETURN)

            # Esperar a que aparezca la tarjeta del producto
            product_card = wait.until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.col-sm-4.list-item.addtocart-qty"))
            )
        except TimeoutException:
            print(f"🔴 No product card found for: {product_code}")
            cub_not_found.append(product_code)
            continue

        # Detectar estado del botón "Add to order" / "OUT OF STOCK" ANTES de cargar cantidad
        try:
            add_button = product_card.find_element(By.CSS_SELECTOR, "button.addToCartButton")
            btn_text = (add_button.text or "").strip().upper()
            is_disabled = (add_button.get_attribute("disabled") is not None) or \
                          (str(add_button.get_attribute("aria-disabled")).lower() == "true")

            if is_disabled or "OUT OF STOCK" in btn_text:
                print(f"⛔ OUT OF STOCK: {product_code}")
                cub_oos.append(product_code)
                continue
        except Exception as e:
            print(f"⚠️ Could not read button state for: {product_code}")

        # 1. Seleccionar unidad si NO es "Case"
        if qty_type in ["Layer", "Pallet"]:
            unit_map = {"Layer": "LAY", "Pallet": "PAL"}
            unit_value = unit_map[qty_type]

            # Hacer clic en el botón select (ej. "Case")
            try:
                select_btn = product_card.find_element(By.CSS_SELECTOR, "div.select-btn")
                driver.execute_script("arguments[0].click();", select_btn)
                time.sleep(1)
                option = product_card.find_element(By.CSS_SELECTOR, f"li[data-value='{unit_value}']")
                driver.execute_script("arguments[0].click();", option)
            except Exception as e:
                print(f"⚠️ Could not select unit {qty_type} for {product_code}: {e}")

        # 2. Completar cantidad (manejar reemplazos de DOM)
        qty_ok = False
        for attempt in range(3):
            try:
                qty_input = wait.until(EC.presence_of_element_located((By.XPATH, "//input[@type='tel']")))
                # reubicar por si el DOM cambió
                qty_input = driver.find_element(By.XPATH, "//input[@type='tel']")
                qty_input.clear()
                qty_input.send_keys(str(qty_value))
                print(f"✅ Quantity entered: {qty_value}")
                qty_ok = True
                break
            except StaleElementReferenceException:
                print("⚠️ Quantity field was replaced in the DOM. Retrying...")
                time.sleep(1)
            except TimeoutException:
                print("⚠️ Timeout locating quantity field. Retrying...")
                time.sleep(1)

        if not qty_ok:
            print(f"🔴 Could not set quantity for: {product_code}")
            cub_qty_failed.append(product_code)
            # seguimos igual para intentar el add (por si el sitio toma qty del form oculto)

        # 3. (Re)comprobar botón antes de click y setear campos ocultos del form
        try:
            # OJO: si está disabled, element_to_be_clickable fallará; por eso primero presence_of_element
            add_button = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "button.addToCartButton")))
            form = add_button.find_element(By.XPATH, "./ancestor::form")

            # Actualizar qty y unit en campos ocultos del form
            driver.execute_script(f"arguments[0].value = '{qty_value}';", form.find_element(By.NAME, "qty"))
            unit_code = {"Case": "CAS", "Layer": "LAY", "Pallet": "PAL"}.get(qty_type, "CAS")
            driver.execute_script(f"arguments[0].value = '{unit_code}';", form.find_element(By.NAME, "unit"))

            # Chequear nuevamente si quedó disabled antes de click (el sitio a veces cambia el estado)
            btn_text = (add_button.text or "").strip().upper()
            is_disabled = (add_button.get_attribute("disabled") is not None) or \
                          (str(add_button.get_attribute("aria-disabled")).lower() == "true")
            if is_disabled or "OUT OF STOCK" in btn_text:
                print(f"⛔ OUT OF STOCK: {product_code}")
                cub_oos.append(product_code)
                continue

            # Hacer clic como usuario (usar JS para evitar overlays sutiles)
            driver.execute_script("arguments[0].click();", add_button)

            # Confirmar visualmente que se mostró el modal
            try:
                wait.until(EC.presence_of_element_located((By.ID, "addToCartLayer")))
                print(f"✅ Product visually confirmed: {product_code}")
            except TimeoutException:
                print(f"⚠️ No se confirmó visualmente addToCart para: {product_code}")

            print(f"🛒 Product added to cart: {product_code}")

        except TimeoutException:
            print(f"🔴 Could not find Add buton for: {product_code}")
            cub_not_found.append(product_code)
        except Exception as e:
            print(f"🛑 Error {product_code}: {e}")

    # --------- FINAL SUMMARY CUB ---------
    print("\n===== CUB SUMMARY =====")
    print(f"🧾 Total items in file: {len(df)}")
    print(f"🟡 Out of stock: {len(cub_oos)} → {', '.join(cub_oos) if cub_oos else '-'}")
    print(f"🔴 Not found / no Add button: {len(cub_not_found)} → {', '.join(cub_not_found) if cub_not_found else '-'}")
    print(f"🟠 Quantity not set: {len(cub_qty_failed)} → {', '.join(cub_qty_failed) if cub_qty_failed else '-'}")


if supplier == "COKE":
    email = os.getenv("COKE_EMAIL")
    password = os.getenv("COKE_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de COKE en el archivo .env.")
        sys.exit()

    # Iniciar Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # Ir a la página
    driver.get("https://www.mycca.com.au/ccrz__CCSiteLogin?cclcl=en_AU")

    # Ingresar email y password
    wait.until(EC.presence_of_element_located((By.ID, "emailField")))
    driver.find_element(By.ID, "emailField").send_keys(email)
    driver.find_element(By.ID, "passwordField").send_keys(password)

    # Hacer clic en el botón de login
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "send2Dsk")))
    driver.execute_script("arguments[0].click();", login_button)

    # Listas de tracking para reportar al final
    coke_not_found = []
    coke_failed_add = []
    coke_qty_failed = []

    for index, row in df.iterrows():
        product_code = str(row["Product Code"]).strip()
        qty_value = int(str(row["Quantity"]).strip())

        print(f"🔍 Buscando producto: {product_code} | Cantidad: {qty_value}")

        added_ok = False
        for attempt in range(3):
            try:
                # 1) Abrir buscador y buscar el código
                wait.until(EC.element_to_be_clickable((By.CSS_SELECTOR, "button.hs-header-toggle.cca-button")))
                search_toggle = driver.find_element(By.CSS_SELECTOR, "button.hs-header-toggle.cca-button")
                js_click(driver, search_toggle)

                wait.until(EC.visibility_of_element_located((By.CSS_SELECTOR, "input.searchInput")))
                search_input = driver.find_element(By.CSS_SELECTOR, "input.searchInput")
                try:
                    search_input.clear()
                except ElementNotInteractableException:
                    driver.execute_script("arguments[0].value='';", search_input)
                # limpiar robusto (CONTROL y COMMAND por compatibilidad)
                search_input.send_keys(Keys.CONTROL, "a")
                search_input.send_keys(Keys.DELETE)
                search_input.send_keys(Keys.COMMAND, "a")
                search_input.send_keys(Keys.DELETE)

                search_input.send_keys(product_code)
                search_input.send_keys(Keys.RETURN)

                # 2) Encontrar el TILE que contenga el código (no cualquier addToCart)
                tile = None
                for _ in range(2):  # darle chance a hidratar
                    tile = find_coke_tile_for_code(driver, product_code, wait_seconds=12)
                    if tile:
                        break
                    time.sleep(0.5)

                if not tile:
                    coke_not_found.append(product_code)
                    print(f"❌ No se pudo encontrar/agregar {product_code}: Timeout → No se encontró tile con el código.")
                    break

                # 3) Click al botón "Add to cart" DENTRO del tile (estado 1), con re-localización y fallback
                try:
                    def _find_add_btn():
                        # Prioriza el botón con la clase exacta que pegaste del HTML
                        btns = tile.find_elements(By.CSS_SELECTOR, "button.cca-button.secondary.addToCart")
                        btn = first_visible(btns)
                        if btn:
                            return btn
                        # Fallback por texto (Case‑insensitive)
                        btns = tile.find_elements(By.XPATH,
                            ".//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'),'add to cart')]"
                        )
                        return first_visible(btns)

                    # Espera a que exista un botón Add dentro del tile
                    add_btn = None
                    for _ in range(3):
                        add_btn = _find_add_btn()
                        if add_btn:
                            break
                        time.sleep(0.3)

                    if not add_btn:
                        # Log de ayuda para depurar el tile
                        try:
                            print(f"ℹ️ Tile sin botón Add: {(tile.text or '')[:200]}")
                        except Exception:
                            pass
                        raise TimeoutException("No se encontró un botón 'Add to cart' dentro del tile.")

                    # Click robusto (scroll + JS click)
                    js_click(driver, add_btn)

                    # Fallback: si el click no cambió el estado, intentá clickear en el span interno .flex-auto
                    try:
                        time.sleep(0.2)
                        # chequear si aún no apareció la UI de cantidad
                        has_qty_ui = (
                            len(tile.find_elements(By.CSS_SELECTOR, "div.number-input-container")) > 0 or
                            len(tile.find_elements(By.CSS_SELECTOR, "input.updateCart")) > 0
                        )
                        if not has_qty_ui:
                            inner = None
                            spans = add_btn.find_elements(By.CSS_SELECTOR, "span.flex-auto")
                            inner = first_visible(spans)
                            if inner:
                                js_click(driver, inner)
                    except Exception:
                        pass

                    # 4) Esperar el estado 2 (UI de cantidad) en el MISMO tile
                    wait_quantity_ui_in_tile(driver, tile, timeout=12)

                    # 5) Setear cantidad si corresponde (input o fallback con '+') con verificación
                    ok_qty, err_qty = set_quantity_in_tile(driver, tile, qty_value)
                    if not ok_qty:
                        coke_qty_failed.append(product_code)
                        print(f"⚠️ No se pudo actualizar cantidad para {product_code}: {err_qty or 'motivo desconocido'}")
                    else:
                        # Doble verificación final (defensiva)
                        final_qty = get_qty_from_tile(tile)
                        if final_qty != qty_value:
                            coke_qty_failed.append(product_code)
                            print(f"⚠️ Cantidad inconsistente para {product_code}: quedó {final_qty}, quería {qty_value}")



                    added_ok = True
                    break

                except ElementClickInterceptedException:
                    print(f"🫣 Click interceptado en {product_code}. Cierro overlays y reintento...")
                    close_common_overlays(driver)
                    time.sleep(0.5)

                except TimeoutException as e:
                    if attempt == 2:
                        coke_not_found.append(product_code)
                        print(f"❌ No se pudo encontrar/agregar {product_code}: Timeout → {e}")
                    else:
                        print(f"⏳ Reintentando {product_code} por layout/tiempo...")
                        time.sleep(0.7)

                except StaleElementReferenceException:
                    print(f"♻️ DOM cambió (stale) para {product_code}, reintentando ({attempt+1}/3)...")
                    time.sleep(0.6)

            except TimeoutException as e:
                if attempt == 2:
                    coke_not_found.append(product_code)
                    print(f"❌ No se pudo encontrar/agregar {product_code}: Timeout general → {e}")
                else:
                    time.sleep(0.5)
            except Exception as e:
                if attempt == 2:
                    coke_failed_add.append(product_code)
                    print(f"❌ Error al agregar {product_code}: {type(e).__name__} → {e}")
                else:
                    time.sleep(0.5)

        time.sleep(0.25)  # ritmo suave para SPA

   # --- COKE Summary ---
    print("\n===== COKE SUMMARY =====")
    print(f"🧾 Total items in file: {len(df)}")
    print(f"🟡 Quantities not updated: {len(coke_qty_failed)} → {', '.join(coke_qty_failed) if coke_qty_failed else '-'}")
    print(f"🔴 Not found / missing 'Add to cart' button: {len(coke_not_found)} → {', '.join(coke_not_found) if coke_not_found else '-'}")
    print(f"🛑 Add errors: {len(coke_failed_add)} → {', '.join(coke_failed_add) if coke_failed_add else '-'}")

    time.sleep(4)


if supplier == "ALM":
    email = os.getenv("ALM_EMAIL")
    password = os.getenv("ALM_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de ALM en el archivo .env.")
        sys.exit()

    # Iniciar Chrome
    chrome_options = Options()
    chrome_options.add_argument("--start-maximized")
    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)
    actions = ActionChains(driver)

    # Ir a la página de login
    driver.get("https://www.askross.com.au/s/login/")

    # Esperar a que cargue algo visible del login
    time.sleep(2)

    # Simular TAB dos veces para llegar al input de email
    actions.send_keys(Keys.TAB)
    actions.send_keys(Keys.TAB)
    actions.send_keys(email)

    # TAB para ir al input de password
    actions.send_keys(Keys.TAB)
    actions.send_keys(password)

    # Dos TABs para llegar al botón de login y ENTER
    actions.send_keys(Keys.TAB)
    actions.send_keys(Keys.TAB)
    actions.send_keys(Keys.ENTER).perform()
    time.sleep(3)


    # Dos TABs para llegar al botón de cart y ENTER
    actions.send_keys(Keys.TAB).pause(0.5)
    actions.send_keys(Keys.TAB).pause(0.5)
    actions.send_keys(Keys.TAB).pause(0.5)
    actions.send_keys(Keys.ENTER).perform()

    # Esperar a que se abra una nueva pestaña y cambiar el foco
    try:
        WebDriverWait(driver, 10).until(lambda d: len(d.window_handles) > 1)
        driver.switch_to.window(driver.window_handles[-1])  # Ir a la nueva pestaña (última)
        print("✅ Cambio a la nueva pestaña realizado correctamente.")
        time.sleep(2)
    except Exception as e:
        print(f"❌ Error al cambiar de pestaña: {e}")
        driver.quit()
        sys.exit(1)

    # Asegurarse de estar en la nueva pestaña
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[-1])
    print("✅ Cambio a la nueva pestaña realizado correctamente.")

    try:
        # Ir directamente a la URL de Upload Order
        upload_url = "https://www.almliquor.com.au/upload-order/"
        driver.get(upload_url)
        print("✅ Navegado directamente a Upload Order.")
    except Exception as e:
        print(f"❌ No se pudo navegar a 'Upload Order': {type(e).__name__} - {e}")
        driver.quit()
        sys.exit(1)

    import csv

    # Crear archivo para subir el pedido a ALM
    alm_csv_path = "alm_upload_ready.csv"

    try:
        # Leer hoja 'ALM' de order_ready.xlsx
        alm_df = pd.read_excel("order_ready.xlsx", sheet_name="ALM")

        if "Product Code" not in alm_df.columns or "Quantity" not in alm_df.columns:
            raise ValueError("❌ Faltan columnas 'Product Code' o 'Quantity' en la hoja 'ALM'.")

        # Filtrar y preparar columnas necesarias
        output_df = alm_df[["Product Code", "Quantity"]].copy()
        output_df["Qty Units"] = ""  # Siempre vacío según lo que me indicaste

        # Renombrar columnas según la plantilla de ALM
        output_df.columns = ["Code", "Qty Ctns", "Qty Units"]

        # Guardar en formato .csv (sin índice)
        output_df.to_csv(alm_csv_path, index=False, quoting=csv.QUOTE_NONNUMERIC)
        print(f"✅ Archivo CSV generado para ALM: {alm_csv_path}")

    except Exception as e:
        print(f"❌ Error al generar archivo alm_upload_ready.csv: {type(e).__name__} - {e}")
        driver.quit()
        sys.exit(1)

    try:
        # Subir el archivo CSV oculto
        file_input = WebDriverWait(driver, 15).until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "input.js-upload-button[type='file']"))
        )
        file_input.send_keys(os.path.abspath(alm_csv_path))
        print("✅ Archivo CSV subido correctamente.")

    except Exception as e:
        print(f"❌ Error al subir o enviar el pedido: {type(e).__name__} - {e}")
        driver.quit()
        sys.exit(1)


    # Esperar unos segundos y cerrar navegador
    time.sleep(5)
    driver.quit()

    # Borrar archivo temporal
    if os.path.exists(alm_csv_path):
        os.remove(alm_csv_path)
        print("🧹 Archivo temporal eliminado: alm_upload_ready.csv")

print(f"✅ Productos a cargar para proveedor '{supplier}':")