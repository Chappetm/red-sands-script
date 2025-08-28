import sys
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from dotenv import load_dotenv
import os
import time

load_dotenv()

# Verificar argumento del proveedor
if len(sys.argv) < 2:
    print("❌ Debes indicar el proveedor. Ejemplo: python order.py cub")
    sys.exit()

supplier = sys.argv[1].upper()

if supplier == "LION":
    email = os.getenv("LION_EMAIL")
    password = os.getenv("LION_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de LION en el archivo .env.")
        sys.exit()

    # Iniciar Chrome (HEADLESS: sin ventana) y forzar descarga en Downloads
    from pathlib import Path
    download_dir = str((Path.home() / "Downloads").resolve())
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")           # 👈 modo headless moderno
    chrome_options.add_argument("--window-size=1920,1080")  # viewport razonable
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")

    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # evita visor interno → fuerza descarga
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)

    # (Opcional para tu flujo COKE si husmeas red con performance logs)
    # chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # 👇 Habilitar descargas en headless vía CDP (clave)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir}
    )

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

    # === Ir a Billing History ===
    driver.get("https://my.lionco.com/billing/history")
    tbody = wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.css-0")))

    # Ventana de semana (lunes-domingo) en Australia/Perth
    from datetime import datetime, timedelta
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Australia/Perth")
    except Exception:
        tz = None  # fallback sin tz
    now = datetime.now(tz) if tz else datetime.now()
    week_start = (now - timedelta(days=now.weekday())).date()      # lunes
    week_end   = (week_start + timedelta(days=6))                   # domingo (date)

    def parse_invoice_date(text):
        # formatos vistos: "11 Aug 25" o "11 Aug 2025"
        text = text.strip()
        for fmt in ("%d %b %y", "%d %b %Y"):
            try:
                dt = datetime.strptime(text, fmt)
                return dt.date()
            except ValueError:
                continue
        return None

    def wait_for_new_pdf(dir_path, before_set, timeout=90):
        """Espera un PDF nuevo, ignorando .crdownload y downloads.html*."""
        import time, os, fnmatch
        end = time.time() + timeout
        while time.time() < end:
            current = set(os.listdir(dir_path))
            new_files = current - before_set

            pdfs = [f for f in new_files if f.lower().endswith(".pdf")]
            if pdfs:
                newest_pdf = max((os.path.join(dir_path, f) for f in pdfs), key=os.path.getmtime)
                return newest_pdf

            # ignorar basura intermedia
            _ = [f for f in new_files if any(
                fnmatch.fnmatch(f.lower(), pat)
                for pat in ("*.crdownload", "downloads.html*", "download.html*")
            )]
            time.sleep(0.5)
        return None
    
    def cleanup_download_junk(dir_path, older_than_sec=2.0):
        """
        Borra restos de descargas HTML intermedias:
        - downloads.html (y variantes: downloads.html (1), etc.)
        - *.crdownload asociados
        Solo elimina archivos con mtime más antiguo que `older_than_sec`
        para no tocar descargas en curso.
        """
        import os, time, fnmatch
        now = time.time()
        for f in os.listdir(dir_path):
            low = f.lower()
            full = os.path.join(dir_path, f)
            try:
                age = now - os.path.getmtime(full)
            except Exception:
                continue

            if age < older_than_sec:
                continue

            if fnmatch.fnmatch(low, "downloads.html") \
            or fnmatch.fnmatch(low, "downloads.html*") \
            or fnmatch.fnmatch(low, "download.html") \
            or fnmatch.fnmatch(low, "download.html*") \
            or low.endswith(".crdownload"):
                try:
                    os.remove(full)
                except Exception:
                    pass




    rows = tbody.find_elements(By.CSS_SELECTOR, "tr.css-1xmxp7q")
    matched = []

    for row in rows:
        try:
            tds = row.find_elements(By.CSS_SELECTOR, "td")
            if len(tds) < 10:
                continue

            # Columns by index:
            # 0: checkbox | 1: invoice no | 2: invoice date | 3: status
            # 4: type     | 5: PO         | 6: order link   | 7: due
            # 8: amount   | 9: download button

            # ---- filter by Type = "Customer Invoice" ----
            type_text = (tds[4].text or "").strip()
            if "customer invoice" not in type_text.lower():
                continue  # skip anything else (e.g., credit note, statement, etc.)

            # ---- date filter: only this week ----
            date_text = (tds[2].text or "").strip()
            inv_date = parse_invoice_date(date_text)
            if not inv_date or not (week_start <= inv_date <= week_end):
                continue

            # ---- invoice number (from <p> if present, else cell text) ----
            try:
                invoice_no_el = tds[1].find_element(By.CSS_SELECTOR, "p, .chakra-text, *")
                invoice_no = (invoice_no_el.text or "").strip()
            except Exception:
                invoice_no = (tds[1].text or "").strip()

            matched.append((row, invoice_no, inv_date))

        except Exception as e:
            print(f"⚠️ Skipping row due to error: {e}")


    if not matched:
        print("No invoices this week")
        # Si quieres cerrar el navegador aquí:
        # driver.quit()
        sys.exit(0)

        # Descargar cada invoice encontrada (clic en el botón de la última columna)
    for row, invoice_no, inv_date in matched:
        before = set(os.listdir(download_dir))
        try:
            download_btn = row.find_elements(By.CSS_SELECTOR, "td")[-1].find_element(By.CSS_SELECTOR, "button")
        except Exception:
            download_btn = row.find_element(By.CSS_SELECTOR, "button")

        download_btn.click()

        downloaded_path = wait_for_new_pdf(download_dir, before, timeout=90)

        # Limpieza final defensiva
        cleanup_download_junk(download_dir, older_than_sec=1.0)


        if not downloaded_path:
            print(f"⚠️ Could not detect downloaded file for invoice {invoice_no}")
            continue

        # Ya no renombramos ni movemos. Dejamos el nombre que ponga el portal.
        print(f"✅ Downloaded to Downloads: {os.path.basename(downloaded_path)}")

        # Opcional: limpiar restos de descargas HTML fantasma
        try:
            for f in os.listdir(download_dir):
                name = f.lower()
                if name.endswith(".crdownload") and name.startswith("downloads.html"):
                    try:
                        os.remove(os.path.join(download_dir, f))
                    except Exception:
                        pass
        except Exception:
            pass

if supplier == "CUB":
    email = os.getenv("CUB_EMAIL")
    password = os.getenv("CUB_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de CUB en el archivo .env.")
        sys.exit()

    # Iniciar Chrome (HEADLESS: sin ventana) y forzar descarga en Downloads
    from pathlib import Path
    download_dir = str((Path.home() / "Downloads").resolve())
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")           # 👈 modo headless moderno
    chrome_options.add_argument("--window-size=1920,1080")  # viewport razonable
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")

    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # evita visor interno → fuerza descarga
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)

    # (Opcional para tu flujo COKE si husmeas red con performance logs)
    # chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # 👇 Habilitar descargas en headless vía CDP (clave)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir}
    )



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
    time.sleep(2)

    # === Ir a Billing History ===
    driver.get("https://online.cub.com.au/sabmStore/en/your-business/billing")

    # === Procesar tabla de Billing y descargar invoices de esta semana ===
    from datetime import datetime, timedelta
    from urllib.parse import urljoin
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Australia/Perth")
    except Exception:
        tz = None

    def now_local():
        return datetime.now(tz) if tz else datetime.now()

    def week_window():
        n = now_local()
        start = (n - timedelta(days=n.weekday())).date()  # lunes
        end = start + timedelta(days=6)                   # domingo
        return start, end

    def parse_cub_date(text):
        # Formato visible: "11/08/25" (dd/mm/yy)
        text = text.strip()
        for fmt in ("%d/%m/%y", "%d/%m/%Y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    def wait_for_new_pdf(dir_path, before_set, timeout=90):
        """Espera un PDF nuevo, ignorando .crdownload y downloads.html*."""
        import time, os, fnmatch
        deadline = time.time() + timeout
        while time.time() < deadline:
            current = set(os.listdir(dir_path))
            new_files = current - before_set

            # si aparece un PDF, lo devolvemos
            pdfs = [f for f in new_files if f.lower().endswith(".pdf")]
            if pdfs:
                newest_pdf = max((os.path.join(dir_path, f) for f in pdfs), key=os.path.getmtime)
                return newest_pdf

            # ignorar basura intermedia
            _junk = [f for f in new_files if any(
                fnmatch.fnmatch(f.lower(), pat)
                for pat in ("*.crdownload", "downloads.html*", "download.html*")
            )]

            time.sleep(0.5)
        return None

    def cleanup_download_junk(dir_path, older_than_sec=2.0):
        """Elimina downloads.html*, download.html* y *.crdownload antiguos (no activos)."""
        import os, time, fnmatch
        now = time.time()
        for f in os.listdir(dir_path):
            full = os.path.join(dir_path, f)
            try:
                if not os.path.isfile(full):
                    continue
                age = now - os.path.getmtime(full)
                if age < older_than_sec:
                    continue
            except Exception:
                continue

            low = f.lower()
            if (fnmatch.fnmatch(low, "downloads.html") or
                fnmatch.fnmatch(low, "downloads.html*") or
                fnmatch.fnmatch(low, "download.html") or
                fnmatch.fnmatch(low, "download.html*") or
                low.endswith(".crdownload")):
                try:
                    os.remove(full)
                except Exception:
                    pass


    # Esperar a que haya filas
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody tr")))
    rows = driver.find_elements(By.CSS_SELECTOR, "tbody tr")

    week_start, week_end = week_window()
    matched = []

    for row in rows:
        tds = row.find_elements(By.CSS_SELECTOR, "td")
        if len(tds) < 11:
            continue

        date_text = tds[4].text.strip()
        inv_date = parse_cub_date(date_text)
        if not inv_date:
            continue
        if not (week_start <= inv_date <= week_end):
            continue

        # href del PDF (aunque el <a> esté oculto, podemos leer el atributo)
        try:
            link_el = tds[10].find_element(By.CSS_SELECTOR, "a")
            href = link_el.get_attribute("href") or ""
        except Exception:
            href = ""

        if not href:
            # si el href es relativo en vez de absoluto, lo resolvemos abajo con urljoin
            # muchos CUB dan href relativo tipo "billing/invoice/pdf/7507882861"
            try:
                href = tds[10].find_element(By.CSS_SELECTOR, "a").get_attribute("href")
            except Exception:
                href = ""

        if not href:
            # último intento: buscar cualquier <a> dentro de la fila con "invoice/pdf"
            try:
                any_link = row.find_element(By.CSS_SELECTOR, "a[href*='invoice/pdf']")
                href = any_link.get_attribute("href")
            except Exception:
                href = ""

        if not href:
            continue

        abs_url = urljoin(driver.current_url, href)
        matched.append((inv_date, abs_url))

    if not matched:
        print("No invoices this week")
        sys.exit(0)

    # Descargar cada PDF abriendo directamente el URL (dispara descarga)
    for inv_date, pdf_url in matched:
        before = set(os.listdir(download_dir))
        driver.get(pdf_url)
        downloaded = wait_for_new_pdf(download_dir, before, timeout=90)
        if downloaded:
            cleanup_download_junk(download_dir, older_than_sec=1.0)
            print(f"✅ Invoice downloaded")
        else:
            print(f"⚠️ Could not detect downloaded file for: {pdf_url}")

if supplier == "COKE":
    email = os.getenv("COKE_EMAIL")
    password = os.getenv("COKE_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de COKE en el archivo .env.")
        sys.exit()

    # Iniciar Chrome (HEADLESS: sin ventana) y forzar descarga en Downloads
    from pathlib import Path
    download_dir = str((Path.home() / "Downloads").resolve())
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")           # 👈 modo headless moderno
    chrome_options.add_argument("--window-size=1920,1080")  # viewport razonable
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")

    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # evita visor interno → fuerza descarga
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)

    # (Opcional para tu flujo COKE si husmeas red con performance logs)
    # chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # 👇 Habilitar descargas en headless vía CDP (clave)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir}
    )



    # Ir a la página
    driver.get("https://www.mycca.com.au/ccrz__CCSiteLogin?cclcl=en_AU")

    # Ingresar email y password
    wait.until(EC.presence_of_element_located((By.ID, "emailField")))
    driver.find_element(By.ID, "emailField").send_keys(email)
    driver.find_element(By.ID, "passwordField").send_keys(password)

    # Hacer clic en el botón de login
    login_button = wait.until(EC.element_to_be_clickable((By.ID, "send2Dsk")))
    driver.execute_script("arguments[0].click();", login_button)

    # === Navegar a "My Account" → Invoices ===
    # Opción A: clic en el menú "My Account"
    try:
        my_account = wait.until(EC.element_to_be_clickable((
            By.CSS_SELECTOR,
            "a.cc_menu_type_url[data-menuid='myAccount'][href*='pageKey=invoices']"
        )))
        driver.execute_script("arguments[0].click();", my_account)
    except Exception:
        # Opción B (fallback): ir directo por URL
        driver.get("https://www.mycca.com.au/ccrz__CCPage?pageKey=invoices")

    # Esperar que carguen las tarjetas de invoices
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "li.CCA_MA_Invoice_Card")))

    from datetime import datetime, timedelta
    from urllib.parse import urljoin
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Australia/Perth")
    except Exception:
        tz = None

    def now_local():
        return datetime.now(tz) if tz else datetime.now()

    def week_window():
        n = now_local()
        start = (n - timedelta(days=n.weekday())).date()  # lunes
        end = start + timedelta(days=6)                   # domingo
        return start, end

    def parse_au_date(text):
        # Formato visible: "12/08/2025" → dd/mm/YYYY
        text = text.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    def wait_for_new_download(dir_path, before_set, timeout=90):
        import time, os
        end = time.time() + timeout
        while time.time() < end:
            current = set(os.listdir(dir_path))
            new_files = [f for f in current - before_set if not f.endswith(".crdownload")]
            if new_files:
                newest = max((os.path.join(dir_path, f) for f in new_files), key=os.path.getmtime)
                return newest
            time.sleep(0.5)
        return None

    week_start, week_end = week_window()

    # Helpers para el DocumentViewer
    from urllib.parse import urljoin

    def try_click_open_deep():
        """Clic en 'Open' (o 'OPEN') en la página o dentro de iframes (hasta 2 niveles)."""
        XPATHS = [
            "//button[normalize-space()='Open' or normalize-space()='OPEN']",
            "//a[normalize-space()='Open' or normalize-space()='OPEN']",
            "//*[@role='button'][contains(translate(., 'open', 'OPEN'), 'OPEN')]",
            "//*[self::a or self::button][contains(translate(., 'open', 'OPEN'), 'OPEN')]",
        ]

        def search_and_click():
            for xp in XPATHS:
                els = driver.find_elements(By.XPATH, xp)
                for el in els:
                    try:
                        driver.execute_script("arguments[0].click();", el)
                        return True
                    except Exception:
                        pass
            return False

        # Documento principal
        if search_and_click():
            return True

        # Iframes (nivel 1 y 2)
        frames1 = driver.find_elements(By.TAG_NAME, "iframe")
        for fr1 in frames1:
            try:
                driver.switch_to.frame(fr1)
                if search_and_click():
                    driver.switch_to.default_content()
                    return True

                frames2 = driver.find_elements(By.TAG_NAME, "iframe")
                for fr2 in frames2:
                    try:
                        driver.switch_to.frame(fr2)
                        if search_and_click():
                            driver.switch_to.default_content()
                            return True
                        driver.switch_to.parent_frame()
                    except Exception:
                        driver.switch_to.parent_frame()
                driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()
        return False

    def find_embedded_pdf_url():
        """
        Busca <embed|object|iframe> con 'application/pdf' o 'pdf' en src/data y devuelve URL absoluta.
        """
        selectors = [
            "embed[type='application/pdf']",
            "object[type='application/pdf']",
            "iframe[src*='pdf']",
            "iframe[src*='Document'][src*='document']",
        ]
        # Documento principal
        for sel in selectors:
            for el in driver.find_elements(By.CSS_SELECTOR, sel):
                src = el.get_attribute("src") or el.get_attribute("data")
                if src:
                    return urljoin(driver.current_url, src)

        # Iframes (1 nivel)
        frames1 = driver.find_elements(By.TAG_NAME, "iframe")
        for fr1 in frames1:
            try:
                driver.switch_to.frame(fr1)
                for sel in selectors:
                    for el in driver.find_elements(By.CSS_SELECTOR, sel):
                        src = el.get_attribute("src") or el.get_attribute("data")
                        if src:
                            driver.switch_to.default_content()
                            return urljoin(driver.current_url, src)
                driver.switch_to.default_content()
            except Exception:
                driver.switch_to.default_content()
        return None


    cards = driver.find_elements(By.CSS_SELECTOR, "li.CCA_MA_Invoice_Card")
    matched = []

    for li in cards:
        # Buscar dentro de la card el dt "Issue date" y su dd siguiente
        try:
            issue_dd = li.find_element(
                By.XPATH,
                ".//dt[normalize-space()='Issue date']/following-sibling::dd[1]"
            )
            date_text = issue_dd.text.strip()
            inv_date = parse_au_date(date_text)
            if not inv_date:
                continue
        except Exception:
            continue

        if not (week_start <= inv_date <= week_end):
            continue

        # Enlace "View invoice" (pdfType=invoice)
        view_links = li.find_elements(By.CSS_SELECTOR, "a[href*='CCADocumentViewer'][href*='pdfType=invoice']")
        if not view_links:
            # fallback: buscar por texto visible
            try:
                view_links = [li.find_element(By.XPATH, ".//a[.//span[contains(., 'View invoice')]]")]
            except Exception:
                continue

        href = view_links[0].get_attribute("href") or ""
        if not href:
            continue

        abs_url = urljoin('https://www.mycca.com.au/', href)

        # Intento de descarga para la ÚNICA invoice de COKE
        before = set(os.listdir(download_dir))

        # 1) Abrir DocumentViewer
        driver.get(abs_url)

        # 2) A veces descarga sola → espera corta
        downloaded = False

        if not downloaded:
            embedded = find_embedded_pdf_url()
            if embedded:
                driver.get(embedded)
                downloaded = wait_for_new_download(download_dir, before, timeout=90)

        if downloaded:
            print(f"✅ Invoices downloaded")
        else:
            print("⚠️ Could not download COKE invoice (no PDF detected).")

        break

if supplier == "ALM":
    email = os.getenv("ALM_EMAIL")
    password = os.getenv("ALM_PASSWORD")

    if not email or not password:
        print("❌ Faltan las credenciales de ALM en el archivo .env.")
        sys.exit()

    # Iniciar Chrome (HEADLESS: sin ventana) y forzar descarga en Downloads
    from pathlib import Path
    download_dir = str((Path.home() / "Downloads").resolve())
    os.makedirs(download_dir, exist_ok=True)

    chrome_options = Options()
    chrome_options.add_argument("--headless=new")           # 👈 modo headless moderno
    chrome_options.add_argument("--window-size=1920,1080")  # viewport razonable
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-popup-blocking")

    chrome_prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "plugins.always_open_pdf_externally": True,  # evita visor interno → fuerza descarga
    }
    chrome_options.add_experimental_option("prefs", chrome_prefs)

    # (Opcional para tu flujo COKE si husmeas red con performance logs)
    # chrome_options.set_capability("goog:loggingPrefs", {"performance": "ALL"})

    driver = webdriver.Chrome(options=chrome_options)
    wait = WebDriverWait(driver, 20)

    # 👇 Habilitar descargas en headless vía CDP (clave)
    driver.execute_cdp_cmd(
        "Page.setDownloadBehavior",
        {"behavior": "allow", "downloadPath": download_dir}
    )
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
        time.sleep(2)
    except Exception as e:
        print(f"❌ Error al cambiar de pestaña: {e}")
        driver.quit()
        sys.exit(1)

    # Asegurarse de estar en la nueva pestaña
    time.sleep(2)
    driver.switch_to.window(driver.window_handles[-1])

    try:
        # Ir directamente a la URL de Upload Order
        upload_url = "https://www.almliquor.com.au/my-report/report?reports=ALMCUSTINV"
        driver.get(upload_url)
    except Exception as e:
        print(f"❌ No se pudo navegar a 'Upload Order': {type(e).__name__} - {e}")
        driver.quit()
        sys.exit(1)    


    # === Helpers (semana actual y descarga) ===
    from datetime import datetime, timedelta
    from urllib.parse import urljoin
    try:
        from zoneinfo import ZoneInfo
        tz = ZoneInfo("Australia/Perth")
    except Exception:
        tz = None

    def now_local():
        return datetime.now(tz) if tz else datetime.now()

    def week_window():
        n = now_local()
        start = (n - timedelta(days=n.weekday())).date()  # lunes
        end = start + timedelta(days=6)                   # domingo
        return start, end

    def parse_au_date_ddmmyyyy(text: str):
        text = text.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y"):
            try:
                return datetime.strptime(text, fmt).date()
            except ValueError:
                continue
        return None

    def parse_from_data_order(s: str):
        # data-order="20250811" → YYYYMMDD
        try:
            return datetime.strptime(s.strip(), "%Y%m%d").date()
        except Exception:
            return None

    def wait_for_new_pdf(dir_path, before_set, timeout=120, drain_timeout=None):
        """
        Espera a que aparezca un PDF nuevo en dir_path (ignora .crdownload y downloads.html*).
        Si se pasa drain_timeout, además espera a que terminen los .crdownload generados
        después del snapshot 'before_set' para evitar residuos.
        """
        import os, time

        deadline = time.time() + timeout
        newest_pdf = None

        # 1) Esperar a que aparezca el PDF NUEVO
        while time.time() < deadline:
            current = set(os.listdir(dir_path))
            new_files = current - before_set

            # Ignorar basura tipo 'downloads.html*'
            pdfs = [
                f for f in new_files
                if f.lower().endswith(".pdf") and not f.lower().startswith("downloads")
            ]
            if pdfs:
                newest_pdf = max(
                    (os.path.join(dir_path, f) for f in pdfs),
                    key=os.path.getmtime
                )

                # Esperar a que el tamaño se estabilice (archivo ya escrito)
                last_size = -1
                stable_ticks = 0
                while time.time() < deadline:
                    try:
                        size = os.path.getsize(newest_pdf)
                    except Exception:
                        size = -1

                    if size > 0 and size == last_size:
                        stable_ticks += 1
                        if stable_ticks >= 3:   # ~0.9s con sleep(0.3)
                            break
                    else:
                        stable_ticks = 0

                    last_size = size
                    time.sleep(0.3)
                break

            time.sleep(0.3)

        if not newest_pdf:
            return None

        # 2) (Opcional) Drenar .crdownload que hayan empezado después de 'before_set'
        if drain_timeout:
            drain_deadline = time.time() + drain_timeout
            while time.time() < drain_deadline:
                current = set(os.listdir(dir_path))
                leftovers = [f for f in (current - before_set) if f.endswith(".crdownload")]
                if not leftovers:
                    break
                time.sleep(0.3)

        return newest_pdf


    def cleanup_download_junk(dir_path, older_than_sec=1.0):
        """Deletes downloads.html*, download.html* and *.crdownload that are not active."""
        import os, time, fnmatch
        now = time.time()
        for f in os.listdir(dir_path):
            full = os.path.join(dir_path, f)
            if not os.path.isfile(full):
                continue
            low = f.lower()
            if (low.endswith(".crdownload") or
                fnmatch.fnmatch(low, "downloads.html") or
                fnmatch.fnmatch(low, "downloads.html*") or
                fnmatch.fnmatch(low, "download.html") or
                fnmatch.fnmatch(low, "download.html*")):
                try:
                    if now - os.path.getmtime(full) > older_than_sec:
                        os.remove(full)
                except Exception:
                    pass

    
    # === Leer tabla y descargar invoices de esta semana ===
    # Esperar a que aparezcan filas dentro de <tbody class="d-xs-block">
    wait.until(EC.presence_of_element_located((By.CSS_SELECTOR, "tbody.d-xs-block tr")))
    rows = driver.find_elements(By.CSS_SELECTOR, "tbody.d-xs-block tr")

    week_start, week_end = week_window()
    matched_links = []

    for row in rows:
        tds = row.find_elements(By.CSS_SELECTOR, "td")
        if len(tds) < 4:
            continue

        # Columna de fecha: td[1], tiene text "11/08/2025" y atributo data-order="20250811"
        td_date = tds[1]
        date_attr = td_date.get_attribute("data-order")
        if date_attr:
            inv_date = parse_from_data_order(date_attr)
        else:
            inv_date = parse_au_date_ddmmyyyy(td_date.text)

        if not inv_date or not (week_start <= inv_date <= week_end):
            continue

        # Link de descarga: hay un <a ...> envolviendo el botón "Download"
        # Puede estar en td[0] (número de invoice) y repetido en td[3] (botón)
        # Buscamos cualquier <a> con /my-report/report-download/
        link = None
        try:
            link = row.find_element(By.CSS_SELECTOR, "a[href*='/my-report/report-download/']")
        except Exception:
            pass
        if not link:
            continue

        href = link.get_attribute("href") or ""
        if not href:
            continue

        abs_url = urljoin(driver.current_url, href)
        matched_links.append(abs_url)

    if not matched_links:
        print("No invoices this week")
        sys.exit(0)

        # Descargar todas las invoices de esta semana (sin clicks extra; solo GET por URL)
    # Deduplicar manteniendo orden, por si alguna fila repite el mismo href
    matched_links = list(dict.fromkeys(matched_links))

    for abs_url in matched_links:
        before = set(os.listdir(download_dir))

        # Abrimos directamente la URL de descarga/visor; con prefs, Chrome descarga el PDF
        driver.get(abs_url)

        downloaded = wait_for_new_pdf(download_dir, before, timeout=120, drain_timeout=90)
        if downloaded:
            print(f"✅ Downloaded to Downloads: {os.path.basename(downloaded)}")
        else:
            # Reintento simple
            driver.get(abs_url)
            downloaded = wait_for_new_pdf(download_dir, before, timeout=90, drain_timeout=60)
            if downloaded:
                print(f"✅ Invoice downloaded")
            else:
                print(f"⚠️ Could not detect downloaded file for: {abs_url}")

        # Pequeña pausa para evitar solapamiento de descargas
        cleanup_download_junk(download_dir)
        time.sleep(1.0)


