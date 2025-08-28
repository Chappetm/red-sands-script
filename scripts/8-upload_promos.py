# scripts/pricelist_playwright.py
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
import os
import time
import random
import re
import json
import pandas as pd
from dotenv import load_dotenv

# =========================
# Paths & Config (project layout)
# =========================
# This script lives in /scripts. Project root is its parent folder.
SCRIPTS_DIR = os.path.dirname(__file__)
PROJECT_ROOT = os.path.abspath(os.path.join(SCRIPTS_DIR, ".."))

# .env is at project root (same level as /assets and /scripts)
dotenv_path = os.path.join(PROJECT_ROOT, ".env")
load_dotenv(dotenv_path)

EMAIL = os.getenv("LIGHTSPEED_EMAIL") or ""
PASSWORD = os.getenv("LIGHTSPEED_PASSWORD") or ""

PRICELIST_URL = "https://my.kounta.com/pricelist"

# Input files relative to project root
PRODUCTS_XLSX = os.path.join(PROJECT_ROOT, "assets", "products.xlsx")
PROMOS_XLSX = os.path.join(PROJECT_ROOT, "bottlemart_promos", "promo_products.xlsx")
PROMOS_SHEET = "Promocionados"

# Name for the new price list
PRICELIST_NAME = "Test"

# Persistent profile (stored at project root)
PROFILE_DIR = os.path.join(PROJECT_ROOT, "chrome-profiles", "Bot-Profile")


# =========================
# Utilities
# =========================
def human_delay(a=0.25, b=0.8):
    time.sleep(random.uniform(a, b))


def norm_code(raw: str) -> str:
    s = str(raw).strip()
    s = re.sub(r"[^\d]", "", s)
    s = s.lstrip("0")
    return s


def read_products_lookup(products_xlsx: str) -> dict:
    """
    Reads assets/products.xlsx and builds: code -> product_name
    Supports multiple codes per row separated by '/', ',', ';', '|'
    """
    if not os.path.exists(products_xlsx):
        raise FileNotFoundError(f"Missing products file: {products_xlsx}")

    code_to_name = {}
    xls = pd.ExcelFile(products_xlsx)
    for sheet in xls.sheet_names:
        df = xls.parse(sheet)
        if "Product Code" not in df.columns or "Product Name" not in df.columns:
            continue
        for _, row in df.iterrows():
            name = str(row["Product Name"]).strip()
            raw_codes = str(row["Product Code"])
            parts = re.split(r"[\/,;|]+", str(raw_codes))
            for part in parts:
                code = norm_code(part)
                if not code:
                    continue
                if code in code_to_name and code_to_name[code] != name:
                    print(f"⚠️ Duplicate code '{code}' already maps to '{code_to_name[code]}', ignoring '{name}'")
                    continue
                code_to_name[code] = name
    return code_to_name


def read_promos(promos_xlsx: str, sheet: str) -> pd.DataFrame:
    if not os.path.exists(promos_xlsx):
        raise FileNotFoundError(f"Missing promos file: {promos_xlsx}")
    return pd.read_excel(promos_xlsx, sheet_name=sheet)


def patch_name_by_pk(base_name: str, promo_name: str) -> str:
    """
    If promo_name ends with '... 4pk' / '10pk', adjust (S|C)<n> suffix in base_name.
    """
    m = re.search(r"(\d+)\s*pk$", str(promo_name).lower())
    if not m:
        return base_name
    pack_number = m.group(1)
    return re.sub(r"(S|C)\d+$", lambda mm: mm.group(1) + pack_number, base_name)


# =========================
# Login like 4-upload.py (anti-detection)
# =========================
def launch_persistent():
    """
    Launch a persistent Chromium context with anti-detection tweaks,
    mirroring the approach from your 4-upload.py.
    """
    os.makedirs(PROFILE_DIR, exist_ok=True)
    p = sync_playwright().start()
    context = p.chromium.launch_persistent_context(
        user_data_dir=PROFILE_DIR,
        headless=False,  # set True if you want to hide the browser window
        viewport={'width': 1366, 'height': 768},
        user_agent='Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        locale='es-ES',
        timezone_id='Europe/Madrid',
        extra_http_headers={
            'Accept-Language': 'es-ES,es;q=0.9,en;q=0.8',
            'Accept-Encoding': 'gzip, deflate, br',
            'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
            'DNT': '1',
            'Connection': 'keep-alive',
            'Upgrade-Insecure-Requests': '1',
        },
        args=[
            '--no-sandbox',
            '--disable-blink-features=AutomationControlled',
            '--disable-web-security',
            '--disable-features=VizDisplayCompositor',
            '--disable-dev-shm-usage',
            '--no-first-run',
            '--disable-extensions',
            '--disable-plugins',
            '--disable-default-apps',
            '--disable-background-mode'
        ]
    )
    return p, context


def human_like_typing(page, selector, text):
    """Simula escritura humana con delays variables"""
    element = page.locator(selector)
    element.click()
    human_delay(0.5, 1)
    
    for char in text:
        element.type(char)
        time.sleep(random.uniform(0.05, 0.15))


def ensure_logged_in(context):
    """
    Reuse session if possible (recent user tile).
    Handle full email+password login with human-like typing.
    """
    page = context.new_page()
    
    # Add anti-detection script like in 4-upload.py
    page.add_init_script("""
        Object.defineProperty(navigator, 'webdriver', {
            get: () => undefined,
        });
        
        window.chrome = {
            runtime: {},
        };
        
        Object.defineProperty(navigator, 'plugins', {
            get: () => [1, 2, 3, 4, 5],
        });
    """)

    print("Navegando a la página de login...")
    page.goto("https://my.kounta.com/login", wait_until='networkidle')
    human_delay(2, 4)

    # Verificar si hay una sesión guardada (usuario reciente)
    print("Verificando si hay sesión guardada...")
    recent_user_selectors = [
        'li.user.recentUser',
        'li[data-username*="matias.chappet"]',
        'li[class*="recentUser"]'
    ]

    user_found = False
    for selector in recent_user_selectors:
        try:
            if page.locator(selector).is_visible(timeout=3000):
                print("✅ Usuario reciente encontrado, haciendo click...")
                page.locator(selector).click()
                human_delay(2, 3)
                user_found = True
                break
        except:
            continue

    if not user_found:
        # Buscar y llenar el campo de email
        print("Buscando campo de email...")
        email_selectors = [
            'input#loginform_username',
            'input[name="email"]',
            'input[id="loginform_username"]'
        ]
        
        email_field = None
        for selector in email_selectors:
            try:
                if page.locator(selector).is_visible(timeout=2000):
                    email_field = selector
                    break
            except:
                continue
        
        if not email_field:
            print("❌ No se pudo encontrar el campo de email")
            raise RuntimeError("Email field not found")
        
        print("Rellenando email...")
        human_like_typing(page, email_field, EMAIL)
        human_delay(1, 2)
        
    else:
        # Flujo normal - email y contraseña
        print("Usuario encontrado...")
    
    # Buscar y llenar el campo de contraseña
    print("Buscando campo de contraseña...")
    password_selectors = [
        'input#loginform_password',
        'input[name="password"]',
        'input[id="loginform_password"]'
    ]
    
    password_field = None
    for selector in password_selectors:
        try:
            if page.locator(selector).is_visible(timeout=2000):
                password_field = selector
                break
        except:
            continue
    
    if not password_field:
        print("❌ No se pudo encontrar el campo de contraseña")
        raise RuntimeError("Password field not found")
    
    print("Rellenando contraseña...")
    human_like_typing(page, password_field, PASSWORD)
    human_delay(1, 2)
    
    # Buscar y hacer click en el botón de login
    print("Buscando botón de login...")
    login_button_selectors = [
        'input#btnLogin',
        'input[id="btnLogin"]',
        'input[type="submit"]',
        'input[value="Log in"]'
    ]
    
    login_button = None
    for selector in login_button_selectors:
        try:
            if page.locator(selector).is_visible(timeout=2000):
                login_button = selector
                break
        except:
            continue
    
    if not login_button:
        print("❌ No se pudo encontrar el botón de login")
        raise RuntimeError("Login button not found")
    
    print("Haciendo click en login...")
    
    # Simular movimiento de mouse antes del click
    button_element = page.locator(login_button)
    box = button_element.bounding_box()
    if box:
        page.mouse.move(
            box['x'] + box['width'] / 2, 
            box['y'] + box['height'] / 2
        )
        human_delay(0.5, 1)
    
    button_element.click()
    
    # Esperar a que se procese el login
    print("Esperando respuesta...")
    page.wait_for_load_state('networkidle', timeout=10000)
    human_delay(3, 5)
    
    # Verificar si el login fue exitoso
    current_url = page.url
    if 'login' not in current_url.lower() or 'dashboard' in current_url.lower():
        print("✅ Login exitoso!")
        print(f"URL actual: {current_url}")
    else:
        print("❌ El login parece haber fallado")
        print(f"URL actual: {current_url}")
        
        # Buscar mensajes de error
        error_selectors = [
            '[class*="error"]',
            '[class*="alert"]',
            '[data-testid*="error"]',
            '.message'
        ]
        
        for selector in error_selectors:
            try:
                error_element = page.locator(selector)
                if error_element.is_visible():
                    error_text = error_element.text_content()
                    print(f"Mensaje de error: {error_text}")
            except:
                continue
    
    return page


# =========================
# Price List flow
# =========================
def open_pricelist(page):
    print("➡️ Navigating to Price Lists…")
    page.goto(PRICELIST_URL, wait_until="load")
    # Some tenants lazy-load; force networkidle then small human pause
    try:
        page.wait_for_load_state("networkidle", timeout=15000)
    except Exception:
        pass
    human_delay(0.8, 1.6)

    # Guard: ensure we actually see the create button or an equivalent anchor in the shell
    # Try multiple selectors (analytics attr can change)
    candidates = [
        'button[data-analytics="btnPriceLists_createList"]',
        'button:has-text("Create price list")',
        'button:has-text("Create Price List")',
    ]
    found = False
    for sel in candidates:
        try:
            if page.locator(sel).first.is_visible(timeout=4000):
                found = True
                break
        except Exception:
            continue

    if not found:
        # Retry once with a hard reload (helps after fresh login)
        print("♻️ Create button not visible yet. Reloading page…")
        page.reload(wait_until="load")
        try:
            page.wait_for_load_state("networkidle", timeout=15000)
        except Exception:
            pass
        human_delay(0.8, 1.2)



def create_price_list(page, name: str):
    """
    1) Click "Create price list"
    2) Fill the modal input
    3) Click "Create"
    """
    print(f"🧾 Creating Price List: {name}")
    # Robust selector set (analytics attr can change, also try by text)
    btn = page.locator(
        'button[data-analytics="btnPriceLists_createList"], '
        'button:has-text("Create price list"), '
        'button:has-text("Create Price List")'
    )
    btn.wait_for(state="visible", timeout=20000)

    btn.wait_for(state="visible", timeout=10000)
    btn.click()
    human_delay(0.5, 1.0)

    name_input = page.locator('input[placeholder="Enter price list name"]')
    name_input.wait_for(state="visible", timeout=8000)
    name_input.fill("")
    human_delay(0.2, 0.4)
    name_input.type(name, delay=random.randint(20, 60))

    create_btn = page.locator('button[type="submit"].btnPrimary')
    create_btn.wait_for(state="visible", timeout=8000)

    try:
        box = create_btn.bounding_box()
        if box:
            page.mouse.move(box["x"] + box["width"]/2, box["y"] + box["height"]/2)
            human_delay(0.2, 0.5)
    except Exception:
        pass

    create_btn.click()
    page.wait_for_load_state("networkidle")
    human_delay(0.6, 1.1)
    print("✅ Price List created.")


def fill_products_and_prices(page, matched_rows):
    """
    For each product:
      - Search by exact name
      - Wait for the exact match block
      - Fill price in the decimal input
    """
    for item in matched_rows:
        name = item["Product Name"]
        price = str(item["Retail Price"]).strip()
        try:
            print(f"🖊️ Typing: {name} -> ${price}")

            search = page.locator('input[data-chaminputid="searchInput"]')
            search.wait_for(state="visible", timeout=10000)
            search.click()
            human_delay(0.1, 0.3)
            search.fill("")
            human_delay(0.1, 0.3)
            search.type(name, delay=random.randint(10, 30))
            search.press("Enter")
            human_delay(0.8, 1.3)

            result = page.locator(
                f'//div[contains(@class, "Alignment__AlignmentContainer") and normalize-space(text())="{name}"]'
            )
            result.wait_for(state="visible", timeout=8000)
            human_delay(0.3, 0.7)

            price_input = page.locator('input[data-chaminputid="textInput"][inputmode="decimal"]')
            price_input.wait_for(state="visible", timeout=8000)
            price_input.fill("")
            human_delay(0.1, 0.2)
            price_input.type(price, delay=random.randint(15, 40))
            human_delay(0.2, 0.4)

            print(f"✅ Set price OK: {name} at ${price}")
            human_delay(0.4, 0.8)

        except PlaywrightTimeoutError as te:
            print(f"❌ Timeout for {name}: {te}")
        except Exception as e:
            print(f"❌ Error for {name}: {e}")


def save_pricelist(page):
    btn = page.locator('button[data-analytics="btnPriceLists_saveList"]')
    btn.wait_for(state="visible", timeout=10000)
    btn.click()
    human_delay(1.0, 1.8)
    print("💾 Save sent.")
    page.wait_for_load_state("networkidle")
    print("✅ Promotions saved.")


# =========================
# Matching / Unmatched
# =========================
def build_matched_rows(products_lookup: dict, promos_df: pd.DataFrame):
    matched, unmatched = [], []
    for _, row in promos_df.iterrows():
        try:
            code = norm_code(int(float(row["Brewer code"])))
        except Exception:
            code = norm_code(row["Brewer code"])

        price = str(row["Retail Price"]).replace("$", "").strip()
        base_name = products_lookup.get(code)

        if base_name:
            promo_name = str(row.get("Promoted product", "")).strip()
            updated_name = patch_name_by_pk(base_name, promo_name)
            matched.append({
                "Product Code": code,
                "Product Name": updated_name,
                "Retail Price": price
            })
        else:
            unmatched.append({
                "Product Code": code,
                "Retail Price": price,
                "Original Name": str(row.get("Promoted product", "")).strip()
            })

    if unmatched:
        print("\n❌ Not found in products.xlsx:")
        for item in unmatched:
            print(f"- Code: {item['Product Code']} | Promo name: {item['Original Name']} | Price: {item['Retail Price']}")

    return matched


# =========================
# Main
# =========================
def main():
    # 1) Data
    print("🔧 Reading Excel sources…")
    products_lookup = read_products_lookup(PRODUCTS_XLSX)
    promos_df = read_promos(PROMOS_XLSX, PROMOS_SHEET)
    matched_rows = build_matched_rows(products_lookup, promos_df)
    print(f"📦 Total matched: {len(matched_rows)}")

    # 2) Browser & login
    p, context = launch_persistent()
    try:
        page = ensure_logged_in(context)

        # 3) Go to Price Lists and create + fill
        open_pricelist(page)
        create_price_list(page, PRICELIST_NAME)
        fill_products_and_prices(page, matched_rows)
        save_pricelist(page)

    finally:
        try:
            for pg in context.pages:
                try:
                    pg.close()
                except Exception:
                    pass
            context.close()
        except Exception as e:
            print(f"⚠️ Error closing context: {e}")
        p.stop()


if __name__ == "__main__":
    if not EMAIL or not PASSWORD:
        print("⚠️ LIGHTSPEED_EMAIL or LIGHTSPEED_PASSWORD missing in .env (will still try to reuse a saved session).")
    main()
