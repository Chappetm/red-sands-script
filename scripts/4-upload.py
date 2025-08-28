from playwright.sync_api import sync_playwright
import time
import random
import os
import pandas as pd
import sys
import re
from dotenv import load_dotenv

class KountaLogin:
    def __init__(self):
        self.email = ""  # Tu email aquí
        self.password = ""  # Tu contraseña aquí
        self.login_url = "https://my.kounta.com/login"  # Ajusta la URL según sea necesario
        
    def random_delay(self, min_seconds=1, max_seconds=3):
        """Simula delays humanos aleatorios"""
        time.sleep(random.uniform(min_seconds, max_seconds))
        
    def setup_browser_context(self, browser):
        """Configura el contexto del navegador para evitar detección"""
        context = browser.new_context(
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
            }
        )
        return context
    
    def human_like_typing(self, page, selector, text):
        """Simula escritura humana con delays variables"""
        element = page.locator(selector)
        element.click()
        self.random_delay(0.5, 1)
        
        for char in text:
            element.type(char)
            time.sleep(random.uniform(0.05, 0.15))
    
    @staticmethod
    def _norm_code(code) -> str:
        """
        Normalize product codes to a canonical form:
        - cast to str
        - strip spaces (incl. NBSP/ZWSP)
        - drop trailing '.0' if came as float-like string
        - remove leading zeros (keep single '0')
        """
        s = str(code).strip()
        # Remove weird spaces
        s = s.replace("\u200b", "").replace("\xa0", " ").strip()
        # If it's like '95725.0' -> '95725'
        if re.fullmatch(r"\d+\.0+", s):
            s = s.split(".", 1)[0]
        # Remove spaces inside code (optional; útil si vienen '95 725')
        s = s.replace(" ", "")
        # Remove leading zeros but keep '0' if that's the whole code
        s = re.sub(r"^0+(?!$)", "", s)
        return s

    def _build_product_lookup(self, lookup_df: pd.DataFrame) -> dict:
        """
        Build a mapping: normalized_code -> Product Name
        Supports multiple codes per row separated by '/', ',', ';' or '|'.
        Warns on duplicate codes mapping to different names.
        """
        code_to_name = {}
        # Ensure expected columns exist
        if "Product Code" not in lookup_df.columns or "Product Name" not in lookup_df.columns:
            raise ValueError("products.xlsx must have columns 'Product Code' and 'Product Name'")

        for _, r in lookup_df.iterrows():
            name = str(r["Product Name"]).strip()
            raw_codes = r["Product Code"]
            if pd.isna(raw_codes):
                continue
            # Split multi-codes: "95109/95725/..." etc.
            parts = re.split(r"[\/,;|]+", str(raw_codes))
            for part in parts:
                code = self._norm_code(part)
                if not code:
                    continue
                if code in code_to_name and code_to_name[code] != name:
                    # Duplicate code pointing to a different name: keep first, log once
                    print(f"⚠️ Código duplicado en products.xlsx: '{code}' ya mapea a '{code_to_name[code]}', ignorando '{name}'")
                    continue
                code_to_name[code] = name

        return code_to_name

    def login(self, profile_name="Bot-Profile", supplier="alm"):
        """Proceso principal de login"""
        with sync_playwright() as p:
            # Crear ruta del perfil personalizado
            profile_path = os.path.join(os.getcwd(), "chrome-profiles", profile_name)
            
            # Crear directorio del perfil si no existe
            os.makedirs(profile_path, exist_ok=True)
            
            # Lanzar navegador con contexto persistente (perfil personalizado)
            context = p.chromium.launch_persistent_context(
                user_data_dir=profile_path,
                headless=False,  # Cambiar a True para modo headless
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
            
            try:
                # Con launch_persistent_context, ya tenemos el contexto directamente
                page = context.new_page()
                
                # Ocultar webdriver
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
                page.goto(self.login_url, wait_until='networkidle')
                self.random_delay(2, 4)

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
                            self.random_delay(2, 3)
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
                        return False
                    
                    print("Rellenando email...")
                    self.human_like_typing(page, email_field, self.email)
                    self.random_delay(1, 2)
                    
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
                    return False
                
                print("Rellenando contraseña...")
                self.human_like_typing(page, password_field, self.password)
                self.random_delay(1, 2)
                
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
                    return False
                
                print("Haciendo click en login...")
                
                # Simular movimiento de mouse antes del click
                button_element = page.locator(login_button)
                box = button_element.bounding_box()
                if box:
                    page.mouse.move(
                        box['x'] + box['width'] / 2, 
                        box['y'] + box['height'] / 2
                    )
                    self.random_delay(0.5, 1)
                
                button_element.click()
                
                # Esperar a que se procese el login
                print("Esperando respuesta...")
                page.wait_for_load_state('networkidle', timeout=10000)
                self.random_delay(3, 5)
                
                # Verificar si el login fue exitoso
                current_url = page.url
                if 'login' not in current_url.lower() or 'dashboard' in current_url.lower():
                    print("✅ Login exitoso!")
                    print(f"URL actual: {current_url}")
                    
                    # Navegar a la página de órdenes
                    print("Navegando a la página de órdenes...")
                    self.random_delay(2, 4)
                    page.goto("https://purchase.kounta.com/purchase#orders", wait_until='networkidle')
                    self.random_delay(3, 5)
                    
                    print(f"✅ Navegado a: {page.url}")

                    # Configure folders and data (no 'processed' folder anymore)
                    excel_invoices_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), "Excel_invoices")
                    input_folder = os.path.join(excel_invoices_folder, supplier)  # Excel_invoices/<supplier>
                    assets_folder = os.path.join(os.path.dirname(os.path.dirname(__file__)), "assets")

                    # Load & normalize product lookup for the selected supplier
                    try:
                        sheet_name = supplier.upper()  # ALM, COKE, CUB, LION
                        lookup_df = pd.read_excel(
                            os.path.join(assets_folder, "products.xlsx"),
                            sheet_name=sheet_name,
                            dtype={"Product Code": str, "Product Name": str}
                        )
                        product_lookup = self._build_product_lookup(lookup_df)
                        print(f"📋 {sheet_name}: {len(lookup_df)} rows → {len(product_lookup)} unique codes (expanded & normalized)")
                    except Exception as e:
                        print(f"❌ Failed to load products.xlsx sheet '{sheet_name}': {e}")
                        return False

                    # Supplier label lookup
                    supplier_lookup = {
                        "alm": "ALM",
                        "cub": "CUB", 
                        "lion": "LION",
                        "coke": "COKE"
                    }

                    print(f"📦 Processing supplier: {supplier_lookup[supplier]}")

                    # Process Excel files (no moving to 'processed' folder here)
                    self.process_excel_files(page, input_folder, supplier, product_lookup, supplier_lookup)


                    return True
                
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
                    
                    return False
                    
            except Exception as e:
                print(f"❌ Error durante el proceso: {str(e)}")
                return False
            
            finally:
                # Always close the browser so the process exits and Streamlit can detect completion
                try:
                    # Close any open pages first (optional but tidy)
                    try:
                        for p in context.pages:
                            try:
                                p.close()
                            except Exception:
                                pass
                    except Exception:
                        pass

                    context.close()
                    # tiny wait to ensure underlying process exits cleanly
                    time.sleep(0.3)
                    print("🔚 Browser context closed.")
                except Exception as e:
                    print(f"⚠️ Failed to close browser context: {e}")

    def process_excel_files(self, page, input_folder, supplier, product_lookup, supplier_lookup):
        """
        Process Excel files and create the orders.
        NOTE: This function no longer moves files to a 'processed' folder.
            The cleanup is managed externally (e.g., by main_gui.py after successful upload).
        """

        # Gather Excel files (xlsx/xls if you really need .xls; recommend .xlsx only)
        excel_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.xlsx', '.xls'))]

        if not excel_files:
            print("No Excel files found to process.")
            return

        for excel_file in excel_files:
            print(f"\n📄 Processing: {excel_file}")
            excel_path = os.path.join(input_folder, excel_file)

            # Read the Excel
            df = pd.read_excel(excel_path, engine="openpyxl")

            # ADMIN FEE logic
            admin_fee_value = None
            if supplier == "alm" and "Admin fee" in df.columns:
                try:
                    value = float(df.at[0, "Admin fee"])
                    if value > 0:
                        admin_fee_value = round(value, 2)
                except Exception as e:
                    print(f"⚠️ Could not read Admin Fee: {e}")

            # Normalize order codes exactly like the lookup
            df["Product Code"] = df["Product Code"].apply(self._norm_code)
            missing = set(df["Product Code"]) - set(product_lookup.keys())
            if missing:
                print(f"⛔ Missing product codes in lookup: {missing}")
                continue

            # Process the single order
            success = self.process_single_order(page, df, supplier, product_lookup, supplier_lookup, admin_fee_value)

            # IMPORTANT: do NOT move/delete the file here.
            if success:
                print("✅ Order processed successfully (no file move; cleanup handled by caller).")

            self.random_delay(3, 5)

    def process_single_order(self, page, df, supplier, product_lookup, supplier_lookup, admin_fee_value):
        """Procesa una orden individual"""
        try:
            po_number_raw = str(df.iloc[0].get("PO Number", "")).strip()
            is_new_invoice = not po_number_raw or po_number_raw == "PO00000000"
            
            if is_new_invoice:
                return self.create_new_order(page, df, supplier, supplier_lookup, product_lookup, admin_fee_value)
            else:
                return self.edit_existing_order(page, df, po_number_raw, product_lookup, admin_fee_value, supplier)
                
        except Exception as e:
            print(f"❌ Error procesando orden: {e}")
            return False

    def create_new_order(self, page, df, supplier, supplier_lookup, product_lookup, admin_fee_value):
        """
        Create a new purchase order in Lightspeed Kounta.

        Flow:
        1) Click "New order"
        2) Select supplier row by exact supplier name
        3) Wait for the PO number input to appear and read back the new PO

        Assumptions:
        - You're already on https://purchase.kounta.com/purchase#orders
        - The supplier table renders a <tr> where a <div> cell contains the supplier name
        - The PO input has selector: input.readonly[data-chaminputid='textInput']

        Returns:
        True if order creation flow reaches the PO number,
        False otherwise (logs the reason).
        """
        try:
            print("🆕 No PO Number detected. Creating a new order...")

            # 1) Click "New order" button (prefer role-based; fallback to XPath contains(text()))
            # Try role-based first for resilience
            new_order_clicked = False
            try:
                btn = page.get_by_role("button", name=lambda n: n and "new order" in n.lower())
                btn.wait_for(state="visible", timeout=5000)
                btn.click()
                new_order_clicked = True
            except Exception:
                # Fallback to XPath
                try:
                    page.locator("//button[contains(., 'New order')]").first.wait_for(state="visible", timeout=5000)
                    page.locator("//button[contains(., 'New order')]").first.click()
                    new_order_clicked = True
                except Exception as e:
                    print(f"❌ Could not locate/click 'New order' button: {e}")
                    return False

            if not new_order_clicked:
                print("❌ 'New order' button was not clicked.")
                return False

            self.random_delay(0.6, 1.1)

            # 2) Select supplier row (robusto con coincidencia parcial/insensible a mayúsculas)
            import re

            supplier_name = supplier_lookup[supplier]  # puede ser "LION", "ALM", etc.

            # 2) Preparar varios localizadores alternativos
            #    A: por rol (fila) con nombre accesible que contenga supplier_name (case-insensitive)
            #    B: XPath flexible con translate() para contains insensible a mayúsculas
            lower_name = supplier_name.lower()
            xpath_ci = (
                "//tr[@role='button']"
                "[.//td[@name='name']//div"
                f"[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), '{lower_name}')]]"
            )

            candidates = [
                page.get_by_role("row", name=re.compile(rf"{re.escape(supplier_name)}", re.I)).first,
                page.locator(xpath_ci).first,
            ]

            # 3) Intentar encontrar una fila visible con cualquiera de los candidatos
            row = None
            for loc in candidates:
                try:
                    loc.wait_for(state="visible", timeout=15000)
                    row = loc
                    break
                except Exception:
                    continue

            if row is None:
                print(f"❌ Could not select supplier row '{supplier_name}': not found/visible")
                return False

            # 4) Desplazar a vista (las tablas virtualizadas requieren esto)
            try:
                row.scroll_into_view_if_needed()
            except Exception:
                pass

            # 5) Clic sobre la celda de nombre (más fiable que clicar el <tr>)
            try:
                cell = row.locator("td[name='name'] div").first
                cell.wait_for(state="visible", timeout=5000)

                # Mover el mouse ayuda en UIs con hover/focus
                try:
                    box = cell.bounding_box()
                    if box:
                        page.mouse.move(box["x"] + box["width"] / 2, box["y"] + box["height"] / 2)
                        self.random_delay(0.2, 0.5)
                except Exception:
                    pass

                try:
                    cell.click()
                except Exception:
                    cell.click(force=True)

                print(f"📦 Supplier selected: {supplier_name}")
            except Exception as e:
                print(f"❌ Could not select supplier row '{supplier_name}': {e}")
                return False

            self.random_delay(0.6, 1.1)


            # 3) Wait for new PO number input and read it
            # Keep two attempts: visible first; if not, presence-only then read attribute.
            po_selector = "input.readonly[data-chaminputid='textInput']"
            new_po_number = None
            try:
                # Prefer visible element so that value is likely populated
                page.wait_for_selector(po_selector, state="visible", timeout=12000)
                po_input = page.locator(po_selector).first
                # Give a short delay to allow value to populate
                self.random_delay(0.3, 0.7)
                new_po_number = po_input.get_attribute("value")
            except Exception:
                # Fallback: presence only
                try:
                    page.wait_for_selector(po_selector, state="attached", timeout=12000)
                    po_input = page.locator(po_selector).first
                    self.random_delay(0.3, 0.7)
                    new_po_number = po_input.get_attribute("value")
                except Exception as e:
                    print(f"⚠️ Unable to obtain new PO Number (no input found): {e}")

            if new_po_number:
                print(f"🆕 New invoice created with PO: PO{new_po_number}")
            else:
                print("⚠️ New PO Number input found but value was empty or unreadable.")

            # If you need to do something here with admin_fee_value later, keep the variable available.
            # For now, we just return True because the creation step completed.
            return self.add_products_and_finalize(page, df, supplier, product_lookup, admin_fee_value)

        except Exception as e:
            print(f"❌ Error while creating new order: {e}")
            return False
    
    def edit_existing_order(self, page, df, po_number_raw, product_lookup, admin_fee_value, supplier):
        """
        Open an existing purchase order and (optionally) clear any pre-populated line item.

        Flow (ported from Selenium logic):
        1) Normalize PO number (strip leading 'PO' and leading zeros via int cast)
        2) Go to "Purchase orders" list
        3) Click the order whose text contains the normalized PO number
        4) Try to remove the pre-loaded product line, if present

        Assumptions:
        - You're already authenticated and on a purchase-related page
        - The "Purchase orders" navigation is available as a link or visible text
        - Each PO entry shows the numeric portion of the PO (e.g., '123456')
        - A pre-loaded line item (if any) is a div with class 'lineInfo'
        - The "Remove" action is a button containing a <span> with text 'Remove'
        """
        try:
            # 1) Normalize PO number (remove 'PO' and cast to int to strip leading zeros)
            try:
                clean_po_number = po_number_raw.upper().replace("PO", "").strip()
                po_number = str(int(clean_po_number))
            except Exception:
                # Fallback: if casting fails, keep the cleaned string (last resort)
                po_number = clean_po_number
            print(f"🔎 Target PO number: {po_number}")

            # 2) Navigate to "Purchase orders" list
            # Prefer role-based locator; fallback to text search
            navigated = False
            try:
                link = page.get_by_role("link", name=lambda n: n and "purchase orders" in n.lower())
                link.wait_for(state="visible", timeout=6000)
                link.click()
                navigated = True
            except Exception:
                try:
                    page.get_by_text("Purchase orders", exact=False).first.wait_for(state="visible", timeout=6000)
                    page.get_by_text("Purchase orders", exact=False).first.click()
                    navigated = True
                except Exception as e:
                    print(f"❌ Could not navigate to 'Purchase orders': {e}")
                    return False

            self.random_delay(0.4, 0.8)

            # 3) Find the order row by exact PO and CHECK 'Total(inc.)' BEFORE opening
            try:
                # Locate the row: <td name="Order no."><div>po_number</div> ... ancestor <tr role="button">
                row_locator = page.locator(
                    f"//td[@name='Order no.'][div[normalize-space(text())='{po_number}']]/ancestor::tr[@role='button'][1]"
                )
                row_locator.wait_for(state="visible", timeout=10000)

                # Scroll into view to avoid header/overlay issues
                try:
                    row_locator.scroll_into_view_if_needed(timeout=4000)
                    self.random_delay(0.2, 0.5)
                except Exception:
                    pass

                # --- Read 'Total(inc.)' amount from the same row BEFORE clicking ---
                try:
                    # Try CSS first (safe for attribute selectors), then fallback to XPath with explicit prefix
                    amount_node = row_locator.locator('td[name="Total(inc.)"] div').first
                    try:
                        amount_node.wait_for(state="visible", timeout=2000)
                    except Exception:
                        amount_node = row_locator.locator("xpath=.//td[@name='Total(inc.)']//div").first
                        amount_node.wait_for(state="visible", timeout=2000)

                    raw_amount = (amount_node.inner_text() or amount_node.text_content() or "").strip()


                    # Normalize "$11,228.26" -> "11228.26" (handle commas/locale gracefully)
                    cleaned = re.sub(r"[^0-9,.\-]", "", raw_amount)
                    if cleaned.count(",") and cleaned.count("."):
                        # Likely comma as thousands sep -> drop commas
                        cleaned = cleaned.replace(",", "")
                    elif cleaned.count(",") and not cleaned.count("."):
                        # Comma as decimal sep -> swap to dot
                        cleaned = cleaned.replace(",", ".")
                    total_inc = float(cleaned) if cleaned else 0.0
                    
                    if total_inc >= 1000.0:
                        print(f"🛑 Invoice PO {po_number} already uploaded (Total inc. = ${total_inc:,.2f}). Skipping.")
                        return True  # short-circuit: treat as done; upstream sees success and stops work
                except Exception as e:
                    print(f"ℹ️ Could not read 'Total(inc.)' before opening PO {po_number}: {e}. Proceeding...")

                # Not above threshold -> open the row
                try:
                    row_locator.click()
                except Exception:
                    row_locator.click(force=True)

                print(f"📄 Opened PO {po_number}")

            except Exception as e:
                print(f"❌ Could not locate/open PO '{po_number}': {e}")
                return False



            # 4) Remove pre-loaded product line if present
            try:
                # The line item container as per your Selenium: //div[contains(@class, 'lineInfo')]
                existing_product = page.locator("//div[contains(@class, 'lineInfo')]").first
                existing_product.wait_for(state="visible", timeout=5000)

                # Scroll into view to avoid header/overlay issues
                try:
                    existing_product.scroll_into_view_if_needed(timeout=3000)
                    self.random_delay(0.2, 0.5)
                except Exception:
                    pass

                # Click the line to focus/select it
                try:
                    existing_product.click()
                except Exception:
                    existing_product.click(force=True)

                self.random_delay(0.4, 0.8)

                # Locate the "Remove" button by span text and click its ancestor button
                remove_btn = page.locator("//span[normalize-space(text())='Remove']/ancestor::button").first
                remove_btn.wait_for(state="visible", timeout=5000)
                try:
                    remove_btn.click()
                except Exception:
                    remove_btn.click(force=True)

                print("🗑️ Pre-loaded line item removed.")
            except Exception as e:
                print(f"ℹ️ No pre-loaded line removed (not found or not clickable): {e}")

            # If you need to apply admin_fee_value or proceed adding items, do it after this point.
            return self.add_products_and_finalize(page, df, supplier, product_lookup, admin_fee_value)

        except Exception as e:
            print(f"❌ Error while editing existing order: {e}")
            return False

    def add_products_and_finalize(self, page, df, supplier, product_lookup, admin_fee_value):
        """
        Add all products from df to the current PO, optionally adjust total price,
        add Admin Fee (ALM), and then click 'Review Order' and go back to the orders list.

        Steps:
        - Click search icon
        - For each product row (Product Code, Order Qty, Total Cost):
            * Search by product name (from product_lookup)
            * Click the product card 'quantity' times to add lines
            * Open the product line and set 'Enter total price'
                (warn if invoice cost >> current cost), apply changes
        - If ALM and admin_fee_value present:
            * Add 'Administration Fee', set its total price, apply changes
        - Click 'Review Order'
        - Navigate back to https://purchase.kounta.com/purchase#orders
        """
        try:
            # --- Click the search icon (fallbacks included) ---
            search_clicked = False
            try:
                # CSS from Selenium
                search_btn = page.locator("button.IconButton__IconButtonStyle-sc-17z8q7c-0").first
                search_btn.wait_for(state="visible", timeout=5000)
                search_btn.click()
                search_clicked = True
            except Exception:
                # Fallback: try a generic "Search" button by role/text
                try:
                    btn = page.get_by_role("button", name=lambda n: n and "search" in n.lower())
                    btn.wait_for(state="visible", timeout=5000)
                    btn.click()
                    search_clicked = True
                except Exception as e:
                    print(f"❌ Could not click search icon: {e}")
                    return False

            self.random_delay(0.3, 0.7)

            # Will ensure the Admin Fee is only added once
            admin_fee_added = False

            # --- Add products from the dataframe ---
            for _, row in df.iterrows():
                product_code = str(row["Product Code"]).strip()
                quantity = int(row["Order Qty"])
                total_cost = float(row["Total Cost"])
                product_name = product_lookup.get(product_code)

                if not product_name:
                    print(f"⚠️ Producto no encontrado para código: {product_code}")
                    continue

                print(f"🛒 Agregando: {product_name} ({quantity}x)")

                try:
                    # Focus the search input and type the product name, then Enter
                    search_input = page.locator("input[placeholder='Search for products']").first
                    search_input.wait_for(state="visible", timeout=8000)
                    # Clear any previous content
                    search_input.fill("")
                    self.random_delay(0.2, 0.5)
                    search_input.fill(product_name)
                    self.random_delay(0.2, 0.5)
                    search_input.press("Enter")
                    self.random_delay(0.5, 0.9)

                    # Try to click the exact match in the results, quantity times
                    # Avoid relying on hashed class names: match by exact text
                    product_option = page.locator(f'//span[normalize-space(text())="{product_name}"]').first
                    product_option.wait_for(state="visible", timeout=8000)

                    for _ in range(quantity):
                        try:
                            product_option.click()
                            self.random_delay(0.1, 0.25)
                        except Exception:
                            product_option.click(force=True)
                            self.random_delay(0.1, 0.25)

                    print(f"✅ Producto agregado {quantity}x: {product_name}")

                    # Open the just-added product line in the PO details to edit total cost
                    product_line_xpath = f"//div[contains(@class, 'lineInfo')]//span[normalize-space(text())=\"{product_name}\"]"
                    try:
                        product_line = page.locator(product_line_xpath).first
                        product_line.wait_for(state="visible", timeout=8000)
                        try:
                            product_line.scroll_into_view_if_needed(timeout=3000)
                        except Exception:
                            pass
                        try:
                            product_line.click()
                        except Exception:
                            product_line.click(force=True)
                        self.random_delay(0.3, 0.7)

                        # 'Enter total price' input
                        total_input = page.locator("input[placeholder='Enter total price']").first
                        total_input.wait_for(state="visible", timeout=8000)

                        # Read current cost in the input (if any)
                        current_cost = None
                        try:
                            # Prefer input_value() for <input> fields
                            val = total_input.input_value()
                            if val is not None and val.strip() != "":
                                current_cost = float(val)
                        except Exception:
                            # Fallback to attribute
                            try:
                                val = total_input.get_attribute("value")
                                if val:
                                    current_cost = float(val)
                            except Exception:
                                current_cost = None

                        # Warn if invoice cost >> current cost
                        if current_cost is not None and (total_cost - current_cost) > 1.00:
                            print(f"🚨 Atención: '{product_name}' tiene costo invoice ${total_cost}, mayor que el actual (${current_cost})")

                        # --- Adjust the total for CUB/LION/COKE (+10%), except specific water SKUs (no tax) ---
                        # Define the no-tax exceptions by exact product name
                        NO_TAX_PRODUCTS = {"Mt FRANKLIN 600ml S1", "Mt FRANKLIN 1.5L S1"}

                        adjusted_cost = float(total_cost)

                        # Apply +10% only if supplier is CUB/LION/COKE AND product is not in the no-tax list
                        if supplier in ["cub", "lion", "coke"] and product_name not in NO_TAX_PRODUCTS:
                            adjusted_cost = round(adjusted_cost * 1.10, 2)


                        # Open keypad, enter price, confirm and apply
                        ok = self.enter_price_via_keypad(page, total_input, adjusted_cost)
                        if not ok:
                            print(f"⚠️ Keypad flow returned False for '{product_name}'. Continuing.")


                    except Exception as e:
                        print(f"⚠️ No se pudo editar el Total Cost de '{product_name}': {e}")

                except Exception as e:
                    print(f"❌ Error con '{product_name}': {e}")

            # --- Admin Fee (ALM) ---
            if supplier == "alm" and admin_fee_value:
                product_name = "Administration Fee"
                total_cost = float(admin_fee_value)
                print(f"➕ Agregando Admin Fee: {product_name} ${total_cost}")

                try:
                    # 1) Search field (robust: placeholder first, fallback to raw input)
                    search_input = page.get_by_placeholder("Search for products").first
                    try:
                        search_input.wait_for(state="visible", timeout=8000)
                    except Exception:
                        search_input = page.locator("input[placeholder='Search for products']").first
                        search_input.wait_for(state="visible", timeout=8000)

                    # Hard clear (select-all + backspace), then type and search
                    for _ in range(2):
                        try:
                            search_input.press("ControlOrMeta+a")
                            search_input.press("Backspace")
                        except Exception:
                            pass
                    self.random_delay(0.2, 0.5)

                    search_input.type(product_name, delay=random.uniform(20, 40))
                    self.random_delay(0.2, 0.5)
                    search_input.press("Enter")
                    self.random_delay(0.5, 0.9)

                    # 2) Click the exact product option by visible text
                    admin_option = page.locator(f"//span[normalize-space(text())='{product_name}']").first
                    admin_option.wait_for(state="visible", timeout=8000)
                    try:
                        admin_option.click()
                    except Exception:
                        admin_option.click(force=True)
                    print("✅ Admin Fee agregado.")

                    # 3) Open the line item panel for this product
                    product_line_xpath = f"//div[contains(@class, 'lineInfo')]//span[normalize-space(text())=\"{product_name}\"]"
                    product_line = page.locator(product_line_xpath).first
                    product_line.wait_for(state="visible", timeout=8000)
                    try:
                        product_line.scroll_into_view_if_needed(timeout=3000)
                    except Exception:
                        pass
                    try:
                        product_line.click()
                    except Exception:
                        product_line.click(force=True)
                    self.random_delay(0.3, 0.7)

                    # 4) Use the SAME helper we already use for products:
                    #    it opens the control/keypad (or focuses the input), types the amount,
                    #    confirms (OK/Enter), and clicks "Apply Changes".
                    total_trigger = None
                    candidates = [
                        page.get_by_placeholder("Enter total price"),                           # preferred (placeholder helper)
                        page.locator("//input[@placeholder='Enter total price']"),             # real <input>
                        page.locator("//button[normalize-space()='Enter total price']"),       # keypad trigger button (if present)
                        page.locator("xpath=//*[contains(@aria-label,'Enter total price')]"),  # any aria-labeled control
                    ]
                    for cand in candidates:
                        try:
                            if cand.count() > 0:
                                total_trigger = cand.first
                                try:
                                    total_trigger.wait_for(state="visible", timeout=4000)
                                except Exception:
                                    pass
                                break
                        except Exception:
                            continue

                    if not total_trigger:
                        print(f"❌ Could not locate the total price control for '{product_name}'")
                    else:
                        ok = self.enter_price_via_keypad(page, total_trigger, float(total_cost))
                        if not ok:
                            print(f"⚠️ Keypad flow returned False for '{product_name}'")

                except Exception as e:
                    print(f"❌ Error al agregar Admin Fee: {e}")

            # --- Review Order ---
            try:
                review_order_button = page.locator("//button[contains(., 'Review Order')]").first
                review_order_button.wait_for(state="visible", timeout=10000)
                try:
                    review_order_button.click()
                except Exception:
                    review_order_button.click(force=True)
                self.random_delay(0.4, 0.9)
            except Exception as e:
                print(f"⚠️ No se pudo clickear 'Review Order': {e}")

            # --- Back to orders list ---
            try:
                page.goto("https://purchase.kounta.com/purchase#orders", wait_until="networkidle")
                self.random_delay(0.5, 1.0)
            except Exception as e:
                print(f"⚠️ No se pudo volver a la lista de órdenes: {e}")

            return True

        except Exception as e:
            print(f"❌ Error en add_products_and_finalize: {e}")
            return False

    def enter_price_via_keypad(self, page, total_input, adjusted_cost: float):
        """
        Open the numeric keypad, clear any previous value, type the new amount, confirm (OK/Enter),
        and click 'Apply Changes'. This handles React-controlled overlays where there is no real <input>.

        Parameters:
        - total_input: locator that OPENS the keypad when clicked
        - adjusted_cost: float value to set (already adjusted, e.g. +10% for CUB/LION/COKE)

        Strategy:
        1) Click 'total_input' to open keypad
        2) Wait for keypad to appear (heuristic: #enter button or a numeric button)
        3) Clear previous value (try 'Clear/AC/C/⌫' buttons; fallback: Backspace spam)
        4) Type digits with the real keyboard (preferred)  OR click digit buttons as fallback
        5) Confirm with OK (#enter) or press Enter
        6) Click 'Apply Changes'
        """
        import re

        # Format amount with 2 decimals; decimal will be mapped to the keypad's available symbol ('.' or ',')
        raw_text = f"{adjusted_cost:.2f}"

        # 1) Open keypad
        try:
            total_input.wait_for(state="visible", timeout=8000)
            try:
                total_input.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            total_input.click()
            self.random_delay(0.15, 0.3)
        except Exception as e:
            print(f"❌ Could not open keypad (click on total_input): {e}")
            return False

        # 2) Wait for keypad to be present (either '#enter' or some digit button)
        keypad_ready = False
        try:
            page.locator("#enter").first.wait_for(state="visible", timeout=2000)
            keypad_ready = True
        except Exception:
            # Try to detect any digit button
            try:
                page.get_by_role("button", name=re.compile(r"^[0-9]$")).first.wait_for(state="visible", timeout=2000)
                keypad_ready = True
            except Exception:
                pass

        if not keypad_ready:
            print("⚠️ Keypad did not show up clearly; proceeding anyway.")
        self.random_delay(0.1, 0.2)

        # 3) Try to CLEAR previous value
        cleared = False
        for label in [r"^clear$", r"^ac$", r"^c$", r"^clr$", r"^reset$", r"^⌫$", r"^del$", r"^delete$"]:
            try:
                btn = page.get_by_role("button", name=re.compile(label, re.I)).first
                if btn.count() > 0:
                    try:
                        btn.click()
                    except Exception:
                        btn.click(force=True)
                    cleared = True
                    break
            except Exception:
                continue

        if not cleared:
            # Backspace several times as a fallback (if the keypad listens to key events)
            for _ in range(8):
                page.keyboard.press("Backspace")
                self.random_delay(0.03, 0.06)

        # 4) Decide which decimal symbol the keypad supports ('.' preferred, fallback ',')
        decimal_symbol = "."
        try:
            dot_btn = page.get_by_role("button", name=re.compile(r"^\.$")).first
            if dot_btn.count() == 0:
                comma_btn = page.get_by_role("button", name=re.compile(r"^,$")).first
                if comma_btn.count() > 0:
                    decimal_symbol = ","
        except Exception:
            # If we can't inspect buttons, assume '.'
            decimal_symbol = "."

        value_text = raw_text.replace(".", decimal_symbol)

        # Try preferred path: TYPE with the physical keyboard (often works on these overlays)
        try:
            page.keyboard.type(value_text, delay=random.uniform(25, 45))
            self.random_delay(0.15, 0.3)
        except Exception as e:
            print(f"⚠️ Keyboard type failed: {e}")

        # Fallback: CLICK digits on keypad (only if we can see digit buttons)
        # We only do this if we suspect typing didn't go through (we can't verify value reliably without DOM,
        # so we conservatively click digits *in addition* if typing was questionable).
        try:
            # Heuristic: if we didn't see #enter earlier, we try clicking buttons to be safe
            if not keypad_ready:
                for ch in value_text:
                    if ch in "0123456789":
                        btn = page.get_by_role("button", name=re.compile(f"^{ch}$")).first
                    elif ch in ".,":  # decimal
                        btn = page.get_by_role("button", name=re.compile(r"^(\.|,)$")).first
                    else:
                        continue

                    btn.wait_for(state="visible", timeout=2000)
                    try:
                        btn.click()
                    except Exception:
                        btn.click(force=True)
                    self.random_delay(0.05, 0.12)
        except Exception as e:
            print(f"⚠️ Clicking keypad digits failed: {e}")

        # 5) Confirm: OK (#enter) or press Enter
        confirmed = False
        try:
            ok_button = page.locator("#enter").first
            if ok_button.count() > 0:
                try:
                    ok_button.click()
                except Exception:
                    ok_button.click(force=True)
                confirmed = True
        except Exception:
            pass

        if not confirmed:
            try:
                # Try a generic OK/Enter/Done button
                ok_generic = page.get_by_role("button", name=re.compile(r"^(ok|enter|done|apply)$", re.I)).first
                if ok_generic.count() > 0:
                    try:
                        ok_generic.click()
                    except Exception:
                        ok_generic.click(force=True)
                    confirmed = True
            except Exception:
                pass

        if not confirmed:
            # Last resort: send Enter key
            page.keyboard.press("Enter")

        self.random_delay(0.2, 0.4)

        # 6) Click 'Apply Changes'
        try:
            apply_btn = page.locator("//button[contains(., 'Apply Changes')]").first
            apply_btn.wait_for(state="visible", timeout=6000)
            try:
                apply_btn.click()
            except Exception:
                apply_btn.click(force=True)
            self.random_delay(0.25, 0.5)
        except Exception as e:
            print(f"⚠️ Could not click 'Apply Changes': {e}")

        return True
 
# Uso del script
if __name__ == "__main__":
    # Verificar argumentos de línea de comandos
    if len(sys.argv) != 2:
        print("❌ Uso: python 4-upload.py <supplier>")
        print("Suppliers disponibles: alm, coke, cub, lion")
        exit(1)
    
    supplier = sys.argv[1].lower()
    valid_suppliers = ["alm", "coke", "cub", "lion"]
    
    if supplier not in valid_suppliers:
        print(f"❌ Supplier '{supplier}' no válido. Usa uno de: {', '.join(valid_suppliers)}")
        exit(1)
    
    # Cargar variables de entorno desde el archivo .env en la carpeta padre
    load_dotenv(dotenv_path=os.path.join(os.path.dirname(os.path.dirname(__file__)), '.env'))
    
    # Configurar credenciales desde variables de entorno
    login_bot = KountaLogin()
    login_bot.email = os.getenv('LIGHTSPEED_EMAIL')
    login_bot.password = os.getenv('LIGHTSPEED_PASSWORD')
    
    # Verificar que las credenciales se cargaron correctamente
    if not login_bot.email or not login_bot.password:
        print("❌ Error: No se pudieron cargar las credenciales desde el archivo .env")
        print("Asegúrate de que el archivo .env contiene:")
        print("EMAIL=tu-email@ejemplo.com")
        print("PASSWORD=tu-contraseña")
        exit(1)
    
    print(f"🔐 Usando email: {login_bot.email}")
    print(f"📦 Supplier seleccionado: {supplier.upper()}")
    
    
    # Ejecutar login con el supplier elegido en CLI
    success = login_bot.login("Bot-Profile", supplier)
    
    if success:
        print("🎉 Proceso completado exitosamente")
    else:
        print("💥 El proceso falló")