import fitz
import re
import os
import pandas as pd


def detect_supplier(product_codes, product_db):
    matches = {supplier: 0 for supplier in product_db.keys()}
    for code in product_codes:
        for supplier, codes_set in product_db.items():
            if code in codes_set:
                matches[supplier] += 1
    best_match = max(matches, key=matches.get)
    if matches[best_match] > 0:
        return best_match
    return None

def extract_lion_invoice_data(pdf_path):

    # Leer líneas del PDF
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        lines.extend(page.get_text().split('\n'))

    # Buscar PO number
    po_number = next((line.strip() for line in lines if re.search(r"PO\d{8}", line)), "PO00000000")

    # Cortar bloques al detectar secuencia "CARRIER LOAD TOTAL"
    bloques = []
    actual = []
    carrier_sequence = ["CARRIER", "LOAD", "TOTAL"]
    carrier_index = 0

    for line in lines:
        line_upper = line.strip().upper()

        if re.match(r"^\d{7} ", line_upper):
            if actual:
                bloques.append(actual)
            actual = [line]
            carrier_index = 0
        elif actual:
            actual.append(line)
            if carrier_sequence[carrier_index] in line_upper:
                carrier_index += 1
                if carrier_index == len(carrier_sequence):
                    bloques.append(actual)
                    break
            else:
                carrier_index = 0

    if actual and carrier_index < len(carrier_sequence):
        bloques.append(actual)

    # Extraer productos
    productos = []

    for bloque in bloques:
        try:
            product_code = bloque[0].split()[0]
            # Buscar QTY inmediatamente después de "CAR" (misma línea o la siguiente),
            # sin limitar a 99, y cortando antes de los valores decimales (precios).
            qty = None
            for idx, line in enumerate(bloque):
                m = re.search(r'\bCAR\b(?:\s+|$)', line)
                if m:
                    # 1) Intentar en la MISMA línea, después de "CAR"
                    tail_tokens = line[m.end():].strip().split()
                    found = None
                    for tok in tail_tokens:
                        if tok.isdigit():
                            found = int(tok)
                            break
                        # si aparece un decimal, ya pasamos la QTY (comienzan precios)
                        if re.match(r'^\d+\.\d{2}$', tok):
                            break

                    # 2) Si no la encontramos, probar en la línea SIGUIENTE
                    if found is None and idx + 1 < len(bloque):
                        next_tokens = bloque[idx + 1].strip().split()
                        for tok in next_tokens:
                            if tok.isdigit():
                                found = int(tok)
                                break
                            if re.match(r'^\d+\.\d{2}$', tok):
                                break

                    if found is not None and found > 0:
                        qty = found
                    break

            # Fallback: primer entero razonable del bloque antes de toparnos con un decimal (precio)
            if qty is None:
                for line in bloque:
                    for tok in line.strip().split():
                        if tok.isdigit():
                            qty = int(tok)
                            break
                        if re.match(r'^\d+\.\d{2}$', tok):
                            # Llegaron precios, dejamos de buscar qty en esta línea
                            break
                    if qty is not None:
                        break


            # Buscar último decimal (LINE VALUE) dentro del bloque
            decimales = []
            for line in bloque:
                decimales += [float(x) for x in line.strip().split() if re.match(r"^\d+\.\d{2}$", x)]

            if qty is not None and decimales:
                line_value = decimales[-1]
                productos.append([po_number, product_code, qty, round(line_value, 2)])

        except Exception:
            continue

    df = pd.DataFrame(productos, columns=["PO Number", "Product Code", "Order Qty", "Total Cost"])
    return df

def extract_cub_invoice_data(pdf_path):
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        lines.extend(page.get_text().split('\n'))

    po_number = next((match.group() for line in lines if (match := re.search(r"PO\d{8}", line))), "PO00000000")

    productos = []
    for i in range(len(lines) - 10):
        if re.match(r"^\d{5,6}$", lines[i].strip()):
            try:
                product_code = lines[i].strip()
                qty = float(lines[i+1].strip())
                for j in range(i+2, i+15):
                    if lines[j].strip() == "Y":
                        total_line = lines[j+1].strip()
                        if re.match(r"^\d+\.\d{2,3}$", total_line):
                            total_cost = float(total_line)
                            productos.append([po_number, product_code, qty, round(total_cost, 2)])
                        break
            except:
                continue

    return pd.DataFrame(productos, columns=["PO Number", "Product Code", "Order Qty", "Total Cost"])

def extract_alm_invoice_data(pdf_path):
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        lines.extend(page.get_text().split('\n'))

    # Detectar PO number (formato: PO12345678)
    po_number_match = next((line for line in lines if re.match(r"PO\d{8}", line)), "PO00000000")
    po_number = po_number_match.strip()
    
    # -------------- TODO: SI NO HAY PO, QUE CREE UNA NUEVA INVOICE EN LIGHTSPEED --------------------

    # Agrupar líneas que parezcan bloques de productos
    productos_bloques = []
    bloque = []
    for line in lines:
        if re.match(r'^.+\d{2,4}(ML|L|GM)$', line):  # empieza con descripción + unidad
            if bloque:
                productos_bloques.append(" ".join(bloque))
            bloque = [line]
        elif re.search(r'\d{5,6}$', line.strip()):  # termina en product code
            bloque.append(line)
            productos_bloques.append(" ".join(bloque))
            bloque = []
        elif re.search(r'\d+\.\d{2}', line) or line.strip().isdigit():
            bloque.append(line)

    # Extraer datos de cada bloque
    productos = []
    for bloque in productos_bloques:
        partes = bloque.strip().split()
        try:
            product_code = next(p for p in reversed(partes) if re.match(r"^\d{5,6}$", p))
            decimales = [float(p) for p in partes if re.match(r"^\d+\.\d{2}$", p)]

            # Detectar el order qty como el último entero (1-99) justo antes del primer decimal
            order_qty = None
            for i in range(len(partes) - 1):
                if re.match(r"^\d+\.\d{2}$", partes[i+1]) and partes[i].isdigit():
                    val = int(partes[i])
                    if 1 <= val <= 99:
                        order_qty = val
                        break

            if order_qty is not None and len(decimales) >= 1:
                total_cost = decimales[-3]  # primer decimal como Total Cost
                productos.append([po_number, product_code, order_qty, round(total_cost, 2)])

        except:
            continue

    # Buscar gasto administrativo
    admin_fee_total = 0.0

    for i, line in enumerate(lines):
        if "SHRINK WRAP" in line.upper() or "ADMINISTRATION FEE" in line.upper():
            # Revisar hasta 3 líneas siguientes
            for j in range(i + 1, min(i + 4, len(lines))):
                next_line = lines[j]
                matches = re.findall(r"\d*\.\d{2}", next_line)
                if matches:
                    admin_fee_total += float(matches[0])
                    break  # cortar cuando encuentra el primer valor válido




    df = pd.DataFrame(productos, columns=["PO Number", "Product Code", "Order Qty", "Total Cost"])

    # Insertar columna Admin fee con valor solo en la primera fila
    df.insert(4, "Admin fee", "")
    if admin_fee_total > 0 and not df.empty:
        df.at[0, "Admin fee"] = round(admin_fee_total, 2)

    return df

def extract_coke_invoice_data(pdf_path):
    doc = fitz.open(pdf_path)
    lines = []
    for page in doc:
        lines.extend(page.get_text().split('\n'))

    match = next((re.search(r"PO\d{8}", line) for line in lines if "PO" in line), None)
    po_number = match.group() if match else "PO00000000"


    productos = []
    for i in range(len(lines) - 6):
        if (
            lines[i].strip().isdigit() and
            re.match(r"\d{6}", lines[i+2].strip()) and
            re.match(r"\d+\.\d{2}", lines[i+6].strip())
        ):
            try:
                qty = int(lines[i].strip())
                product_code = lines[i+2].strip()
                total_cost = float(lines[i+6].strip())
                productos.append([po_number, product_code, qty, round(total_cost, 2)])
            except:
                continue

    return pd.DataFrame(productos, columns=["PO Number", "Product Code", "Order Qty", "Total Cost"])


if __name__ == "__main__":
    base_dir = os.path.dirname(__file__)
    input_base = os.path.join(base_dir, "../PDF_invoices")
    output_base = os.path.join(base_dir, "../Excel_invoices")

    # 1. Cargar productos por hoja desde products.xlsx
    products_path = os.path.join(base_dir, "../assets/products.xlsx")
    xl = pd.read_excel(products_path, sheet_name=None)
    product_db = {
        supplier.lower(): set(df["Product Code"].astype(str)) for supplier, df in xl.items()
    }

    # 2. Crear carpetas de salida
    for supplier in product_db:
        os.makedirs(os.path.join(output_base, supplier), exist_ok=True)

    # 3. Procesar todos los PDF de la carpeta PDF_invoices
    for filename in os.listdir(input_base):
        if not filename.lower().endswith(".pdf"):
            continue

        pdf_path = os.path.join(input_base, filename)

        # Probar todos los extractores
        extractors = {
            "alm": extract_alm_invoice_data,
            "coke": extract_coke_invoice_data,
            "cub": extract_cub_invoice_data,
            "lion": extract_lion_invoice_data
        }

        best_supplier = None
        best_df = None
        best_match_count = 0

        for supplier, extractor in extractors.items():
            try:
                df = extractor(pdf_path)
                product_code_col = next((col for col in df.columns if col.strip().lower() == "product code"), None)
                if not product_code_col:
                    raise ValueError("❌ No se encontró la columna 'Product Code' en el DataFrame extraído.")
                product_codes = set(df[product_code_col].astype(str))

                match_supplier = detect_supplier(product_codes, product_db)

                if match_supplier and len(product_codes & product_db[match_supplier]) > best_match_count:
                    best_supplier = match_supplier
                    best_df = df
                    best_match_count = len(product_codes & product_db[match_supplier])
            except Exception as e:
                print(f"❌ Error processing {filename} with extractor {supplier}: {e}")
                continue

        if best_supplier and best_df is not None:
            output_path = os.path.join(output_base, best_supplier, os.path.splitext(filename)[0] + ".xlsx")
            best_df.to_excel(output_path, index=False)
            print(f"✅ {best_supplier.upper()}: {filename} → {output_path}")

        else:
            print(f"❌ Could not detect the supplier for {filename}")



