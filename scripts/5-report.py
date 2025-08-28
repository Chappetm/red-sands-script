import re
import pandas as pd
from rapidfuzz import process, fuzz


EXCEPTIONS_C30_TO_C1 = {
    "EMU EXPORT", "EMU BITTER", "IRON JACK MID", "CARLTON DRY", "CARLTON DRY 3.5", "CARLTON MID",
    "GREAT NORTHERN ORIGINAL", "GREAT NORTHERN 3.5", "VB", "HAHN 3.5", "HAHN SUPER DRY", "XXXX GOLD"
}

def map_suffix(product_name):
    name = product_name.strip().upper()

    # Excepción específica para "CORONA S24 (12PK)"
    if name == "CORONA S24 (12PK)":
        return "CORONA S12"

    if name.endswith("C30 PK"):
        base_name = name[:-7].strip()
        return f"{base_name} C10"

    base_name = re.sub(r'\s+[CS]\d{1,2}$', '', name, flags=re.IGNORECASE)
    suffix_match = re.search(r'(C|S)(\d{2})$', name)

    if suffix_match:
        kind = suffix_match.group(1)
        size = int(suffix_match.group(2))

        if size == 30 and base_name in EXCEPTIONS_C30_TO_C1:
            return f"{base_name} {kind}1"
        elif size in {24, 16}:
            return f"{base_name} {kind}1"
        elif size in {20, 30}:
            return f"{base_name} {kind}10"

    return name

def fuzzy_match_products(report_df, products_df):
    matched = []
    products_df["Product Name Normalized"] = products_df["Product Name"].str.strip().str.upper()
    product_choices = products_df["Product Name Normalized"].tolist()

    for _, row in report_df.iterrows():
        name = row["Product Name"]
        quantity = row["Quantity"]
        normalized_name = map_suffix(name)

        match, score, idx = process.extractOne(
            normalized_name, product_choices, scorer=fuzz.token_sort_ratio
        )

        if score > 85:
            product_code = products_df.iloc[idx]["Product Code"]
            supplier = products_df.iloc[idx]["Supplier"]

            matched.append({
                "Product Code": product_code,
                "Product Name": name,
                "Quantity": quantity,
                "Supplier": supplier
            })
        else:
            print(f"⚠️ No match confiable para: {name} → {normalized_name} (score: {score})")

    return pd.DataFrame(matched)

def load_report(filepath):
    all_data = []
    sheets = pd.read_excel(filepath, sheet_name=None, header=1)

    for sheet_name, df in sheets.items():
        df.columns = [col.strip().lower() for col in df.columns]

        if "products" in df.columns and "qty" in df.columns:
            df = df[["products", "qty"]].copy()
            df.rename(columns={"products": "Product Name", "qty": "Quantity"}, inplace=True)
            df.dropna(subset=["Product Name", "Quantity"], inplace=True)
            all_data.append(df)
        else:
            print(f"⚠️ Skipping sheet '{sheet_name}': missing required columns.")

    if not all_data:
        print("❌ No valid data found.")
        return pd.DataFrame(columns=["Product Name", "Quantity"])

    return pd.concat(all_data, ignore_index=True)

def load_products(filepath):
    sheets = pd.read_excel(filepath, sheet_name=None)
    all_products = []

    def clean_code(x, supplier):
        # Normalizamos el nombre del proveedor
        sup = (supplier or "").strip().upper()

        if pd.isna(x):
            return ""
        
        # Si ya es string, lo dejamos tal cual (salvo espacios)
        if isinstance(x, str):
            return x.strip()

        # Si es numérico, convertimos a entero y luego a string
        if isinstance(x, (int, float)):
            code_str = str(int(x))
            if sup == "ALM":
                # Solo ALM: rellenar a 6 si tiene menos de 6 dígitos
                return code_str if len(code_str) >= 6 else code_str.zfill(6)
            else:
                # Otros proveedores: dejar tal cual
                return code_str

        # Cualquier otro tipo (por si acaso): convertir a string y strip
        return str(x).strip()

    for sheet_name, df in sheets.items():
        if "Product Name" in df.columns and "Product Code" in df.columns:
            df = df[["Product Name", "Product Code"]].copy()
            df.dropna(subset=["Product Name"], inplace=True)
            df["Supplier"] = sheet_name
            df["Product Code"] = df["Product Code"].apply(lambda x: clean_code(x, sheet_name))
            df["Product Name Normalized"] = df["Product Name"].str.strip().str.upper()
            all_products.append(df)
        else:
            print(f"⚠️ Sheet '{sheet_name}' skipped: missing 'Product Name' or 'Product Code'.")

    return pd.concat(all_products, ignore_index=True)

# Cargar datos
report_path = "report.xlsx"
report_df = load_report(report_path)
products_df = load_products("products.xlsx")

# Generar archivo final
final_df = fuzzy_match_products(report_df, products_df)

if final_df.empty:
    print("⚠️ No se encontraron coincidencias. No se generó ningún archivo.")
else:
    with pd.ExcelWriter("order_ready.xlsx", engine="openpyxl") as writer:
        for supplier, group in final_df.groupby("Supplier"):
            if supplier.upper() == "LION":
                keg_mask = group["Product Name"].str.upper().str.endswith("KEG")
                lion_kegs = group[keg_mask]
                lion_others = group[~keg_mask]

                if not lion_kegs.empty:
                    lion_kegs[["Product Code", "Product Name", "Quantity"]].to_excel(writer, sheet_name="LION KEGS", index=False)

                if not lion_others.empty:
                    lion_others[["Product Code", "Product Name", "Quantity"]].to_excel(writer, sheet_name="LION", index=False)
            else:
                group[["Product Code", "Product Name", "Quantity"]].to_excel(writer, sheet_name=supplier[:31], index=False)

    print("✅ Archivo 'order_ready.xlsx' generado con hojas por proveedor.")
