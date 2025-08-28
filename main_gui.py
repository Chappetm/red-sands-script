import streamlit as st
import os
import sys
import subprocess
import tempfile
from pathlib import Path
from datetime import datetime


st.set_page_config(page_title="Red Sands Panel", layout="centered")

st.markdown("""
<h1 style='font-size: 2.5em'>
  THE <span style='color: red'>RED</span> SANDS - Automation Panel 🍺
</h1>
""", unsafe_allow_html=True)


# Navigation state
if "active_page" not in st.session_state:
    st.session_state.active_page = "Home"

# Function to change view
def change_page(name):
    st.session_state.active_page = name

# -----------------------------
# SIDEBAR WITH BUTTONS
# -----------------------------
st.sidebar.title("📋 Main Menu")

# Invoices & uploads
with st.sidebar.expander("🧾 Invoices & Uploads", expanded=False):
    if st.button("📥 Download Invoices", use_container_width=True):
        change_page("download")
    if st.button("🚚 Upload Alcohol Order", use_container_width=True):
        change_page("upload_order")
    if st.button("🧾 Invoice to Lightspeed", use_container_width=True):
        change_page("parse_upload")

# Inventory & products
with st.sidebar.expander("📦 Inventory & Products", expanded=False):
    if st.button("➕ Add New Product", use_container_width=True):
        change_page("add_product")
    if st.button("🧮 Stocktake", use_container_width=True):
        change_page("stocktake")

# Reports & lists
with st.sidebar.expander("📊 Reports & Lists", expanded=False):
    if st.button("📊 Order Report", use_container_width=True):
        change_page("order_report")
    if st.button("🍽️ Meals List", use_container_width=True):
        change_page("meals")
    if st.button("📦 Delivery Checklist", use_container_width=True):
        change_page("delivery")

# Help
with st.sidebar.expander("❓ Help", expanded=False):
    if st.button("❓ Help & FAQ", use_container_width=True):
        change_page("help")

st.sidebar.divider()
if st.sidebar.button("🏠 Back to Home", use_container_width=True):
    change_page("Home")

# -----------------------------
# HOME PAGE
# -----------------------------
if st.session_state.active_page == "Home":
    st.subheader("Welcome")
    st.markdown("Select an option from the menu on the left.")


# -----------------------------
# PAGE: DOWNLOAD INVOICES ✅
# -----------------------------
elif st.session_state.active_page == "download":
    st.header("📥 Download Invoices")
    st.caption("Choose the supplier and automatically download this week's invoices. PDFs are saved to your local Downloads folder.")

    supplier_map = {
        "LION": "lion",
        "CUB": "cub",
        "COKE": "coke",
        "ALM": "alm",
    }

    supplier_label = st.selectbox("Supplier", list(supplier_map.keys()), index=0, key="dl_supplier")

    # Actions (solo descarga simple)
    if st.button("⬇️ Download now", key="btn_dl_now"):
        supplier_arg = supplier_map[supplier_label]
        with st.spinner(f"Downloading invoices from {supplier_label}..."):
            result = subprocess.run(
                ["python", "scripts/11-download_invoice.py", supplier_arg],
                capture_output=True,
                text=True
            )
        if result.returncode == 0:
            st.success("✅ Download complete. Check your **Downloads** folder.")
        else:
            st.error("❌ Error during download.")
            if result.stderr.strip():
                st.code(result.stderr)

    st.info("📌 Tip: you can then upload these PDFs in **Delivery Checklist** to generate the checklist.")


# -----------------------------
# PAGE: DELIVERY CHECKLIST ✅
# -----------------------------
elif st.session_state.active_page == "delivery":
    from pathlib import Path
    import subprocess, os

    st.header("📦 Create Delivery Checklist")
    st.caption("Upload the invoice PDF for each supplier. The system will generate a delivery checklist.")

    suppliers = ["alm", "coke", "lion", "cub"]

    pdf_dir = Path("PDF_invoices"); pdf_dir.mkdir(parents=True, exist_ok=True)
    xlsx_base = Path("Excel_invoices"); xlsx_base.mkdir(parents=True, exist_ok=True)

    # --- Estado de control para evitar re-escrituras en reruns ---
    st.session_state.setdefault("delivery_uploader_nonce", 0)      # cambia la key de los uploaders -> resetea selección
    st.session_state.setdefault("delivery_freeze_uploads", False)  # True mientras generamos checklist

    # === UPLOAD PDFs por proveedor ===
    for supplier in suppliers:
        st.markdown(f"### {supplier.upper()} Invoice")

        uploaded_files = st.file_uploader(
            f"Upload new {supplier.upper()} invoice PDFs",
            type=["pdf"],
            accept_multiple_files=True,
            key=f"{supplier}_uploader_{st.session_state.delivery_uploader_nonce}",
            help="You can add multiple PDFs. They won’t be re-saved during generation."
        )

        # ⛔️ No guardar nada si estamos generando (evita que reaparezcan tras vaciar la carpeta)
        if uploaded_files and not st.session_state.delivery_freeze_uploads:
            new_files_saved = 0
            for file in uploaded_files:
                file_bytes = file.read()
                stem = Path(file.name).stem.replace(" ", "_")
                out_path = pdf_dir / f"invoice_{supplier}_{stem}.pdf"
                if not out_path.exists():
                    out_path.write_bytes(file_bytes)
                    new_files_saved += 1
            if new_files_saved:
                st.success(f"✅ {new_files_saved} new {supplier.upper()} invoice(s) uploaded")
            else:
                st.info(f"ℹ️ No new {supplier.upper()} invoices to upload (all were already saved)")

        existing = sorted(pdf_dir.glob(f"invoice_{supplier}_*.pdf"))
        label = f"✅ {len(existing)} {supplier.upper()} invoice(s) in folder" if existing else f"⚠️ No {supplier.upper()} invoice found in folder"
        with st.expander(label, expanded=not existing):
            if existing:
                for p in existing:
                    c1, c2 = st.columns([0.75, 0.25])
                    c1.write(f"• {p.name}")
                    if c2.button("🗑️ Delete", key=f"del_{supplier}_{p.name}".replace(" ", "_").replace("/", "_")):
                        try:
                            p.unlink(missing_ok=True)
                            st.success(f"🗑️ Deleted {p.name}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"❌ Could not delete {p.name}: {e}")
                if st.button(f"🧹 Delete ALL {supplier.upper()} PDFs", key=f"del_all_{supplier}"):
                    removed = 0
                    for p in list(existing):
                        try:
                            p.unlink(missing_ok=True); removed += 1
                        except Exception:
                            pass
                    st.success(f"🧹 Removed {removed} PDF(s)")
                    st.rerun()
            else:
                st.info("Please upload at least one invoice for this supplier.")

    # === Condición: al menos 1 PDF por proveedor ===
    all_suppliers_ready = all(any(pdf_dir.glob(f"invoice_{s}_*.pdf")) for s in suppliers)

    if all_suppliers_ready:
        if st.button("📋 Generate Delivery Checklist"):
            # 🔒 CONGELAR uploads desde este momento para que NO se re-escriban en reruns
            st.session_state.delivery_freeze_uploads = True

            # Asegurar subcarpetas de Excel
            for s in suppliers:
                (xlsx_base / s).mkdir(parents=True, exist_ok=True)

            # ¿Faltan excels? -> parser
            missing = [s for s in suppliers if not any((xlsx_base / s).glob("*.xlsx"))]
            if missing:
                st.info(f"🔄 Missing parsed Excel files for: {', '.join(s.upper() for s in missing)}")
                with st.spinner("🧾 Parsing invoices..."):
                    p = subprocess.run([sys.executable, "scripts/1-parser.py"], capture_output=True, text=True)
                    if p.returncode != 0:
                        st.error("❌ Error parsing invoices:"); st.code(p.stderr); st.stop()
                    st.success("✅ Invoices parsed successfully")
            else:
                st.success("✅ All suppliers already parsed. Skipping parser.")

            # Generar delivery (este script VACÍA PDF_invoices al terminar)
            with st.spinner("📦 Generating delivery checklist..."):
                p = subprocess.run([sys.executable, "scripts/2-delivery.py"], capture_output=True, text=True)
                if p.returncode != 0:
                    st.error("❌ Error generating delivery checklist:"); st.code(p.stderr); st.stop()

                st.success("✅ Delivery checklist generated successfully")
                st.code(p.stdout)

                # Obtener OUTPUT_FILE=
                generated_path = None
                for line in p.stdout.splitlines():
                    if line.strip().startswith("OUTPUT_FILE="):
                        generated_path = line.split("OUTPUT_FILE=", 1)[1].strip()
                        break

                if not generated_path or not Path(generated_path).exists():
                    st.warning("⚠️ Delivery checklist generated but file not found."); st.stop()

                sheet_path = Path(generated_path)
                sheet_bytes = sheet_path.read_bytes()

                clicked = st.download_button(
                    "⬇️ Download Delivery Sheet",
                    data=sheet_bytes,
                    file_name=sheet_path.name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_new_delivery",
                )

                if clicked:
                    # Opcional: borrar el xlsx generado
                    try: sheet_path.unlink(missing_ok=True)
                    except Exception: pass

                    # ✅ RESET: soltar selección de los uploaders y “descongelar” uploads
                    st.session_state.delivery_uploader_nonce += 1   # cambia las keys -> los uploaders quedan vacíos
                    st.session_state.delivery_freeze_uploads = False
                    st.success("✅ Delivery Sheet downloaded.")
                    st.rerun()
    else:
        st.warning("🚫 Cannot generate delivery checklist: missing supplier invoices.")


# -----------------------------
# PAGE: ALCOHOL ORDER REPORT
# -----------------------------
elif st.session_state.active_page == "order_report":
    st.header("📊 Generate Order Report")
    st.caption("Upload your weekly sales report and generate a new order report using the template.")

    script_path = "scripts/3-sell_report.py"
    template_path = "assets/report_template.xlsx"

    # Nombre con fecha tipo 01_08_2025
    from datetime import datetime
    today_str = datetime.today().strftime("%d_%m_%Y")
    output_file = os.path.abspath(f"order_report_{today_str}.xlsx")

    # Subida del archivo sale_report.csv
    uploaded_csv = st.file_uploader("📥 Upload sales report CSV", type=["csv"], key="sale_csv")

    if uploaded_csv:
        with open("sale_report.csv", "wb") as f:
            f.write(uploaded_csv.read())
        st.success("✅ Sales report uploaded successfully.")

        # Solo si la plantilla existe
        if os.path.exists(template_path):
            if st.button("📊 Generate Report"):
                with st.spinner("Generating report from template..."):
                    result = subprocess.run(["python", script_path, output_file], capture_output=True, text=True)

                if result.returncode == 0:
                    st.success("✅ Report generated successfully.")
                    st.code(result.stdout)

                    if os.path.exists(output_file):
                        import base64

                        if os.path.exists(output_file):
                            with open(output_file, "rb") as f:
                                bytes_data = f.read()
                                b64 = base64.b64encode(bytes_data).decode()

                                href = f'<a href="data:application/octet-stream;base64,{b64}" download="{os.path.basename(output_file)}">⬇️ Download Report</a>'
                                st.markdown(href, unsafe_allow_html=True)

                                # Eliminar archivo después de mostrar el enlace
                                os.remove(output_file)

                    else:
                        st.warning("⚠️ Script ran but no output file was found.")
                else:
                    st.error("❌ Error running the script:")
                    st.code(result.stderr)
        else:
            st.error("❌ Template file not found: 'assets/report_template.xlsx'")
    else:
        st.info("Please upload sales report (.csv) to begin.")


# -----------------------------
# PAGE: SUPPLIER ORDER
# -----------------------------
elif st.session_state.active_page == "upload_order":
    st.header("🚚 Upload Alcohol Order")
    st.caption("Upload your completed **Order Report**, choose the supplier, and the system will submit your order.")

    script_generate = "scripts/5-report.py"
    script_upload = "scripts/6-order.py"
    report_path = "report.xlsx"
    order_ready_path = "order_ready.xlsx"

    # Subir archivo report.xlsx
    uploaded_file = st.file_uploader("📁 Upload your **Order Report**", type=["xlsx"])

    if uploaded_file:
        with open(report_path, "wb") as f:
            f.write(uploaded_file.read())
        st.success("✅ Report uploaded successfully.")
        st.markdown("---")

        st.markdown("### Select a Supplier to Submit the Order")

        supplier_options = {
            "ALM": "ALM",
            "COKE": "COKE",
            "CUB": "CUB",
            "LION (Alcohol)": "LION",
            "LION (Kegs only)": "LION KEGS",
        }


        selected_label = st.selectbox("Choose a supplier:", list(supplier_options.keys()))
        selected_supplier = supplier_options[selected_label]

        def submit_order(supplier):
            try:
                # 1. Ejecutar 5-report.py
                with st.spinner("🔄 Generating order file (order_ready.xlsx)..."):
                    result_report = subprocess.run(["python", script_generate], capture_output=True, text=True)
                    if result_report.returncode != 0:
                        st.error("❌ Failed to generate order file.")
                        st.code(result_report.stderr)
                        return
                    st.success("✅ order_ready.xlsx generated.")
                    st.code(result_report.stdout)

                # 2. Ejecutar 6-order.py
                with st.spinner(f"🚚 Submitting order to {supplier}..."):
                    result_order = subprocess.run(["python", script_upload, supplier], capture_output=True, text=True)
                    if result_order.returncode == 0:
                        st.success(f"✅ Order successfully submitted to {supplier}.")

                        # === Descarga de order_ready.xlsx para comparar con el carrito ===
                        order_ready_path = Path("order_ready.xlsx")

                        if order_ready_path.exists():
                            # Persistir los bytes en sesión para evitar que se pierdan en reruns
                            if "order_ready_bytes" not in st.session_state:
                                st.session_state["order_ready_bytes"] = order_ready_path.read_bytes()

                            st.info("🧾 Tu archivo **order_ready.xlsx** está listo. Puedes compararlo con el carrito generado.")
                            st.download_button(
                                "⬇️ Descargar order_ready.xlsx",
                                data=st.session_state["order_ready_bytes"],
                                file_name=f"order_ready_{datetime.now():%Y%m%d_%H%M}.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                key="dl_order_ready_supplier",
                            )
                        else:
                            st.warning("⚠️ No encuentro **order_ready.xlsx**. Verifica que el reporte se haya generado correctamente.")

                        st.code(result_order.stdout)
                    else:
                        st.error(f"❌ Failed to submit order to {supplier}.")
                        st.code(result_order.stderr)
            finally:
                # 🧹 Borrar archivos temporales
                for path in [report_path, order_ready_path]:
                    if os.path.exists(path):
                        os.remove(path)
                st.info("🧹 Temporary files cleaned up.")

        if st.button("🚀 Submit Order"):
            submit_order(selected_supplier)

    else:
        st.info("Please upload a report.xlsx file to begin.")


# -----------------------------
# PAGE: MEAL LIST
# -----------------------------
elif st.session_state.active_page == "meals":
    st.header("🍽️ Generate Meals List")
    st.caption("Upload the **bookings report** and get a Meal list.")

    # Subir el CSV
    uploaded_csv = st.file_uploader("📥 Upload bookings report CSV", type=["csv"], key="bookings_csv")

    if uploaded_csv:
        # Guardamos con un nombre fijo para que el script lo lea fácil
        bookings_csv_path = "bookings_report.csv"
        with open(bookings_csv_path, "wb") as f:
            f.write(uploaded_csv.read())
        st.success("✅ Bookings report uploaded successfully.")

        # Nombre de salida sugerido
        date_txt = datetime.now().strftime("%d-%m-%Y")
        output_xlsx = f"meal_list_{date_txt}.xlsx"

        if st.button("🍽️ Generate Meals List"):
            with st.spinner("Generating meals list..."):
                result = subprocess.run(
                    ["python", "scripts/10-meal_list.py", "--input", bookings_csv_path, "--out", output_xlsx],
                    capture_output=True, text=True
                )

            if result.returncode == 0:
                st.success("✅ Meals list generated successfully.")
                st.code(result.stdout)

                if os.path.exists(output_xlsx):
                    with open(output_xlsx, "rb") as f:
                        st.download_button("⬇️ Download Meals List", f, file_name=output_xlsx)
                    # Limpieza opcional después de mostrar el botón
                    try:
                        os.remove(output_xlsx)
                        os.remove(bookings_csv_path)
                        st.info("🧹 Temporary files cleaned up.")
                    except Exception as e:
                        st.warning(f"Couldn't clean temporary files: {e}")
                else:
                    st.warning("⚠️ Script ran but no output file was found.")
            else:
                st.error("❌ Error generating meals list:")
                st.code(result.stderr)
    else:
        st.info("Please upload the bookings CSV to begin.")


# -----------------------------
# PAGE: STOCKTAKE
# -----------------------------
elif st.session_state.active_page == "stocktake":
    import tempfile

    st.header("🧮 Stocktake")
    st.caption("Upload two scanner files and your products file. The system will generate a final count and the list of unmatched barcodes.")

    # Script a ejecutar
    script_stocktake = "scripts/9-stocktake.py"

    # ---- Estado persistente para evitar perder los archivos al descargar ----
    if "stk_final_bytes" not in st.session_state:
        st.session_state["stk_final_bytes"] = None
    if "stk_unmatched_bytes" not in st.session_state:
        st.session_state["stk_unmatched_bytes"] = None
    if "stk_logs" not in st.session_state:
        st.session_state["stk_logs"] = ""

    col1, col2 = st.columns(2)
    with col1:
        up_scanner1 = st.file_uploader("📥 Upload scanner1 (xlsx/csv)", type=["xlsx", "xls", "csv"], key="stk_sc1")
        up_products = st.file_uploader("📥 Upload products (csv/xlsx)", type=["csv", "xlsx", "xls"], key="stk_prd", help="See Help & FAQ")
    with col2:
        up_scanner2 = st.file_uploader("📥 Upload scanner2 (xlsx/csv)", type=["xlsx", "xls", "csv"], key="stk_sc2")

    run_btn = st.button("▶️ Run Stocktake")

    if run_btn:
        # Limpiar resultados anteriores
        st.session_state["stk_final_bytes"] = None
        st.session_state["stk_unmatched_bytes"] = None
        st.session_state["stk_logs"] = ""

        if not (up_scanner1 and up_scanner2 and up_products):
            st.error("❌ Please upload all three files: scanner1, scanner2 and products.")
        elif not os.path.exists(script_stocktake):
            st.error(f"❌ Script not found: {script_stocktake}")
        else:
            with st.spinner("Processing stocktake..."):
                try:
                    with tempfile.TemporaryDirectory() as tmpdir:
                        tmpdir = os.path.abspath(tmpdir)
                        sc1_ext = up_scanner1.name.split(".")[-1].lower()
                        sc2_ext = up_scanner2.name.split(".")[-1].lower()
                        prd_ext = up_products.name.split(".")[-1].lower()

                        p_sc1 = os.path.join(tmpdir, f"scanner1.{sc1_ext}")
                        p_sc2 = os.path.join(tmpdir, f"scanner2.{sc2_ext}")
                        p_prd = os.path.join(tmpdir, f"products.{prd_ext}")
                        outdir = os.path.join(tmpdir, "out")
                        os.makedirs(outdir, exist_ok=True)

                        with open(p_sc1, "wb") as f:
                            f.write(up_scanner1.read())
                        with open(p_sc2, "wb") as f:
                            f.write(up_scanner2.read())
                        with open(p_prd, "wb") as f:
                            f.write(up_products.read())

                        result = subprocess.run(
                            [
                                "python",
                                script_stocktake,
                                "--scanner1", p_sc1,
                                "--scanner2", p_sc2,
                                "--products", p_prd,
                                "--outdir",   outdir,
                            ],
                            capture_output=True,
                            text=True,
                        )

                        # Guardar logs para mostrarlos persistentes
                        logs = ""
                        if result.stdout:
                            logs += result.stdout
                        if result.stderr:
                            logs += ("\n" + result.stderr)
                        st.session_state["stk_logs"] = logs.strip()

                        if result.returncode != 0:
                            st.error("❌ Stocktake failed.")
                        else:
                            final_path = os.path.join(outdir, "final_count.csv")
                            unmatched_path = os.path.join(outdir, "unmatched_barcodes.xlsx")

                            if os.path.exists(final_path):
                                with open(final_path, "rb") as f:
                                    st.session_state["stk_final_bytes"] = f.read()
                                st.success("✅ final_count.csv generated.")
                            else:
                                st.warning("⚠️ Script ran but final_count.csv was not found.")

                            if os.path.exists(unmatched_path):
                                with open(unmatched_path, "rb") as f:
                                    st.session_state["stk_unmatched_bytes"] = f.read()
                                st.info("📄 unmatched_barcodes.csv generated.")
                            else:
                                st.session_state["stk_unmatched_bytes"] = None
                                st.info("🎉 No unmatched barcodes.")
                except Exception as e:
                    st.error(f"❌ Error executing stocktake: {e}")

    # ---- Mostrar logs persistentes (si hubo) ----
    if st.session_state["stk_logs"]:
        st.text_area("Log", st.session_state["stk_logs"], height=160)

    # ---- Botones de descarga persistentes (no se pierden tras el primer click) ----
    if st.session_state["stk_final_bytes"] is not None:
        st.download_button(
            "⬇️ Download final_count.csv",
            data=st.session_state["stk_final_bytes"],
            file_name="final_count.csv",
            mime="text/csv",
            key="dl_final_stocktake",
        )

    if st.session_state["stk_unmatched_bytes"] is not None:
        st.download_button(
            "⬇️ Download unmatched_barcodes.xlsx",
            data=st.session_state["stk_unmatched_bytes"],
            file_name="unmatched_barcodes.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="dl_unmatched_stocktake",
        )


# -----------------------------
# PAGE: HELP
# -----------------------------
elif st.session_state.active_page == "help":
    st.title("Help & FAQ")

    # --- Desplegable: STOCKTAKE ---
    with st.expander("🧮 Stocktake", expanded=False):
        st.markdown("### How to download the **Products** file from Lightspeed")

        st.markdown("**1. Log in to Lightspeed**")

        st.markdown("**2. Go to Products (left sidebar)**")

        st.markdown("**3. Click the blank checkbox and choose All**")
        st.image("assets/help/step_3.png", width=200)

        st.markdown("**4. Click the dropdown next to it and select Export products**")
        st.image("assets/help/step_5.png", width=300)

        st.markdown("**5. Click Export, then Export products**")

        st.markdown("**6. Wait for the export to finish, then click Download the export**")
      

# -----------------------------
# PAGE: ADD NEW PRODUCT
# -----------------------------
elif st.session_state.active_page == "add_product":
    st.header("➕ Add Product")
    st.caption("Fill the fields below. The summary updates live. We’ll generate Excel/CSV in the next step.")

    # ---- Estado persistente ----
    if "addp_data" not in st.session_state:
        st.session_state["addp_data"] = {}
    if "last_ptype" not in st.session_state:
        st.session_state["last_ptype"] = None

    # ---- Opciones maestras ----
    SUPPLIERS   = ["ALM", "CUB", "LION", "COKE"]
    CATEGORIES  = ["Beers", "Beer on tap", "Ciders", "RTDs", "Wines", "Spirits", "Soft drinks", "Snacks"]
    TYPES_RAW   = ["", "Cans", "Stubbies", "Bottle (Wine/Spirit)"]  # "" = vacío por defecto
    NUM_UNIT_CHOICES = [1, 3, 4, 6, 10, 12, 16, 24, 30]                  # para Cans/Stubbies -> Cn/Sn
    BOTTLE_SIZES = ["700ml", "750ml", "1L", "500ml"]                 # único size permitido (no custom)

    def fmt_type(x: str) -> str:
        return "— Select type —" if x == "" else x

    with st.expander("ℹ️ Quick help", expanded=False):
        st.info(
            "• **Type** defines the base unit: C1 (can), S1 (stubbie) or the **bottle size** (wine/spirit).\n"
            "• **Carton size** = how many units per carton (e.g., 24).\n"
            "• **Sell in** are the presentations you will sell (e.g., C1/C6/C24 or S1/S4/S24, or a bottle size)."
        )

    # ---------- Inputs "live" ----------
    colA, colB = st.columns(2)
    with colA:
        name = st.text_input("Product name*", placeholder="e.g., Great Northern Original", key="addp_name")
        code_carton = st.text_input("Product code (carton)*", placeholder="Supplier/Carton code", key="addp_code_carton")
        supplier = st.selectbox("Supplier*", SUPPLIERS, index=0, key="addp_supplier")
        category = st.selectbox("Category*", CATEGORIES, index=0, key="addp_category")

    with colB:
        ptype = st.selectbox("Type*", TYPES_RAW, index=0, format_func=fmt_type, key="addp_ptype")
        carton_size = st.number_input("Carton size (units per carton)*", min_value=1, step=1, value=24, key="addp_carton_size")
        cost_per_carton = st.number_input("Cost per carton", min_value=0.0, step=0.01, format="%.2f", key="addp_cost")
        barcode = st.text_input("Barcode (optional)", placeholder="EAN/UPC", key="addp_barcode")

    # --- Reset dependientes si cambia Type ---
    if st.session_state["last_ptype"] != ptype:
        for k in ["sell_units_numbers", "sell_units_bottle", "bottle_size_select"]:
            if k in st.session_state:
                del st.session_state[k]
        st.session_state["last_ptype"] = ptype

    # ------- Dependientes (Sell in SIEMPRE visible) -------
    prefix_map = {"Cans": "C", "Stubbies": "S", "Bottle (Wine/Spirit)": ""}

    sell_units = []
    base_unit_display = None
    bottle_size = None

    if ptype == "Bottle (Wine/Spirit)":
        # Solo una medida (sin custom ni extras)
        bottle_size = st.selectbox(
            "Bottle size*",
            BOTTLE_SIZES,
            index=0,
            key="bottle_size_select"
        )

        # Sell in bloqueado con la única medida
        locked_opts = [bottle_size] if bottle_size else []
        st.multiselect(
            "Sell in (choose unit sizes)",
            locked_opts,
            default=locked_opts,
            key="sell_units_bottle",
            disabled=True
        )
        sell_units = locked_opts
        base_unit_display = bottle_size or "(size)"

    elif ptype in ("Cans", "Stubbies"):
        # Sell in por números -> mapeamos a Cn / Sn
        raw_sel = st.multiselect(
            "Sell in (choose unit sizes)",
            NUM_UNIT_CHOICES,
            default=[],
            key="sell_units_numbers"
        )
        pref = prefix_map.get(ptype, "")
        sell_units = [f"{pref}{n}" for n in raw_sel]
        base_unit_display = f"{pref}1" if pref else None

    else:
        # Type vacío: mostrar Sell in vacío (sin opciones)
        st.multiselect(
            "Sell in (choose unit sizes)",
            [],
            default=[],
            key="sell_units_numbers"
        )
        sell_units = []
        base_unit_display = None

    # ------- Validaciones -------
    errors = []
    if not name.strip():
        errors.append("Product name is required.")
    if not code_carton.strip():
        errors.append("Product code (carton) is required.")
    if cost_per_carton <= 0:
        errors.append("Cost per carton must be greater than 0.")
    if carton_size <= 0:
        errors.append("Carton size must be greater than 0.")
    if ptype == "":
        errors.append("Type is required.")
    if ptype == "Bottle (Wine/Spirit)" and not (bottle_size and str(bottle_size).strip()):
        errors.append("Bottle size is required for Bottle (Wine/Spirit).")
    if ptype != "" and len(sell_units) == 0:
        errors.append("Select at least one 'Sell in' unit.")

    is_valid = (len(errors) == 0)
    st.session_state["addp_is_valid"] = is_valid

    # ------- Guardar en session_state -------
    st.session_state["addp_data"] = {
        "name": name.strip(),
        "code_carton": code_carton.strip(),
        "supplier": supplier,
        "category": category,
        "type": (ptype if ptype else None),
        "carton_size": int(carton_size) if carton_size else None,
        "sell_units": sell_units,
        "cost_per_carton": float(cost_per_carton) if cost_per_carton else None,
        "barcode": barcode.strip() if barcode else "",
        "base_unit_display": base_unit_display,
        "bottle_size": (str(bottle_size).strip() if bottle_size else None),
    }

    # ------- Summary -------
    import html

    # CSS para el estilo rojo (una sola vez por render)
    st.markdown(
        "<style>.summary-value{color:#d32f2f;font-weight:700}</style>",
        unsafe_allow_html=True
    )

    def red(val):
        """Devuelve el valor envuelto en un span rojo; maneja vacíos."""
        if val is None:
            txt = "-"
        elif isinstance(val, str) and not val.strip():
            txt = "-"
        else:
            txt = str(val)
        return f"<span class='summary-value'>{html.escape(txt)}</span>"

    data = st.session_state["addp_data"]
    sell_in_txt = ", ".join(data["sell_units"]) if data.get("sell_units") else "-"
    cost_txt = f"{data['cost_per_carton']:.2f}" if data.get("cost_per_carton") is not None else "-"

    st.markdown("### 📋 Summary")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown(f"**Name:** {red(data.get('name'))}", unsafe_allow_html=True)
        st.markdown(f"**Product Code:** {red(data.get('code_carton'))}", unsafe_allow_html=True)
        st.markdown(f"**Supplier:** {red(data.get('supplier'))}", unsafe_allow_html=True)
        st.markdown(f"**Category:** {red(data.get('category'))}", unsafe_allow_html=True)
        st.markdown(f"**Type:** {red(data.get('type'))}", unsafe_allow_html=True)

    with col2:
        st.markdown(f"**Carton size:** {red(data.get('carton_size'))}", unsafe_allow_html=True)
        st.markdown(f"**Sell in:** {red(sell_in_txt)}", unsafe_allow_html=True)
        st.markdown(f"**Cost/carton:** {red(cost_txt)}", unsafe_allow_html=True)
        st.markdown(f"**Barcode:** {red(data.get('barcode') or '-')}", unsafe_allow_html=True)
        st.markdown(f"**Base unit for name suffix:** {red(data.get('base_unit_display'))}", unsafe_allow_html=True)

    # Bottle size (si aplica)
    if data.get("bottle_size"):
        st.markdown(f"**Bottle size:** {red(data['bottle_size'])}", unsafe_allow_html=True)


    # ------- Config fija: paths NO editables por el usuario -------
    from pathlib import Path
    import sys, subprocess, os

    APP_DIR = Path(__file__).resolve().parent
    PRODUCTS_XLSX = os.getenv("PRODUCTS_XLSX", str(APP_DIR / "assets" / "products.xlsx"))
    REPORT_XLSX   = os.getenv("REPORT_XLSX",   str(APP_DIR / "assets" / "report_template.xlsx"))

    # Chequeo de existencia del template
    if not Path(REPORT_XLSX).exists():
        st.error(f"Template not found: {REPORT_XLSX}")

    # ------- Botón Next: agrega a products.xlsx y al report_template.xlsx -------
    d = st.session_state["addp_data"]
    is_beer_on_tap = (d.get("category") == "Beer on tap")

    if is_beer_on_tap:
        st.info("‘Beer on tap’ no se modifica desde aquí.", icon="🛑")

    clicked_next = st.button(
        "➡️ Next: Add to products.xlsx & report_template.xlsx",
        disabled=is_beer_on_tap or not st.session_state.get("addp_is_valid", False),
        help="Agrega el producto al supplier sheet y al reporte (sobre el template en assets)."
    )

    if clicked_next:
        # --------- Unidades: calcular UNA sola vez según CATEGORÍA ----------
        cat = d["category"]
        typ = d["type"] if d["type"] != "Bottle (Wine/Spirit)" else "Bottle"

        if cat in ("Beers", "Ciders", "RTDs"):
            units_csv = ",".join(d["sell_units"])  # p.ej. C1,C6,C24 / S1,S24
        elif cat == "Wines":
            units_csv = "S1"                       # 750 ml
        elif cat == "Spirits":
            units_csv = (d.get("bottle_size") or "").upper()  # 700ML o 1L
        elif cat == "Soft drinks":
            units_csv = "S1"                       # fila simple; la unidad aquí es simbólica
        elif cat == "Snacks":
            units_csv = ""                         # no se usa
        else:
            st.error(f"Category not supported: {cat}")
            st.stop()

        # Validación específica Soft drinks (necesita carton_size)
        if cat == "Soft drinks" and not d.get("carton_size"):
            st.error("For Soft drinks, Carton size is required.")
            st.stop()

        script_path = APP_DIR / "scripts" / "12-add_to_products.py"
        if not script_path.exists():
            st.error(f"❌ Script not found: {script_path}")
        else:
            cmd = [
                sys.executable, str(script_path),
                "-w", str(PRODUCTS_XLSX),
                "-r", str(REPORT_XLSX),
                "-s", d["supplier"],
                "-c", d["code_carton"],
                "-n", d["name"],
                "--category", cat,
                "--ptype", typ,
                "--units", units_csv
            ]
            if cat == "Soft drinks":
                cmd += ["--carton-size", str(d["carton_size"])]

            result = subprocess.run(cmd, capture_output=True, text=True)

            logs = (result.stdout or "") + (("\n" + result.stderr) if result.stderr else "")
            st.text_area("Log", logs.strip(), height=200)

            if result.returncode == 0:
                st.success("✅ Added to products.xlsx & assets/report_template.xlsx")
                st.toast("Workbooks updated", icon="✅")
            elif result.returncode == 2:
                st.warning("⚠️ Code already exists in products sheet. Report was updated.")
            else:
                st.error("❌ Failed to update one or both workbooks")


# -----------------------------
# PAGE: UPLOAD INVOICE (LIGHTSPEED) ✅
# -----------------------------
elif st.session_state.active_page == "parse_upload":
    # Inline: upload PDF → auto-parse (no button) → detect XLSX → allow upload to Lightspeed
    import sys, time, re, subprocess
    from pathlib import Path

    st.header("🧾 Upload Invoice")

    # Base paths
    base_dir = Path(__file__).resolve().parent
    pdf_dir = base_dir / "PDF_invoices"
    xlsx_base = base_dir / "Excel_invoices"
    pdf_dir.mkdir(parents=True, exist_ok=True)
    xlsx_base.mkdir(parents=True, exist_ok=True)



    # State (persist across reruns)
    st.session_state.setdefault("pu_pdf_path", None)            # saved PDF path (str)
    st.session_state.setdefault("pu_xlsx_path", None)           # parsed XLSX path (str)
    st.session_state.setdefault("pu_detected_supplier", None)   # supplier deduced from XLSX folder
    st.session_state.setdefault("pu_parsed", False)             # whether parser already ran successfully for current PDF

    # Used to reset the file_uploader widget after a successful upload
    st.session_state.setdefault("pu_uploader_key", "pu_uploader_v1")
    # Fingerprint of the currently selected file (name+size) to avoid re-saving on reruns
    st.session_state.setdefault("pu_pdf_token", None)


    # Upload a single PDF (parser will auto-run after save)
    pdf_file = st.file_uploader(
        "Upload a PDF invoice",
        type=["pdf"],
        accept_multiple_files=False,
        key=st.session_state.pu_uploader_key  # stable key so we can reset it later
    )


    # Save the uploaded PDF exactly once per selected file (avoid duplicates on reruns)
    if pdf_file is not None:
        # Build a fingerprint that survives reruns while the same file remains selected
        file_size = getattr(pdf_file, "size", None)
        cur_token = f"{pdf_file.name}:{file_size}"

        # Only write the file if this is a new selection
        if st.session_state.pu_pdf_path is None or st.session_state.pu_pdf_token != cur_token:
            ts = time.strftime("%Y%m%d_%H%M%S")
            safe_name = re.sub(r"[^A-Za-z0-9_.-]+", "_", pdf_file.name)
            pdf_path = pdf_dir / f"{Path(safe_name).stem}_{ts}.pdf"

            with open(pdf_path, "wb") as f:
                f.write(pdf_file.read())

            # Persist state
            st.session_state.pu_pdf_path = str(pdf_path)
            st.session_state.pu_pdf_token = cur_token
            st.session_state.pu_xlsx_path = None
            st.session_state.pu_detected_supplier = None
            st.session_state.pu_parsed = False

            st.success(f"✅ PDF saved: `{pdf_path.name}` → `{pdf_dir}`")
        # else: same file still selected → do NOT write again


    # Parser will run automatically when you click "Upload to Lightspeed"
    if st.session_state.pu_pdf_path and not st.session_state.pu_parsed:
        st.info("Invoice uploaded. Click **Upload to Lightspeed** to process and upload.")


    # Upload to Lightspeed: run parser now, then upload
    can_upload = st.session_state.pu_pdf_path is not None
    if st.button("🚚 Upload to Lightspeed", disabled=not can_upload, key="btn_upload"):
        if not st.session_state.pu_pdf_path:
            st.error("Upload a PDF first.")
        else:
            # --- Step 1: Run parser (silent; just a status spinner) ---
            parser_path = base_dir / "scripts" / "1-parser.py"
            parser_cmd = [sys.executable, str(parser_path)]
            with st.status("Processing invoice (parser)...", state="running") as status:
                proc = subprocess.Popen(
                    parser_cmd,
                    cwd=str(base_dir),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    bufsize=1,
                    universal_newlines=True
                )
                for _ in iter(proc.stdout.readline, ''):
                    pass
                proc.wait()
                if proc.returncode == 0:
                    status.update(label="Parser: Done", state="complete")
                else:
                    status.update(label="Parser: Failed", state="error")

            if proc.returncode != 0:
                st.error("❌ Parser failed. See logs in terminal.")
                st.stop()

            # --- Step 2: Find resulting XLSX for this PDF stem (pick newest if multiple leftovers) ---
            pdf_stem = Path(st.session_state.pu_pdf_path).stem
            candidates = list(xlsx_base.rglob(f"{pdf_stem}.xlsx"))
            if not candidates:
                st.error("❌ Excel not found after parsing.")
                st.stop()

            def _mtime(p: Path) -> float:
                try:
                    return p.stat().st_mtime
                except Exception:
                    return 0.0

            chosen = sorted(candidates, key=_mtime, reverse=True)[0]
            effective_supplier = chosen.parent.name  # Excel_invoices/<supplier>/<file.xlsx>

            st.session_state.pu_xlsx_path = str(chosen)
            st.session_state.pu_detected_supplier = effective_supplier
            st.session_state.pu_parsed = True

            # --- Step 3: Run uploader (stream full console) ---
            st.info(f"Uploading using supplier: **{effective_supplier.upper()}**")
            uploader_path = base_dir / "scripts" / "4-upload.py"

            # Use -u (unbuffered) so prints show up immediately
            uploader_cmd = [sys.executable, "-u", str(uploader_path), effective_supplier]

            # Live log box
            log_box = st.empty()
            lines = []

            with st.status("Uploading to Lightspeed...", state="running") as status:
                up = subprocess.Popen(
                    uploader_cmd,
                    cwd=str(base_dir),
                    stdout=subprocess.PIPE,
                    stderr=subprocess.STDOUT,
                    bufsize=1,
                    universal_newlines=True
                )
                # Stream every line to the UI
                for line in iter(up.stdout.readline, ''):
                    lines.append(line.rstrip("\n"))
                    # show last N lines to keep UI snappy
                    log_box.code("\n".join(lines[-500:]))

                # make sure pipe is closed and process finished
                try:
                    up.stdout.close()
                except Exception:
                    pass
                up.wait()

                if up.returncode == 0:
                    status.update(label="Uploader: Done", state="complete")
                else:
                    status.update(label="Uploader: Failed", state="error")


            # --- Step 4: Cleanup and reset UI (only on success) ---
            if up.returncode == 0:
                # Cleanup PDF_invoices and Excel_invoices contents (keep roots)
                # 1) Delete files inside PDF_invoices recursively
                for p in pdf_dir.rglob("*"):
                    if p.is_file():
                        try:
                            p.unlink()
                        except Exception as e:
                            print(f"⚠️ Could not delete {p}: {e}")
                # 2) Remove empty subdirectories in PDF_invoices
                for p in sorted(pdf_dir.rglob("*"), reverse=True):
                    if p.is_dir():
                        try:
                            p.rmdir()
                        except OSError:
                            pass

                # 3) Delete files inside Excel_invoices recursively
                for p in xlsx_base.rglob("*"):
                    if p.is_file():
                        try:
                            p.unlink()
                        except Exception as e:
                            print(f"⚠️ Could not delete {p}: {e}")
                # 4) Remove empty subdirectories in Excel_invoices
                for p in sorted(xlsx_base.rglob("*"), reverse=True):
                    if p.is_dir():
                        try:
                            p.rmdir()
                        except OSError:
                            pass

                st.success("✅ Upload flow completed. 🧹 Cleaned PDF_invoices/ and Excel_invoices/.")

                # Reset UI state
                st.session_state.pu_pdf_path = None
                st.session_state.pu_xlsx_path = None
                st.session_state.pu_detected_supplier = None
                st.session_state.pu_parsed = False

                # IMPORTANT: reset the uploader widget so it doesn't re-save on next rerun
                st.session_state.pu_uploader_key = f"pu_uploader_{int(time.time())}"
            else:
                st.error("❌ Upload failed. See logs in terminal.")








