import subprocess
import sys
import os

# Configuración de scripts y proveedores
SCRIPTS = {
    "1": ("Extract product data from PDF invoice ✅", "scripts/1-parser.py"),
    "4": ("Upload invoice to Lightspeed", "scripts/4-upload.py", True),  # Requires supplier
    "7": ("Extract promos from Bottlemart PDF", "scripts/7-promos_parser.py"),
    "8": ("Upload promo prices to Lightspeed", "scripts/8-upload_promos.py"),
}

SUPPLIERS = ["ALM", "COKE", "CUB", "LION"]


def run_script(script_name, supplier=None):
    try:
        cmd = [sys.executable, script_name]
        if supplier:
            cmd.append(supplier)

        print(f"\n🚀 Running: {' '.join(cmd)}\n")
        subprocess.run(cmd, check=True)
        print("✅ Script completed successfully.\n")
    except subprocess.CalledProcessError as e:
        print(f"❌ Error while running {script_name}: {e}\n")


def choose_supplier():
    print("\nSelect supplier:")
    for idx, name in enumerate(SUPPLIERS, 1):
        print(f"{idx}. {name}")
    choice = input("Enter number: ").strip()
    return SUPPLIERS[int(choice) - 1] if choice in map(str, range(1, len(SUPPLIERS) + 1)) else None

def clear_console():
    os.system("cls" if os.name == "nt" else "clear")

def main():
    while True:
        print("""
========= Red Sands Order Manager 🍻 =========
0. Clear console
1. 🧾 Extract product data from PDF invoice
4. Upload invoice to Lightspeed
7. 💰 Extract promos from Bottlemart PDF (Beta 🤖)
8. Upload promo prices to Lightspeed (Beta 🤖)  
10. 👋 Exit
""")

        choice = input("Select an option: ").strip()

        if choice == "10":
            print("👋 Exiting.")
            break

        elif choice == "0":
            clear_console()
            continue

        option = SCRIPTS.get(choice)
        if not option:
            print("⚠️  Invalid option. Try again.")
            continue

        name, script = option[0], option[1]
        requires_supplier = option[2] if len(option) > 2 else False
        supplier = choose_supplier() if requires_supplier else None

        if requires_supplier and not supplier:
            print("⚠️  Invalid supplier. Returning to menu.\n")
            continue

        run_script(script, supplier)


if __name__ == "__main__":
    main()
