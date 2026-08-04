"""
Academix DSS — Maintenance Tool
Consolidated script for common maintenance tasks.

Usage:
    python tools/maintenance.py              # Show menu
    python tools/maintenance.py fix-pwd      # Fix passwords on all sheets
    python tools/maintenance.py fix-accueil  # Re-import mod_AccueilButtons
    python tools/maintenance.py fix-all      # Fix passwords + re-import modules
    python tools/maintenance.py verify       # Verify all sheets accessible
    python tools/maintenance.py nuclear      # Delete and recreate locked sheets (DANGER)
"""
import sys
import os

try:
    import win32com.client
except ImportError:
    print("ERROR: pip install pywin32")
    sys.exit(1)

# Configuration
XLSM = r"C:\Users\Admin\projects\active\academix-dss-hardware\ERP_dss_v13.4_hardware_store.xlsm"
MOD_DIR = r"C:\Users\Admin\projects\active\academix-dss-hardware\modules"
MASTER_PWD = "erp_secure_pwd_2026"
OLD_PWD = "DSS_260803_4711"


def connect_excel():
    """Connect to running Excel or launch new instance."""
    try:
        excel = win32com.client.GetActiveObject("Excel.Application")
        print("Connected to running Excel")
    except:
        excel = win32com.client.Dispatch("Excel.Application")
        print("Launched new Excel")
    excel.Visible = True
    excel.DisplayAlerts = False
    return excel


def find_workbook(excel):
    """Find the hardware store workbook."""
    wb = None
    for w in excel.Workbooks:
        if "hardware" in w.Name.lower() or "erp_dss" in w.Name.lower():
            wb = w
            break
    if not wb:
        wb = excel.Workbooks.Open(XLSM)
        print("Opened workbook")
    return wb


def unprotect_sheet(ws):
    """Try multiple passwords to unprotect a sheet."""
    for pw in [OLD_PWD, MASTER_PWD, ""]:
        try:
            if pw:
                ws.Unprotect(Password=pw)
            else:
                ws.Unprotect()
            return True
        except:
            pass
    return False


def fix_passwords(wb):
    """Fix passwords on all sheets."""
    print("\n=== Fixing Passwords ===")
    for ws in wb.Sheets:
        unprotect_sheet(ws)
        ws.Protect(Password=MASTER_PWD, UserInterfaceOnly=True)
        print(f"  {ws.Name}: OK")

    # Update CONFIG MASTER_PWD
    ws_cfg = wb.Sheets("CONFIG")
    unprotect_sheet(ws_cfg)
    for r in range(1, 25):
        if ws_cfg.Cells(r, 1).Value == "MASTER_PWD":
            old = ws_cfg.Cells(r, 2).Value
            ws_cfg.Cells(r, 2).Value = MASTER_PWD
            print(f"  CONFIG MASTER_PWD: '{old}' -> '{MASTER_PWD}'")
            break
    ws_cfg.Protect(Password=MASTER_PWD, UserInterfaceOnly=True)


def fix_accueil(excel, wb):
    """Re-import mod_AccueilButtons."""
    print("\n=== Fixing mod_AccueilButtons ===")
    vbproj = wb.VBProject

    # Remove old
    try:
        old = vbproj.VBComponents("mod_AccueilButtons")
        vbproj.VBComponents.Remove(old)
        print("  Removed old mod_AccueilButtons")
    except:
        pass

    # Import fresh
    fp = os.path.join(MOD_DIR, "mod_AccueilButtons_v14.bas")
    try:
        vbproj.VBComponents.Import(fp)
        print(f"  Imported: {fp}")
    except Exception as e:
        print(f"  Import failed: {e}")
        # Fallback: strip Attribute
        with open(fp, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        clean = ''.join(l for l in lines if not l.startswith('Attribute'))
        comp = vbproj.VBComponents.Add(1)
        comp.Name = "mod_AccueilButtons"
        comp.CodeModule.AddFromString(clean)
        print("  Imported via AddFromString (Attribute stripped)")


def fix_all(excel, wb):
    """Fix passwords + re-import modules."""
    fix_passwords(wb)
    fix_accueil(excel, wb)

    # Try running KPI refresh
    try:
        excel.ExecuteExcel4Macro('Application.Run("RefreshAccueilKPIs")')
        print("\nRefreshAccueilKPIs executed")
    except Exception as e:
        print(f"\nKPI refresh failed: {e}")


def verify(wb):
    """Verify all sheets are accessible."""
    print("\n=== Verification ===")
    for ws in wb.Sheets:
        try:
            unprotect_sheet(ws)
            ws.Protect(Password=MASTER_PWD, UserInterfaceOnly=True)
            print(f"  {ws.Name}: OK")
        except:
            print(f"  {ws.Name}: FAIL")


def nuclear(wb):
    """Delete and recreate locked sheets (DANGER)."""
    print("\n=== NUCLEAR: Recreating Sheets ===")
    print("WARNING: This will delete and recreate all data sheets!")

    SHEETS = {
        "ARTICLES": ["Code", "Designation", "Stock", "Seuil Min", "Categorie", "Classe ABC",
                     "Stock Actuel", "PU", "Fournisseur", "Stock Securite", "Notes",
                     "CMUP", "Methode Appro", "Delai"],
        "MOUVEMENTS": ["Date", "Code Article", "Designation", "Type", "Quantite", "Valeur",
                       "Ref Document", "PU", "Tiers", "Notes", "Utilisateur", "Horodatage"],
        "FOURNISSEURS": ["Code", "Raison Sociale", "Adresse", "Telephone", "NIF", "NIS", "RC", "Art. Impot"],
        "CONFIG": ["Parameter", "Value", "Description"],
        "DASHBOARD": [],
        "FACTURES": ["Numero", "Date", "Client", "Adresse", "NIF", "NIS", "RC",
                     "Total HT", "Total TVA", "Total TTC", "Reglement", "Notes", "Utilisateur", "Horodatage"],
        "BARCODES": ["Barcode", "Code Article", "Designation"],
        "BONS_COMMANDE": ["Numero", "Date", "Fournisseur", "Code Article", "Designation",
                          "Quantite", "PU", "Total HT", "Total TVA", "Total TTC",
                          "Etat", "Notes", "Utilisateur", "Horodatage"],
    }

    excel.DisplayAlerts = False

    # Delete locked sheets
    for name in SHEETS:
        try:
            ws = wb.Sheets(name)
            ws.Delete()
            print(f"  Deleted: {name}")
        except:
            pass

    # Create fresh sheets
    for name, headers in SHEETS.items():
        ws = wb.Sheets.Add(After=wb.Sheets(wb.Sheets.Count))
        ws.Name = name
        if headers:
            for col, h in enumerate(headers, 1):
                ws.Cells(1, col).Value = h
            ws.Range("A1:" + chr(64 + len(headers)) + "1").Font.Bold = True
            ws.Range("A1:" + chr(64 + len(headers)) + "1").Interior.Color = 0x7F4600
            ws.Range("A1:" + chr(64 + len(headers)) + "1").Font.Color = 0xFFFFFF
        ws.Protect(Password=MASTER_PWD, UserInterfaceOnly=True)
        print(f"  Created: {name}")

    print("\nDone. Save the workbook now.")


def show_menu():
    """Show usage menu."""
    print("""
Academix DSS — Maintenance Tool
================================
Usage: python tools/maintenance.py <command>

Commands:
  fix-pwd       Fix passwords on all sheets
  fix-accueil   Re-import mod_AccueilButtons
  fix-all       Fix passwords + re-import modules
  verify        Verify all sheets accessible
  nuclear       Delete and recreate locked sheets (DANGER)
  (no args)     Show this menu
""")


if __name__ == "__main__":
    cmd = sys.argv[1] if len(sys.argv) > 1 else None

    if not cmd:
        show_menu()
        sys.exit(0)

    excel = connect_excel()
    wb = find_workbook(excel)
    print(f"Workbook: {wb.Name}")

    if cmd == "fix-pwd":
        fix_passwords(wb)
    elif cmd == "fix-accueil":
        fix_accueil(excel, wb)
    elif cmd == "fix-all":
        fix_all(excel, wb)
    elif cmd == "verify":
        verify(wb)
    elif cmd == "nuclear":
        nuclear(wb)
    else:
        print(f"Unknown command: {cmd}")
        show_menu()
