# HARDWARE STORE DSS — Master Context for Claude

> **Purpose:** Full context for Claude GUI sessions working on the Academix DSS v14.0 hardware store deployment.
> **Last Updated:** 2026-07-19

---

## 1. Project Identity

| Field          | Details                                                                                   |
| -------------- | ----------------------------------------------------------------------------------------- |
| **Project**    | Academix DSS v14.0 — Hardware Store (Quincaillerie El Bayadh)                             |
| **Target**     | `C:\Users\Admin\Projects\active\ERP_dss_v13.4_hardware_store.xlsm`                        |
| **VBA Source** | `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\Academix_DSS\VBA_NEW\`         |
| **Author**     | MAHI Kamel Abdelghani                                                                     |
| **Location**   | El Bayadh, Algeria                                                                        |
| **Business**   | Small hardware store (quincaillerie) — iron, cement, PVC, electrical, tools, paint, tiles |

---

## 2. Architecture — Two Layers

### Layer 1: Data Sheets (CURRENT — 7 sheets)
These are the **raw data tables**. Standard Excel, no UI.

| Sheet | Columns | Purpose |
|-------|---------|---------|
| **ACCUEIL** | Dashboard branded homepage | System landing page |
| **ARTICLES** | 14 cols (Code, Designation, Stock, Seuil Min, Categorie, Classe ABC, Stock Actuel, PU, Fournisseur, Stock Securite, Notes, CMUP, Methode Appro, Delai) | Product database |
| **MOUVEMENTS** | 10 cols (Date, Code Article, Designation, Type, Quantite, Valeur, Ref Document, PU, Tiers, Notes) | Stock movements (ENTREE/SORTIE) |
| **FOURNISSEURS** | 8 cols (Code, Raison Sociale, Adresse, Telephone, NIF, NIS, RC, Art. Impot) | Supplier database |
| **CONFIG** | Key-Value pairs (A:B) | All system parameters |
| **DASHBOARD** | KPIs, projection table, ABC summary | Live metrics |
| **BON_RECEPTION** | Arabic/French bilingual template | Reception form template |

### Layer 2: UserForms (TO BUILD — the real UI)
Professional VBA UserForms that replace manual sheet editing.

| Form | Purpose | Priority |
|------|---------|----------|
| **frmStockEntry** | Add/edit stock movements (ENTREE/SORTIE) | ⭐⭐⭐ Critical |
| **frmArticleEditor** | Add/edit products | ⭐⭐⭐ Critical |
| **frmSupplierEditor** | Add/edit suppliers | ⭐⭐ High |
| **frmDashboard** | Main navigation + KPIs | ⭐⭐ High |
| **frmSearch** | Search articles by code/name/category | ⭐⭐ High |
| **frmReception** | Print reception document | ⭐ Medium |
| **frmReports** | Generate reports (ABC, stock aging, etc.) | ⭐ Medium |
| **frmConfig** | Edit system parameters | ⭐ Low |

---

## 3. Data Structures

### ARTICLES (14 columns)
```
Col 1:  Code           — FER-001, FER-002, CEM-001, etc. (unique ID)
Col 2:  Designation    — Product name (Arabic/French)
Col 3:  Stock          — Current stock quantity
Col 4:  Seuil Min      — Minimum threshold (alert if below)
Col 5:  Categorie      — Category (FER, CEM, PVC, ELEC, OUTIL, PEINT, CARO)
Col 6:  Classe ABC     — A/B/C classification (auto-calculated)
Col 7:  Stock Actuel   — Actual physical count
Col 8:  PU             — Unit price (DZD)
Col 9:  Fournisseur    — Supplier code
Col 10: Stock Securite — Safety stock level
Col 11: Notes          — Free text
Col 12: CMUP           — Weighted average cost (auto-calculated)
Col 13: Methode Appro  — Reorder method (COMMANDER / COMMANDE_AUTO)
Col 14: Delai          — Lead time (days)
```

### MOUVEMENTS (10 columns)
```
Col 1:  Date           — Movement date (DD/MM/YYYY)
Col 2:  Code Article   — Links to ARTICLES.Code
Col 3:  Designation    — Auto-filled from ARTICLES
Col 4:  Type           — ENTREE or SORTIE
Col 5:  Quantite       — Quantity moved
Col 6:  Valeur         — Total value (PU × Quantite)
Col 7:  Ref Document   — Invoice/receipt number
Col 8:  PU             — Unit price at time of movement
Col 9:  Tiers          — Third party (client/supplier name)
Col 10: Notes          — Free text
```

### FOURNISSEURS (8 columns)
```
Col 1:  Code           — SUP-001, SUP-002, etc.
Col 2:  Raison Sociale — Company name
Col 3:  Adresse        — Full address
Col 4:  Telephone      — Phone number
Col 5:  NIF            — Tax ID (Numéro d'Identification Fiscale)
Col 6:  NIS            — Statistical ID
Col 7:  RC             — Trade register number
Col 8:  Art. Impot     — Tax article
```

### CONFIG (Key-Value pairs)
```
WORKING_DAYS      — Working days per year (default: 300)
ORDER_COST        — Fixed order cost in DZD (default: 500)
HOLDING_RATE      — Annual holding rate (default: 0.25)
LEAD_TIME         — Default lead time in days (default: 7)
TAX_RATE          — VAT rate (default: 0.19)
REVIEW_PERIOD     — Review period in days (default: 7)
SERVICE_LEVEL     — Target service level (default: 0.95)
```

---

## 4. Product Categories (Quincaillerie)

| Code | Category | Products |
|------|----------|----------|
| FER | Fer & Acier | Iron bars, sheets, pipes, wire |
| CEM | Ciment | Cement, plaster, lime |
| PVC | Plastique | PVC pipes, fittings, connectors |
| ELEC | Électricité | Wire, switches, outlets, breakers |
| OUTIL | Outils | Hand tools, power tools |
| PEINT | Peinture | Paint, brushes, rollers, thinner |
| CARO | Céramique | Tiles, grout, adhesive |

---

## 5. Algerian Business Context

- **Currency:** DZD (Dinar Algérien)
- **Tax:** TVA 19% (standard rate)
- **Documents:** BON DE RÉCEPTION (delivery note), FACTURE (invoice)
- **Language:** Arabic (primary) + French (secondary)
- **Regulations:** NIF, NIS, RC required for suppliers
- **Working hours:** Typically 8:00-12:00, 14:00-18:00
- **Weekend:** Friday-Saturday (some shops open Saturday morning)

---

## 6. VBA Modules (Current)

| Module | Purpose | Status |
|--------|---------|--------|
| `mod_Config_v14.bas` | Constants, sheet names, column indices | ✅ Complete |
| `mod_DemoData_v14.bas` | Seed 40 articles, 9 suppliers, 90 days | ✅ Complete |
| `mod_StockEngine_v14.bas` | EOQ, ROP, ABC, CMUP calculations | ✅ Complete |
| `mod_Dashboard_v14.bas` | KPI updates, projection table | ✅ Complete |
| `mod_SupplierRegistry_v14.bas` | Supplier CRUD operations | ✅ Complete |
| `mod_Cleanup.bas` | Nuclear reset, sheet management | ✅ Complete |

---

## 7. What Claude Should Know

### DO:
- Always use `mod_Config` constants for column indices (never hardcode numbers)
- Protect sheets after writing (use `ws.Unprotect` / `ws.Protect` pattern)
- Use `Application.ScreenUpdating = False` for bulk operations
- Handle errors with `On Error GoTo ErrHandler`
- Use bilingual labels (Arabic + French) for all UI elements
- Test with 40 articles + 90 days of movements

### DON'T:
- Never hardcode sheet names (use `SHEET_ARTICLES`, `SHEET_MOUVEMENTS`, etc.)
- Never skip error handling
- Never leave sheets unprotected
- Never assume English-only context
- Never use `Select` or `Activate` (use direct object references)
- Never delete the CONFIG sheet

### PATTERNS:
```vba
' Standard movement entry
Dim ws As Worksheet
Set ws = ThisWorkbook.Sheets(SHEET_MOUVEMENTS)
ws.Unprotect Password:=MASTER_PWD
' ... write data ...
ws.Protect Password:=MASTER_PWD, DrawingObjects:=True, Contents:=True

' Get article code from user input
Dim articleCode As String
articleCode = UCase(Trim(txtCode.Value))

' Validate article exists
If Not ArticleExists(articleCode) Then
    MsgBox "Article not found: " & articleCode, vbExclamation
    Exit Sub
End If
```

---

## 8. Screenshots (Current State)

| Screenshot | Sheet | Shows |
|------------|-------|-------|
| `accuil_sheet.PNG` | ACCUEIL | Branded homepage |
| `articles_sheet.PNG` | ARTICLES | 40 products with FER-001 codes |
| `mouvements_sheet.PNG` | MOUVEMENTS | 90 days of ENTREE/SORTIE |
| `fournisseurs_sheet.PNG` | FOURNISSEURS | 9 suppliers with tax IDs |
| `config_sheet.PNG` | CONFIG | System parameters |
| `dashboard_sheet.PNG` | DASHBOARD | KPIs, 64M DZD stock value |
| `bon_reception_sheet.PNG` | BON_RECEPTION | Arabic/French template |

**Note:** These show the **data layer** (standard sheets). The **UI layer** (UserForms) is what we build next.

---

## 9. Next Steps

1. **Build frmStockEntry** — Main data entry form
2. **Build frmArticleEditor** — Product management
3. **Build frmSupplierEditor** — Supplier management
4. **Build frmDashboard** — Main navigation form
5. **Build frmSearch** — Search functionality
6. **Polish ACCUEIL** — Add buttons to launch forms
7. **Take new screenshots** — Show the UI layer

---

## 10. File Locations Summary

| What | Where |
|------|-------|
| Active workbook | `C:\Users\Admin\Projects\active\ERP_dss_v13.4_hardware_store.xlsm` |
| VBA source | `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\Academix_DSS\VBA_NEW\` |
| Screenshots | `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\Academix_DSS\Screenshots\Hardware_Store\` |
| Brain Map | `C:\Users\Admin\My Drive\LifeWorkspace\00-Brain-Map.md` |
| Academix MOC | `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\Academix_DSS\00-MOC-Academix.md` |
| Completed Projects | `C:\Users\Admin\My Drive\LifeWorkspace\04_Ideas_&_Projects\Completed_Projects\` |

---

**Last Updated:** 2026-07-19
**Version:** 1.0

#hardware-store #project/dss #tech/vba #claude-context #quincaillerie
