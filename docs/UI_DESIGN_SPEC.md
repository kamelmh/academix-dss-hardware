# UI Design Spec — Academix DSS v14.0 Hardware Store

> **Purpose:** Detailed design for all UserForms (VBA UI layer)
> **Last Updated:** 2026-07-19
> **Status:** Ready for Claude to implement

---

## Overview

The current DSS has **7 data sheets** (Layer 1). This spec defines the **UserForms** (Layer 2) that provide professional UI for data entry, editing, and reporting.

### Design Principles
- **Bilingual:** Arabic (primary) + French (secondary) labels
- **Professional:** Blue header, consistent spacing, proper alignment
- **Efficient:** Tab order flows naturally, keyboard shortcuts work
- **Safe:** Validation before save, confirmation dialogs, undo support
- **Offline:** No internet required, all VBA local

---

## Form 1: frmStockEntry (CRITICAL)

### Purpose
Add/edit stock movements (ENTREE/SORTIE). This is the **most-used form** — staff enters every stock movement here.

### Layout (600×450 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 حركة مخزون — Mouvement de Stock                    [X] │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  التاريخ / Date:    [___________]  Type: [ENTREE ▼]        │
│                                                             │
│  كود المادة / Code:  [___________]  [🔍 Search]            │
│  التسمية / Designation: [AUTO-FILLED — read-only]          │
│  الكمية / Quantite:  [___________]  PU: [___________]      │
│  القيمة / Valeur:    [AUTO-CALCULATED — read-only]         │
│                                                             │
│  مرجع الوثيقة / Ref Document:  [___________]               │
│  الطرف الثالث / Tiers:          [___________]               │
│  ملاحظات / Notes:               [___________]               │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│  [💾 Enregistrer]  [🗑️ Supprimer]  [📄 Nouveau]  [❌ Fermer] │
└─────────────────────────────────────────────────────────────┘
```

### Controls

| Name | Type | Size | Source/Validation |
|------|------|------|-------------------|
| `lblTitle` | Label | 600×30 | "حركة مخزون — Mouvement de Stock" |
| `txtDate` | TextBox | 150×25 | Default: `Format(Date, "DD/MM/YYYY")`, validate date format |
| `cboType` | ComboBox | 120×25 | List: "ENTREE", "SORTIE" |
| `txtCode` | TextBox | 150×25 | Validate against ARTICLES.Code |
| `btnSearch` | CommandButton | 80×25 | Opens frmSearch, returns selected code |
| `lblDesignation` | Label | 250×25 | Auto-filled when code is entered |
| `txtQuantite` | TextBox | 100×25 | Must be > 0, numeric |
| `txtPU` | TextBox | 100×25 | Auto-filled from ARTICLES.PU, editable |
| `lblValeur` | Label | 150×25 | Auto-calculated: Quantite × PU |
| `txtRefDoc` | TextBox | 200×25 | Invoice/receipt number |
| `txtTiers` | TextBox | 200×25 | Client/supplier name |
| `txtNotes` | TextBox | 300×25 | Free text |
| `btnSave` | CommandButton | 120×35 | Save movement, update stock |
| `btnDelete` | CommandButton | 120×35 | Delete selected movement (with confirmation) |
| `btnNew` | CommandButton | 120×35 | Clear form for new entry |
| `btnClose` | CommandButton | 120×35 | Close form |

### Behavior

#### On Code Exit (`txtCode_Exit`)
```
1. Validate code exists in ARTICLES
2. If found: fill lblDesignation, txtPU from ARTICLES sheet
3. If not found: show warning, clear fields
```

#### On Quantite/PU Change
```
1. Calculate Valeur = Quantite × PU
2. Update lblValeur
```

#### On Save (`btnSave_Click`)
```
1. Validate all required fields (Date, Code, Type, Quantite)
2. Confirm: "Enregistrer cette ENTREE/SORTIE de [Quantite] × [Designation]?"
3. Write to MOUVEMENTS sheet (next empty row)
4. Update ARTICLES.Stock (ENTREE: +Quantite, SORTIE: -Quantite)
5. Update ARTICLES.Stock_Actuel
6. Refresh DASHBOARD
7. Show success: "Movement saved. Stock: [new stock level]"
8. Clear form for next entry
```

#### On Delete (`btnDelete_Click`)
```
1. Confirm: "Supprimer ce mouvement?"
2. Reverse stock change (ENTREE: -Quantite, SORTIE: +Quantite)
3. Delete row from MOUVEMENTS
4. Refresh DASHBOARD
5. Show: "Movement deleted."
```

### Keyboard Shortcuts
- `Ctrl+S` — Save
- `Ctrl+N` — New
- `Ctrl+F` — Search
- `Escape` — Close
- `Enter` — Move to next field (or Save if on last field)

---

## Form 2: frmArticleEditor

### Purpose
Add/edit products in the ARTICLES database.

### Layout (650×500 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 إدارة المواد — Gestion des Articles                [X] │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  كود المادة / Code:        [___________]                    │
│  التسمية / Designation:    [___________]                    │
│  الفئة / Categorie:        [FER ▼]                          │
│  المورد / Fournisseur:     [SUP-001 ▼]                     │
│                                                             │
│  ──── الأسعار / Prix ────                                    │
│  سعر الوحدة / PU:          [___________] DZD               │
│  وحدة القياس / Unite:      [___________] (kg, m, pcs, etc.)│
│                                                             │
│  ──── المخزون / Stock ────                                   │
│  الحد الأدنى / Seuil Min:   [___________]                   │
│  مخزون الأمان / Stock Sec:  [___________]                   │
│  طريقة الطلب / Methode:     [COMMANDER ▼]                   │
│  مدة التوريد / Delai:        [___________] jours             │
│                                                             │
│  ملاحظات / Notes:           [___________]                   │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│  [💾 Enregistrer]  [🗑️ Supprimer]  [📄 Nouveau]  [❌ Fermer] │
└─────────────────────────────────────────────────────────────┘
```

### Categories (ComboBox)
```
FER — Fer & Acier
CEM — Ciment
PVC — Plastique
ELEC — Électricité
OUTIL — Outils
PEINT — Peinture
CARO — Céramique
```

### Suppliers (ComboBox — from FOURNISSEURS)
```
Dynamic: Load from FOURNISSEURS.Raison_Sociale
```

### Behavior
- **New Article:** Auto-generate code (e.g., FER-003)
- **Save:** Validate unique code, write to ARTICLES
- **Delete:** Check no movements exist for this article
- **Search:** Open frmSearch filtered to ARTICLES

---

## Form 3: frmSupplierEditor

### Purpose
Add/edit suppliers in the FOURNISSEURS database.

### Layout (600×400 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 إدارة الموردين — Gestion des Fournisseurs          [X] │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  كود المورد / Code:          [___________]                  │
│  الاسم / Raison Sociale:     [___________]                  │
│  العنوان / Adresse:          [___________]                  │
│  الهاتف / Telephone:         [___________]                  │
│                                                             │
│  ──── الضرائب / Fiscalité ────                               │
│  NIF:                        [___________]                  │
│  NIS:                        [___________]                  │
│  RC:                         [___________]                  │
│  مقال الضريبة / Art. Impot:  [___________]                  │
│                                                             │
├─────────────────────────────────────────────────────────────┤
│  [💾 Enregistrer]  [🗑️ Supprimer]  [📄 Nouveau]  [❌ Fermer] │
└─────────────────────────────────────────────────────────────┘
```

---

## Form 4: frmDashboard

### Purpose
Main navigation form with KPIs and quick actions.

### Layout (800×600 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 لوحة التحكم — Tableau de Bord                    [X]  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ──── KPIs ────                                             │
│  ┌──────────┐ ┌──────────┐ ┌──────────┐ ┌──────────┐       │
│  │ Articles │ │ Stock    │ │ Alerts   │ │ Value    │       │
│  │    40    │ │  1,250   │ │    3     │ │ 64.1M    │       │
│  └──────────┘ └──────────┘ └──────────┘ └──────────┘       │
│                                                             │
│  ──── Quick Actions ────                                    │
│  [📦 حركة مخزون]  [📋 إدارة المواد]  [🏢 إدارة الموردين]   │
│                                                             │
│  [📊 التقارير]  [🔍 بحث]  [⚙️ الإعدادات]  [🚪 خروج]       │
│                                                             │
│  ──── Alerts ────                                           │
│  ┌─────────────────────────────────────────────────────┐   │
│  │ ⚠️ FER-001: Stock (5) < Seuil Min (10)            │   │
│  │ ⚠️ CEM-003: Stock (2) < Seuil Securite (5)        │   │
│  │ ✅ All other articles OK                           │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Form 5: frmSearch

### Purpose
Universal search across articles, suppliers, and movements.

### Layout (500×350 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 بحث — Recherche                                   [X]  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  بحث في / Chercher dans: [الكل ▼]                           │
│  كلمة البحث / Mot-clé:    [___________]                     │
│                                                             │
│  ┌─────────────────────────────────────────────────────┐   │
│  │ ListView: Search results                            │   │
│  │ Code    | Designation  | Stock  | Category         │   │
│  │ FER-001 | Barre fer 10 |   15   | FER              │   │
│  │ FER-002 | Barre fer 8  |   45   | FER              │   │
│  └─────────────────────────────────────────────────────┘   │
│                                                             │
│  [✅ Sélectionner]  [❌ Fermer]                              │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

### Search Filters
- **الكل (All):** Search Code, Designation, Notes
- **المقالات (Articles):** Search ARTICLES only
- **الموردين (Suppliers):** Search FOURNISSEURS only
- **الحركات (Movements):** Search MOUVEMENTS

---

## Form 6: frmReception

### Purpose
Generate and print BON DE RÉCEPTION (delivery note).

### Layout (700×500 pixels — Print Preview)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 بrine de réception                              [X]     │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  ──── HEADER ────                                           │
│  Réception N°: [___________]  Date: [___________]           │
│  Fournisseur:  [___________]                                │
│  Référence:    [___________]                                │
│                                                             │
│  ──── ITEMS ────                                            │
│  ┌─────┬───────────┬──────┬──────┬──────────┐             │
│  │ N°  │ Code      │ Qte  │ PU   │ Valeur   │             │
│  ├─────┼───────────┼──────┼──────┼──────────┤             │
│  │  1  │ FER-001   │  50  │ 150  │  7,500   │             │
│  │  2  │ CEM-001   │ 100  │ 500  │ 50,000   │             │
│  └─────┴───────────┴──────┴──────┴──────────┘             │
│                                                             │
│  Total TVA (19%):              10,925 DZD                  │
│  Total HT:                     57,500 DZD                  │
│  Total TTC:                    68,425 DZD                  │
│                                                             │
│  [🖨️ Imprimer]  [📧 Exporter]  [❌ Fermer]                  │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Form 7: frmReports

### Purpose
Generate reports (ABC analysis, stock aging, supplier performance).

### Layout (600×400 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 التقارير — Rapports                               [X]  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  نوع التقرير / Type:                                       │
│  (●) تحليل ABC — ABC Analysis                              │
│  ( ) تحليل حركة المخز Stock Aging                          │
│  ( ) أداء الموردين — Supplier Performance                   │
│  ( ) حركات الفترة — Period Movements                        │
│                                                             │
│  الفترة / Période: [__/__/__] إلى [__/__/__]               │
│                                                             │
│  [📊 Générer]  [🖨️ Imprimer]  [❌ Fermer]                  │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Form 8: frmConfig

### Purpose
Edit system parameters (CONFIG sheet).

### Layout (500×350 pixels)

```
┌─────────────────────────────────────────────────────────────┐
│  🔹 الإعدادات — Configuration                        [X]  │
├─────────────────────────────────────────────────────────────┤
│                                                             │
│  أيام العمل / Working Days:     [300]                       │
│  تكلفة الطلب / Order Cost:      [500] DZD                  │
│  معدل الاحتفاظ / Holding Rate:  [0.25]                     │
│  مدة التوريد / Lead Time:        [7] jours                  │
│  ضريبة TVA / Tax Rate:           [0.19]                     │
│  فترة المراجعة / Review Period:  [7] jours                  │
│  مستوى الخدمة / Service Level:   [0.95]                     │
│                                                             │
│  [💾 Enregistrer]  [🔄 Réinitialiser]  [❌ Fermer]          │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Color Scheme

| Element | Color | RGB |
|---------|-------|-----|
| Header Background | Dark Blue | RGB(0, 70, 127) |
| Header Text | White | RGB(255, 255, 255) |
| Button Background | Light Blue | RGB(0, 120, 215) |
| Button Text | White | RGB(255, 255, 255) |
| Alert Background | Light Yellow | RGB(255, 255, 224) |
| Error Background | Light Red | RGB(255, 200, 200) |
| Success Background | Light Green | RGB(200, 255, 200) |
| Form Background | Light Gray | RGB(240, 240, 240) |

---

## Implementation Order

1. **frmStockEntry** — Most used, start here
2. **frmArticleEditor** — Product management
3. **frmSupplierEditor** — Supplier management
4. **frmDashboard** — Main navigation
5. **frmSearch** — Search functionality
6. **frmReception** — Print documents
7. **frmReports** — Reporting
8. **frmConfig** — Settings

---

## VBA Pattern for All Forms

```vba
' Form module pattern
Option Explicit

Private Sub UserForm_Initialize()
    Me.Caption = "حركة مخزون — Mouvement de Stock"
    Call LoadComboBoxes
    Call ClearForm
End Sub

Private Sub LoadComboBoxes()
    ' Populate dropdowns from sheets
End Sub

Private Sub ClearForm()
    ' Reset all controls to defaults
End Sub

Private Sub btnSave_Click()
    ' Validate
    If Not ValidateForm Then Exit Sub
    ' Confirm
    If MsgBox("Enregistrer?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
    ' Save
    Call SaveData
    ' Refresh
    Call RefreshDashboard
    ' Clear
    Call ClearForm
End Sub

Private Function ValidateForm() As Boolean
    ' Check required fields
    ValidateForm = True
End Function

Private Sub SaveData()
    ' Write to sheet
End Sub
```

---

**Last Updated:** 2026-07-19
**Version:** 1.0

#hardware-store #project/dss #tech/vba #ui-design #userforms
