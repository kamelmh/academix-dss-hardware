# Academix v14.0 - Hardware Store Deployment

## Overview
5 VBA modules updated for hardware store deployment (Quincaillerie El Bayadh).

## Files

| File | Purpose | Changes |
|------|---------|---------|
| `mod_Config_v14.bas` | System configuration | Reads from CONFIG sheet, 14 new parameters, version "v14.0" |
| `mod_DemoData_v14.bas` | Demo data generator | 40 hardware articles, 9 suppliers, 90-day movements |
| `mod_StockEngine_v14.bas` | Stock management | Reads ORDER_COST/HOLDING_RATE/LEAD_TIME from CONFIG (no hardcoded values) |
| `mod_SupplierRegistry_v14.bas` | Supplier database | 9 hardware suppliers with Algerian tax IDs (NIF/NIS/RC) |
| `mod_Dashboard_v14.bas` | Dashboard KPIs | Updated for 40 articles, top 10 critical items |

## Import Instructions

1. Open `ERP_v13.4_CLEAN.xlsm`
2. Press `Alt+F11` to open VBA Editor
3. For each file:
   - Right-click on Modules folder
   - Select Import File
   - Navigate to `VBA_NEW\` folder
   - Select the `.bas` file
4. The new module will appear in the project
5. **IMPORTANT**: Delete the old module with the same name (right-click > Remove)

## Post-Import Steps

1. Open CONFIG sheet and verify parameters:
   - `WORKING_DAYS` = 250
   - `ORDER_COST` = 801.45
   - `HOLDING_RATE` = 0.2
   - `LEAD_TIME` = 2
   - `TAX_RATE` = 0.19
   - `BUSINESS_NAME` = Quincaillerie El Bayadh

2. Run `GenerateDemoData` macro:
   - Press `Alt+F8`
   - Select `GenerateDemoData`
   - Click Run

3. Run `RefreshDashboard` macro to verify

## 40 Hardware Articles

| Category | Count | Examples |
|----------|-------|---------|
| Fer/Acier | 8 | Fer a beton 10-16mm, Acier filete, Tole |
| Ciment/Granulat | 7 | CPA 55, CPA 45, Sable, Gravier |
| Plomberie | 8 | Tuyau PVC 32-160mm, Raccords, Robinetterie |
| Electricite | 7 | Cable 2.5-6mm2, Disjoncteurs, Prises |
| Outillage/Peinture/Carrelage | 10 | Perceuse, Peinture 25L, Carrelage |

## 9 Suppliers

| Code | Name | Location | Specialty |
|------|------|----------|-----------|
| SIDERAL | SIDERAL SPA | El Bayadh | Acier et metaux |
| CIMENTAL | CIMENTAL SPA | Alger | Ciment |
| GRANULATS | GRANULATS SA | Tipaza | Sable et Gravier |
| PLASTIQUE | PLASTIQUE PRO | Oran | Tuyauterie PVC |
| ELECTRO | ELECTRO PLUS | Constantine | Cablage |
| DISTRIBUTION | DISTRIBUTION ELECTRIQUE | Setif | Electricite |
| OUTILMAG | OUTILMAG | Alger | Outillage |
| PEINTURE | PEINTURE PLUS | Blida | Peinture |
| CERAMIQUE | CERAMIQUE EL BAYADH | El Bayadh | Carrelage |

## Key Changes from v13.4

1. **No hardcoded values** - All parameters read from CONFIG sheet
2. **Hardware store data** - 40 items across 6 categories
3. **Algerian suppliers** - Realistic tax IDs (NIF/NIS/RC)
4. **Spring season** - 90-day movement patterns (April-June 2026)
5. **Updated dashboard** - Top 10 critical items instead of 5
6. **Configurable business info** - Name, address, phone, tax IDs
