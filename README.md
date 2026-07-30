# Academix DSS — Hardware Store

VBA inventory management system for Quincaillerie El Bayadh (hardware store).

## Domain

Hardware store — fer/acier, ciment, plomberie, electricite, outillage, peinture, carrelage.

## Modules

### Shared (6) — canonical copies
`mod_Dashboard` (12.7 KB), `mod_DemoData` (38.3 KB), `mod_StockEngine` (22.6 KB), `mod_SupplierRegistry` (13.6 KB), `mod_Barcode` (10.4 KB), `mod_Reports` (13.9 KB)

### Hardware-Specific (21)
| Module | Purpose |
|--------|---------|
| `MAIN_MACROS_CLEAN.bas` | Entry point — macro runner |
| `mod_AccueilButtons_v14.bas` | Home page buttons |
| `mod_AccueilDesign_v14.bas` | Home page design |
| `mod_Backup_v14.bas` | Backup/restore |
| `mod_BuildArticleEditor_v14.bas` | Article editor form |
| `mod_BuildConfig_v14.bas` | Config form builder |
| `mod_BuildDashboard_v14.bas` | Dashboard form builder |
| `mod_BuildForm_v14.bas` | Generic form builder |
| `mod_BuildReception_v14.bas` | Reception form builder |
| `mod_BuildReports_v14.bas` | Reports form builder |
| `mod_BuildSearch_v14.bas` | Search form builder |
| `mod_BuildSupplierEditor_v14.bas` | Supplier editor form |
| `mod_Cleanup.bas` | Cleanup utilities |
| `mod_ErrorHandler_v14.bas` | Error handling |
| `mod_FormHelpers_v14.bas` | Form helpers |
| `mod_Invoice_v14.bas` | Invoice generation |
| `mod_MasterSetup_v14.bas` | Master setup |
| `mod_PurchaseOrder_v14.bas` | Purchase orders |
| `mod_StockEntryHelpers.bas` | Stock entry helpers |
| `mod_Branding.bas` | Version info, about dialog, branding colors |
| `mod_Splash.bas` | Splash screen, loading indicators, branded messages |

## Data

- 40 hardware articles across 6 categories
- 9 Algerian suppliers with NIF/NIS/RC tax IDs
- 90-day movement patterns (spring season)

## Import

See `docs/IMPORT_INSTRUCTIONS.md` for VBA import steps.

## Dependencies

- Excel 2010+ (VBA host)
- No external libraries

## License

MIT — Mahi Kamel Abdelghani
