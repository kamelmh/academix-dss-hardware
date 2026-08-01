# Academix DSS — HyperAgent Context

## Project Overview

**Academix DSS** is a complete VBA/Excel inventory management system for Algerian hardware stores (quincailleries). Built by Mahi Kamel Abdelghani.

## Git Repos

| Repo | URL | Purpose |
|------|-----|---------|
| `academix-dss-hardware` | https://github.com/kamelmh/academix-dss-hardware | **Sellable product** — 30 modules, 9,254 lines |
| `logistics-public-sector-refactor` | https://github.com/kamelmh/logistics-public-sector-refactor | **Main dev** — 74 modules, 16,831 lines |

## Current Version: v14.0

### What's Built
- First-run wizard (demo data + clean start)
- 8 UserForms (frmStockEntry, frmArticleEditor, frmSupplierEditor, frmDashboard, frmSearch, frmReception, frmReports, frmConfig)
- CMUP (weighted average cost), ABC classification, Wilson EOQ
- BTS formulas: DSI, Turnover Ratio, Stock Coverage, Variance
- Invoice generation with barcode printing
- Auto-backup system
- Dashboard with live KPIs
- Debug logging in mod_FirstRun and mod_AccueilButtons_v14

### Test Results (v14.0)
| Test | Status |
|------|--------|
| Wizard "Charger donnees demo" | ✅ 40 articles, 9 suppliers, 90 days |
| Wizard "Demarrer a vide" | ✅ Empty ledger, config saved |
| All 8 UserForms | ✅ Load without errors |
| ACCUEIL KPI refresh | ✅ FIXED — shows 0/0/0 after clean start |
| Auto-backup | ✅ 5 backups in session |
| XHTML exports | ✅ Demo (464KB) + Clean (33KB) |

### VBA Project Structure (from .xlsm)
**Sheets (14):** AUDIT_LOG, BON_RECEPTION, ACCUEIL, CONFIG, MOUVEMENTS, FACTURES, FOURNISSEURS, ARTICLES, DASHBOARD, BARCODES, BONS_COMMANDE, + 3 more

**UserForms (9):** frmArticleEditor, frmConfig, frmDashboard, frmFirstRun, frmReception, frmReports, frmSearch, frmStockEntry, frmSupplierEditor

**Modules (30):** Listed in `academix-dss-hardware/modules/` directory

## Files for Analysis

### In This Directory
- `test_academix/ERP_dss_v13.4_hardware_store.xlsm` — Working .xlsm with all modules imported (611KB)
- `test_academix/ERP_dss_v13.4_hardware_store_with_demo_data.xhtml` — Full demo export (464KB)
- `test_academix/ERP_dss_v13.4_hardware_store_without_demo_data.xhtml` — Clean version export (33KB)

### In Git Repo
- `docs/TEST_REPORT_v14.md` — Full test results
- `docs/FIRST_RUN.md` — First-run flow documentation
- `docs/USER_GUIDE.md` — 14-section FR/EN guide
- `docs/SELLABLE_PACKAGE.md` — Pricing, distribution, support
- `docs/QUICK_START.md` — 30-second setup guide
- `docs/LINKEDIN_POST.md` — EN/FR launch posts
- `modules/*.bas` — All 30 VBA modules

## Pricing
- Basic: 45,000 DZD (single-user, 40 demo articles)
- Professional: 75,000 DZD (unlimited, phone support)
- Enterprise: 120,000 DZD (multi-user, priority support)

## What HyperAgent Should Analyze

### 1. Code Quality Audit
- Review all 30 modules for bugs, edge cases, error handling
- Check for hardcoded values that should be configurable
- Verify CMUP formula correctness
- Check ABC classification logic
- Validate Wilson EOQ calculations

### 2. Missing Features
- What's needed for production sale?
- Payment tracking (CCP/BaridiMob integration?)
- Multi-currency support?
- User authentication/permissions?
- Data import/export (CSV, Excel)?
- Print templates optimization?

### 3. Documentation Gaps
- Is USER_GUIDE.md complete?
- Are there missing setup instructions?
- Is the pricing strategy correct for Algeria?

### 4. Distribution Strategy
- How to deliver the .xlsm to customers?
- Trial version vs full version?
- License key system needed?
- Update mechanism?

### 5. Web App Conversion
- What FastAPI endpoints are needed?
- Database schema design (PostgreSQL?)
- Frontend framework choice?
- Authentication flow?

## Collaboration

HyperAgent should:
1. Analyze the .xhtml exports (structured HTML with all VBA code)
2. Review the git repos for code quality
3. Identify missing features and bugs
4. Suggest improvements with specific file references
5. Create action items prioritized by impact

## Contact
- **Name:** Mahi Kamel Abdelghani
- **Email:** kamelmahi71@gmail.com
- **Phone:** +213 676 77 38 92
- **Location:** El Bayadh, Algeria
