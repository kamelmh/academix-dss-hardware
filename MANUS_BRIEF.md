# MANUS BRIEF — Academix DSS v14.0

## Project Overview
**Academix DSS** is a VBA/Excel Decision Support System for inventory management in Algerian hardware stores (quincailleries). Built by Mahi Kamel Abdelghani.

**Current Version:** v14.0 (Hardware Store Edition)
**GitHub:** https://github.com/kamelmh/academix-dss-hardware
**Landing Page:** https://academixdss-9jayh829.manus.space
**Demo Dashboard:** https://academix-dss-hardware.streamlit.app

---

## Git Status (as of 2026-08-04)

### Branch: master (up to date with origin/master)

### Modified Files (12)
```
.gitignore
ERP_dss_v13.4_hardware_store.xlsm
FACEBOOK_POSTS.md
docs/SELLABLE_PACKAGE.md
modules/mod_AccueilButtons_v14.bas
modules/mod_Backup_v14.bas
modules/mod_Cleanup.bas
modules/mod_Config.bas
modules/mod_Dashboard.bas
modules/mod_DemoData.bas
modules/mod_PurchaseOrder_v14.bas
modules/mod_Reports.bas
streamlit-dashboard/app.py
```

### Untracked Files (60+)
- **Screenshots:** 20+ PNG files (Facebook group posts, banner previews)
- **Marketing:** ACADEMIX_SALES_PLAYBOOK.md, DM_TEMPLATES.md, WHATSAPP_TEMPLATES.md, PROSPECTION_ELBAYADH.md, AI_MOAT_STRATEGY.md
- **Distribution:** dist/ folder (ZIP, .xlsm, HTML guides, README)
- **Tools:** tools/ (import scripts, fix scripts, packaging)
- **Modules:** modules/ (FixPwd, MasterSetup1, UsageAnalytics, ValidationGate, DiagnoseModConfig, FixModConfig, FixMyDSS, ImportClean, check_features.py)
- **Contact Form:** contact-form/ (Formspree integration)
- **Streamlit Demo:** streamlit-dashboard/pages/3_Demo_Interactif.py
- **Landing Page Source:** compressed/academix-dss_1/ (HTML/CSS/JS)

---

## Current State (v14.0)

### Technical Architecture
- **Frontend:** Excel VBA + UserForms (8 forms, created dynamically)
- **Backend:** 32 VBA modules
- **Database:** Excel sheets (11 sheets)
- **Deploy:** Local .xlsm files
- **Demo:** Streamlit dashboard (Python, deployed on Streamlit Cloud)

### Features Completed
- ✅ First-run wizard (setup on first open)
- ✅ 40 demo articles (pre-loaded)
- ✅ 9 Algerian suppliers (with NIF/NIS/RC)
- ✅ Wilson EOQ (Economic Order Quantity)
- ✅ CMUP (Weighted Average Cost, SCF compliant)
- ✅ ABC Classification (automatic)
- ✅ Barcode printing
- ✅ Invoice generation
- ✅ Delivery notes
- ✅ Dashboard (real-time KPIs)
- ✅ Reports (ABC, aging, supplier performance, stock summary)
- ✅ Backup system
- ✅ Mod_Config (MASTER_PWD = erp_secure_pwd_2026)
- ✅ SafeUnprotect/SafeProtect helpers
- ✅ RefreshAccueilKPIs (fixed)
- ✅ Score-based pricing (15k/25k/45k)

### Landing Page (Manus)
- ✅ Hero section with 2 CTAs
- ✅ Features section
- ✅ Pricing section
- ✅ Contact section
- ✅ Bilingual FR/AR
- ✅ Professional blue/white design
- ✅ Live on https://academixdss-9jayh829.manus.space

### Streamlit Demo Dashboard
- ✅ Multipage app (Home, Dashboard, Interactive Demo, Contact)
- ✅ 40 articles with real Algerian product names
- ✅ Interactive charts (Plotly)
- ✅ Cost calculator
- ✅ ROI calculator
- ✅ Contact form (Gmail SMTP)
- ✅ Mobile-responsive

---

## What Needs Improvements

### 1. Code Quality (Priority: HIGH)
- [ ] Remove unused modules (DiagnoseModConfig, FixModConfig, FixMyDSS, ImportClean)
- [ ] Consolidate fix scripts (tools/ has 10+ redundant fix scripts)
- [ ] Add proper error handling to all modules
- [ ] Document all public functions
- [ ] Add unit tests for critical functions (CMUP, EOQ, ABC)

### 2. Documentation (Priority: HIGH)
- [ ] Create comprehensive README.md for GitHub
- [ ] Add inline code comments (currently minimal)
- [ ] Create API documentation for module interfaces
- [ ] Update USER_GUIDE.md with screenshots
- [ ] Create CHANGELOG.md

### 3. Distribution Package (Priority: MEDIUM)
- [ ] Clean up dist/ folder (remove redundant files)
- [ ] Create professional installer (or self-extracting ZIP)
- [ ] Add license agreement PDF
- [ ] Create quick-start video (screen recording)
- [ ] Add Arabic user guide

### 4. Streamlit Dashboard (Priority: MEDIUM)
- [ ] Add real data import (currently uses synthetic data)
- [ ] Add user authentication
- [ ] Add data export (CSV/Excel)
- [ ] Add more chart types (trend lines, forecasts)
- [ ] Optimize for mobile viewing

### 5. Marketing Materials (Priority: LOW)
- [ ] Create professional flyer (PDF)
- [ ] Record demo video (60 seconds)
- [ ] Create LinkedIn post templates
- [ ] Design business cards
- [ ] Create Google My Business listing

---

## Scaling Plan

### Phase 1: Local Sales (Months 1-3)
**Goal:** 10 customers in El Bayadh
- Sell .xlsm files directly
- 15,000 DZD per customer
- WhatsApp support
- Target: 150,000 DZD revenue

### Phase 2: Regional Expansion (Months 4-6)
**Goal:** 50 customers across Algeria
- Expand to Oran, Algiers, Constantine
- Hire 2 sales reps (commission-based)
- Create referral program
- Target: 750,000 DZD revenue

### Phase 3: SaaS Transition (Months 7-12)
**Goal:** Convert to web-based SaaS
- Build FastAPI backend
- Deploy PostgreSQL database
- Add user authentication
- Launch subscription model (5k-25k DZD/month)
- Target: 600,000 DZD MRR

### Phase 4: AI Integration (Year 2)
**Goal:** Add AI-powered features
- Demand forecasting (Prophet + LightGBM)
- Smart reorder suggestions
- Anomaly detection
- Natural language queries
- Target: 1,200,000 DZD MRR

---

## Files in This Package

### Core Files
| File | Purpose |
|------|---------|
| ERP_dss_v13.4_hardware_store.xlsm | Main application (0.8 MB) |
| modules/*.bas | 32 VBA modules |
| streamlit-dashboard/ | Demo dashboard (Python) |
| contact-form/ | Contact form (HTML) |

### Documentation
| File | Purpose |
|------|---------|
| docs/SELLABLE_PACKAGE.md | Package guide |
| docs/AI_ROADMAP.md | 3-tier AI strategy |
| docs/USER_GUIDE.md | User manual |
| docs/FIRST_RUN.md | Setup guide |
| ACADEMIX_SALES_PLAYBOOK.md | Sales scripts |
| FACEBOOK_POSTS.md | Marketing content |
| DM_TEMPLATES.md | Outreach templates |
| WHATSAPP_TEMPLATES.md | WhatsApp scripts |
| PROSPECTION_ELBAYADH.md | Store list |

### Marketing Materials
| File | Purpose |
|------|---------|
| dist/banner.html | Facebook banner |
| dist/FLYER.html | Printable flyer |
| dist/TARIFICATION.md | Pricing guide |
| dist/DEMARRAGE_RAPIDE.html | Quick start |
| dist/USER_GUIDE.html | User guide |
| dist/README.md | Package README |

### Tools
| File | Purpose |
|------|---------|
| tools/import_nuclear.py | Module importer |
| tools/package_sellable.py | Package builder |
| tools/fix_*.py | Various fix scripts |

---

## Specific Tasks for Manus

### Task 1: Clean Up Repository
1. Remove unused modules (DiagnoseModConfig, FixModConfig, FixMyDSS, ImportClean)
2. Consolidate fix scripts into single tool
3. Remove redundant screenshots from root
4. Create proper .gitignore entries

### Task 2: Create Professional README.md
1. Project description
2. Features list
3. Installation instructions
4. Usage guide
5. Contributing guidelines
6. License
7. Contact info

### Task 3: Optimize Distribution Package
1. Clean dist/ folder
2. Create single ZIP with all deliverables
3. Add professional packaging (folder structure)
4. Include all documentation

### Task 4: Enhance Streamlit Dashboard
1. Add real data import capability
2. Add more interactive features
3. Optimize for mobile
4. Add screenshots to README

### Task 5: Create Marketing Assets
1. Professional flyer (HTML → PDF)
2. Demo video script
3. LinkedIn post templates
4. Business card design

---

## Contact

**Mahi Kamel Abdelghani**
- Email: kamelmahi71@gmail.com
- Phone: +213 676 77 38 92
- GitHub: https://github.com/kamelmh
- Portfolio: https://kamelmahi.netlify.app

---

**Last Updated:** 2026-08-04
**Version:** 1.0
