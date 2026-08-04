# MASTER PROMPT FOR MANUS

## You are helping Mahi Kamel Abdelghani improve and scale Academix DSS v14.0

### Project Context
Academix DSS is a VBA/Excel inventory management system for Algerian hardware stores (quincailleries). It's a working product with 32 VBA modules, 40 demo articles, and a Streamlit demo dashboard. The goal is to clean up the codebase, create professional documentation, and prepare for scaling from local sales to SaaS.

### Files Provided (MANUS_PACKAGE_ACADEMIX_DSS.zip)
The ZIP contains:
1. **Core Application:** ERP_dss_v13.4_hardware_store.xlsm (main .xlsm file)
2. **MANUS_BRIEF.md:** Comprehensive project overview with git status, current state, improvements needed, and scaling plan
3. **Documentation:** SELLABLE_PACKAGE.md, AI_ROADMAP.md, USER_GUIDE.md, FIRST_RUN.md, README.md
4. **Sales Materials:** ACADEMIX_SALES_PLAYBOOK.md, FACEBOOK_POSTS.md, DM_TEMPLATES.md, WHATSAPP_TEMPLATES.md, PROSPECTION_ELBAYADH.md
5. **Strategy:** AI_MOAT_STRATEGY.md (3-tier AI roadmap)
6. **Distribution:** dist/ folder (ZIP, flyer, pricing, quick-start guide)
7. **Streamlit Dashboard:** Multipage Python app (Home, Dashboard, Interactive Demo, Contact)
8. **Key VBA Modules:** mod_Config, mod_AccueilButtons, mod_Dashboard, mod_StockEngine, mod_AuditTrail, mod_FixPwd
9. **Tools:** import_nuclear.py, package_sellable.py
10. **Contact Form:** HTML form with Formspree integration

### What You Need To Do

#### Task 1: Repository Cleanup (Priority: HIGH)
- Remove unused VBA modules: DiagnoseModConfig, FixModConfig, FixMyDSS, ImportClean
- Consolidate the 10+ fix scripts in tools/ into a single maintenance tool
- Clean up screenshots from root directory (move to docs/screenshots/)
- Update .gitignore to exclude build artifacts and temporary files

#### Task 2: Professional README.md (Priority: HIGH)
Create a comprehensive README.md that includes:
- Project banner/header
- One-line description
- Features list with icons
- Quick start guide (3 steps)
- Installation instructions
- Usage examples
- screenshots section
- Contributing guidelines
- License (MIT)
- Contact info + links

#### Task 3: Distribution Package Optimization (Priority: MEDIUM)
- Create clean dist/ folder structure
- Package all deliverables into single ZIP
- Add professional folder structure (docs/, assets/, scripts/)
- Include all documentation in PDF-ready format

#### Task 4: Streamlit Dashboard Enhancement (Priority: MEDIUM)
- Add real data import capability (Excel → PostgreSQL)
- Add more interactive features (filters, date ranges)
- Optimize for mobile viewing
- Add screenshot capture for README

#### Task 5: Marketing Assets (Priority: LOW)
- Create professional flyer (HTML → PDF conversion)
- Write demo video script (60 seconds)
- Create LinkedIn post templates (3 variations)
- Design business card template

### Success Criteria
1. Clean repository with no redundant files
2. Professional README that impresses visitors
3. Optimized distribution package
4. Enhanced demo dashboard
5. Complete marketing asset library

### Constraints
- Keep the VBA/Excel architecture (don't rewrite in Python)
- Maintain backward compatibility with existing customers
- Keep the landing page on Manus (don't migrate)
- Keep the Streamlit demo separate from the main product

### Contact
**Mahi Kamel Abdelghani**
- Email: kamelmahi71@gmail.com
- Phone: +213 676 77 38 92
- GitHub: https://github.com/kamelmh
- Portfolio: https://kamelmahi.netlify.app
