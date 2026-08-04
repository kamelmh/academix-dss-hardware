# Academix DSS — Hardware Store Edition

> **Le premier système de gestion de stock** spécialement conçu pour les quincailleries en Algérie.

[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Excel: 2016+](https://img.shields.io/badge/Excel-2016%2B-blue.svg)](https://www.microsoft.com/excel)
[![VBA: Built-in](https://img.shields.io/badge/VBA-Built--in-green.svg)](https://docs.microsoft.com/office/vba/api/overview/)

---

## Features

| Feature | Description |
|---------|-------------|
| **Real-Time Dashboard** | Track stocks, alerts, and performance instantly |
| **Smart Alerts** | Automatic warnings before stockouts |
| **ABC Classification** | Identify your most valuable items |
| **CMUP Calculation** | Weighted average cost (SCF compliant) |
| **Wilson EOQ** | Economic Order Quantity optimization |
| **Invoice Generation** | Create and print professional invoices |
| **Barcode Printing** | Generate barcodes for articles |
| **Supplier Management** | Track suppliers with NIF/NIS/RC tax IDs |
| **Backup System** | Automatic backups on every open |
| **Bilingual UI** | French and Arabic interface |

---

## Quick Start

### 1. Download
```bash
git clone https://github.com/kamelmh/academix-dss-hardware.git
```

### 2. Open
- Open `ERP_dss_v13.4_hardware_store.xlsm` in Excel 2016+
- Enable macros when prompted
- Trust VBA project access: `File > Options > Trust Center > Macro Settings > Trust access to VBA project object model`

### 3. Start
- Follow the first-run wizard
- 40 demo articles pre-loaded
- Start managing your inventory!

---

## Screenshots

| Dashboard | Reports | Alerts |
|-----------|---------|--------|
| ![Dashboard](docs/screenshots/EXCEL_0bdi9zE94p.png) | ![Reports](docs/screenshots/EXCEL_4BmvZMjwIV.png) | ![Alerts](docs/screenshots/EXCEL_95N1mRFEij.png) |

---

## Demo

### Live Streamlit Dashboard
**https://academix-dss-hardware.streamlit.app**

- **Home** — Overview and quick links
- **Dashboard** — Real-time KPIs, charts, alerts
- **Interactive Demo** — Step-by-step walkthrough
- **Contact** — Get in touch

### Landing Page
**https://academixdss-9jayh829.manus.space**

---

## Installation

### System Requirements
- Windows 7 or later
- Microsoft Excel 2016 or later
- 50 MB free disk space
- Macro support enabled

### First-Time Setup
1. Copy `ERP_dss_v13.4_hardware_store.xlsm` to your computer
2. Open with Excel 2016 or later
3. Enable macros when prompted
4. **Important:** Trust VBA project access:
   - `File > Options > Trust Center > Trust Center Settings > Macro Settings`
   - Check **"Trust access to VBA project object model"**
5. Follow the setup wizard
6. Start using the system!

---

## Project Structure

```
academix-dss-hardware/
├── ERP_dss_v13.4_hardware_store.xlsm    # Main application
├── modules/                              # 32 VBA modules
│   ├── mod_Config.bas                   # Configuration & constants
│   ├── mod_Dashboard.bas                # Dashboard & KPIs
│   ├── mod_StockEngine.bas              # Stock management
│   ├── mod_AuditTrail.bas               # Audit logging
│   └── ...                              # 28 more modules
├── streamlit-dashboard/                 # Demo dashboard (Python)
│   ├── app.py                           # Main entry
│   └── pages/                           # 4 pages
├── tools/                               # Maintenance scripts
│   ├── maintenance.py                   # Consolidated maintenance
│   ├── import_nuclear.py                # Module importer
│   └── package_sellable.py              # Package builder
├── dist/                                # Distribution files
├── docs/                                # Documentation
└── contact-form/                        # Contact form (HTML)
```

---

## Documentation

| Document | Description |
|----------|-------------|
| [USER_GUIDE.md](docs/USER_GUIDE.md) | Complete user manual |
| [FIRST_RUN.md](docs/FIRST_RUN.md) | First-run setup guide |
| [SELLABLE_PACKAGE.md](docs/SELLABLE_PACKAGE.md) | Package & pricing guide |
| [AI_ROADMAP.md](docs/AI_ROADMAP.md) | 3-tier AI strategy |
| [ACADEMIX_SALES_PLAYBOOK.md](ACADEMIX_SALES_PLAYBOOK.md) | Sales scripts & templates |

---

## Pricing

### Score-Based Pricing (Launch)
| Milestone | Price | Articles |
|-----------|-------|----------|
| **First 10 customers** | 15,000 DZD | Up to 100 |
| After 10 customers | 25,000 DZD | Up to 500 |
| After 25 customers | 45,000 DZD | Unlimited |
| After 50 customers | 60,000+ DZD | Unlimited |

### Add-Ons
| Service | Price |
|---------|-------|
| On-site training (2h) | 5,000 DZD |
| Custom branding | 10,000 DZD |
| Additional user license | 10,000 DZD |
| Annual support renewal | 8,000 DZD |

---

## Tech Stack

- **Frontend:** Excel VBA + UserForms (8 forms)
- **Backend:** 32 VBA modules
- **Database:** Excel sheets (11 sheets)
- **Demo:** Streamlit (Python)
- **Deploy:** Local .xlsm files

---

## Contributing

1. Fork the repository
2. Create your feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

---

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

---

## Contact

**Mahi Kamel Abdelghani**
- Email: kamelmahi71@gmail.com
- Phone: +213 676 77 38 92
- GitHub: [kamelmh](https://github.com/kamelmh)
- Portfolio: [kamelmahi.netlify.app](https://kamelmahi.netlify.app)
- LinkedIn: [Mahi Kamel](https://linkedin.com/in/mahikamel)

---

## Support

If you find this project helpful, please give it a star on GitHub!
