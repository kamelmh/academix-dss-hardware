# Academix DSS — Web App Conversion Plan

## Current State
- **Platform:** VBA/Excel (.xlsm)
- **Users:** Single-user, local only
- **Data:** Excel sheets (ARTICLES, MOUVEMENTS, FOURNISSEURS, CONFIG, etc.)
- **Features:** Wilson EOQ, CMUP, ABC classification, barcode printing, invoicing

## Web App Options

### Option 1: Full Web App (React + Node.js + PostgreSQL)
**Effort:** 3-6 months  
**Cost:** High  
**Pros:**
- Modern, scalable, accessible anywhere
- Multi-user support
- Real-time collaboration
- Mobile-friendly

**Cons:**
- Complete rewrite from scratch
- Loses Excel familiarity
- Requires hosting infrastructure
- Ongoing maintenance costs

### Option 2: Excel Online Integration
**Effort:** 1-2 months  
**Cost:** Low  
**Pros:**
- Keeps Excel interface
- Works in browser
- No rewrite needed

**Cons:**
- Requires Microsoft 365 subscription
- Limited VBA support in Excel Online
- Not truly web-native

### Option 3: Hybrid (Excel + Web API) ⭐ RECOMMENDED
**Effort:** 2-3 months  
**Cost:** Medium  
**Pros:**
- Keeps Excel as frontend (familiar)
- Adds web API for data sync
- Multi-device access
- No rewrite of business logic

**Cons:**
- Complex architecture
- Requires API hosting
- Excel must be installed

### Option 4: Power Apps Migration
**Effort:** 2-4 months  
**Cost:** Medium  
**Pros:**
- Microsoft ecosystem
- Low-code development
- Mobile app support

**Cons:**
- Licensing costs
- Limited customization
- Learning curve

### Option 5: Google Sheets Migration
**Effort:** 2-3 months  
**Cost:** Low  
**Pros:**
- Free
- Accessible anywhere
- Collaborative

**Cons:**
- Different scripting language (Apps Script)
- Limited advanced features
- Performance concerns

## Recommended Approach: Hybrid (Option 3)

### Architecture

```
┌─────────────────────────────────────────────────────────────┐
│                    Client Layer                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  Excel Desktop (VBA Frontend)                       │    │
│  │  - UserForms (UI)                                   │    │
│  │  - Business Logic (Wilson EOQ, CMUP, ABC)           │    │
│  │  - Local Cache (Offline Mode)                       │    │
│  └─────────────────────────────────────────────────────┘    │
│                          │                                   │
│                          ▼                                   │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  HTTP/REST API Client (VBA)                         │    │
│  │  - XMLHTTP for REST calls                           │    │
│  │  - JSON parsing via VBA-JSON                        │    │
│  └─────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────┘
                          │
                          ▼ HTTP/REST
┌─────────────────────────────────────────────────────────────┐
│                    Server Layer                              │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  Node.js/Express API Server                         │    │
│  │  - REST endpoints                                   │    │
│  │  - Authentication (JWT)                             │    │
│  │  - Rate limiting                                    │    │
│  └─────────────────────────────────────────────────────┘    │
│                          │                                   │
│                          ▼                                   │
│  ┌─────────────────────────────────────────────────────┐    │
│  │  PostgreSQL Database                                │    │
│  │  - Articles, Movements, Suppliers                   │    │
│  │  - Audit log                                        │    │
│  │  - User accounts                                    │    │
│  └─────────────────────────────────────────────────────┘    │
└─────────────────────────────────────────────────────────────┘
```

### Phase 1: API Layer (Month 1)
1. **Create REST API endpoints:**
   - `GET /api/articles` — List articles
   - `POST /api/articles` — Create article
   - `PUT /api/articles/:id` — Update article
   - `DELETE /api/articles/:id` — Delete article
   - `GET /api/movements` — List movements
   - `POST /api/movements` — Create movement
   - `GET /api/suppliers` — List suppliers
   - `POST /api/suppliers` — Create supplier
   - `GET /api/dashboard` — Dashboard KPIs
   - `GET /api/reports/:type` — Generate reports

2. **Database schema:**
   ```sql
   CREATE TABLE articles (
     id SERIAL PRIMARY KEY,
     reference VARCHAR(20) UNIQUE,
     designation VARCHAR(100),
     category VARCHAR(50),
     supplier_id INTEGER,
     unit_price DECIMAL(10,2),
     current_stock INTEGER,
     minimum_stock INTEGER,
     created_at TIMESTAMP DEFAULT NOW(),
     updated_at TIMESTAMP DEFAULT NOW()
   );
   
   CREATE TABLE movements (
     id SERIAL PRIMARY KEY,
     article_id INTEGER,
     movement_type ENUM('IN', 'OUT'),
     quantity INTEGER,
     unit_price DECIMAL(10,2),
     reference VARCHAR(50),
     created_at TIMESTAMP DEFAULT NOW()
   );
   
   CREATE TABLE suppliers (
     id SERIAL PRIMARY KEY,
     code VARCHAR(20) UNIQUE,
     name VARCHAR(100),
     address TEXT,
     phone VARCHAR(20),
     email VARCHAR(100),
     nif VARCHAR(20),
     nis VARCHAR(20),
     rc VARCHAR(20),
     created_at TIMESTAMP DEFAULT NOW()
   );
   ```

### Phase 2: Excel API Client (Month 2)
1. **VBA HTTP Client:**
   ```vba
   ' modApiClient.bas
   Public Function ApiGet(ByVal endpoint As String) As String
       Dim http As Object
       Set http = CreateObject("MSXML2.XMLHTTP")
       http.Open "GET", BASE_URL & endpoint, False
       http.setRequestHeader "Authorization", "Bearer " & GetToken()
       http.Send
       ApiGet = http.responseText
   End Function
   ```

2. **JSON Parser:**
   - Use VBA-JSON library for parsing
   - Convert JSON to VBA collections/arrays

3. **Sync Engine:**
   - Push local changes to API
   - Pull remote changes to local
   - Conflict resolution (last write wins)

### Phase 3: Web Dashboard (Month 3)
1. **React Dashboard:**
   - Real-time KPIs
   - Charts and graphs
   - Mobile-responsive

2. **Features:**
   - View-only access for managers
   - Print reports from browser
   - Export to PDF/Excel

## Implementation Timeline

### Month 1: API Development
- Week 1: Setup Node.js project, database schema
- Week 2: Implement article endpoints
- Week 3: Implement movement and supplier endpoints
- Week 4: Authentication, rate limiting, testing

### Month 2: Excel Integration
- Week 1: VBA HTTP client module
- Week 2: JSON parser integration
- Week 3: Sync engine (push/pull)
- Week 4: Testing, error handling

### Month 3: Web Dashboard
- Week 1: React project setup
- Week 2: Dashboard components
- Week 3: Charts and reporting
- Week 4: Deployment, documentation

## Cost Breakdown

### Development (One-time)
- API development: 200,000 DZD
- Excel integration: 150,000 DZD
- Web dashboard: 100,000 DZD
- **Total: 450,000 DZD**

### Hosting (Monthly)
- VPS (DigitalOcean/Hetzner): $5-10/month
- Domain name: $10/year
- SSL certificate: Free (Let's Encrypt)
- **Total: ~$60-130/year**

### Maintenance (Annual)
- Updates and bug fixes: 50,000 DZD
- Support: 30,000 DZD
- **Total: 80,000 DZD/year**

## Revenue Model

### Subscription Pricing
| Tier | Price/month | Users | Features |
|------|-------------|-------|----------|
| Basic | 2,000 DZD | 1 | Core features |
| Professional | 5,000 DZD | 3 | + Reports, API |
| Enterprise | 10,000 DZD | 10 | + Multi-site, priority |

### One-time Purchase
| Tier | Price | Users | Support |
|------|-------|-------|---------|
| Basic | 45,000 DZD | 1 | 30 days |
| Professional | 75,000 DZD | 1 | 90 days |
| Enterprise | 120,000 DZD | 3 | 1 year |

## Risks and Mitigations

| Risk | Impact | Mitigation |
|------|--------|------------|
| Excel compatibility | High | Test on Excel 2016, 2019, 365 |
| API downtime | High | Implement offline mode, retry logic |
| Data loss | High | Regular backups, audit trail |
| Security breach | Medium | JWT auth, rate limiting, input validation |
| Performance | Medium | Caching, pagination, optimization |

## Success Metrics

### Technical
- API response time: < 200ms
- Uptime: 99.9%
- Data sync reliability: 99.99%

### Business
- Customer adoption: 10+ customers in Year 1
- Revenue: 500,000 DZD in Year 1
- Support tickets: < 5 per customer/month

## Next Steps

1. **Validate demand** — Survey potential customers
2. **Choose tech stack** — Node.js vs Python (FastAPI)
3. **Setup infrastructure** — VPS, database, domain
4. **Start Phase 1** — API development
5. **Build Excel client** — VBA HTTP integration
6. **Launch beta** — 3-5 pilot customers
7. **Iterate** — Based on feedback
8. **Scale** — Marketing and sales

---

**Created:** July 30, 2026  
**Author:** Mahi Kamel Abdelghani  
**Contact:** kamelmahi71@gmail.com | +213 676 77 38 92
