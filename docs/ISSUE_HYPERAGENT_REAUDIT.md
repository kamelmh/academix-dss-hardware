# HyperAgent — Post-Fix Audit Request

## All Critical Findings Fixed (commits `38aaf98` → `79eb966`)

### 🔴 CRITICAL Fixes

1. **MASTER_PWD time bomb** — `mod_Config.bas`: generates random password once at first-run wizard, stores in CONFIG sheet, reads from there. No more `CHANGEME_YYYYMMDD` midnight expiry.

2. **Dashboard Error 91** — 4 error handlers guarded with `If Not ws Is Nothing Then`:
   - `mod_Dashboard.bas:35`
   - `mod_AccueilButtons_v14.bas:287`
   - `mod_AccueilButtons_v14.bas:346`
   - `mod_Invoice_v14.bas:135`

3. **Stock valuation → CMUP** — All 6 stock value paths now use CMUP (col 12) with PU fallback:
   - `mod_Dashboard.bas:64`
   - `mod_AccueilDesign_v14.bas:541`
   - `mod_AccueilButtons_v14.bas:315`
   - `mod_Reports.bas:38`, `:148`
   - `mod_BuildReports_v14.bas:92`, `:111`
   - `mod_BuildDashboard_v14.bas:134`

4. **Turnover ratio** — New `GetCMUP()` helper in `mod_StockEngine.bas`. `CalculateTurnoverRatio` uses CMUP for average inventory instead of PU.

5. **Previous fixes preserved** — `"OUT"` → `"SORTIE"`, NuclearClear re-protect, `mod_AuditTrail` stub, English → French messages, VBA trust docs.

### 🟡 HIGH Fixes

6. **All Build* modules** — French messages (14 total), bare `frmXxx.Show` → `VBA.UserForms.Add()`

7. **mod_DemoData.bas** — NuclearClear re-protects sheets after clear

8. **mod_AuditTrail.bas** — Stub module created for future expansion

### ✅ VERIFIED (No Change Needed)

9. **31 modules, exact 1:1 sync** — Confirmed (52 components = 31 .bas + 21 forms)
10. **BTS formulas correct** — Validated against curriculum
11. **Pricing 45/75/120k DZD appropriate** — No change needed

---

## Verification Requested

Please re-run your full audit on `master` at `79eb966` and confirm:

1. MASTER_PWD no longer changes at midnight (reads from CONFIG sheet)
2. All Error 91 sites are guarded (`ws` Nothing check before `.Protect`)
3. Stock value uses CMUP (col 12) everywhere with PU fallback
4. No other `Cells(i, 8)` used in stock valuation contexts
5. Any new findings from the latest code

---

## Open Questions (Your Recommendation)

### Turnover COGS
MOUVEMENTS has no CMUP column. Options:
- **A)** Add a CMUP column to MOUVEMENTS (computed at movement time)
- **B)** Compute COGS at query time: `SUM(Qty WHERE SORTIE) × article.CMUP`
- **C)** Accept current overstatement (sale value > cost) for MVP

### Average Inventory
Current uses `stock × CMUP` (point-in-time). Options:
- **A)** Track monthly snapshots for proper average
- **B)** Keep current approximation for MVP

### Missing Formulas
Worth adding for Excel MVP or save for web?
- Safety Stock (currently static 50)
- Variation de Stock
- Marge Commerciale
- Taux de rupture / service level
- Fill rate
- TCO (Total Cost of Ownership)
- Weighted supplier score

---

## .xlsm Status

| Metric | Value |
|--------|-------|
| Size | 542KB |
| Modules | 31 .bas + 21 forms = 52 total |
| Sync | Exact 1:1 with repo |
| Rebuilt | After every fix via win32com |

---

**Commits:** `c3db4db` → `38aaf98` → `79eb966`
**Branch:** `master`
**Repo:** https://github.com/kamelmh/academix-dss-hardware
