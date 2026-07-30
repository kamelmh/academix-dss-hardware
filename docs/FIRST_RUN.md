# First-run flow

Opening the workbook used to present the thesis case study: 40 pre-loaded
articles, 90 days of generated movements, and a placeholder identity
(`Quincaillerie` / `Algerie` / NIF `000100000000000`). A store needs the
opposite - an empty ledger carrying its own identity - while the demo set has to
stay reachable for training and for demonstrating the system.

## What happens on first open

`Auto_Open` calls `mod_FirstRun.FirstRunCheck`, which:

1. Calls `mod_Config.SeedDefaultConfig` to fill in any missing CONFIG row. This
   never overwrites an existing value, so it is also what brings a workbook
   built before this flow existed up to date.
2. Reads `FIRST_RUN`. Absent counts as `TRUE`, so every workbook shipped before
   this change presents the wizard once.
3. If pending, shows the wizard. Otherwise returns silently.

The wizard has two pages: the business identity, then the operating parameters.
It closes on one of two buttons.

| Button | Effect |
|---|---|
| **Demarrer a vide** | Writes the configuration, empties the ledgers, sets `FIRST_RUN=FALSE`. The default, and what a real store wants. |
| **Charger donnees demo** | Writes the configuration, loads 40 articles / 9 suppliers / 90 days, sets `FIRST_RUN=FALSE`. The identity entered is preserved. |

Closing the wizard with the window X leaves `FIRST_RUN` set, so it returns on the
next open rather than leaving a half-configured system.

## Entry points

| Macro | Purpose |
|---|---|
| `mod_FirstRun.FirstRunSetup` | Reopen the wizard at any time, before or after first run. |
| `mod_FirstRun.PrepareCleanStart` | Empty the ledgers, keep the configuration. Prompts first. |
| `mod_FirstRun.LoadDemoDataKeepingIdentity` | Load the demo set without losing the configuration. |
| `mod_Config.ResetConfigToDefaults` | Factory reset of the configuration, re-arms the wizard. Prompts first. |
| `mod_MasterSetup.SystemStatus` | Reports whether setup is complete, the business name, and the observation window. |

## Configuration keys

Seeded by `SeedDefaultConfig`. Business identity is seeded **blank** on purpose:
shipping a plausible but fictitious NIF is how the thesis placeholders reached
printed documents.

| Key | Default | Notes |
|---|---|---|
| `FIRST_RUN` | `TRUE` | Cleared by the wizard. Absent counts as `TRUE`. |
| `WORKING_DAYS` | 300 | Divides annual demand in `mod_StockEngine`; feeds every reorder point. |
| `ORDER_COST` | 300 | Numerator of the Wilson formula. |
| `HOLDING_RATE` | 0.2 | EOQ divisor. The wizard refuses zero. |
| `LEAD_TIME` | 2 | Days. |
| `TAX_RATE` | 0.19 | TVA. |
| `CURRENCY` | DZD | |
| `PU_INCLUDES_TVA` | TRUE | |
| `INCLUDE_FREIGHT_IN_CMUP` | FALSE | |
| `BUSINESS_NAME` | *(blank)* | Required by the wizard - it appears on invoices. |
| `BUSINESS_ADDRESS` / `_PHONE` / `_NIF` / `_NIS` / `_RC` | *(blank)* | |

### Removed

- **`SEASON`** - a thesis parameter that nothing ever read.
- **`OBSERVATION_DAYS`** - no longer seeded by default. The getter remains for
  compatibility, but consumption is now derived from the data by
  `mod_Config.ObservationDaysEffective`, which measures the span actually
  covered by MOUVEMENTS. The fixed 90 was correct only for the demo generator,
  which really does produce 90 days; for a real store three weeks in, dividing
  by 90 understated daily consumption roughly fourfold and the stockout
  projection reported about four times more runway than the shelf held.
  `mod_DemoData` still seeds 90 explicitly, where it is accurate.

## Safety notes

**Config is only ever written through `mod_Config.WriteConfig`.** It upserts on
the key and handles unprotect/reprotect, so nothing depends on a row position or
on `UserInterfaceOnly` surviving a workbook reopen.

**`SeedDefaultConfig` is no longer destructive.** It used to delete rows 2:last
and rewrite all sixteen keys, which meant calling it after setup silently erased
the identity the owner had just entered - a routine nobody could safely call, and
nothing did. It is now a gap-filling upsert, safe on every open. The destructive
behaviour lives in `ResetConfigToDefaults`, named for what it does.

**Loading the demo set no longer costs the configuration.** `NuclearClear` does
`Cells.Clear` on CONFIG, so `GenerateDemoData` now snapshots the parameters
before that clear and writes them back over the demo seeds afterwards. Whatever
the owner set wins, `FIRST_RUN` included, so a demo load cannot resurrect the
wizard on a configured system.

**Numbers are parsed for a French locale.** `Val` accepts only a period as the
decimal separator, so a typed `0,2` would read as `0` - and that value is the EOQ
divisor. `mod_FirstRun.ParseNumber` normalises the separator first. A rate typed
as `19` rather than `0,19` is offered as a percentage reading rather than being
converted silently.

## Optional: the Workbook_Open hook

`Auto_Open` already covers the normal case and needs no manual step.
`ThisWorkbook` is a document module, so it cannot be imported as a `.bas` - the
code has to be pasted into its code pane by hand. Add this only if you also want
the hook to fire when the workbook is opened by another macro
(`Workbooks.Open`), which does not trigger `Auto_Open`:

```vba
Private Sub Workbook_Open()
    mod_FirstRun.FirstRunCheck
End Sub
```

Running both hooks is harmless: `FirstRunCheck` guards itself against showing the
wizard twice in one session.

## Test checklist

Requires Excel; none of this can be verified statically.

- [ ] Fresh workbook opens the wizard
- [ ] **Demarrer a vide** writes the entered values into CONFIG
- [ ] Reopening does not show the wizard again
- [ ] ARTICLES, MOUVEMENTS and FOURNISSEURS contain headers only
- [ ] Config form shows the entered values, not the defaults
- [ ] Changing a value in the config form round-trips to the correct CONFIG row
- [ ] **Charger donnees demo** loads 40 articles *and* keeps the entered identity
- [ ] Invoice and purchase-order headers print the entered identity
- [ ] Barcode generation works against entered data
- [ ] Dashboard shows zeros on a clean start
- [ ] With a few days of real movements, `Cons./Jour` reflects the real span and
      not a fixed 90 days
- [ ] ABC classification and CMUP compute on entered data
- [ ] `FirstRunSetup` reopens the wizard mid-life without data loss
- [ ] `ResetConfigToDefaults` clears the configuration and re-arms the wizard
