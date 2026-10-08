# AGENTS.md — RetailStoreAccess

Instructions for any AI coding assistant (Claude, ChatGPT / Codex, Gemini, Cursor, Copilot …) and for
human developers working in this repository. Claude reads it through `CLAUDE.md`.

**Read first:** `docs/dev/01-Architecture.md` (how it is built), `docs/dev/02-Database.md`,
`docs/dev/03-Forms.md`, `docs/dev/04-Workflow.md` (how to work and think, every Access rule with its
reason) and the generated indexes `docs/02-Tables-Reference.md`, `docs/dev/Screens-Index.md`.

## What this project is
A retail-store ERP built on **Microsoft Access** (split front-end / back-end), Arabic and right-to-left
(each front-end can also be built in English, left-to-right: `docs/38-English-Interface.md`),
Saudi VAT 15 % and ZATCA QR. Nothing is drawn by hand in Access: **Python generates VBA modules**, the
user imports them into Access and runs the `Build*` procedures, which create the tables, relationships,
queries, screens and reports. Logic that Access cannot run outside Windows is mirrored in Python and
tested on Linux/macOS/Windows with `unittest`.

- `RetailStore_BE.accdb` — the data (created/upgraded by `BuildSchema`, never overwritten).
- `RetailStore_FE.accdb` — the program (built in one step by `dist/tools/BuildFrontEnd.vbs`).

## Working conventions (the owner's rules)
- **Reply in Arabic.** Code, comments, commit messages and identifiers stay in English.
- Work **one feature at a time**. After each feature, report in Arabic:
  1. what was made and the problems found,
  2. which modules to import or replace and which `Build*` commands to run (in order),
  3. how to test it (the in-Access `Test*` procedure and manual steps),
  then **wait for approval** before starting the next feature.
- When the user reports an error from Access (a message box, `BuildFrontEnd.log`, `RunAllTests.log`):
  find the root cause, fix it, **add a test that fails without the fix**, look for the same pattern
  elsewhere in the code, then report.
- Commit with a clear English message and push to the working branch; never open a pull request
  unless asked. Never put an AI model name or identifier in commits or files.
- **Keep the Markdown files up to date with every addition or change** (the owner's rule):
  a new feature gets `docs/NN-Name.md` and a `README.md` row ("✅ بانتظار الموافقة"); a change to a
  feature updates its document; a new table / screen / module / rule updates `docs/dev/*.md` and this
  file. The references (`docs/02..04-*-Reference.md`, `docs/dev/Screens-Index.md`) are generated —
  run `tools/generate.py`, never edit them.
- **Never publish or print `LICENSE_SECRET`** (in `src/vba/modActivation.bas`). Before delivery the
  owner changes it and delivers an **ACCDE**, keeping an ACCDB for development.

## Commands
```bash
python3 tools/generate.py                     # regenerate src/vba, dist/vba and the reference docs
python3 -m unittest discover -s tests         # all tests (~750, about 2 minutes)
python3 -m unittest discover -s tests -p "test_budget.py"     # one file
cd tests && python3 -m unittest test_forms.LayoutTests        # one class (run from tests/)
```
On Windows use `python` instead of `python3`. **Always run `tools/generate.py` before the tests**:
the tests compare `src/vba` and `dist/vba` with what the generator produces.

Optional test dependencies (tests skip when missing): `pip install -r requirements-dev.txt`
(qrcode, Pillow, numpy, opencv-python-headless, zxing-cpp) and LibreOffice with python3-uno
(`soffice` on the PATH) for the tests that really execute VBA modules.

## Repository map
| Path | Contents |
|---|---|
| `tools/schema.py` | **Source of truth**: tables, fields, rules, indexes, seed data, `SCREEN_LIST` (screens + permissions), account tree |
| `tools/relations.py` | Relationships derived from `fk=`; `UNENFORCED` fields; Access 32-index limit |
| `tools/queries.py` | Saved queries (Access SQL) |
| `tools/forms.py` | Screen models: `DataScreen` (list + detail), `NAV_ITEMS` (side menu), dashboard, `SCREEN_PERMISSIONS`, `REPORTS` |
| `tools/forms_*.py` | Larger screens per area (sales, purchases, cash, journal, bank, assets, payroll, budget, accounting hub …) |
| `tools/reports*.py` | Report layouts (`ListSpec`) |
| `tools/gen_*.py`, `tools/generate.py` | Generators → `modBuildSchema`, `modBuildRelations`, `modBuildQueries`, `modBuildForms`, `modBuildReports`, `modAppData`, `modDemoData`, test modules |
| `tools/sim.py`, `tools/*_reference.py` | Python mirrors of the VBA logic (journal, stock, VAT, QR, SHA-256, activation …) |
| `src/vba/*.bas` | VBA, UTF-8 (readable copy). **Hand-written** modules are listed in `STATIC_MODULES` in `tools/generate.py`; all others are generated — never edit those by hand |
| `dist/vba/*.bas` | The same modules in **Windows-1256 with CRLF** — what the user imports. Written by the generator only (`.gitattributes`: binary) |
| `dist/tools/BuildFrontEnd.vbs` | Builds the whole front-end in Access in one step; `EnableShiftKey.vbs` re-enables SHIFT |
| `tests/` | `access_sqlite.py` (SQLite mirror of Access SQL: `Nz`, `IIf`, `DateAdd`, `Year` …), `helpers.py` (`VbaModuleChecks`), `vba_harness.py` (LibreOffice Basic runner) |
| `tools/i18n.py`, `tools/i18n_en.py`, `tools/gen_lang.py` | The English interface: Arabic -> English dictionary, its Python mirror of `Tr`, and the generated `modLang` + `modLangData*` |
| `docs/NN-*.md` | One Arabic document per feature; `README.md` lists them with their approval status |

## Adding a feature (checklist)
1. Tables / fields in `tools/schema.py`. New fields on existing tables must be optional or have a
   default (`BuildSchema` upgrades an existing back-end). Add the screen to `SCREEN_LIST` and a
   permission if needed.
2. Queries in `tools/queries.py`; screens in `tools/forms*.py` (register new form modules in
   `all_forms()` and `SCREEN_PERMISSIONS`); reports in `tools/reports*.py` / `REPORTS`.
3. Hand-written logic in a new `src/vba/modX.bas` (UTF-8, `Attribute VB_Name`, `Option Compare Database`,
   `Option Explicit`), added to `STATIC_MODULES`. Include an in-Access test `Public Function TestX() As Boolean`
   that works inside `ws.BeginTrans` … `ws.Rollback` and add it to the list in `modTestAll`.
4. If it posts accounting entries: add the journal query and keep `JOURNAL_SOURCE_QUERIES` (`tools/queries.py`) and
   `SOURCE_QUERIES` (`modJournal`) identical; respect closed periods (`ClosedPeriodProblem`).
5. Money: every amount field stays in the program currency (SAR). A document entered in another currency
   also stores `CurrencyCode`, `ExchangeRate`, `ForeignAmount` (`tools/schema.fx_fields`); its screen uses
   `currency_pair` + `CurrencyChoice`, and posting converts with `ToBase` (`modCurrency`). Add the
   document to `StampJournalCurrencies` so its journal entry carries the currency.
6. Every new Arabic text shown to the user needs its English text in `tools/i18n_en.py` (`test_i18n` lists the
   missing keys). Text set by code in a `.Caption`, `.RowSource`, `MsgBox` or `InputBox` goes through `Tr(...)`.
7. Audit trail (`modAudit`): data screens and bound grids are audited by `modForms` /
   `AuditFormBefore` + `AuditFormAfter`. Code that deletes a document takes
   `Set auditBefore = AuditSnapshot(table, key, id)` before the `DELETE` and calls
   `AuditDeleted "X_DELETE", table, id, auditBefore` after it; code that edits a saved record uses
   `AuditSnapshot` + `AuditEdited`. A test checks every document delete.
8. `python3 tools/generate.py`, add `tests/test_x.py` (Python mirror + code checks + `VbaModuleChecks`),
   run all tests. Update tests that check exact code text when the change is intentional
   (relation count in `test_relations.py`, document-screen set in `test_permissions.py` …).
9. Write `docs/NN-Name.md` (Arabic) and add the row to `README.md` with "✅ بانتظار الموافقة";
   update `docs/dev/*.md` when tables, screens, modules or rules changed.

## Access / VBA rules (each one was a real failure in Access)
VBA is never compiled here, so these are enforced by tests and must be followed by hand:
- **VBA evaluates both sides of `And` / `Or` and both branches of `IIf`.** Never write
  `IsDate(x) And DateValue(x)…` or `IIf(IsNull(x), …, SqlDate(x))`; use `If` blocks or `Nz`.
- A **Private** procedure cannot be called from another module (compile error). Make it Public.
- Block `If … End If` instead of single-line `If … Then A Else B` that calls Subs; no `Switch()`.
- Do not name variables like built-in functions or keywords: `left`, `month`, `line`, `text`, `now`,
  `dir`, `sub`, `exp_`, `base` … No procedure name may exist Public in two modules.
- Lines < 1000 characters and < 24 line continuations (`long_const`, `form_names_const` split strings).
- Only Windows-1256 characters in VBA (no `−`, smart symbols outside cp1256).
- Access SQL: an alias may not be reused inside its own expression; every `UNION` branch needs a
  `FROM`; no `&` in saved queries; no column named `[ ]`; parameters come from
  `QDate`/`QLong` (`modQueryParams`) and must be declared; no subquery in the column list of a
  totalled report source.
- Access allows **32 indexes per table**, and every enforced relationship counts on **both** tables
  (`Employees` "recorded by" fields are in `relations.UNENFORCED`).
- Controls: a check box has no `BackColor`; set only properties the control type has.
  A button paints over a label lying on it — put a **transparent** button on top of icon + label.
- Local temp tables (`tmp*`, created in `modPOS.EnsureLocalTables`) have no defaults: write 0 explicitly.
- Two screens on one table (`frmUsers` / `frmEmployeePay` on `Employees`): table handlers in
  `modForms` must check `frm.Name`.
- Dates: set `Calendar = vbCalGreg` (Saudi PCs may default to Hijri); show dates with `GDate`.
- English interface: never compare SQL with an Arabic literal (`= 'نقدي'`), the literals are translated.
  A list column alias must not equal, once translated, a field of its own expression, even qualified
  (`Format(q.Cost, ...) AS [التكلفة]` becomes `AS [Cost]`: circular reference) - choose a caption whose
  English differs (`[تكلفة الأصل]` = Asset cost). `test_i18n` checks it.
- Master data names: SQL that shows the name of an account, payment method, journal source or stock move
  type, role, permission, screen, category, unit, expense type, customer, supplier, cash box, bank, cost centre
  or sales rep reads `[@Accounts] AS a` (also `[@PaymentMethods]`, `[@JournalSourceTypes]`, `[@TransactionTypes]`,
  `[@Roles]`, `[@Permissions]`, `[@Screens]`, `[@Categories]`, `[@Units]`, `[@ExpenseTypes]`, `[@Customers]`,
  `[@Suppliers]`, `[@CashBoxes]`, `[@Banks]`, `[@CostCenters]`, `[@SalesReps]`; not for a look-up by the typed name,
  not in `qrySalesDocPrint`, the tax invoice) and goes through `Tr`: an English front-end reads `qryLocAccounts`
  (same columns, English name). No `DLookup` of such a name (it cannot go through `Tr`): `DbValue(Tr("SELECT ..."))`.
  A new master table with an English name goes into `tools/master_en.py` (`docs/39-English-Master-Data.md`, `docs/40`).
- Text the user typed (a search) is added to SQL **after** `Tr`, never before: `Tr` would translate an Arabic word
  of it (`sql = Tr(template)`, then `Replace(sql, "{LIKE}", ...)`).
  `MSG_RTL` is a function of `modLang` (0 in English). The English texts contain no `" ' [ ] ; | & = < >`.
- Ratios: divide `CDbl(...)` values, not `Currency` (LibreOffice keeps 4 decimals; a test harness
  cannot return a `Currency` either — wrap it in a `Double` function, see `tests/test_indicators.py`).

## Accounting model (short)
Journal entries are rebuilt from the documents by `SyncJournal` (`modJournal`), one source query per
document type; an entry's signature decides whether it changed. Key accounts: cash boxes 110000+box,
banks 120000+bank, 1200 Mada/wallet clearing, 1300 customers, 2100 suppliers, 1500 input VAT,
1250 cheques under collection, 2110 notes payable, 1790/5600 depreciation, 2310/2320 payroll,
1600 employee advances, 2330/5530 sales rep commissions. Cost centres travel on `JournalLines.CostCenterID`.

## Status and roadmap
Phases 1–12, accounting, VAT return, aging, banks, cheques, fixed assets, payroll, cost centres,
budget, recurring expenses, the audit trail, the financial indicators of the dashboard, multiple
currencies, sales reps with commissions and the English interface are done
(see `README.md` for approval status): phase 3 is complete. Next steps are decided by the owner.

## Standard procedures
**Run the tests:** `python3 tools/generate.py`, then `python3 -m unittest discover -s tests`; report the
count, failures and skips. Change an exact-text test only for an intentional change, and say so.

**New feature:** follow "Adding a feature" above and `docs/dev/04-Workflow.md` §1, one feature per run,
then the Arabic report and stop until the owner approves.

**Error reported from Access:** locate it (build step, screen, or the check after the last passing one
in a `Test*` procedure), find the root cause, fix every place with the same pattern, add a test that
fails without the fix (prove it), add the rule to "Access / VBA rules" here and to
`docs/dev/04-Workflow.md` §3, regenerate, run all tests, commit, report in Arabic. If the evidence is not
enough, ask for the exact line shown by Debug > Compile, the log file or a screenshot.
