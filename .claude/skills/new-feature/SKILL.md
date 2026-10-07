---
name: new-feature
description: Implement one new feature of the RetailStoreAccess Access ERP end to end (schema, queries, screens, reports, VBA logic, in-Access test, Python tests, Arabic doc), then report in Arabic and wait for approval. Use when the user asks for a new feature, screen, report or phase.
---

# Implement a feature

Follow `CLAUDE.md` ("Adding a feature" and "Access / VBA rules"). One feature per run.

1. **Understand first.** Read the related `docs/NN-*.md`, the tables in `tools/schema.py` and the
   nearest similar feature (e.g. `modBudget`, `modRecurring`, `forms_budget.py`) and follow its patterns.
   If a business decision is truly unclear, ask once in Arabic before writing code.
2. **Schema** (`tools/schema.py`): new tables / fields (fields added to existing tables need a default
   or must be optional), indexes, `SCREEN_LIST` row and permission. Check the 32-index limit
   (`relations.index_load`).
3. **Queries / screens / reports**: `tools/queries.py`, `tools/forms*.py` (register in `all_forms()`,
   `SCREEN_PERMISSIONS`, and reach the screen from the side menu or the accounting hub
   `tools/forms_accounting.py`), `tools/reports*.py`.
4. **VBA logic** in a hand-written module `src/vba/modX.bas`, listed in `STATIC_MODULES`
   (`tools/generate.py`), with `Public Function TestX() As Boolean` running inside a rolled-back
   transaction, added to `modTestAll`. Accounting entries: journal source query + `SOURCE_QUERIES`,
   closed periods, cost centre.
5. `python3 tools/generate.py`, write `tests/test_x.py` (logic mirrored in Python where possible,
   code checks, `VbaModuleChecks`), run all tests (`/run-tests`) until green.
6. Re-read the diff against the Access / VBA rules in `CLAUDE.md` (IIf / And evaluate both sides,
   Private procedures, reserved names, line length, cp1256, Access SQL limits).
7. Write `docs/NN-Name.md` in Arabic and add the README row with "✅ بانتظار الموافقة".
8. Commit (clear English message, no AI model names) and push to the current branch. No pull request.
9. **Report in Arabic** and stop:
   - ما الذي أُضيف، والمشكلات التي ظهرت وكيف حُلّت
   - الوحدات التي تُستورد أو تُستبدل، وأوامر Build بالترتيب
     (أو: شغّل `dist\tools\BuildFrontEnd.vbs` من جديد)
   - طريقة الاختبار: `TestX` / `RunAllTests` وخطوات يدوية
   - اقتراح الخطوة التالية، ثم انتظر الموافقة
