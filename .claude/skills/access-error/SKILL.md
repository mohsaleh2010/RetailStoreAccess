---
name: access-error
description: Diagnose and fix an error the user saw in Microsoft Access (a message box, a compile error, BuildFrontEnd.log or RunAllTests.log) for RetailStoreAccess. Use when the user pastes an Access error, a screenshot of one, or a failing in-Access test.
---

# Fix an error reported from Access

VBA is never compiled or run in Access here, so the user's message is the only evidence. Be exact.

1. Identify where it happened: the build step (`BuildSchema`, `BuildRelationships`, `BuildQueries`,
   `BuildForms`, `BuildReports`, the compile step), the screen / report, or the `Test*` procedure and
   the check after the last passing one (count the checks in the test to locate the failing line).
2. Find the **root cause** in the source (`tools/*.py` for generated modules, `src/vba/*.bas` for
   hand-written ones). Typical causes are listed in `AGENTS.md` under "Access / VBA rules" (with reasons in `docs/dev/04-Workflow.md`):
   IIf / And evaluating both sides, Private procedure called from another module, property not
   supported by a control type, 32-index limit, Null in a temp table, alias / bracket rules in
   Access SQL, a shared table handled by two screens.
3. Fix it, then **search the whole code base for the same pattern** and fix those too.
4. Add a test that fails without the fix (prove it: stash the fix, run the test, restore).
5. `python3 tools/generate.py` and the full test suite.
6. Commit and push. Report in Arabic: the cause, the fix, other places fixed, which modules to
   re-import (or re-run `dist\tools\BuildFrontEnd.vbs`; the back-end is kept), and how to verify.
   If the error cannot be explained from the evidence, say what is needed (the exact line shown by
   Debug > Compile, the log file, a screenshot) instead of guessing.
