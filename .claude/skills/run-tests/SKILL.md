---
name: run-tests
description: Regenerate the VBA modules and docs from tools/*.py and run the whole test suite of RetailStoreAccess. Use after any change, before committing, or when the user asks to run or check the tests.
---

# Run the tests

1. Regenerate everything (the tests compare `src/vba` and `dist/vba` with the generator output):
   `python3 tools/generate.py` (on Windows: `python tools/generate.py`).
2. Run all tests from the repository root:
   `python3 -m unittest discover -s tests` — about 750 tests, about 2 minutes.
   To run one file: `python3 -m unittest discover -s tests -p "test_name.py"`;
   one class: `cd tests && python3 -m unittest test_forms.LayoutTests`.
3. Skipped tests are normal when LibreOffice / python3-uno or the packages of
   `requirements-dev.txt` are missing; say how many were skipped.
4. On a failure: read the assertion, decide whether the code or an exact-text test is wrong.
   Change a test only when the code change was intentional (relation counts, expected screen sets,
   exact VBA text), and say so in the report.
5. Report in Arabic: the number of tests, OK / failures, and what was fixed.
