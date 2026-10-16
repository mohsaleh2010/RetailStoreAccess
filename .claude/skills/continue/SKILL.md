---
name: continue
description: Resume the RetailStoreAccess work where the last session stopped, on any machine. Reads docs/dev/05-Progress.md, handles a pending approval, then starts the next phase of the plan. Use when the user says continue, resume, تابع, كمّل, or approves with موافق.
---

# Continue the plan

1. `git status` and `git pull` on the working branch (`docs/dev/05-Progress.md` §1 names it).
2. Read `docs/dev/05-Progress.md`, then `AGENTS.md` and the plan document it points to.
3. **Something waits for approval?**
   - If the user's message approves it ("موافق"): mark it "✅ تمت الموافقة" in `README.md` and in the log
     at the end of `docs/51-...md`, update `docs/dev/05-Progress.md` (§1 and a line in §5), commit
     `Mark phase X approved`, push. Then go on to step 4.
   - Otherwise tell the user in Arabic what is waiting and stop.
4. Start the next phase with `/new-feature` (one phase only). Its last steps also update
   `docs/dev/05-Progress.md`.
5. Report in Arabic in three parts (what was made and the problems; modules and `Build*` commands in order;
   how to test), then stop until the user approves.
