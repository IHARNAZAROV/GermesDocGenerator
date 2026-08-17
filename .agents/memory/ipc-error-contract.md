---
name: IPC error contract
description: Electron IPC handlers resolve {ok:false,error} instead of throwing — callers must check result.ok
---

Rule: main-process IPC handlers (e.g. Excel write) catch all errors and resolve `{ ok: false, error }` — they never reject.

**Why:** renderer code that only wraps `await` in try/catch silently treats failed writes as success; this once caused save/close flows to commit dirty state and clear the crash-recovery draft on a failed write.

**How to apply:** every renderer call to an Excel/file IPC channel must check `result.ok === false` and throw/handle before committing state (commitCurrentValues, clearing drafts, closing the window).
