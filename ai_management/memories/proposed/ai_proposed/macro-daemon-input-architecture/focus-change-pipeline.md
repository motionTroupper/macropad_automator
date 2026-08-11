---
name: macro-daemon-focus-change-pipeline
description: The focus/tab-change path in the macropad daemon and where its worker threads are started
metadata:
  type: project
---

`host-scripts/macro-daemon.py` learns about app and tab changes through two
`SetWinEventHook` subscriptions installed by the `winhook` thread (`monitor_windows()`):
`EVENT_SYSTEM_FOREGROUND` and `EVENT_OBJECT_NAMECHANGE` (the latter filtered to
`OBJID_WINDOW`/`CHILDID_SELF` and to the foreground HWND, which is how Chrome and VS Code
**tab** switches are detected). Both land in `window_change_callback()`, which resolves the
program name and, when it differs from `current_program`, calls
`setup_program(active_program, hwnd)` → `setup_program_macropad()` (config lookup + serial
send) then `setup_program_layout()` (keyboard layout).

`setup_program()` is therefore the single choke point where any "what should be active for
this app" state must be recomputed.

Worker threads are all daemon threads started in `launch_program()` before the pystray icon
runs: `macropad`, `winhook`, `kblayout`, `teams`, `store_layouts`, `hotkeys`, `power`,
`watchdog`, plus `heartbeat` and `winevt` only when `DIAGNOSTICS_ENABLED`. `watchdog` is
mandatory — `supervisor.py` uses it to detect hangs.
