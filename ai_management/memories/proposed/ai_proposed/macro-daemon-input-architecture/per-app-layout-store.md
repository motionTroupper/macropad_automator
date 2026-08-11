---
name: macro-daemon-per-app-layout-store
description: Where per-app state lives in the macropad daemon — config.json profiles, APP_LAYOUTS runtime store, and the substring force table
metadata:
  type: project
---

Per-app state in `host-scripts/macro-daemon.py` is resolved in three layers, all keyed by the
program identity of [[macro-daemon-active-program-identity]]:

| Layer | Where | Role |
|---|---|---|
| Profile defaults | `host-scripts/config.json`, top-level keys used as **regexes** (`"outlook\|mail"`, `"vscode"`, `"."` as fallback) | merged by `lookup_config()` into `running_config`: `keys`, `colors`, `toggles`, `symbols`, `layout`, `layouts`, `programs` |
| Learned state | `APP_LAYOUTS` dict → `host-scripts/app_layouts.json` | one entry per program: `{layout, last_used}`; written by a 60 s `store_layouts()` thread, entries older than `LAYOUT_DROP_DAYS = 30` dropped at load |
| Hard override | `FORCE_LAYOUT_BY_SUBSTRING` python dict | substring match on the program name; wins over both `APP_LAYOUTS` and observed OS changes (for Electron/WebView2 hosts where detection is unreliable) |

`lookup_config()` re-reads `config.json` only when its mtime changed, so editing it takes
effect on the next focus change without restarting the daemon.
