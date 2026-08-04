---
name: macro-daemon-active-program-identity
description: How macro-daemon.py turns the foreground window into the single "active program" key used by every per-app feature
metadata:
  type: project
---

`active_program_name()` in `host-scripts/macro-daemon.py` collapses the foreground window
into one lower-cased, punctuation-stripped, 50-char key, and that key is what every per-app
feature in the daemon is indexed by:

- default: the executable name (`dbeaver.exe`, `windowsterminal.exe`)
- `chrome.exe`: **replaced** by the first ` - ` segment of the window title — so a Chrome tab's
  key is the page title, and the string `chrome` never appears in it
- `code.exe`: `vscode ` + first title segment — VS Code tabs keep a common `vscode ` prefix
- `msrdc.exe`: the whole window title

Consequence for any new per-app behaviour: a regex over this key can address *all* VS Code
(prefix) but **cannot** address *all* Chrome — Chrome must be keyed off the executable
instead. See [[macro-daemon-per-app-layout-store]] and [[macro-daemon-focus-change-pipeline]].
