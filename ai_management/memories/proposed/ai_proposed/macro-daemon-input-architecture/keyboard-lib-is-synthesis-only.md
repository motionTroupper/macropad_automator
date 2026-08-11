---
name: macro-daemon-keyboard-lib-is-synthesis-only
description: The daemon uses the `keyboard` library only to synthesize input; global hotkeys deliberately go through WinAPI RegisterHotKey because low-level hooks broke the Win key
metadata:
  type: project
---

In `host-scripts/macro-daemon.py` the `keyboard` package is used **only to emit** input —
`keyboard.write()` in `type_chars()`, `keyboard.press()`/`release()` in `toggle_key()` and in
the Teams recording logic. No hook, no `add_hotkey`, no `suppress=True` anywhere.

Global hotkeys (`ctrl+alt+shift+left/right` for zone cycling) are registered instead with
WinAPI `RegisterHotKey` plus a private message pump in `hotkey_listener()`. Its docstring
states the reason explicitly: the `keyboard` library installs a low-level keyboard hook, which
on Windows interferes with the Win key and swallows non-matching keystrokes.

So any feature that needs to *intercept and suppress* a key cannot reuse `keyboard`, and
`RegisterHotKey` cannot suppress a bare key either — it would need an explicit
`SetWindowsHookEx(WH_KEYBOARD_LL)`, which is the exact mechanism the codebase rejected for
hotkeys. See [[macro-daemon-focus-change-pipeline]].
