# 🧠 Smart Contextual Macropad for Framework Laptop

This project implements a context-aware macropad system that changes its behavior depending on the currently active window on a Windows PC.

It combines a CircuitPython-based firmware running on a physical macropad with a Windows daemon that provides real-time dynamic configuration.

---

## 🧩 Project Structure

```
.
├── code.py              # CircuitPython firmware for the macropad
├── default.json         # Default configuration loaded on startup
├── config.json          # Contextual mappings for keys and colors
├── macro-daemon.py      # Windows daemon that detects active window and syncs config
```

---

## 🔧 Components

### `code.py`
Firmware running on the macropad:
- Scans a key matrix using analog multiplexing.
- Sends HID keypresses or serial messages depending on key configuration.
- Controls per-key RGB backlighting.
- Dynamically reloads configuration via USB serial (CDC) when received.

### `macro-daemon.py`
Python background process on Windows:
- Monitors the active window (title and process).
- Looks up the appropriate key/color configuration in `config.json`.
- Sends the merged configuration to the macropad via serial (COM4).
- Interprets special messages (`MSG:TYPE`, `MSG:OPEN`) received from the macropad and executes them.

### `config.json`
JSON file defining context-specific key behavior and LED color:
- Supports multiple application profiles using regex window matching (e.g., `"outlook|mail"`).
- Allows composite key mappings, color layouts, and symbol overrides.
- Falls back to `"."` profile for defaults.

### `default.json`
The base configuration loaded by the macropad at startup if no context is yet known.

---

## ⚙️ Features

- 🔄 Real-time configuration switching based on active window
- ⌨️ HID key support with modifiers, delays, and multi-key sequences
- 🎨 Per-key RGB color customization
- 🧪 UUID key support (generates and re-types a UUID string)
- 🪟 Auto-detection of app/window context using regex titles
- ⌨️ Keyboard layout switching (EN/ES) based on context

---

## 🖥️ Requirements

- **Hardware**: A CircuitPython-compatible board with HID + CDC support (e.g., Raspberry Pi Pico)
- **OS**: Windows
- **Python packages**:
  ```bash
  pip install pyserial pystray pygetwindow pywin32 psutil keyboard
  ```

---

## 🚀 Usage

1. Flash `code.py` and `default.json` to the macropad.
2. Connect the macropad to your PC (ensure it appears as a serial device).
3. Run `macro-daemon.py` on Windows.
4. When you switch applications, the daemon detects the window and sends the corresponding layout to the macropad.
5. Press a key on the macropad to trigger its contextual action.

---

## 📝 Example `keys` mapping

```json
"b1": "\e\e\C\S1\s\c\sgi",
"d1": "c\pp3\n",
"e3": "MSG:TYPE:#UUID#",
"e4": "MSG:TYPE:#NEW_UUID##UUID#"
```

---

## 📄 License
MIT License – Based on original by Daniel Schaefer (2023), modified by Raul Martinez (2025)

