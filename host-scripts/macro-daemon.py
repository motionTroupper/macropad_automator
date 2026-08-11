import os
import sys

import time
import pystray
from pystray import MenuItem as item, Icon

import serial
import queue

from PIL import Image
import threading
import pygetwindow as gw

import win32api
import win32gui
import win32con
import win32com
import win32com.client
import win32process

import json
import re
from pathlib import Path
import datetime
import subprocess
import keyboard

import ctypes
import ctypes.wintypes
import psutil
import uuid
import traceback
import socket
import atexit
import faulthandler
import tempfile
import winreg


program_name_pattern = r'[^a-z0-9._\-\s]'
normalize_program_name_re = re.compile(program_name_pattern, re.IGNORECASE)   


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # 2 = PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1) # 1 = PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()

## --- Feature flags --------------------------------------------------------
## Verbose logging: per-window-event prints (config lookups, macropad sends,
## layout changes). Off for normal operation; flip True when diagnosing.
debug = False

## Diagnostic threads: heartbeat (every-2s thread census) and the hidden-window
## Windows broadcast listener (DISPLAYCHANGE/POWERBROADCAST/...). Both purely
## informative — leave off unless you're hunting a freeze/crash.
DIAGNOSTICS_ENABLED = False

## --- Paths ----------------------------------------------------------------
base_path = Path(sys.argv[0]).resolve().parent
os.chdir(base_path)

## Logs and the watchdog file live under %TEMP%/macropad-automator so the
## script's working directory stays clean and supervisor.py sees the same
## paths without having to coordinate.
log_base = Path(tempfile.gettempdir()) / "macropad-automator"
log_base.mkdir(parents=True, exist_ok=True)


class _TimestampedWriter:
    def __init__(self, fp, console=None):
        self._fp = fp
        ## When launched from a terminal (python.exe) we also tee to the
        ## original console stream so logs are visible live without tailing
        ## the file. Under pythonw there is no console and `console` is None,
        ## so we silently write only to the file.
        self._console = console
        self._lock = threading.Lock()
        self._partial = ""
    def write(self, s):
        if not s:
            return
        try:
            with self._lock:
                self._partial += s
                if "\n" in self._partial:
                    lines = self._partial.split("\n")
                    self._partial = lines.pop()
                    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S.%f")[:-3]
                    for line in lines:
                        stamped = f"[{ts}] {line}\n"
                        self._fp.write(stamped)
                        if self._console is not None:
                            try:
                                self._console.write(stamped)
                            except Exception:
                                pass
                    ## Force OS-level flush so a sudden process termination
                    ## cannot swallow the last few lines.
                    try:
                        self._fp.flush()
                        os.fsync(self._fp.fileno())
                    except Exception:
                        pass
                    if self._console is not None:
                        try:
                            self._console.flush()
                        except Exception:
                            pass
        except Exception:
            pass
    def flush(self):
        try:
            with self._lock:
                self._fp.flush()
                if self._console is not None:
                    try:
                        self._console.flush()
                    except Exception:
                        pass
        except Exception:
            pass
    def fileno(self):
        return self._fp.fileno()
    def isatty(self):
        return False


## "w" so each daemon run starts with a fresh log. The supervisor relaunches
## us on death, so accumulated history just makes it harder to find the
## latest cause; cross-run history lives in macro-supervisor.log instead.
_log_fp = open(log_base / "macro-daemon.log", "w", buffering=1, encoding="utf-8", errors="replace")
## Capture the original console streams before reassigning; these are None
## under pythonw (no console) and a real stream when launched from a terminal.
sys.stdout = _TimestampedWriter(_log_fp, console=sys.__stdout__)
sys.stderr = _TimestampedWriter(_log_fp, console=sys.__stderr__)

## Dumps a Python traceback to the log if the interpreter dies from a C-level
## fault (segfault, abort, access violation). Crucial under pythonw where the
## OS would otherwise swallow the crash silently.
faulthandler.enable(file=_log_fp, all_threads=True)


@atexit.register
def _log_exit():
    print(f"=== atexit: interpreter shutting down (pid={os.getpid()}) ===")
    try:
        _log_fp.flush()
    except Exception:
        pass


def _sys_excepthook(exctype, value, tb):
    print("=== UNHANDLED EXCEPTION (main thread) ===")
    traceback.print_exception(exctype, value, tb)


def _thread_excepthook(args):
    if issubclass(args.exc_type, SystemExit):
        return
    thread_name = args.thread.name if args.thread else "?"
    print(f"=== UNHANDLED EXCEPTION in thread '{thread_name}' ===")
    traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback)


sys.excepthook = _sys_excepthook
threading.excepthook = _thread_excepthook
print(f"=== macro-daemon starting (pid={os.getpid()}) ===")


latest_uuid = None
was_teams_running = False
was_teams_to_be_recorded = False

serial_port = None
serial_lock = threading.Lock()
serial_queue = queue.Queue()

APP_OVERRIDES = {}
ZONE_DEFINITIONS = {}
MONITOR_ALIASES = {}
BORDER_OFFSET = {}
HARDWARE_ID_MAP = {}
TEAMS_RECORDING_TOP = 0
TEAMS_RECORDING_LEFT = 0
TEAMS_TOP = 0
TEAMS_LEFT = 0
## Pixel slack when matching a live Teams window against the stored zone
## corner. A DPI-aware, multi-process window can settle a couple of pixels off
## the requested position, so exact equality is too brittle. The two Teams
## zones (recording vs paused) are a full monitor-third apart, so this slack
## can never make one match the other.
TEAMS_ZONE_MATCH_TOLERANCE = 20
icon_global = None

current_tray_layout = None
current_program_hwnd = None
current_program = None

## Cycle of zones for ctrl+win+left / ctrl+win+right (loaded from zones.json -> "cycle")
ZONE_CYCLE = [
    "left-top",
    "left-top+mid",
    "left-full",
    "left-mid",
    "left-mid+bottom",
    "left-bottom",
    "top-left",
    "top-full",
    "top-right",
    "laptop-left",
    "laptop-full",
    "laptop-right",
]
ZONE_CYCLE_RESET_SECONDS = 5
WINDOW_ZONE_INDEX = {}        ## hwnd -> last index in ZONE_CYCLE
WINDOW_ZONE_LAST_PRESS = {}   ## hwnd -> last hotkey timestamp (monotonic)

LAYOUT_MAP = {
    0x4090409: {
        "code":"0x04090409",
        "name":"US English",
        "icon":"us.png"
    },
    0x40a0c0a: {
        "code":"0x040a0c0a",
        "name":"Spanish (Spain)",
        "icon":"es.png"
    }
}

## Programs whose layout we always force, regardless of APP_LAYOUTS state.
## Useful for Electron/WebView2 hosts (Teams, the new Outlook, ...) where in-
## process layout detection is unreliable — we'd rather guarantee a layout
## than let stale or partial signals win. Substring match against the lower-
## cased active program name.
FORCE_LAYOUT_BY_SUBSTRING = {
    'ateams': 0x040a0c0a,   # Spanish (Spain)
}


def force_layout_for_program(active_program):
    if not active_program:
        return None
    name = active_program.lower()
    for substring, hkl in FORCE_LAYOUT_BY_SUBSTRING.items():
        if substring in name:
            return hkl
    return None

## Windows Event Hook constants
EVENT_SYSTEM_FOREGROUND = 0x0003
WINEVENT_OUTOFCONTEXT = 0x0000
EVENT_OBJECT_NAMECHANGE = 0x800C

## Used to filter SetWinEventHook callbacks to actual window events
## (idObject=OBJID_WINDOW, idChild=CHILDID_SELF). Without this filter
## NAMECHANGE fires for every UI element in the system — toolbars, list
## items, icons — which is a huge amount of work for nothing.
OBJID_WINDOW = 0
CHILDID_SELF = 0

## Define the callback function type
WinEventProcType = ctypes.WINFUNCTYPE(
    None, 
    ctypes.wintypes.HANDLE, 
    ctypes.wintypes.DWORD, 
    ctypes.wintypes.HWND, 
    ctypes.wintypes.LONG, 
    ctypes.wintypes.LONG, 
    ctypes.wintypes.DWORD, 
    ctypes.wintypes.DWORD
)

## Load app layouts from json file
PERSIST_APP_LAYOUTS = True
LAYOUT_DROP_DAYS = 30
APP_LAYOUTS_FILE = "./app_layouts.json"

## Load existing layouts if persistence is enabled
if PERSIST_APP_LAYOUTS and os.path.exists(APP_LAYOUTS_FILE):
    with open(APP_LAYOUTS_FILE, 'r') as file:
        APP_LAYOUTS = json.load(file)
        ## Remove layouts not used in the last 30 days
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=LAYOUT_DROP_DAYS)
        APP_LAYOUTS = {app: data for app, data in APP_LAYOUTS.items() if 'last_used' in data and datetime.datetime.fromisoformat(data['last_used']) >= cutoff_date}
else:
    APP_LAYOUTS = {}

running_config={}       ## Current active configuration
macropad_config={}      ## Last sent configuration to macropad
all_configurations={}   ## All loaded configurations
all_configurations_version=None
all_configurations_keys=[]  ## Keys of all_configurations, sorted; cached on reload
toggles={}              ## Toggles state


def print_monitor_ids():
    print("\n--- SCANNING MONITORS ---")
    monitors = win32api.EnumDisplayMonitors()
    for i, (hMonitor, hdc, rect) in enumerate(monitors):
        monitor_info = win32api.GetMonitorInfo(hMonitor)
        adapter_name = monitor_info['Device']
        
        try:
            # Get the display device associated with the adapter
            # Second parameter 0 is the device index for that adapter
            device = win32api.EnumDisplayDevices(adapter_name, 0, 0)
            device_id = device.DeviceID
            print(f"Monitor {i}:")
            print(f"  Handle: {hMonitor}")
            print(f"  Adapter: {adapter_name}")
            print(f"  DeviceID: {device_id}") 
        except Exception as e:
            print(f"  Error reading ID: {e}")
    print("---------------------------------------\n")

def active_monitors():
    monitors = win32api.EnumDisplayMonitors()
    active_monitors = []
    for hMonitor, hdc, rect in monitors:
        try:
            monitor_info = win32api.GetMonitorInfo(hMonitor)
            adapter_name = monitor_info['Device']
            device = win32api.EnumDisplayDevices(adapter_name, 0, 0)
            real_device_id = device.DeviceID
            active_monitors.append((real_device_id.split('\\')[1], monitor_info['Work']))
        except Exception as e:
            print(f"Error obtaining monitor information: {e}")
    return active_monitors


def load_zones_config():
    global ZONE_DEFINITIONS, HARDWARE_ID_MAP, BORDER_OFFSET, APP_OVERRIDES, ZONE_CYCLE
    global CUSTOM_INACTIVITY_CHECK
    try:
        with open("zones.json", "r") as f:
            data = json.load(f)
            ZONE_DEFINITIONS = data.get("areas", {})
            HARDWARE_ID_MAP = data.get("hardware_mapping", {}) # <--- NUEVO
            APP_OVERRIDES = data.get("app_overrides", {})
            CUSTOM_INACTIVITY_CHECK = data.get("custom_inactivity_check", {})

            cycle = data.get("cycle")
            if isinstance(cycle, list) and cycle:
                ZONE_CYCLE = cycle

            hostname = socket.gethostname()
            offset_key = f"offsets-{hostname}"
            BORDER_OFFSET = data.get(offset_key, data.get("offsets-default", {}))

            print(f"Loaded {len(ZONE_DEFINITIONS)} zones, {len(HARDWARE_ID_MAP)} hardware monitors, cycle of {len(ZONE_CYCLE)} zones, {len(CUSTOM_INACTIVITY_CHECK)} custom inactivity budgets.")
    except Exception as e:
        print(f"Error loading zones.json: {e}")

def get_monitor_rect_by_alias(target_alias):

    # Lookup the monitor rectangle by its alias
    global HARDWARE_ID_MAP
    target_hw_id_part = None
    active_monitors_list = active_monitors()
    for hw_id, alias in HARDWARE_ID_MAP.items():
        # Allow same monitor with multiple indices
        hw_id = hw_id.split('_')[0]  

        # Discard monitors that are not currently active
        if hw_id not in [dev_id for dev_id, _ in active_monitors_list]:
            continue
        # Look for the target
        if alias == target_alias:
            target_hw_id_part = hw_id
            break
    
    ## Try to find the monitor by its hardware ID part
    if target_hw_id_part:
        for dev_id, work_rect in active_monitors_list:
            if target_hw_id_part in dev_id:
                # ¡Encontrado el legítimo dueño!
                return work_rect
            
    # If not found, try to fallback to any unknown monitor
    if target_alias: 
        print(f"Monitor oficial para '{target_alias}' no encontrado. Buscando monitor extraño...")
        
        known_ids = list(HARDWARE_ID_MAP.keys())
        for dev_id, work_rect in active_monitors_list:
            # Is this monitor 'dev_id' one of my known ones?
            is_known = False
            for kid in known_ids:
                if kid in dev_id:
                    is_known = True
                    break
            
            if not is_known:
                print(f"FALLBACK: Asignando monitor desconocido ({dev_id}) a '{target_alias}'")
                return work_rect

    print(f"Monitor para '{target_alias}' no encontrado ni reemplazable.")
    return None


def move_window_to_zone(zone_key):
    global TEAMS_RECORDING_TOP, TEAMS_RECORDING_LEFT, BORDER_OFFSET, APP_OVERRIDES, TEAMS_LEFT, TEAMS_TOP

    zone = ZONE_DEFINITIONS.get(zone_key)
    if not zone:
        print(f"Zona {zone_key} no existe")
        return

    hwnd = win32gui.GetForegroundWindow()
    if not hwnd: return

    # --- RESTORE IF MAXIMIZED OR MINIMIZED ---
    placement = win32gui.GetWindowPlacement(hwnd)
    is_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)
    is_minimized = (placement[1] == win32con.SW_SHOWMINIMIZED)

    if is_maximized or is_minimized:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # GET START MONITOR (REQUIRED) ---
    start_rect = get_monitor_rect_by_alias(zone['monitor'])
    if not start_rect: 
        print(f"Start monitor '{zone['monitor']}' not found.")
        return
    
    # GET END MONITOR (OPTIONAL / FALLBACK) ---
    end_alias = zone.get('monitor_end', zone['monitor'])
    end_rect = get_monitor_rect_by_alias(end_alias)

    # FALLBACK:
    if end_rect:
        # Home scenario: Both monitors exist
        # Calculate the union of both monitors as canvas
        s_left, s_top, s_right, s_bottom = start_rect
        e_left, e_top, e_right, e_bottom = end_rect
        
        canvas_left = min(s_left, e_left)
        canvas_top = min(s_top, e_top)
        canvas_right = max(s_right, e_right)
        canvas_bottom = max(s_bottom, e_bottom)
        # print(f"Dual Monitor Mode: {zone['monitor']} -> {end_alias}")
    else:
        # Work scenario: The end monitor is not connected
        # Gracefully degrade: The total canvas is ONLY the start monitor
        canvas_left, canvas_top, canvas_right, canvas_bottom = start_rect
        print(f"Single Monitor Fallback: '{end_alias}' not detected. Using only '{zone['monitor']}'.")

    canvas_width = canvas_right - canvas_left
    canvas_height = canvas_bottom - canvas_top

    # CALCULATE COORDINATES ---
    # Percentages are applied over the calculated canvas (whether dual or single)
    raw_x = canvas_left + int(canvas_width * (zone['min_x'] / 100))
    raw_y = canvas_top + int(canvas_height * (zone['min_y'] / 100))
    
    raw_x2 = canvas_left + int(canvas_width * (zone['max_x'] / 100))
    raw_y2 = canvas_top + int(canvas_height * (zone['max_y'] / 100))

    raw_w = raw_x2 - raw_x
    raw_h = raw_y2 - raw_y

    # APPLY BORDER CORRECTION AND OVERRIDES ---
    app_name = get_active_window()[0].lower()
    app_adj = APP_OVERRIDES.get(app_name, {})

    final_x = raw_x + BORDER_OFFSET["x"] + app_adj.get("x",0)
    final_y = raw_y + BORDER_OFFSET["y"] + app_adj.get("y",0)
    final_w = raw_w + BORDER_OFFSET["w"] + app_adj.get("w",0)
    final_h = raw_h + BORDER_OFFSET["h"] + app_adj.get("h",0)

    # EXECUTE
    try:
        ## Move twice on purpose. When a window crosses to a monitor with a
        ## different DPI scale, Windows fires WM_DPICHANGED *during* the move
        ## and rescales the window, so a single MoveWindow lands it in the
        ## right spot but too small. The first call settles the window on the
        ## target monitor (and thus its DPI); the second applies the final
        ## size at that DPI. On a same-monitor move the second call is a
        ## harmless no-op.
        win32gui.MoveWindow(hwnd, final_x, final_y, final_w, final_h, True)
        win32gui.MoveWindow(hwnd, final_x, final_y, final_w, final_h, True)
        win32gui.SetForegroundWindow(hwnd)

        ## Record where the window ACTUALLY ended up, not where we asked it to
        ## go. A DPI-aware target (Teams is WebView2/multi-process) can settle
        ## a few pixels off `final_x/final_y` after the WM_DPICHANGED reflow,
        ## and check_teams_window matches these against the live window rect —
        ## so storing the requested coords would make the (tolerant) match miss.
        actual_left, actual_top, _, _ = win32gui.GetWindowRect(hwnd)

        if zone.get("is_teams_recording_zone", False):
            TEAMS_RECORDING_LEFT = actual_left
            TEAMS_RECORDING_TOP = actual_top

        if zone.get("is_teams_zone", False):
            TEAMS_LEFT = actual_left
            TEAMS_TOP = actual_top

    except Exception as e:
        print(f"Error: {e}")



def toggle_monitor_timeout(spec):
    """Toggle the display-off idle timeout between two values (minutes),
    applied to both AC (plugged) and DC (battery) power states.

    `spec` is "low,high" (e.g. "5,30"). The current value is read from the
    active power scheme rather than kept in memory, so the toggle survives a
    daemon restart."""
    CREATE_NO_WINDOW = 0x08000000
    try:
        low_str, high_str = spec.split(",")
        low, high = int(low_str), int(high_str)
    except Exception as e:
        print(f"Invalid IDLE spec '{spec}': {e}")
        return

    try:
        out = subprocess.run(
            ["powercfg", "/query", "SCHEME_CURRENT", "SUB_VIDEO", "VIDEOIDLE"],
            capture_output=True, text=True, creationflags=CREATE_NO_WINDOW,
        ).stdout
        m = re.search(r"Current AC Power Setting Index:\s*0x([0-9a-fA-F]+)", out)
        current_min = int(m.group(1), 16) // 60 if m else None
    except Exception as e:
        print(f"Error querying monitor timeout: {e}")
        current_min = None

    ## Switch to `high` only when we are currently at `low`; anything else
    ## (unknown, or already high) drops back to `low`.
    target = high if current_min == low else low

    try:
        for opt in ("monitor-timeout-ac", "monitor-timeout-dc"):
            subprocess.run(
                ["powercfg", "-change", opt, str(target)],
                creationflags=CREATE_NO_WINDOW,
            )
        print(f"Monitor idle timeout set to {target} min (was {current_min})")
    except Exception as e:
        print(f"Error setting monitor timeout: {e}")


def open_window(regexp_filter):

    # Check for comma to extract second part
    if ',' in regexp_filter:
        parts = regexp_filter.split(',')
        regexp_filter = parts[1]

    # Get program info from running config
    programs = running_config.get('programs', {})
    if regexp_filter not in programs:
        print (f"Program {regexp_filter} was not recognized")
        return 

    # Extract program details   
    program_name = programs[regexp_filter]['program']
    window_name = programs[regexp_filter]['window']
    multiple_instances = programs[regexp_filter].get('multiple_instances',False)

    # Search for existing windows
    def callback(hwnd, lista):
        if win32gui.IsWindowVisible(hwnd):
            _, pid = win32process.GetWindowThreadProcessId(hwnd)
            try:
                proceso = psutil.Process(pid)
                nombre_ejecutable = proceso.name()
                if re.search(window_name, nombre_ejecutable, re.IGNORECASE):
                    lista.append(hwnd)
            except psutil.NoSuchProcess:
                pass

    windows=[]
    win32gui.EnumWindows(callback, windows)

    if len(windows)==0:
        print (f"Launching program {program_name}")
        subprocess.Popen(f"start {program_name}", shell=True)
        time.sleep(1)
    elif win32gui.GetForegroundWindow() in windows:
        print (f"Window for {regexp_filter} is already active")
        if multiple_instances:
            print (f"Launching another instance of {program_name}")
            subprocess.Popen(f"start {program_name}", shell=True)

    if len(windows)==0:
        win32gui.EnumWindows(callback, windows)

    for hwnd in windows:
        if win32gui.IsIconic(hwnd):  
            print (f"Restoring minimized window for {regexp_filter}")
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            continue
        else:
            print (f"Bringing to front window for {regexp_filter}")
            win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
            time.sleep(0.05)
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
            continue


def lookup_config(window_title):
    global all_configurations, all_configurations_version, all_configurations_keys, toggles

    try:
        config_version = datetime.datetime.fromtimestamp(Path("./config.json").stat().st_mtime)

        ## Reload (and re-sort keys) only when config.json has actually changed.
        ## On steady state this is a single stat() call per focus change.
        if config_version != all_configurations_version:
            all_configurations_version = config_version
            with open("./config.json", 'r') as file:
                all_configurations = json.load(file)
            all_configurations_keys = sorted(all_configurations.keys(), key=len, reverse=False)
            print(f"Loaded {len(all_configurations_keys)} config keys: {all_configurations_keys}")

        new_config = {
            "window": None,
            "colors": {},
            "keys": {}
        }
        for config_key in all_configurations_keys:
            if re.search(config_key, window_title, re.IGNORECASE) or config_key == '.':
                debug and print(f"{config_key} matched for {window_title}")
                if not new_config['window']:
                    new_config['window'] = config_key

                for key, value in all_configurations[config_key]['keys'].items():
                    new_config['keys'][key]=value

                for key, value in all_configurations[config_key]['colors'].items():
                    new_config['colors'][key]=value

                for key, value in all_configurations[config_key].get('toggles',{}).items():
                    toggle = toggles.setdefault(key, {})
                    toggle['config'] = value
                    toggle.setdefault('pos', 0)

                if (all_configurations[config_key]).get('symbols',None):
                    new_config['symbols'] = all_configurations[config_key]['symbols'] 

                if (all_configurations[config_key]).get('layout',None):
                    new_config['layout']=all_configurations[config_key]['layout']

                if (all_configurations[config_key]).get('programs',None):
                    new_config['programs']=all_configurations[config_key]['programs']
                
                if (all_configurations[config_key]).get('layouts',None):
                    new_config['layouts']=all_configurations[config_key]['layouts']

        return new_config
    except Exception as e:
        print(f"Error loading json: {e}")
        traceback.print_exc()
        
    
    return {
        "window": window_title,
        "colors": {},
        "keys": {}
    }

def type_chars(cadena):
    global latest_uuid
    if '#NEW_UUID#' in cadena:
        latest_uuid=str(uuid.uuid4())
        cadena = cadena.replace('#NEW_UUID#','')

    if '#UUID#' in cadena:
        if not latest_uuid:
            latest_uuid=str(uuid.uuid4())
        cadena = cadena.replace("#UUID#",latest_uuid)
    keyboard.write(cadena)

def toggle_key(toggle_name):
    global toggles, running_config

    print ("toggle key called for "+toggle_name)

    cur_pos = toggles[toggle_name].get('pos',0)
    options = toggles[toggle_name]['config']
    num_options = len(options)
    next_pos = (cur_pos+1) % num_options
    toggles[toggle_name]['pos']=next_pos
    next_leds = toggles[toggle_name]['config'][next_pos]['color']
    next_strokes = toggles[toggle_name]['config'][next_pos]['strokes']
    next_key = toggles[toggle_name]['config'][next_pos]['key']

    print (f"Toggling {toggle_name} to position {next_pos} with key {next_key}, strokes {next_strokes} and leds {next_leds}")

    for stroke in next_strokes:
        print (f"Pressing {stroke}")
        keyboard.press(stroke)
        time.sleep(0.05)
        keyboard.release(stroke)

    running_config['colors'][next_key]=next_leds
    send_command_to_macropad(running_config)


## Get the current active window title and process name
def get_active_window():
    window = win32gui.GetForegroundWindow()
    if not window:
        return None, None

    window_title = win32gui.GetWindowText(window)
    _, pid = win32process.GetWindowThreadProcessId(window)

    try:
        executable = psutil.Process(pid).name()
    except psutil.NoSuchProcess:
        return None, None

    return executable, window_title


def active_program_name():
    try:
        active_program, active_window = get_active_window()
    except Exception:
        active_program, active_window = None, None

    if not active_program:
        active_program = 'explorer.exe'
    if not active_window:
        active_window = ''

    active_program = re.sub(normalize_program_name_re, '', active_program.lower()).strip()[:50]
    active_window = re.sub(normalize_program_name_re, '', active_window.lower()).strip()[:50]

    if active_program == 'chrome.exe':
        active_program = active_window.split(' - ')[0]
    elif active_program == 'msrdc.exe':
        active_program = active_window
    elif active_program == 'code.exe':
        ## VS Code, like Chrome, is one process hosting many tabs, so the process
        ## name alone can't tell a Claude conversation from a code editor.
        ## Identify by the tab (the window title's first segment) with a
        ## 'vscode ' prefix: each tab gets its own APP_LAYOUTS entry and keeps
        ## whatever layout you set for it (config's 'vscode' -> EN by default;
        ## switch a Claude tab to ES once and it sticks).
        active_program = 'vscode ' + active_window.split(' - ')[0]

    return active_program

def send_command_to_macropad(command_dict):
    global serial_port, macropad_config

    ## Avoid spamming the macropad with identical configs on every window
    ## focus event. This is also where the `Sending command` print used to
    ## spam the log; gated behind `debug`.
    if command_dict == macropad_config:
        debug and print("Configuration unchanged, skipping macropad send")
        return

    macropad_config = command_dict.copy()
    with serial_lock:
        debug and print("Sending command to macropad")
        command = json.dumps(macropad_config) + '\n'
        try:
            serial_port.write(command.encode())
            serial_port.flush()
        except Exception as e:
            print(f"Serial write failed: {e}")

def setup_program_macropad(active_program):
    global running_config
    debug and print(f"Setting up macropad for program: {active_program}")
    running_config = lookup_config(active_program)
    send_command_to_macropad(running_config)

def get_app_layout(active_program):
    global APP_LAYOUTS

    ## Use the program the switch was decided on (passed in), never re-resolved
    ## here — re-resolving used to race and hand back the wrong program.

    ## Forced layouts (e.g. Teams) win over anything saved in APP_LAYOUTS.
    forced = force_layout_for_program(active_program)
    if forced is not None:
        debug and print(f"Forcing layout {LAYOUT_MAP.get(forced,{}).get('name', '?')} for {active_program}")
        return forced

    if active_program in APP_LAYOUTS:
        app_layout = APP_LAYOUTS[active_program]['layout']
        APP_LAYOUTS[active_program]['last_used'] = datetime.datetime.now().isoformat()
    else:
        app_layout = running_config.get('layouts', {}).get(running_config.get('layout', current_tray_layout), current_tray_layout)
        APP_LAYOUTS[active_program]= {
            "layout": app_layout,
            "last_used": datetime.datetime.now().isoformat()
        }
    debug and print(f"Layout for {active_program} is {LAYOUT_MAP.get(app_layout,{}).get('name', 'Unknown')}")
    return app_layout

def setup_program_layout(active_program):
    debug and print(f"Setting up layout for program: {active_program}")

    ## Get required layout
    required_layout_id = get_app_layout(active_program)
    if not required_layout_id:
        debug and print("No layout required")
        return
    
    ## Get current layout
    hwnd = win32gui.GetForegroundWindow()
    thread_id = ctypes.windll.user32.GetWindowThreadProcessId(hwnd, None)
    if not hwnd:
        print ("No active window found")
        return
    
    ## Post layout
    win32gui.PostMessage(
        hwnd, win32con.WM_INPUTLANGCHANGEREQUEST, 0, required_layout_id
    )

    ## Check and try with send if needed
    current_layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
    if current_layout == required_layout_id:
        ## Wait a bit
        time.sleep(0.05)

        ## Verify change and retry with SendMessage if needed
        current_layout = ctypes.windll.user32.GetKeyboardLayout(thread_id)
        if current_layout != required_layout_id:
            win32gui.SendMessage(hwnd, win32con.WM_INPUTLANGCHANGEREQUEST, 0, required_layout_id)

    change_tray_icon_layout(required_layout_id)


def setup_program(active_program, hwnd):
    global current_program, current_program_hwnd
    print(f"Switching to program: {active_program}")
    ## Mark the new program BEFORE applying its layout. setup_program_layout is
    ## slow (posts the layout change and redraws the tray icon) and the keyboard
    ## monitor runs concurrently: if current_program still pointed at the OLD
    ## program when the monitor saw the new layout land, it saved that layout
    ## under the old program (e.g. EN under the Claude tab on a tab switch).
    ## Setting it first attributes the observed change to the right program.
    current_program = active_program
    current_program_hwnd = hwnd
    setup_program_macropad(active_program)
    setup_program_layout(active_program)


def serial_port_initialize():
    global serial_port
    ## Initialize serial port
    if serial_port:
        serial_port.close()
        serial_port = None
    serial_port = serial.Serial('COM4', 115200, timeout=1)  
    print(f"Serial port COM4 initialized.")

# Función principal que monitorea el cambio de window 
def monitor_macropad():
    global serial_port

    threading.Thread(target=process_serial_queue, daemon=True).start()

    while True:
        try:
            if serial_port and serial_port.is_open:
                line = serial_port.readline().decode('utf-8').strip()
                if line:
                    data = json.loads(line) 
                    serial_queue.put(data)
            else:
                serial_port_initialize()
        except Exception as ex:
            print(f"Serial port error: {ex}")
            time.sleep(2)

def enter_suspend_state(data):
    print ("Entering suspend state as requested by macropad")

    code_hibernate = data['code'][6]
    code_critical = data['code'][7]
    code_wakeup = data['code'][8]

    if code_hibernate=='0' and code_critical=='1' and code_wakeup=='0':
        ## Sleep monitor
        ctypes.windll.user32.SendMessageW(
            0xFFFF,  # HWND_BROADCAST
            0x0112,  # WM_SYSCOMMAND
            0xF170,  # SC_MONITORPOWER
            2        # monitor off
        )
    else:
        ## Sleep system
        ctypes.windll.powrprof.SetSuspendState(int(code_hibernate), int(code_critical), int(code_wakeup))


def process_serial_queue():
    global serial_queue

    while True:
        data = serial_queue.get()
        print(f"{data} received")
        if data['code'][:5]=='OPEN:':
            app = data['code'][5:]
            print(f"Told to open [{app}]")
            open_window(app)
        elif data['code'][:5]=='TYPE:':
            to_type = data['code'][5:]
            print(f"Told to type {to_type}")
            type_chars(to_type)
        elif data['code'][:7]=='TOGGLE:':
            toggle_name = data['code'][7:]
            toggle_key(toggle_name)
        elif data['code'][:7]=='SCREEN:':
            screen_code = data['code'][7:]
            move_window_to_zone(screen_code)
        elif data['code'][:6]=='SLEEP:':
            enter_suspend_state(data)
        elif data['code'][:5]=='IDLE:':
            toggle_monitor_timeout(data['code'][5:])


# Function to quit the application
def quit(icon, item):
    icon.stop()
    sys.exit()

def chat_title(title):
    parts = [part.strip() for part in title.split("|")]
    for i, part in enumerate(parts):
        if part == "Bosonit" and i > 0:
            return parts[i - 1]
    return None

def at_teams_zone(window, zone_top, zone_left):
    ## A zone corner of (0, 0) means "never set this run" — don't match stray
    ## windows sitting near the screen origin against an unconfigured zone.
    if zone_top == 0 and zone_left == 0:
        return False
    return (abs(window.top - zone_top) <= TEAMS_ZONE_MATCH_TOLERANCE and
            abs(window.left - zone_left) <= TEAMS_ZONE_MATCH_TOLERANCE)

def check_teams_window():
    global was_teams_running, was_teams_to_be_recorded, TEAMS_RECORDING_TOP, TEAMS_RECORDING_LEFT, TEAMS_LEFT, TEAMS_TOP
    print ("Starting Teams window monitor")
    while True:
        is_teams_running = False
        is_teams_to_be_recorded = False
        for window in gw.getAllWindows():
            title = window.title or ""
            title_lower = title.lower()
            if "teams" in title_lower and at_teams_zone(window, TEAMS_RECORDING_TOP, TEAMS_RECORDING_LEFT):
                print (f"All window info: {window}")
                teams_app = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}_{chat_title(title) or 'Meeting'}"
                print (f"Found Teams window: {teams_app}")
                is_teams_running = True
                is_teams_to_be_recorded = True

            if "teams" in title_lower and at_teams_zone(window, TEAMS_TOP, TEAMS_LEFT):
                print (f"All window info: {window}")
                teams_app = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}_{chat_title(title) or 'Meeting'}"
                print (f"Found Teams window: {teams_app}")
                is_teams_running = True
                is_teams_to_be_recorded = False

        if is_teams_running and not was_teams_running:
            print ("Teams started running")

            # Switch to scene to record
            keyboard.press('control+windows+shift+f1')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f1')

            # Switch to scene with camera
            keyboard.press('control+windows+shift+f8')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f8')

            # Start virtual camera
            keyboard.press('control+windows+shift+f11')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f11')

            # Switch camera off
            keyboard.press('control+windows+shift+f10')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f10')
        
        if is_teams_to_be_recorded and not was_teams_to_be_recorded:
            # Unpause recording
            keyboard.press('control+windows+shift+f4')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f4')
            # Start recording
            keyboard.press('control+windows+shift+f6')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f6')
        elif not is_teams_to_be_recorded and was_teams_to_be_recorded:
            # Pause recording
            keyboard.press('control+windows+shift+f5')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f5')

        if not is_teams_running and was_teams_running:
            print ("Teams stopped running")

            # Switch camera off
            keyboard.press('control+windows+shift+f10')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f10')

            # Stop virtual camera
            keyboard.press('control+windows+shift+f2')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f2')

            # Stop recording
            keyboard.press('control+windows+shift+f7')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f7')

            print ("Waiting for previous recording to be released...")
            moved = not os.path.exists("c:\\Users\\raul.mzabala\\Videos\\latest.mp4")
            while not moved:
                print ("Trying to rename the previous recording...")
                try:
                    os.replace(
                        "c:\\Users\\raul.mzabala\\Videos\\latest.mp4",
                        f"c:\\Users\\raul.mzabala\\Videos\\Captures\\{teams_app}.mp4"
                    )
                    moved = True
                except Exception as e:
                    print (f"Could not rename: {e}") 
                    time.sleep(1) 
            print ("Recording file renamed successfully.")

            # Switch to scene to record
            keyboard.press('control+windows+shift+alt+f1')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+alt+f1')

        was_teams_running = is_teams_running
        was_teams_to_be_recorded = is_teams_to_be_recorded
        time.sleep(1)   

def heartbeat():
    tick = 0
    while True:
        time.sleep(2)
        tick += 1
        try:
            alive = [t.name for t in threading.enumerate() if t.is_alive()]
            print(f"[heartbeat#{tick}] alive threads={alive}")
        except Exception as e:
            print(f"[heartbeat] error: {e}")


def watchdog():
    """Refreshes a dedicated file every 500 ms; supervisor.py watches its
    mtime to detect a hung daemon. Uses its own file handle (no shared lock)
    and skips fsync — the OS updates mtime on write, which is all the
    supervisor checks via stat()."""
    watchdog_path = log_base / "macro-daemon.watchdog"
    while True:
        try:
            with open(watchdog_path, "w", encoding="utf-8") as f:
                f.write(datetime.datetime.now().isoformat())
        except Exception:
            pass
        time.sleep(0.5)


def system_event_listener():
    """Hidden window that logs critical Windows broadcast messages so we can
    see what Windows is sending around the time the process dies (monitor
    disconnect, power transition, end-session, etc.)."""
    WM_DISPLAYCHANGE   = 0x007E
    WM_POWERBROADCAST  = 0x0218
    WM_DEVICECHANGE    = 0x0219
    WM_QUERYENDSESSION = 0x0011
    WM_ENDSESSION      = 0x0016
    WM_SETTINGCHANGE   = 0x001A

    NAMES = {
        WM_DISPLAYCHANGE: "DISPLAYCHANGE",
        WM_POWERBROADCAST: "POWERBROADCAST",
        WM_DEVICECHANGE: "DEVICECHANGE",
        WM_QUERYENDSESSION: "QUERYENDSESSION",
        WM_ENDSESSION: "ENDSESSION",
        WM_SETTINGCHANGE: "SETTINGCHANGE",
    }

    def wndproc(hwnd, msg, wparam, lparam):
        if msg in NAMES:
            print(f"[winevt] {NAMES[msg]} wparam=0x{wparam:X} lparam=0x{lparam & 0xFFFFFFFF:X}")
            try:
                _log_fp.flush()
                os.fsync(_log_fp.fileno())
            except Exception:
                pass
        return win32gui.DefWindowProc(hwnd, msg, wparam, lparam)

    try:
        wc = win32gui.WNDCLASS()
        wc.lpfnWndProc = wndproc
        wc.lpszClassName = "MacropadDaemonEventSink"
        wc.hInstance = win32api.GetModuleHandle(None)
        class_atom = win32gui.RegisterClass(wc)
        hwnd = win32gui.CreateWindow(
            class_atom, "MacropadDaemonEventSink", 0, 0, 0, 0, 0, 0, 0,
            wc.hInstance, None
        )
        print(f"[winevt] hidden event window created hwnd={hwnd}")
        win32gui.PumpMessages()
    except Exception as e:
        print(f"[winevt] listener died: {e}")
        traceback.print_exc()


def run_wsl_command(cmd):
    try:
        CREATE_NO_WINDOW = 0x08000000
        subprocess.Popen(
            ["wsl", "--"] + cmd,
            creationflags=CREATE_NO_WINDOW,
            stdin=subprocess.DEVNULL,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
            close_fds=True,
        )
        print(f"WSL command launched: {' '.join(cmd)}")
    except Exception as e:
        print(f"Error running WSL command {cmd}: {e}")


def monitor_power_state():
    print("Starting power state monitor")
    last_plugged = None
    while True:
        try:
            battery = psutil.sensors_battery()
            if battery is not None:
                plugged = bool(battery.power_plugged)
                if last_plugged is None:
                    last_plugged = plugged
                    print(f"Initial power state: {'plugged' if plugged else 'unplugged'}")
                elif plugged != last_plugged:
                    if plugged:
                        print("Charger connected -> starting geoserver container")
                        run_wsl_command(["docker", "container", "start", "geoserver"])
                    else:
                        print("Charger disconnected -> stopping geoserver container")
                        run_wsl_command(["docker", "container", "stop", "geoserver"])
                    last_plugged = plugged
        except Exception as e:
            print(f"Error monitoring power state: {e}")
        time.sleep(2)


## --- Anti-idle -------------------------------------------------------------
## Domain policy locks the session after a fixed idle period and the timeout is
## not ours to change. `custom_inactivity_check` in zones.json maps a monitor
## hardware id to a longer inactivity budget: while that monitor is attached we
## keep the session alive by synthesizing a phantom keypress before the system
## counter expires, and we STOP doing it once the human has been away for
## (budget - system limit) — so the session still locks, just at `budget`
## instead of at the policy value. With no listed monitor attached the budget
## equals the system limit and the loop never injects anything.
CUSTOM_INACTIVITY_CHECK = {}      ## hardware id -> inactivity budget in seconds
DEFAULT_INACTIVITY_LIMIT = 300    ## fallback when the registry says nothing
VK_NONAME = 0xFC                  ## reserved virtual key with no meaning
KEYEVENTF_KEYUP = 0x0002
INACTIVITY_TICK_SECONDS = 5
INACTIVITY_BUDGET_REFRESH_SECONDS = 30
## Ceiling relative to the system limit, so a policy change scales the cadence
## on its own. In practice the absolute cap below is the binding constraint.
INACTIVITY_INJECT_AT = 0.6
## Hard cap on the gap between injections, regardless of the fraction above.
## This is what actually governs the cadence: the 0.6 fraction alone (180 s
## against a 300 s limit) did NOT keep the session alive in testing even though
## the injections were provably registering in GetLastInputInfo. 10 s did. If
## the session ever locks again, lower this before suspecting anything else.
INACTIVITY_MAX_INTERVAL_SECONDS = 120


def read_system_inactivity_limit():
    """Seconds of inactivity after which policy locks the session.

    Two independent mechanisms can do it and both read the same idle counter:
    the secure screensaver (`ScreenSaveTimeOut`, but only when it is both
    active AND secure) and Winlogon's machine inactivity limit
    (`InactivityTimeoutSecs`). Whichever fires first is the real deadline, so
    take the minimum. Read at runtime rather than hardcoded because the GPO
    owns these values and can change them under us."""
    limits = []

    ## The Policies branch is what the GPO writes; the plain Control Panel key
    ## is the user's own setting and only matters when no policy applies.
    for path in (r"Software\Policies\Microsoft\Windows\Control Panel\Desktop",
                 r"Control Panel\Desktop"):
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, path) as key:
                def value(name):
                    try:
                        return str(winreg.QueryValueEx(key, name)[0])
                    except OSError:
                        return None

                timeout = value("ScreenSaveTimeOut")
                if timeout and value("ScreenSaveActive") == "1" and value("ScreenSaverIsSecure") == "1":
                    limits.append(int(timeout))
                    break
        except OSError:
            pass

    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                            r"SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System") as key:
            seconds = int(winreg.QueryValueEx(key, "InactivityTimeoutSecs")[0])
            if seconds > 0:
                limits.append(seconds)
    except OSError:
        pass

    if not limits:
        print(f"Anti-idle: no lock timeout found in registry, assuming {DEFAULT_INACTIVITY_LIMIT}s")
        return DEFAULT_INACTIVITY_LIMIT
    return min(limits)


def inactivity_budget(system_limit):
    """Inactivity budget for the monitors currently attached.

    Any monitor listed in `custom_inactivity_check` raises the budget to its
    value; with several attached the most permissive one wins."""
    if not CUSTOM_INACTIVITY_CHECK:
        return system_limit

    try:
        attached = {dev_id for dev_id, _ in active_monitors()}
    except Exception as e:
        print(f"Anti-idle: could not enumerate monitors: {e}")
        return system_limit

    budget = system_limit
    for hw_id, seconds in CUSTOM_INACTIVITY_CHECK.items():
        ## Same normalization as get_monitor_rect_by_alias: the `_N` suffix in
        ## zones.json distinguishes two heads of one model, irrelevant here.
        if hw_id.split('_')[0] in attached:
            budget = max(budget, int(seconds))
    return budget


def press_phantom_key():
    """Reset the system idle counter with the least intrusive input available.

    VK_NONAME is documented as reserved and carries no meaning, so it produces
    no character and no scancode anything translates — yet the injection still
    updates GetLastInputInfo, which is what both lock mechanisms read.

    Do NOT go back to F13-F24: this used F15 first, on the reasoning that no
    modern PC keyboard has the key. Wrong reasoning — xterm-style terminals,
    Windows Terminal among them, DO map those to ANSI escape sequences, and a
    WSL shell echoed each phantom press as stray characters.

    Emitted via keybd_event rather than the `keyboard` package on purpose: the
    daemon switches keyboard layouts constantly and we want a raw virtual-key
    event, not a layout-translated one."""
    user32 = ctypes.windll.user32
    user32.keybd_event(VK_NONAME, 0, 0, 0)
    user32.keybd_event(VK_NONAME, 0, KEYEVENTF_KEYUP, 0)


class LASTINPUTINFO(ctypes.Structure):
    _fields_ = [("cbSize", ctypes.wintypes.UINT),
                ("dwTime", ctypes.wintypes.DWORD)]


def monitor_inactivity():
    """Keeps the session alive up to the monitor-conditioned budget.

    GetLastInputInfo cannot distinguish our injections from real typing, so
    measuring how long the *human* has been away needs the absolute tick of the
    last input: after each injection we record the tick it produced, and any
    later tick that isn't that one is real user activity. Two clocks therefore
    run here — the system's idle time (which our own keypresses reset, and which
    decides WHEN to inject) and the human's idle time (which only real input
    resets, and which decides WHETHER to keep injecting)."""
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32

    lii = LASTINPUTINFO()
    lii.cbSize = ctypes.sizeof(LASTINPUTINFO)

    def last_input_tick():
        if not user32.GetLastInputInfo(ctypes.byref(lii)):
            return None
        return lii.dwTime

    def seconds_since(tick):
        ## Masking to 32 bits makes this correct across GetTickCount's ~49-day
        ## wrap-around, where a plain subtraction would go negative.
        return ((kernel32.GetTickCount() - tick) & 0xFFFFFFFF) / 1000.0

    system_limit = read_system_inactivity_limit()
    inject_after = system_limit * INACTIVITY_INJECT_AT
    if INACTIVITY_MAX_INTERVAL_SECONDS:
        inject_after = min(inject_after, INACTIVITY_MAX_INTERVAL_SECONDS)
    print(f"Anti-idle: policy locks after {system_limit}s; will inject past {int(inject_after)}s idle")

    human_tick = last_input_tick()
    injected_tick = None
    ## Set when an injection's tick could not be read back yet, so the next new
    ## tick we see is adopted as ours instead of read as the user returning.
    adopt_next_tick = False
    budget = inactivity_budget(system_limit)
    budget_checked_at = time.monotonic()
    budget_exhausted = False
    print(f"Anti-idle: initial budget {budget}s")

    while True:
        time.sleep(INACTIVITY_TICK_SECONDS)
        try:
            now = time.monotonic()
            if now - budget_checked_at >= INACTIVITY_BUDGET_REFRESH_SECONDS:
                budget_checked_at = now
                new_budget = inactivity_budget(system_limit)
                if new_budget != budget:
                    print(f"Anti-idle: budget {budget}s -> {new_budget}s (monitors changed)")
                    budget = new_budget
                    budget_exhausted = False

            tick = last_input_tick()
            if tick is None or human_tick is None:
                human_tick = tick
                continue

            if tick != injected_tick and tick != human_tick:
                if adopt_next_tick:
                    injected_tick = tick
                    adopt_next_tick = False
                else:
                    human_tick = tick
                    if budget_exhausted:
                        print("Anti-idle: user is back, budget re-armed")
                        budget_exhausted = False

            ## No extended budget for this monitor set: leave the policy alone.
            if budget <= system_limit:
                continue

            human_idle = seconds_since(human_tick)

            ## Stop `system_limit` short of the budget so the natural lock lands
            ## exactly at `budget` — the whole point of the config value.
            if human_idle >= budget - system_limit:
                if not budget_exhausted:
                    print(f"Anti-idle: human idle {int(human_idle)}s exhausted the {budget}s "
                          f"budget — letting the session lock")
                    budget_exhausted = True
                continue

            if seconds_since(tick) >= inject_after:
                press_phantom_key()
                ## Give the raw input thread a moment to register the event
                ## before reading back the tick it produced.
                time.sleep(0.05)
                new_tick = last_input_tick()
                if new_tick is None or new_tick == tick:
                    adopt_next_tick = True
                else:
                    injected_tick = new_tick
                ## Logging the tick on both sides of the injection is the whole
                ## diagnostic: an unchanged tick means keybd_event did nothing,
                ## a changed one means the system did register our keystroke and
                ## any remaining lock is decided by something other than
                ## GetLastInputInfo.
                debug and print(f"Anti-idle: phantom key sent, tick {tick} -> {new_tick} "
                                f"({'NO CHANGE' if new_tick == tick else 'idle reset'}), "
                                f"human idle {int(human_idle)}s of {budget}s")
        except Exception as e:
            print(f"Anti-idle error: {e}")
            traceback.print_exc()


def cycle_window_zone(direction):
    global WINDOW_ZONE_INDEX, WINDOW_ZONE_LAST_PRESS

    hwnd = win32gui.GetForegroundWindow()
    if not hwnd:
        return

    now = time.monotonic()
    last_press = WINDOW_ZONE_LAST_PRESS.get(hwnd)
    cur_idx = WINDOW_ZONE_INDEX.get(hwnd)
    if last_press is None or (now - last_press) > ZONE_CYCLE_RESET_SECONDS:
        cur_idx = None

    if cur_idx is None:
        new_idx = 0 if direction > 0 else len(ZONE_CYCLE) - 1
    else:
        new_idx = (cur_idx + direction) % len(ZONE_CYCLE)

    WINDOW_ZONE_INDEX[hwnd] = new_idx
    WINDOW_ZONE_LAST_PRESS[hwnd] = now
    zone_key = ZONE_CYCLE[new_idx]
    print(f"Cycling hwnd {hwnd} to zone[{new_idx}] = {zone_key}")
    move_window_to_zone(zone_key)


def on_zone_cycle_next():
    cycle_window_zone(+1)


def on_zone_cycle_prev():
    cycle_window_zone(-1)


def hotkey_listener():
    """Register ctrl+alt+shift+left/right via WinAPI RegisterHotKey and pump messages.

    Using the WinAPI instead of the `keyboard` library avoids installing a
    low-level keyboard hook (which on Windows interferes with the Win key and
    swallows non-matching keystrokes).
    """
    user32 = ctypes.windll.user32

    MOD_ALT = 0x0001
    MOD_CONTROL = 0x0002
    MOD_SHIFT = 0x0004
    MOD_NOREPEAT = 0x4000
    WM_HOTKEY = 0x0312
    VK_LEFT = 0x25
    VK_RIGHT = 0x27

    HOTKEY_PREV = 1
    HOTKEY_NEXT = 2

    mods = MOD_CONTROL | MOD_ALT | MOD_SHIFT | MOD_NOREPEAT

    if not user32.RegisterHotKey(None, HOTKEY_PREV, mods, VK_LEFT):
        print("Failed to register hotkey ctrl+alt+shift+left")
    if not user32.RegisterHotKey(None, HOTKEY_NEXT, mods, VK_RIGHT):
        print("Failed to register hotkey ctrl+alt+shift+right")

    msg = ctypes.wintypes.MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
        if msg.message == WM_HOTKEY:
            try:
                if msg.wParam == HOTKEY_PREV:
                    on_zone_cycle_prev()
                elif msg.wParam == HOTKEY_NEXT:
                    on_zone_cycle_next()
            except Exception as e:
                print(f"Error handling zone hotkey: {e}")
                traceback.print_exc()
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))


def monitor_keyboard_layout():
    """Polls the foreground window's keyboard layout every 250 ms and
    persists changes for the current program.

    Known limitation: doesn't reliably catch in-place layout changes inside
    Electron / WebView2 hosts (Teams, the new Outlook, ...) because their
    actual keyboard input is dispatched in a renderer subprocess and the
    foreground (host) thread's HKL doesn't consistently reflect what the
    user just selected. Tried TSF (ITfInputProcessorProfiles::GetCurrent
    Language) — turns out that interface reports *this* process's TSF
    state, not the system's, so it doesn't help for cross-process layout
    detection. For Teams, the previously saved layout is still applied
    on focus change via setup_program_layout; only Win+Space switches
    made while staying inside the app are not auto-detected.

    Race-safe: only persists a change when the foreground HWND hasn't moved
    since the previous tick. A different HWND means the layout difference
    is just the new window's input language, not a user-triggered switch,
    so we resync silently and leave APP_LAYOUTS untouched."""
    user32 = ctypes.windll.user32

    last_layout = None
    last_hwnd = None
    last_program = None
    while True:
        time.sleep(0.25)
        try:
            fg_hwnd = user32.GetForegroundWindow()
            if not fg_hwnd:
                continue
            fg_thread = user32.GetWindowThreadProcessId(fg_hwnd, None)
            if not fg_thread:
                continue
            layout = user32.GetKeyboardLayout(fg_thread) & 0xFFFFFFFF
            prog = current_program

            if last_hwnd is None:
                last_layout = layout
                last_hwnd = fg_hwnd
                last_program = prog
                continue

            if fg_hwnd != last_hwnd:
                last_layout = layout
                last_hwnd = fg_hwnd
                last_program = prog
                continue

            ## A program switch (even within the same window, e.g. a VS Code tab
            ## change) means any layout difference we see now is us applying the
            ## new program's layout — or the previous program's layout landing
            ## late — not a user Win+Space. Re-baseline and DON'T persist;
            ## otherwise a fast A->B->A switch saves B's layout under A. We only
            ## learn when both the window AND the program stayed put.
            if prog != last_program:
                last_layout = layout
                last_program = prog
                continue

            if layout == last_layout:
                continue

            prev_name = LAYOUT_MAP.get(last_layout, {}).get('name', f'0x{last_layout:08X}')
            new_name = LAYOUT_MAP.get(layout, {}).get('name', f'0x{layout:08X}')
            print(f"Keyboard layout changed for HWND {fg_hwnd}: {prev_name} -> {new_name}")
            last_layout = layout

            if not prog:
                continue

            ## Programs with a forced layout ignore observed OS changes: the
            ## forced value (e.g. ESP for Teams) re-applies on next focus
            ## event via setup_program_layout, and we don't want stale or
            ## unreliable WebView2-side detection to clobber that.
            if force_layout_for_program(prog) is not None:
                continue

            entry = APP_LAYOUTS.setdefault(prog, {})
            entry['layout'] = layout
            entry['last_used'] = datetime.datetime.now().isoformat()
            try:
                change_tray_icon_layout(layout)
            except Exception as e:
                print(f"Failed to update tray icon: {e}")
        except Exception as e:
            print(f"Layout monitor error: {e}")


# Cargar una imagen para el icono
def store_layouts():
    global APP_LAYOUTS
    if not PERSIST_APP_LAYOUTS:
        return
    while True:
        time.sleep(60)
        try:
            with open(APP_LAYOUTS_FILE, 'w') as file:
                json.dump(APP_LAYOUTS, file, indent=4)
                print (f"App layouts saved to {APP_LAYOUTS_FILE}")
        except Exception as e:
            print (f"Error saving app layouts: {e}")

def launch_program():

    global icon_global

    image = Image.open("default_icon.png")  # Reemplaza con tu icono
    menu = (item('Quit', quit),)
    icon_global = Icon("Macropad", image, menu=menu)

    ## Worker threads. Watchdog is mandatory because supervisor.py uses it to
    ## detect hangs. Heartbeat and winevt listener are pure diagnostics —
    ## only spawned when DIAGNOSTICS_ENABLED is True.
    threads = [
        threading.Thread(target=monitor_macropad, daemon=True, name="macropad"),
        threading.Thread(target=monitor_windows, daemon=True, name="winhook"),
        threading.Thread(target=monitor_keyboard_layout, daemon=True, name="kblayout"),
        threading.Thread(target=check_teams_window, daemon=True, name="teams"),
        threading.Thread(target=store_layouts, daemon=True, name="store_layouts"),
        threading.Thread(target=hotkey_listener, daemon=True, name="hotkeys"),
        threading.Thread(target=monitor_power_state, daemon=True, name="power"),
        threading.Thread(target=monitor_inactivity, daemon=True, name="antiidle"),
        threading.Thread(target=watchdog, daemon=True, name="watchdog"),
    ]
    if DIAGNOSTICS_ENABLED:
        threads.append(threading.Thread(target=heartbeat, daemon=True, name="heartbeat"))
        threads.append(threading.Thread(target=system_event_listener, daemon=True, name="winevt"))

    for t in threads:
        t.start()


    # Initialize system tray icon
    try:
        icon_global.run()
        print("[main] icon_global.run() returned cleanly — main thread will exit")
    except Exception as e:
        print(f"[main] icon_global.run() raised: {e}")
        traceback.print_exc()

def change_tray_icon_layout(layout_id):
    global current_tray_layout, icon_global
    new_icon = LAYOUT_MAP.get(layout_id,{}).get('icon','default_icon.png')
    icon_global.icon = Image.open(new_icon)
    current_tray_layout = layout_id

def kill_other_instances_same_script():
    me = os.getpid()

    target = os.path.abspath(sys.argv[0]).lower()
    target_script = target.split(os.sep)[-1]

    print (f"Checking for other instances of {target_script}...")
    for p in psutil.process_iter(["pid", "cmdline"]):
        try:

            ## Ignore myself
            pid = p.info["pid"]
            if pid == me:
                continue

            ## Get command line
            cmdline = p.info["cmdline"] or []

            if len(cmdline) < 2:
                continue

            if "python" in cmdline[0].lower() and target_script in cmdline[1].lower():
                # Mata árbol (hijos) primero
                for child in p.children(recursive=True):
                    try:
                        child.terminate()
                    except psutil.Error:
                        pass

                try:
                    p.terminate()  # educado
                except psutil.Error:
                    continue

                # Si no muere rápido, kill
                try:
                    p.wait(timeout=2)
                except psutil.TimeoutExpired:
                    try:
                        p.kill()
                    except psutil.Error:
                        pass


        except (psutil.NoSuchProcess, psutil.AccessDenied):
            pass

def respawn():
    # Flags para Windows: Proceso separado, nueva consola, sin heredar del padre
    DETACHED_PROCESS = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    
    subprocess.Popen(
        [sys.executable] + sys.argv,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP,
        close_fds=True
    )
    sys.exit()

def window_change_callback(hWinEventHook, event, hwnd, idObject, idChild, dwEventThread, dwmsEventTime):
    global current_program, current_program_hwnd
    try:
        ## EVENT_OBJECT_NAMECHANGE fires for *every* UI element in the system —
        ## toolbar items, listbox entries, taskbar buttons. We only care about
        ## actual top-level window events. This filter drops ~95% of callbacks
        ## before we touch psutil/win32 at all.
        if idObject != OBJID_WINDOW or idChild != CHILDID_SELF:
            return

        ## For title changes, ignore background windows; we only react when
        ## the foreground app retitles itself (e.g. Chrome tab switch).
        if event == EVENT_OBJECT_NAMECHANGE:
            if hwnd != win32gui.GetForegroundWindow():
                return
        elif event != EVENT_SYSTEM_FOREGROUND:
            return

        ## Cheap HWND comparison before the expensive psutil lookup. Only
        ## valid for FOREGROUND (NAMECHANGE on the same HWND can still mean
        ## the resolved program changed, e.g. a Chrome retitle).
        if event == EVENT_SYSTEM_FOREGROUND and hwnd == current_program_hwnd:
            return

        active_program = active_program_name()
        if active_program == current_program:
            current_program_hwnd = hwnd
            return

        print(f"Focus -> {active_program} (HWND {hwnd}, event {event:#x})")
        setup_program(active_program, hwnd)
    except Exception as e:
        print(f"Error in window_change_callback: {e}")
        traceback.print_exc()

def monitor_windows():
    global _event_proc, current_program, current_program_hwnd

    ## Initial setup
    current_program_hwnd = win32gui.GetForegroundWindow()
    active_program = active_program_name()
    setup_program(active_program, current_program_hwnd)

    _event_proc = WinEventProcType(window_change_callback)

    # Hook for FOCUS change
    ctypes.windll.user32.SetWinEventHook(
        EVENT_SYSTEM_FOREGROUND, EVENT_SYSTEM_FOREGROUND,
        0, _event_proc, 0, 0, WINEVENT_OUTOFCONTEXT
    )

    ## Hook for name change (to detect title changes)
    ctypes.windll.user32.SetWinEventHook(
        EVENT_OBJECT_NAMECHANGE, EVENT_OBJECT_NAMECHANGE,
        0, _event_proc, 0, 0, WINEVENT_OUTOFCONTEXT
    )


    msg = ctypes.wintypes.MSG()
    while ctypes.windll.user32.GetMessageW(ctypes.byref(msg), 0, 0, 0) != 0:
        ctypes.windll.user32.TranslateMessageW(ctypes.byref(msg))
        ctypes.windll.user32.DispatchMessageW(ctypes.byref(msg))


if __name__ == "__main__":
    ## Kill other instances of this same script
    kill_other_instances_same_script()

    ## Globally initialize serial port
    serial_port_initialize()

    ## Inventory of monitors at startup is purely informational — log it only
    ## when debugging because it can be noisy on multi-monitor setups.
    print_monitor_ids()
    load_zones_config()

    ## Create system tray icon and start monitoring
    launch_program()

