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


program_name_pattern = r'[^a-z0-9._\-\s]'
normalize_program_name_re = re.compile(program_name_pattern, re.IGNORECASE)   


try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2) # 2 = PROCESS_PER_MONITOR_DPI_AWARE
except Exception:
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1) # 1 = PROCESS_SYSTEM_DPI_AWARE
    except Exception:
        ctypes.windll.user32.SetProcessDPIAware()

debug = True

base_path = Path(sys.argv[0]).resolve().parent
os.chdir(base_path)

latest_uuid = None
was_teams_running = False

serial_port = None
serial_lock = threading.Lock()
serial_queue = queue.Queue()

APP_OVERRIDES = {}
ZONE_DEFINITIONS = {}
MONITOR_ALIASES = {}
BORDER_OFFSET = {}
HARDWARE_ID_MAP = {}
TEAMS_TOP = 0
TEAMS_LEFT = 0
icon_global = None

current_tray_layout = None
current_program_hwnd = None
current_program = None

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

## Windows Event Hook constants
EVENT_SYSTEM_FOREGROUND = 0x0003
WINEVENT_OUTOFCONTEXT = 0x0000
EVENT_OBJECT_NAMECHANGE = 0x800C

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
    global ZONE_DEFINITIONS, HARDWARE_ID_MAP, BORDER_OFFSET, APP_OVERRIDES
    try:
        with open("zones.json", "r") as f:
            data = json.load(f)
            ZONE_DEFINITIONS = data.get("areas", {})
            HARDWARE_ID_MAP = data.get("hardware_mapping", {}) # <--- NUEVO
            APP_OVERRIDES = data.get("app_overrides", {})

            hostname = socket.gethostname()
            offset_key = f"offsets-{hostname}"
            BORDER_OFFSET = data.get(offset_key, data.get("offsets-default", {}))

            print(f"Loaded {len(ZONE_DEFINITIONS)} zones and {len(HARDWARE_ID_MAP)} hardware monitors.")
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
    global TEAMS_TOP, TEAMS_LEFT, BORDER_OFFSET

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
        win32gui.MoveWindow(hwnd, final_x, final_y, final_w, final_h, True)
        win32gui.SetForegroundWindow(hwnd)
        
        if zone.get("is_teams_zone", False):
            TEAMS_LEFT = final_x
            TEAMS_TOP = final_y
        
    except Exception as e:
        print(f"Error: {e}")



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
    global all_configurations, all_configurations_version, toggles

    try:
        config_version = datetime.datetime.fromtimestamp(Path("./config.json").stat().st_mtime)

        if config_version != all_configurations_version:
            all_configurations_version = config_version
            with open("./config.json", 'r') as file:
                all_configurations = json.load(file)

        config_keys = sorted(all_configurations.keys(), key=len, reverse=False)

        print (f"Keys are {config_keys}")

        new_config = {
            "window": None,
            "colors": {},
            "keys": {}  
        }
        print (f"config keys are {config_keys}")
        for config_key in config_keys:
            #print (f"Procesando {config_key} para {window_title}")
            if re.search(config_key, window_title,re.IGNORECASE) or config_key=='.':
                print(f"{config_key} matched for {window_title}")  
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
        return 'None'

    window_title = win32gui.GetWindowText(window)
    _, pid = win32process.GetWindowThreadProcessId(window)

    try:
        process = psutil.Process(pid)
        executable = process.name() 
        window_title = win32gui.GetWindowText(window)
    except psutil.NoSuchProcess:
        return None,None

    return executable,window_title


def active_program_name():
    try:
        active_program, active_window = get_active_window()
    except Exception as ex:
        active_program = 'explorer.exe'

    active_program = re.sub(normalize_program_name_re, '', str(active_program.lower())).strip()[:50]
    active_window = re.sub(normalize_program_name_re, '', str(active_window.lower())).strip()[:50]

    if active_program == 'chrome.exe':
        active_program = active_window.split(' - ')[0]
    elif active_program == 'msrdc.exe':
        active_program = active_window

    return active_program

def send_command_to_macropad(command_dict):
    global serial_port, macropad_config
    print (f"Sending configuration to macropad")

    ## Avoid sending same config again
    if command_dict == macropad_config:
        print (f"Configuration unchanged, not sending to macropad")
        return
    
    with serial_lock:
        print (f"Sending command to macropad")
        macropad_config = command_dict.copy()
        command = json.dumps(macropad_config) + '\n'
        serial_port.write(command.encode())
        serial_port.flush()

def setup_program_macropad(active_program):
    global running_config
    print (f"Setting up macropad for program: {active_program}")
    running_config = lookup_config(active_program)
    send_command_to_macropad(running_config)

def get_app_layout():
    global APP_LAYOUTS
    active_program = active_program_name()

    if active_program in APP_LAYOUTS:
        app_layout = APP_LAYOUTS[active_program]['layout']
        APP_LAYOUTS[active_program]['last_used'] = datetime.datetime.now().isoformat()
    else:
        app_layout = running_config.get('layouts', {}).get(running_config.get('layout', current_tray_layout), current_tray_layout)
        APP_LAYOUTS[active_program]= {
            "layout": app_layout,
            "last_used": datetime.datetime.now().isoformat()
        }   
    debug and print (f"Layout for {active_program} is {LAYOUT_MAP.get(app_layout,{}).get('name', 'Unknown')}")
    return app_layout

def setup_program_layout(active_program):
    print (f"Setting up layout for program: {active_program}")

    ## Get required layout
    required_layout_id = get_app_layout()
    if not required_layout_id:
        print ("No layout required")
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
    setup_program_macropad(active_program)
    setup_program_layout(active_program)
    current_program = active_program
    current_program_hwnd = hwnd


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

def check_teams_window():
    global was_teams_running, TEAMS_TOP, TEAMS_LEFT
    print ("Starting Teams window monitor")
    while True:
        is_teams_running = False
        for window in gw.getAllWindows():
            title = window.title or ""
            title_lower = title.lower()
            if "teams" in title_lower and window.top == TEAMS_TOP and window.left == TEAMS_LEFT:
                print (f"All window info: {window}")
                teams_app = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}_{chat_title(title) or 'Meeting'}"
                print (f"Found Teams window: {teams_app}")
                is_teams_running = True
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

            # Stop recording (just in case)
            keyboard.press('control+windows+shift+f7')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f7')

            if os.path.exists("c:\\Users\\raul.mzabala\\Videos\\latest.mp4"):
                print ("Stopping recording...")
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
                            f"c:\\Users\\raul.mzabala\\Videos\\Captures\\{teams_app}_orphan_prev_meeting.mp4"
                        )
                        moved = True
                    except Exception as e:
                        print (f"Could not rename: {e}") 
                        time.sleep(1) 

            print ("Recording file renamed successfully.")

            # Start recording
            keyboard.press('control+windows+shift+f6')
            time.sleep(0.1)
            keyboard.release('control+windows+shift+f6')

        elif not is_teams_running and was_teams_running:
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
        time.sleep(3)   

def on_layout_shortcut():
    global current_tray_layout, current_program, APP_LAYOUTS

    print ("Layout change detected, saving current program layout...")
    if not current_program:
        print ("No current program to save layout for.")
        return
    
    if current_tray_layout == 0x04090409:
        print ("Current layout is US English, switching to Spanish.")
        target_layout = 0x040a0c0a
    else:
        print ("Current layout is not US English, switching to US English.")
        target_layout = 0x04090409

    APP_LAYOUTS[current_program]['layout'] = target_layout
    current_tray_layout = target_layout
    setup_program_layout(current_program)
    change_tray_icon_layout(target_layout)
    save_current_program()

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

    ## Hook for layout change
    keyboard.add_hotkey('windows+space', on_layout_shortcut, suppress=False)
    keyboard.add_hotkey('left alt+shift', on_layout_shortcut, suppress=False)

    ## Configure threads
    thread_macropad = threading.Thread(target=monitor_macropad, daemon=True)
    thread_teams = threading.Thread(target=check_teams_window, daemon=True)
    thread_window_hook = threading.Thread(target=monitor_windows, daemon=True)    
    thread_store_layouts = threading.Thread(target=store_layouts, daemon=True)

    # Start threads
    thread_macropad.start()
    thread_window_hook.start()
    thread_teams.start()
    thread_store_layouts.start()


    # Initialize system tray icon
    icon_global.run()

def change_tray_icon_layout(layout_id):
    global current_tray_layout, icon_global
    new_icon = LAYOUT_MAP.get(layout_id,{}).get('icon','default_icon.png')
    icon_global.icon = Image.open(new_icon)
    current_tray_layout = layout_id

def save_current_program():
    global APP_LAYOUTS, current_tray_layout, icon_global, current_program_hwnd, current_program

    ## Prevent null program
    if not current_program:
        return

    ## Check for layout change
    layout_id = current_tray_layout

    print (f"Saving layout {LAYOUT_MAP.get(layout_id,{}).get('name', 'Unknown')} for program {current_program}")

    APP_LAYOUTS[current_program]['layout']=layout_id
    APP_LAYOUTS[current_program]['last_used']=datetime.datetime.now().isoformat()
    

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
    global current_program, current_program_hwnd, current_tray_layout

    active_program = active_program_name()
    if active_program == current_program:
        ## No change
        return
    elif event == EVENT_SYSTEM_FOREGROUND:
        ## Focus changed
        print(f"Detected focus change to HWND: {hwnd}")
        setup_program(active_program, hwnd)
    elif event == EVENT_OBJECT_NAMECHANGE:
        ## Title changed
        print(f"Detected title change in HWND: {hwnd}")
        setup_program(active_program, hwnd)

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

    ## Print monitor IDs and load zones
    print_monitor_ids()
    load_zones_config()

    ## Create system tray icon and start monitoring
    launch_program()

