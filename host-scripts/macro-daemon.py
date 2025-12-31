import os
import sys

import serial
import time
import pystray
from pystray import MenuItem as item, Icon
from PIL import Image
import threading
import pygetwindow as gw

import win32api
import win32gui
import win32con
import win32process

import json
import re
from pathlib import Path
import datetime
import subprocess
import keyboard

import ctypes
import psutil
import uuid
import traceback
import socket

base_path = Path(sys.argv[0]).resolve().parent
os.chdir(base_path)

latest_uuid = None
was_teams_running = False
serial_port = None

APP_OVERRIDES = {}
ZONE_DEFINITIONS = {}
MONITOR_ALIASES = {}
BORDER_OFFSET = {}
HARDWARE_ID_MAP = {}
TEAMS_TOP = 0
TEAMS_LEFT = 0
LAYOUT_DROP_DAYS = 30

## Load app layouts from json file
PERSIST_APP_LAYOUTS = True
APP_LAYOUTS_FILE = "./app_layouts.json"

if PERSIST_APP_LAYOUTS and os.path.exists(APP_LAYOUTS_FILE):
    with open(APP_LAYOUTS_FILE, 'r') as file:
        APP_LAYOUTS = json.load(file)
        ## Remove layouts not used in the last 30 days
        cutoff_date = datetime.datetime.now() - datetime.timedelta(days=LAYOUT_DROP_DAYS)
        APP_LAYOUTS = {app: data for app, data in APP_LAYOUTS.items() if 'last_used' in data and datetime.datetime.fromisoformat(data['last_used']) >= cutoff_date}
else:
    APP_LAYOUTS = {}

LAST_APP_SWITCH_TIME = datetime.datetime.now()

running_config={}
configs={}
toggles={}


def print_monitor_ids():
    print("\n--- ESCANEANDO MONITORES CONECTADOS ---")
    monitors = win32api.EnumDisplayMonitors()
    for i, (hMonitor, hdc, rect) in enumerate(monitors):
        monitor_info = win32api.GetMonitorInfo(hMonitor)
        adapter_name = monitor_info['Device']
        
        try:
            # Obtenemos el dispositivo MONITOR asociado al adaptador
            # El segundo 0 es el índice del monitor en ese adaptador
            device = win32api.EnumDisplayDevices(adapter_name, 0, 0)
            device_id = device.DeviceID
            print(f"Monitor {i}:")
            print(f"  Handle: {hMonitor}")
            print(f"  Adapter: {adapter_name}")
            print(f"  DeviceID: {device_id}") # <--- ESTO ES LO QUE NECESITAS COPIAR
        except Exception as e:
            print(f"  Error leyendo ID: {e}")
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
            print(f"Error al obtener información del monitor: {e}")
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

            print(f"Cargadas {len(ZONE_DEFINITIONS)} zonas y {len(HARDWARE_ID_MAP)} monitores hardware.")
    except Exception as e:
        print(f"Error cargando zones.json: {e}")

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
            # ¿Es este monitor 'dev_id' uno de los míos conocidos?
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


def get_process_name(hwnd):
    try:
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        proc = psutil.Process(pid)
        return proc.name().lower()
    except:
        return ""
    
def move_window_to_zone(zone_key):
    global TEAMS_TOP, TEAMS_LEFT, BORDER_OFFSET

    zone = ZONE_DEFINITIONS.get(zone_key)
    if not zone:
        print(f"Zona {zone_key} no existe")
        return

    hwnd = win32gui.GetForegroundWindow()
    if not hwnd: return

    # --- RESTAURAR SI MAXIMIZADA ---
    placement = win32gui.GetWindowPlacement(hwnd)
    is_maximized = (placement[1] == win32con.SW_SHOWMAXIMIZED)
    is_minimized = (placement[1] == win32con.SW_SHOWMINIMIZED)

    if is_maximized or is_minimized:
        win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

    # --- 1. OBTENER MONITOR DE INICIO (OBLIGATORIO) ---
    start_rect = get_monitor_rect_by_alias(zone['monitor'])
    if not start_rect: 
        print(f"Monitor de inicio '{zone['monitor']}' no encontrado.")
        return
    
    # --- 2. OBTENER MONITOR DE FIN (OPCIONAL / FALLBACK) ---
    end_alias = zone.get('monitor_end', zone['monitor'])
    end_rect = get_monitor_rect_by_alias(end_alias)

    # AQUÍ ESTÁ LA MAGIA DEL FALLBACK:
    if end_rect:
        # Escenario Casa: Ambos monitores existen
        # Calculamos la unión de ambos
        s_left, s_top, s_right, s_bottom = start_rect
        e_left, e_top, e_right, e_bottom = end_rect
        
        canvas_left = min(s_left, e_left)
        canvas_top = min(s_top, e_top)
        canvas_right = max(s_right, e_right)
        canvas_bottom = max(s_bottom, e_bottom)
        # print(f"Dual Monitor Mode: {zone['monitor']} -> {end_alias}")
    else:
        # Escenario Trabajo: El monitor final no está conectado
        # Degradamos suavemente: El lienzo total es SOLO el monitor de inicio
        canvas_left, canvas_top, canvas_right, canvas_bottom = start_rect
        print(f"Single Monitor Fallback: '{end_alias}' no detectado. Usando solo '{zone['monitor']}'.")

    canvas_width = canvas_right - canvas_left
    canvas_height = canvas_bottom - canvas_top

    # --- 3. CALCULAR COORDENADAS ---
    # Los porcentajes se aplican sobre el canvas calculado (sea doble o simple)
    
    raw_x = canvas_left + int(canvas_width * (zone['min_x'] / 100))
    raw_y = canvas_top + int(canvas_height * (zone['min_y'] / 100))
    
    raw_x2 = canvas_left + int(canvas_width * (zone['max_x'] / 100))
    raw_y2 = canvas_top + int(canvas_height * (zone['max_y'] / 100))

    raw_w = raw_x2 - raw_x
    raw_h = raw_y2 - raw_y

    # --- 4. APLICAR CORRECCIÓN DE BORDES Y OVERRIDES ---
    app_name = get_process_name(hwnd)
    app_adj = APP_OVERRIDES.get(app_name, {})

    final_x = raw_x + BORDER_OFFSET["x"] + app_adj.get("x",0)
    final_y = raw_y + BORDER_OFFSET["y"] + app_adj.get("y",0)
    final_w = raw_w + BORDER_OFFSET["w"] + app_adj.get("w",0)
    final_h = raw_h + BORDER_OFFSET["h"] + app_adj.get("h",0)

    # --- 5. EJECUTAR ---
    try:
        win32gui.MoveWindow(hwnd, final_x, final_y, final_w, final_h, True)
        win32gui.SetForegroundWindow(hwnd)
        
        if zone.get("is_teams_zone", False):
            TEAMS_LEFT = final_x
            TEAMS_TOP = final_y
        
    except Exception as e:
        print(f"Error: {e}")

def get_running_layout():
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    threadID = ctypes.windll.user32.GetWindowThreadProcessId(hWnd, None)
    hkl = ctypes.windll.user32.GetKeyboardLayout(threadID)
    layout_id = hkl & 0xFFFFFFFF
    return layout_id


def switch_layout(delay = 0.05, tries=2):
    if tries <=0:
        ## For some reason, Microsoft Notepad does not switch layout properly
        print ("Max tries exceeded for layout switch")
        return

    required_layout = get_app_layout()
    starting_layout = get_running_layout()
    if starting_layout != required_layout:
        keyboard.press('windows+space')
        time.sleep(delay)
        keyboard.release('windows+space')
        time.sleep(delay)
        print (f"Switched layout from {hex(starting_layout)} to {hex(get_running_layout())} seeking {hex(required_layout)}")

    ## If not correct (fast windows switch or other issues), try again
    required_layout = get_app_layout()
    resulting_layout = get_running_layout()
    if resulting_layout != required_layout:
        print (f"Layout missed from {hex(starting_layout)} to {hex(resulting_layout)} seeking {hex(required_layout)}")
        switch_layout(delay= 2*delay, tries=tries-1)



def open_window(filtro_regex):

    if ',' in filtro_regex:
        parts = filtro_regex.split(',')
        filtro_regex = parts[1]

    programs = running_config.get('programs', {})
    if filtro_regex not in programs:
        print (f"Program {filtro_regex} was not recognized")
        return 
    
    program_name = programs[filtro_regex]['program']
    window_name = programs[filtro_regex]['window']
    multiple_instances = programs[filtro_regex].get('multiple_instances',False)

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

    ventanas=[]
    win32gui.EnumWindows(callback, ventanas)

    if len(ventanas)==0:
        print (f"Launching program {program_name}")
        subprocess.Popen(f"start {program_name}", shell=True)
    elif win32gui.GetForegroundWindow() in ventanas:
        print (f"Window for {filtro_regex} is already active")
        if multiple_instances:
            print (f"Launching another instance of {program_name}")
            subprocess.Popen(f"start {program_name}", shell=True)
    else:
        print (f"Found existing window(s) for {filtro_regex}, bringing to front")
        for hwnd in ventanas:
            if win32gui.IsIconic(hwnd):  
                print (f"Restoring minimized window for {filtro_regex}")
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                continue
            else:
                print (f"Bringing to front window for {filtro_regex}")
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                time.sleep(0.05)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                continue

    ## Cambiar layout si es necesario 
    switch_layout()

def get_app_layout():
    global APP_LAYOUTS
    active_program = active_program_name()

    if active_program in APP_LAYOUTS:
        APP_LAYOUTS[active_program]['last_used'] = datetime.datetime.now().isoformat()
    else:
        APP_LAYOUTS[active_program]= {
            "layout": running_config.get('layouts', {}).get(running_config['layout'],None),
            "last_used": datetime.datetime.now().isoformat()
        }   

    return APP_LAYOUTS.get(
        active_program,
        None
    )['layout']

# Función para obtener el nombre de la ventana activa
def get_active_window():
    window = win32gui.GetForegroundWindow()
    if not window:
        return 'None'

    window_title = win32gui.GetWindowText(window)
    _, pid = win32process.GetWindowThreadProcessId(window)
    try:
        proc = psutil.Process(pid)
        exe = proc.name()  # Nombre del ejecutable, por ejemplo: Teams.exe
        window_title = win32gui.GetWindowText(window)
    except psutil.NoSuchProcess:
        return None,None
    return exe,window_title

def lookup_config(window_title):
    global configs
    global toggles

    try:
        config_version = datetime.datetime.fromtimestamp(Path("./config.json").stat().st_mtime)

        if not configs or config_version > configs['version']:
            with open("./config.json", 'r') as file:
                configs = json.load(file)
                configs['version'] = config_version

        claves_ordenadas = sorted(configs.keys(), key=len, reverse=False)

        new_config = {
            "window": None,
            "colors": {},
            "keys": {}  
        }
        for clave in claves_ordenadas:
            #print (f"Procesando {clave} para {window_title}")
            if re.search(clave, window_title,re.IGNORECASE) or clave=='.':
                print(f"{clave} matched for {window_title}")  
                if not new_config['window']:
                    new_config['window'] = clave

                for key, value in configs[clave]['keys'].items():
                    new_config['keys'][key]=value

                for key, value in configs[clave]['colors'].items():
                    new_config['colors'][key]=value

                for key, value in configs[clave].get('toggles',{}).items():
                    toggle = toggles.setdefault(key, {})
                    toggle['config'] = value
                    toggle.setdefault('pos', 0)

                if (configs[clave]).get('symbols',None):
                    new_config['symbols'] = configs[clave]['symbols'] 

                if (configs[clave]).get('layout',None):
                    new_config['layout']=configs[clave]['layout']

                if (configs[clave]).get('programs',None):
                    new_config['programs']=configs[clave]['programs']
                
                if (configs[clave]).get('layouts',None):
                    new_config['layouts']=configs[clave]['layouts']

        # prettyprint new_config
        #print (f"Configuración compuesta: {new_config}") # en prettyprint

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
    global toggles
    global running_config
    global serial_port 

    print ("toggle key called for "+toggle_name)

    cur_pos = toggles[toggle_name].get('pos',0)
    options = toggles[toggle_name]['config']
    num_options = len(options)
    next_pos = (cur_pos+1) % num_options
    toggles[toggle_name]['pos']=next_pos
    next_leds = toggles[toggle_name]['config'][next_pos]['color']
    next_strokes = toggles[toggle_name]['config'][next_pos]['strokes']
    next_key = toggles[toggle_name]['config'][next_pos]['key']

    running_config['colors'][next_key]=next_leds
    for stroke in next_strokes:
        print (f"Pressing {stroke}")
        keyboard.press(stroke)
        time.sleep(0.05)
        keyboard.release(stroke)
    command = json.dumps(running_config) + '\n'
    serial_port.write(command.encode())  # Enviar el comando al puerto (debe ser codificado en bytes)

def active_program_name():
    try:
        active_program, active_window = get_active_window()
    except Exception as ex:
        active_program = 'explorer.exe'

    if active_program == 'chrome.exe':
        active_program = active_window.split(' - ')[0]
    elif active_program == 'msrdc.exe':
        active_program = active_window

    return active_program

def save_running_layout(prev_program=None):
    global APP_LAYOUTS, LAST_APP_SWITCH_TIME

    ## Prevent fast switch wrong saves
    if LAST_APP_SWITCH_TIME + datetime.timedelta(seconds=2) > datetime.datetime.now():
        print ("Skipping save due to fast switch")
        return
    
    LAST_APP_SWITCH_TIME = datetime.datetime.now()

    # Save current layout for previous program
    running_layout = get_running_layout()

    if not prev_program:
        return 

    ## Save layout for previous program
    if APP_LAYOUTS.get(prev_program,None)!=running_layout:
        print (f"Saving layout {running_layout} for {prev_program}")
        APP_LAYOUTS[prev_program] = {
            "layout": running_layout,
            "last_used": datetime.datetime.now().isoformat()
        }
        if PERSIST_APP_LAYOUTS:
            with open(APP_LAYOUTS_FILE, 'w') as file:
                json.dump(APP_LAYOUTS, file, indent=4)

    return


# Función principal que monitorea el cambio de ventana 
def monitor_window_focus():
    global configs, serial_port, splits, running_config, APP_LAYOUTS

    while True:
        try:
            configs = {}
            prev_program = ''

            if serial_port:
                serial_port.close()
                serial_port = None

            serial_port = serial.Serial('COM4', 115200, timeout=1)  
            while True:
                if serial_port.in_waiting:
                    data = json.loads(serial_port.readline().decode('utf-8').strip())
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

                active_program = active_program_name()
                if  active_program != prev_program:

                    ## Save layout for previous program
                    save_running_layout(prev_program)

                    # Switch to new program
                    prev_program = active_program

                    # Load new config and send to pad
                    running_config = lookup_config(active_program)
                    command = json.dumps(running_config) + '\n'
                    serial_port.write(command.encode())  # Enviar el comando al puerto (debe ser codificado en bytes)

                    # Change keyboard layout if needed
                    if active_program!= 'explorer.exe':
                        switch_layout()

                # Wait for a while before checking again
                time.sleep(0.5)  

        except Exception as ex:
            print(f"Process failed {ex}")
            time.sleep(5)

# Función para salir del programa
def salir(icon, item):
    icon.stop()
    sys.exit()

def chat_title(texto):
    partes = [parte.strip() for parte in texto.split("|")]
    for i, parte in enumerate(partes):
        if parte == "Bosonit" and i > 0:
            return partes[i - 1]
    return None

def check_teams_window():
    global was_teams_running, TEAMS_TOP, TEAMS_LEFT
    print ("Starting Teams window monitor")
    while True:
        is_teams_running = False
        for ventana in gw.getAllWindows():
            titulo = ventana.title or ""
            titulo_minus = titulo.lower()
            if "teams" in titulo_minus and ventana.top == TEAMS_TOP and ventana.left == TEAMS_LEFT:
                print (f"All ventana info: {ventana}")
                teams_app = f"{datetime.datetime.now().strftime('%Y%m%d_%H%M')}_{chat_title(titulo) or 'Meeting'}"
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

# Cargar una imagen para el icono
def crear_icono():
    image = Image.open("icono.png")  # Reemplaza con tu icono
    menu = (item('Salir', salir),)
    icon = Icon("MiApp", image, menu=menu)

    # Iniciar el proceso en segundo plano
    hilo = threading.Thread(target=monitor_window_focus, daemon=True)
    hilo_teams = threading.Thread(target=check_teams_window, daemon=True)

    hilo.start()
    hilo_teams.start()

    icon.run()

def kill_other_instances_same_script():
    me = os.getpid()

    target = os.path.abspath(sys.argv[0]).lower()
    target_script = target.split(os.sep)[-1]

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

            print (f"checking target {target_script} against cmdline: {cmdline}")

            if "python" in cmdline[0].lower() and target_script in cmdline[1].lower():
                print (f"Will kill PID {pid} with cmdline: {cmdline} for target {target}")

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

if __name__ == "__main__":
    kill_other_instances_same_script()
    print_monitor_ids()
    load_zones_config()

    # Flags de Windows para lanzar el proceso sin ventana y desacoplado
    DETACHED_PROCESS         = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    CREATE_NO_WINDOW         = 0x08000000

    # Lanzar WSL oculto
    script_dir = os.path.dirname(os.path.abspath(__file__))
    vbs_path = os.path.join(script_dir, "wsl_hidden.vbs")

    p = subprocess.Popen(
        ["wscript.exe", vbs_path],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW
    )


    print(f"WSL lanzado con PID {p.pid}, ejecutando sleep infinity")
    crear_icono()

