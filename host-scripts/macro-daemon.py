import serial
import time
import pystray
from pystray import MenuItem as item, Icon
from PIL import Image
import sys
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
import os
import subprocess
import keyboard

import ctypes
import psutil
import uuid
import traceback
import msvcrt


latest_window = ''
latest_uuid = None
was_teams_running = False

ZONE_DEFINITIONS = {}
MONITOR_ALIASES = {}
TEAMS_TOP = 0
TEAMS_LEFT = 0

def load_zones_config():
    global ZONE_DEFINITIONS, MONITOR_ALIASES
    try:
        with open("zones.json", "r") as f:
            data = json.load(f)
            ZONE_DEFINITIONS = data.get("areas", {})
            MONITOR_ALIASES = data.get("monitors", {})
            print(f"Cargadas {len(ZONE_DEFINITIONS)} zonas.")
    except Exception as e:
        print(f"Error cargando zones.json: {e}")

load_zones_config()


# Cambiar al directorio donde está el script
os.chdir(os.path.dirname(os.path.abspath(__file__)))

# Configurar el puerto COM4
ser = None
KLF_ACTIVATE = 0x00000001

layouts = {
    "EN": 67699721,
    "ES": 67767306
}

programs={
    "chrome":{
        "program":"chrome.exe",
        "window":"chrome.exe"
    },
    "Ubuntu":{
        "program":"wt -w 0 nt -p \"Ubuntu\" -- bash -lc \"tmux\"",
        "window":"WindowsTerminal.exe"
    }
}

running_config={}
configs={}
toggles={}

def get_monitor_rect(alias_name):
    # Función auxiliar para obtener la geometría de un monitor por su alias
    idx = int(MONITOR_ALIASES.get(alias_name, -1))
    if idx == -1: return None

    monitors = win32api.EnumDisplayMonitors()
    if idx >= len(monitors): return None

    handle = monitors[idx][0]
    info = win32api.GetMonitorInfo(handle)
    return info['Work'] # (left, top, right, bottom) Globales

# AJUSTE FINO PARA BORDES INVISIBLES (Windows 10/11)
# Valores típicos: Left/Right = 7px, Bottom = 7px. Top = 0px.
# Esto hace que la ventana sea un poco más grande para ocultar la sombra.
BORDER_OFFSET = {
    "x": -8,   # Mover a la izquierda para comerse el borde izq
    "y": -1,    # Arriba suele estar bien (la barra de título es sólida)
    "w": 16,   # Sumar 7 de izq + 7 de der
    "h": 9     # Sumar 9 de abajo
}

def move_window_to_zone(zone_key):
    global TEAMS_TOP, TEAMS_LEFT

    zone = ZONE_DEFINITIONS.get(zone_key)
    if not zone:
        print(f"Zona {zone_key} no existe")
        return

    hwnd = win32gui.GetForegroundWindow()
    if not hwnd: return

    # 1. Obtener monitores
    start_rect = get_monitor_rect(zone['monitor'])
    if not start_rect: return
    
    end_alias = zone.get('monitor_end', zone['monitor'])
    end_rect = get_monitor_rect(end_alias)
    if not end_rect: return

    # Geometría monitores
    s_left, s_top, s_right, s_bottom = start_rect
    s_width, s_height = (s_right - s_left), (s_bottom - s_top)

    e_left, e_top, e_right, e_bottom = end_rect
    e_width, e_height = (e_right - e_left), (e_bottom - e_top)

    # 2. CALCULAR PÍXELES TEÓRICOS (Lógica de % original)
    raw_x = s_left + int(s_width * (zone['min_x'] / 100))
    raw_y = s_top + int(s_height * (zone['min_y'] / 100))
    
    # El punto final X2/Y2 se calcula sobre el monitor de destino
    raw_x2 = e_left + int(e_width * (zone['max_x'] / 100))
    raw_y2 = e_top + int(e_height * (zone['max_y'] / 100))

    raw_w = raw_x2 - raw_x
    raw_h = raw_y2 - raw_y

    # 3. APLICAR CORRECCIÓN DE BORDES (FUDGE FACTOR)
    # Solo aplicamos corrección si la ventana toca los bordes? 
    # Generalmente se aplica siempre para que las ventanas adyacentes se toquen visualmente.
    
    final_x = raw_x + BORDER_OFFSET["x"]
    final_y = raw_y + BORDER_OFFSET["y"]
    final_w = raw_w + BORDER_OFFSET["w"]
    final_h = raw_h + BORDER_OFFSET["h"]

    # Seguridad: Si maximizas (0-100%), a veces es mejor usar el comando nativo de maximizar
    # Pero si es multi-monitor (span), usamos MoveWindow.
    
    # 4. EJECUTAR
    try:
        if win32gui.IsIconic(hwnd):
            win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
        
        win32gui.MoveWindow(hwnd, final_x, final_y, final_w, final_h, True)
        win32gui.SetForegroundWindow(hwnd)
        print(f"Posicionado en zona {zone_key} en monitor {zone['monitor']}")
        print(f"Desde ({raw_x},{raw_y}) {raw_w}x{raw_h} a ({final_x},{final_y}) {final_w}x{final_h} con bordes")
        print(f"Movido con ajuste de bordes: {final_w}x{final_h}")

        if zone.get("is_teams_zone", False):
            TEAMS_LEFT = final_x
            TEAMS_TOP = final_y
        
    except Exception as e:
        print(f"Error: {e}")

def obtener_layout_actual():
    # Obtiene el ID del thread con foco (ventana activa)
    hWnd = ctypes.windll.user32.GetForegroundWindow()
    threadID = ctypes.windll.user32.GetWindowThreadProcessId(hWnd, None)
    # Obtiene el layout del teclado (HKL)
    hkl = ctypes.windll.user32.GetKeyboardLayout(threadID)
    layout_id = hkl & 0xFFFFFFFF
    return layout_id


def cambiar_layout(layout,recheck):
    curr_layout = obtener_layout_actual()
    if curr_layout != layouts.get(layout,None):
        if recheck:
            time.sleep(0.1)
            cambiar_layout(layout,False)
        else:
            print (f"Cambiando layout a {layout} desde {curr_layout}")
            keyboard.press_and_release('windows+space')

def open_window(filtro_regex):
    if ',' in filtro_regex:
        parts = filtro_regex.split(',')
        cambiar_layout(parts[0],True)
        filtro_regex = parts[1]
        time.sleep(0.1)

    if filtro_regex not in programs:
        print (f"Program {filtro_regex} was not recognized")
        return 
    
    program_name = programs[filtro_regex]['program']
    window_name = programs[filtro_regex]['window']

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
        subprocess.Popen(f"start {program_name}", shell=True)
    else:
        for hwnd in ventanas:
            if hwnd == win32gui.GetForegroundWindow():
                # Abrir una nueva pestana si ya está en primer plano
                subprocess.Popen(f"start {program_name}", shell=True)
            elif win32gui.IsIconic(hwnd):  
                # Si la ventana está minimizada, restaurarla
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)
                time.sleep(0.1)  # Pequeña pausa para asegurar la restauración
            else:
                # Si no está minimizada, minimizarla y restaurarla para traerla al frente
                win32gui.ShowWindow(hwnd, win32con.SW_MINIMIZE)
                time.sleep(0.1)
                win32gui.ShowWindow(hwnd, win32con.SW_RESTORE)

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
    global ser

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
    ser.write(command.encode())  # Enviar el comando al puerto (debe ser codificado en bytes)


# Función principal que monitorea el cambio de ventana 
def monitor_window_focus():
    global configs
    global ser
    global splits
    global running_config
    while True:
        try:
            configs = {}
            current_program = ''

            if ser:
                ser.close()
                ser = None

            ser = serial.Serial('COM4', 115200, timeout=1)  
            while True:
                if ser.in_waiting:
                    data = json.loads(ser.readline().decode('utf-8').strip())
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

                try:
                    active_program, active_window = get_active_window()
                except Exception as ex:
                    print (f"Could not get active program")

                if not active_program:
                    continue
                elif active_program == 'chrome.exe':
                    active_program = active_window.split(' - ')[0]
                elif active_program == 'msrdc.exe':
                    active_program = active_window

                if  active_program != current_program:
                    current_program = active_program
                    running_config = lookup_config(active_program)
                    command = json.dumps(running_config) + '\n'
                    ser.write(command.encode())  # Enviar el comando al puerto (debe ser codificado en bytes)
                    if current_program!='explorer.exe' and running_config.get('layout'):
                        cambiar_layout(running_config['layout'],False)

                time.sleep(0.5)  # Espera un poco antes de volver a comprobar

        except Exception as ex:
            print(f"Process failed {ex}")
        finally:
            try:
                if ser and ser.is_open:
                    ser.close()
            except Exception as e:
                pass
            ser = None
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


if __name__ == "__main__":

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

