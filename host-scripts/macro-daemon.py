import serial
import time
import pystray
from pystray import MenuItem as item, Icon
from PIL import Image
import sys
import threading
import pygetwindow as gw
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

latest_window = ''
latest_uuid = None


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

# Cargar una imagen para el icono
def crear_icono():
    image = Image.open("icono.png")  # Reemplaza con tu icono
    menu = (item('Salir', salir),)
    icon = Icon("MiApp", image, menu=menu)

    # Iniciar el proceso en segundo plano
    hilo = threading.Thread(target=monitor_window_focus, daemon=True)
    hilo.start()

    icon.run()

if __name__ == "__main__":

    # Flags de Windows para lanzar el proceso sin ventana y desacoplado
    DETACHED_PROCESS         = 0x00000008
    CREATE_NEW_PROCESS_GROUP = 0x00000200
    CREATE_NO_WINDOW         = 0x08000000

    # Lanzar WSL oculto
    p = subprocess.Popen(
        ["wscript.exe", "C:\\Users\\raul.mzabala\\Local data\\scripts\"\wsl_hidden.vbs"],
        stdin=subprocess.DEVNULL,
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
        creationflags=DETACHED_PROCESS | CREATE_NEW_PROCESS_GROUP | CREATE_NO_WINDOW
    )

    print(f"WSL lanzado con PID {p.pid}, ejecutando sleep infinity")
    crear_icono()

