# SPDX-FileCopyrightText: Daniel Schaefer 2023 for Framework Computer
# SPDX-FileCopyrightText: Modified by Raul Martinez Zabala 2025
# SPDX-License-Identifier: MIT
#
# Handle button pressed on the macropad
# FIXED: Ghosting (Hi-Z) AND Aliasing (Active Discharge)
#
import time
import board
import busio
import digitalio
import analogio
import usb_hid
import usb_cdc
from adafruit_hid.keyboard import Keyboard
from adafruit_hid.keycode import Keycode
from framework_is31fl3743 import IS31FL3743
import json
import traceback

# === Matrix and Threshold Configuration ===
MATRIX_COLS = 8
MATRIX_ROWS = 4 
ADC_THRESHOLD = 0.55  
DEBOUNCE_DELAY = 0.05  

# List of currently pressed keys
pressed = []

# Matrix layout 
# CORREGIDO: En tu código original 'b4' aparecía dos veces.
# He cambiado el de la fila 1 (index 5) a "c4" para que coincida con tu reporte.
MATRIX = [
    ["f1", "b3", "c3", "d3", "e3", "f3", "b4", "d4"],
    ["f4", "a1", "a2", None, "a4", "c4", "e4", "f2"], # <-- c4 corregido aquí
    ["b1", "c1", "d1", "e1", "b2", "c2", "d2", "e2"],
    [None, None, None, None, "a3", None, None, None]
]

# Mapping LEDs
MATRIX_LED_MAP = {
    "a1" : 40,  "a2" : 37,  "a3" : 52,  "a4" : 49,
    "b1" : 4,   "b2" : 1,   "b3" : 16,  "b4" : 13,
    "c1" : 22,  "c2" : 19,  "c3" : 34,  "c4" : 31, # c4 tiene LED map, faltaba en MATRIX
    "d1" : 58,  "d2" : 55,  "d3" : 70,  "d4" : 67,
    "e1" : 25,  "e2" : 61,  "e3" : 64,  "e4" : 28,
    "f1" : 7,   "f2" : 43,  "f3" : 46,  "f4" : 10
}

SYMBOLS = {}
MATRIX_COLORS = {}
MATRIX_COMMANDS = {}

# Init Keyboard
try:
    keyboard = Keyboard(usb_hid.devices)
except:
    keyboard = None

# === Hardware Setup ===
gp6 = digitalio.DigitalInOut(board.GP6)
gp6.direction = digitalio.Direction.INPUT
gp7 = digitalio.DigitalInOut(board.GP7)
gp7.direction = digitalio.Direction.INPUT

# Mux Pins
mux_enable = digitalio.DigitalInOut(board.MUX_ENABLE)
mux_enable.direction = digitalio.Direction.OUTPUT
mux_enable.value = False 
mux_a = digitalio.DigitalInOut(board.MUX_A)
mux_a.direction = digitalio.Direction.OUTPUT
mux_b = digitalio.DigitalInOut(board.MUX_B)
mux_b.direction = digitalio.Direction.OUTPUT
mux_c = digitalio.DigitalInOut(board.MUX_C)
mux_c.direction = digitalio.Direction.OUTPUT

# KSO Pins
kso_pins = [
    digitalio.DigitalInOut(x)
    for x in [
        board.KSO0, board.KSO1, board.KSO2, board.KSO3,
        board.KSO4, board.KSO5, board.KSO6, board.KSO7,
        board.KSO8, board.KSO9, board.KSO10, board.KSO11,
        board.KSO12, board.KSO13, board.KSO14, board.KSO15
    ]
]
for kso in kso_pins:
    kso.direction = digitalio.Direction.INPUT

adc_in = analogio.AnalogIn(board.GP28)

boot_done = digitalio.DigitalInOut(board.BOOT_DONE)
boot_done.direction = digitalio.Direction.OUTPUT
boot_done.value = False

# === Helper Functions ===
def mux_select_row(row):
    mux_a.value = row & 0x01
    mux_b.value = row & 0x02
    mux_c.value = row & 0x04

# === CORRECCIÓN CRÍTICA DE ALIASING ===
def drive_col(col, value):
    pin = kso_pins[col]
    if value == 0:
        # ACTIVAR: Modo salida y valor 0
        pin.direction = digitalio.Direction.OUTPUT
        pin.value = False
    else:
        # DESACTIVAR (Con descarga activa)
        # 1. Forzamos a 1 (HIGH) brevemente para borrar la capacitancia de 0V
        pin.direction = digitalio.Direction.OUTPUT
        pin.value = True
        # 2. Ahora que está limpia a 3.3V, la dejamos flotando
        pin.direction = digitalio.Direction.INPUT

def to_voltage(adc_sample):
    return (adc_sample * 3.3) / 65536

# === LED Driver ===
sdb = digitalio.DigitalInOut(board.GP29)
sdb.direction = digitalio.Direction.OUTPUT
sdb.value = True
i2c = busio.I2C(board.SCL, board.SDA)
i2c.try_lock()
i2c.scan()
i2c.unlock()
is31 = IS31FL3743(i2c)
is31.set_led_scaling(0x20)
is31.global_current = 0x20
is31.enable = True

sleep_pin = digitalio.DigitalInOut(board.GP0)
sleep_pin.direction = digitalio.Direction.INPUT

def matrix_paint():
    global MATRIX_LED_MAP, MATRIX_COLORS
    for key in MATRIX_LED_MAP.keys():
        value = MATRIX_COLORS.get(key,None)
        if value:
            try:
                r = int(value[:2], 16)
                g = int(value[2:4], 16)
                b = int(value[-2:], 16)
                idx = MATRIX_LED_MAP[key]
                is31[idx + 2] = r
                is31[idx + 1] = g
                is31[idx + 0] = b
            except: pass
        else:
            idx = MATRIX_LED_MAP[key]
            is31[idx + 2] = 0; is31[idx + 1] = 0; is31[idx + 0] = 0


def process_strokes(code,press):
    global keyboard, SYMBOLS

    if not keyboard:
        return

    escaped = False
    for key_char in code:
        release = True
        if escaped:
            escaped = False
            if key_char == key_char.upper(): 
                ## Make it release within sequence only if we are in release mode
                release = not press
            else:
                ## Make it release within sequence only if we are in press mode
                release = press
            if key_char.upper() =='P': 
                time.sleep(0.15)
            else:  
                key_char = "\\" + key_char
        else:
            if key_char == '\\': 
                escaped = True
            else: 
                escaped = False
        if key_char:
            symbol = SYMBOLS.get(key_char.upper(), None)
            if symbol:
                key_code = getattr(Keycode, symbol, None)
                if key_code:
                    if press and not release:
                        keyboard.press(key_code)
                    elif press and release:
                        keyboard.press(key_code)
                        time.sleep(0.05)
                        keyboard.release(key_code)
                    elif not press and release:
                        keyboard.release(key_code)


def process_key(pressed, released):
    global MATRIX_COMMANDS, SYMBOLS, usb_serial
    sorted_pressed = sorted(pressed)
    sorted_released = sorted(released)

    ## Pressed part
    lookup_key = "-".join(sorted_pressed)
    code = MATRIX_COMMANDS.get(lookup_key, None)
    if code:
        if code.startswith("MSG:"):
            to_send = {"key": lookup_key, "code": code[4:], "pressed": True}
            print (f"Sending message: {to_send}")
            if usb_serial:
                usb_serial.write((json.dumps(to_send) + '\n').encode())
                usb_serial.flush()
        else:
            process_strokes(code, True)
    
    ## Released part
    lookup_key = "-".join(sorted_released)
    code = MATRIX_COMMANDS.get(lookup_key,None)
    if code and not code.startswith("MSG:"):
        process_strokes(code, False)

    ## Release all if nothing is pressed
    if not pressed:
        keyboard.release_all()

def get_raw_matrix_state():
    current_state = []
    for col in range(MATRIX_COLS):
        drive_col(col, 0) 
        for row in range(MATRIX_ROWS):
            mux_select_row(row)
            # Aumentamos ligeramente la pausa para estabilizar
            time.sleep(0.00005) 
            if to_voltage(adc_in.value) < ADC_THRESHOLD:
                key_name = MATRIX[row][col]
                if key_name: current_state.append(key_name)
        drive_col(col, 1) 
    return current_state

def load_config(config):
    global MATRIX_COLORS, MATRIX_COMMANDS, SYMBOLS
    MATRIX_COLORS = config.get('colors', {})
    MATRIX_COMMANDS = config.get('keys', {})
    if config.get('symbols', None): SYMBOLS = config['symbols']
    matrix_paint()

# === Main Loop ===
print("Starting Anti-Ghosting Engine V3")

last_read_state = []
stable_pressed = []

while True:
    for col in range(MATRIX_COLS): drive_col(col, 1) # Reset All to Hi-Z

    try: usb_serial = usb_cdc.data
    except: usb_serial = None
        
    while True:
        try:
            is31.enable = sleep_pin.value
            if not sleep_pin.value: time.sleep(5); continue
            
            if usb_serial and usb_serial.in_waiting:
                try:
                    data = usb_serial.readline().decode().strip()
                    if data: load_config(json.loads(data))
                except: pass

            raw_state = get_raw_matrix_state()
            
            if sorted(raw_state) != sorted(last_read_state):
                last_read_state = raw_state
                time.sleep(DEBOUNCE_DELAY)
                continue 
            
            raw_set = set(raw_state)
            stable_set = set(stable_pressed)
            
            if raw_set != stable_set:
                stable_pressed = list(raw_set)
                to_press = raw_set - stable_set
                to_release = stable_set - raw_set
                process_key(to_press, to_release)
                
            time.sleep(0.01)

        except Exception as e:
            print(f"Error: {e}")
            try: keyboard.release_all()
            except: pass
            time.sleep(1)