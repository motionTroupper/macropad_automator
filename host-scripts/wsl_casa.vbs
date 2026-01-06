Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "wsl.exe -d Ubuntu-24.04 --exec sleep infinity", 0, False
WshShell.Run "python L:\repos\macropad_automator\host-scripts\macro-daemon.py", 0, False
