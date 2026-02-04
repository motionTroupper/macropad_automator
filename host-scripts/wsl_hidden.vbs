Set WshShell = CreateObject("WScript.Shell")
WshShell.Run "wsl.exe -d Ubuntu -u root /root/startup.sh", 0, True
WshShell.Run "python.exe c:\users\raulm\LocalData\macropad", 0, True
