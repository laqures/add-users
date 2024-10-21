Dim objShell
Set objShell = CreateObject("WScript.Shell")

' Путь к PowerShell скрипту
Dim scriptPath
scriptPath = "\\dc1\cmd$\Add-User2.ps1" ' Замените на фактический путь к вашему скрипту

' Команда для запуска PowerShell скрипта
Dim command
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

' Запуск PowerShell скрипта
objShell.Run command, 1, True

' Освобождение объекта
Set objShell = Nothing
