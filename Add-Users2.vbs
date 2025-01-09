On Error Resume Next

' Получаем LDAP путь из аргументов
strDN = Wscript.Arguments(0)

' Удаляем префикс LDAP:// и доменное имя
strDN = Replace(strDN, "LDAP://", "")
strDN = Mid(strDN, InStr(strDN, "/") + 1)

' Создаем объект WScript.Shell
Set objShell = CreateObject("WScript.Shell")

' Получаем путь к временной директории
strTempPath = objShell.ExpandEnvironmentStrings("%temp%")

' Создаем объект FileSystemObject
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Открываем файл для записи с перезаписью, если он существует
Set objFile = objFSO.CreateTextFile(strTempPath & "\ou100med.txt", True)

' Записываем LDAP путь в файл
objFile.WriteLine(strDN)

' Закрываем файл
objFile.Close

' Путь к PowerShell скрипту
Dim scriptPath
scriptPath = "\\dc1\NETLOGON\Scripts\Add-User.ps1" ' Замените на фактический путь к вашему скрипту

' Команда для запуска PowerShell скрипта
Dim command
command = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File """ & scriptPath & """"

' Запуск PowerShell скрипта
objShell.Run command, 1, True

' Освобождение объектов
Set objFile = Nothing
Set objFSO = Nothing
Set objShell = Nothing
