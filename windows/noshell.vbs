Set WshShell = CreateObject("WScript.Shell") 
WshShell.Run chr(34) & "D:\backup\backup_singgle_database.bat" & Chr(34), 0
Set WshShell = Nothing