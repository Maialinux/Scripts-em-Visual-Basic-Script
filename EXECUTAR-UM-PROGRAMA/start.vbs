dim wshShell
Set wshShell = CreateObject("WScript.Shell")

WScript.Sleep(3600000)
wshShell.Run("programa.bat")