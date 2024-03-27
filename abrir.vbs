' Crear el objeto Shell
Set WshShell = CreateObject("WScript.Shell")

' Ruta al archivo .hta
Dim rutaHta
rutaHta = "\scripts\index.hta"

' Ejecutar el archivo .hta
WshShell.Run Chr(34) & rutaHta & Chr(34), 1, False

'WScript.Sleep 500
'WshShell.SendKeys "+{TAB}"
'WScript.Sleep 200
'WshShell.SendKeys "{ENTER}"

'MsgBox "OK"
