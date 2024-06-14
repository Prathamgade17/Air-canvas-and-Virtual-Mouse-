set oShell = CreateObject ("Wscript.shell")
Dim strArgs
strArgs = "cmd /c AirCanvasDemo2.bat"
oShell.Run strArgs, 0, false