

	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "7z a -t zip D:\123.zip D:\SCRIPTS" ,,True
	Set objShell = Nothing

	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "7z a -t zip D:\456.zip D:\456" ,,True
	Set objShell = Nothing

	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "7z a -t zip D:\789.zip D:\S789" ,,True
	Set objShell = Nothing











