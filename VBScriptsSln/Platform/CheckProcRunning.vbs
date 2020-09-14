Option Explicit
	Dim objShell', Running
	Set objShell = CreateObject("WScript.Shell")
		'Running = 0
		'Running = CheckProcRunning(objShell,"notepad.exe")
		Do While CheckProcRunning(objShell,"notepad.exe")
			Wscript.Sleep 3000
			'Running = CheckProcRunning(objShell,"notepad.exe")
		Loop
	Set objShell = Nothing
	Wscript.Echo "Exit"


' Function: Use tasklist Command To Check running program
' Input: 
' (1)objShell : [Set objShell = CreateObject("WScript.Shell")]
' (2)procName : Program Name For example: "notepad.exe"
' OutPut:
' (1) 1 : Program is Running
' (2) 0 : Program is Not Running
Function CheckProcRunning(objShell,procName)
	Dim procLen
	procLen = Len(procName)

	Dim objWshScriptExec, objStdOut, strLine, Do_Loop, ThisTimeRunning

			Set objWshScriptExec = objShell.Exec("tasklist")
			Set objStdOut = objWshScriptExec.StdOut
			strLine = ""
			ThisTimeRunning = 0
			
				'WScript.Echo objStdOut.ReadAll
			Do While Not objStdOut.AtEndOfStream
				strLine = objStdOut.ReadLine
				If Left(Lcase(strLine),procLen) = procName Then
					WScript.Echo strLine
					ThisTimeRunning = 1
					Exit Do
				End If
			Loop
			
			Set objStdOut = Nothing
			Set objWshScriptExec = Nothing

	CheckProcRunning = ThisTimeRunning
	
End Function