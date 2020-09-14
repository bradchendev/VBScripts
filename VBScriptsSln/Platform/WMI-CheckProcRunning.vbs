
' For Example:
'		Do While ChkProcRunning("notepad.exe")
'			Wscript.Sleep 3000
'		Loop
'		
'		WScript.Echo "Out"

' Function: Check a program Running or Not
' Input: (1)procName: program Name that you want to check
' Putput: 1 or 0
' 1 = program is running
' 0 = program is not running
Function ChkProcRunning(procName)
	Dim strComputer, objWMIService, colProcesses, chk
	strComputer = "."
	Set objWMIService = GetObject("winmgmts:\\"& strComputer & "\root\cimv2")
	Set colProcesses = objWMIService.ExecQuery ("Select * from Win32_Process Where Name = '" & procName & "'")

	If colProcesses.Count = 0 Then
		Wscript.Echo procName & " is not running."
		ChkProcRunning = 0
	Else
		Wscript.Echo procName & " is running."
		ChkProcRunning = 1
	End If
	
End Function