' Function:Debug�ΩΦۭq����L�{�P�O��
' Input:
' (1)Set objFSO = CreateObject("Scripting.FileSystemObject")
' (2)LogFile = "C:\Temp\MyLogfile.log"
' (3)LogMessages = "�g�J�O���ɤ�r"
' Putput:none
' For Example:
'	Const ForAppending = 8
'	Dim objFSO, LogFile
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	LogFile = "C:\Temp\MyLogfile.log"
'
'	WriteToLog objFSO, LogFile, "�g�J�O���ɤ�r"
'
'	Set objFSO = Nothing
'
Sub WriteToLog(objFSO, LogFile, LogMessages)
	Dim objTextFile
	Set objTextFile = objFSO.OpenTextFile(LogFile, ForAppending, True)
	objTextFile.WriteLine(LogMessages)
	Set objTextFile = Nothing
End Sub