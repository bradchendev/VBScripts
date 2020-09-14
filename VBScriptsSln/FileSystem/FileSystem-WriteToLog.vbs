' Function:Debug用或自訂執行過程與記錄
' Input:
' (1)Set objFSO = CreateObject("Scripting.FileSystemObject")
' (2)LogFile = "C:\Temp\MyLogfile.log"
' (3)LogMessages = "寫入記錄檔文字"
' Putput:none
' For Example:
'	Const ForAppending = 8
'	Dim objFSO, LogFile
'	Set objFSO = CreateObject("Scripting.FileSystemObject")
'	LogFile = "C:\Temp\MyLogfile.log"
'
'	WriteToLog objFSO, LogFile, "寫入記錄檔文字"
'
'	Set objFSO = Nothing
'
Sub WriteToLog(objFSO, LogFile, LogMessages)
	Dim objTextFile
	Set objTextFile = objFSO.OpenTextFile(LogFile, ForAppending, True)
	objTextFile.WriteLine(LogMessages)
	Set objTextFile = Nothing
End Sub