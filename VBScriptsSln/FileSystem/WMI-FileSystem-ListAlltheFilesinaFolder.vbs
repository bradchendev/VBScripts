

' Function: List All the Files in a Folder
' input:
' (1)strComputer: "."
' (2)FolderPath: "C:\Scripts"
' Output: none
'
' For Example
'
'strComputer = "."
'FolderPath = "C:\Scripts"
'
'ListFilesinFolder strComputer, FolderPath
'
Sub ListFilesinFolder(strComputer,FolderPath)
	Dim objWMIService,colFileList

	Set objWMIService = GetObject("winmgmts:" _
			& "{impersonationLevel=impersonate}!\\" & strComputer & "\root\cimv2")

	Set colFileList = objWMIService.ExecQuery _
			("ASSOCIATORS OF {Win32_Directory.Name='" & FolderPath & "'} Where " _
			& "ResultClass = CIM_DataFile")

	For Each objFile In colFileList
		Wscript.Echo objFile.Name
	Next

End Sub
