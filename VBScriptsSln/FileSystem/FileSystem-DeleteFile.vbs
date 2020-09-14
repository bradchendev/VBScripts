Option Explicit


' Function: Delete a file
' input:
' (1)TargetFile: For example: "D:\Backup\backup.bak"
' Output: none
'
' For Example
'DeleteFile "D:\20070815.bak"
'
Sub DeleteFile(TargetFile)
	Dim ofile
	Set ofile=CreateObject("Scripting.FileSystemObject")
	'ofile.GetFile(TargetFile)
	ofile.DeleteFile TargetFile
	Set ofile= Nothing
End Sub

