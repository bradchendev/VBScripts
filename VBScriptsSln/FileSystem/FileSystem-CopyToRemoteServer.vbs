Option Explicit


' Function: Copy And Override To RemoteServer
' input:
' (1)SourceFile: For example: "D:\Backup\backup.bak"
' (2)Destinationfile: For example: "G:\Backup\backup.bak"
' Output: none
'
' For Example
'CopyAndOverrideToRemoteServer "D:\Backup\20070815-2.bak", "D:\20070815.bak"
'
Sub CopyAndOverrideToRemoteServer(SourceFile,Destinationfile)
	Dim ofile
	Set ofile=CreateObject("Scripting.FileSystemObject")
	ofile.GetFile(SourceFile)
	' Copy to file server
	ofile.CopyFile SourceFile , Destinationfile , True
	Set ofile= Nothing

End Sub

