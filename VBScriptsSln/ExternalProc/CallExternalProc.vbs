
CompressFileWith7z "zip", "D:\123.zip", "D:\SCRIPTS"


CompressFileWith7z "zip", "D:\123.zip", "D:\456"


CompressFileWith7z "zip", "D:\123.zip", "D:\789"



Sub CompressFileWith7z(compressType,compressFile, backupFile)
	Dim objShell
	Set objShell = CreateObject("WScript.Shell")
	objShell.Run "7z a -t" & compressType & " " & compressFile & " " & backupFile ,,True
	Set objShell = Nothing
	
	
End Sub





