Option Explicit
' Function: Compress file with 7z
' Input:
' (1)compressType: "zip" or "7z"
' (2)passwd: Specifies password
' (3)compressFile: compress file name with full path
' (4)backupFile: backup file name with full path
' Output: none
'
' For Example
'CompressFileWith7z "7z", "myPassword", "D:\test.7z", "D:\test.bak"
'
'CompressFileWith7z "zip", "", "D:\test.zip", "D:\test.bak"

Sub CompressFileWith7z(compressType, _
							passwd, _
							compressFile, _
							backupFile)
	Dim objShell, cmdString
	Set objShell = CreateObject("WScript.Shell")
	
	If passwd = "" Then ' ¤£¥[±K
		cmdString = "7z.exe a -t" & _
						compressType & _
						" " & _
						compressFile & _
						" " & _
						backupFile
	Else ' ¥[±K½X
		cmdString = "7z.exe a -t" & _
						compressType & _
						" -p" & _
						passwd & _
						" " & _
						compressFile & _
						" " & _
						backupFile

	End If
	'Wscript.Echo cmdString
	objShell.Run cmdString ,,True
					
	Set objShell = Nothing
End Sub