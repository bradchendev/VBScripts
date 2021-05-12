' Program: CopyRecentBakFile2RemoteFileServer_v2.vbs
' Create by: Brad Chen
' Create Date: 2021/05/11
' Modify Date: 2021/05/12

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0

	Dim LogFile
	LogFile = "C:\BackupScript\" & Wscript.ScriptName & ".log"
	
	'Set s = CreateObject("Scripting.Dictionary")
	's.Add "0", "D:\DBBackup\Tfs_ITTEAM" 
	's.Add "1", "D:\DBBackup\Tfs_configuration"
	's.Add "2", "D:\DBBackup\MyDB"
	
	' Declare a class for RecentFile
	Class RecentFile
	   Public Path
	   Public DateLastModified
	End Class
	
	' Backup Source folder
	Set s = CreateObject("Scripting.Dictionary")
	s.Add "0", "D:\DBBackup\Tfs_ITTEAM" 
	s.Add "1", "D:\DBBackup\Tfs_configuration"

	' File Server UNC Path
	Dim FileSrvUnc
	FileSrvUnc = "\\10.0.0.xx\xxx"

	For Each I in s
		'WScript.Echo "D.Item(" & I & ") : " & d.Item(I)
		CopyRecentBackup2FileSrv s.Item(I), FileSrvUnc
		WriteLog WScript.ScriptName & " Complete!!!", LogFile
	Next

	
	
Sub CopyRecentBackup2FileSrv(SourceFolder, FileSrvUncPath)
	Dim SourceFile
	' 01.Get recentFile from folder
	SourceFile = GetRecentFileInFolder(SourceFolder)
	WriteLog "recent Backup: " & SourceFile, LogFile
	
	' 02.Create today folder if the folder not exist
	Dim DestFldr, todayFldr
	todayFldr= DateTimeConvert(NOW,1,1)
	' for example: todayFldr= "20210510"
	
	DestFldr = FileSrvUncPath & "\" & todayFldr
	' for example: DestFldr= "\\10.13.0.120\psp\20210510"
	
	' create remote folder if the folder not exist
	If Not(FolderExist(DestFldr)) Then
		WriteLog "Create " & DestFldr & " folder", LogFile
		'WScript.Echo "need to create the folder"
		folder = CreateFolder(DestFldr)
	End If
	
	'WScript.Echo SourceFile
	'WScript.Echo DestFldr
	'Wscript.Sleep 3000
	
	' 03.Copy Sourcefile to Destination folder
	CopyAndOverrideToRemoteServer SourceFile, DestFldr & "\"
	'CopyAndOverrideToRemoteServer "G:\Backup\db20210510112301.bak", "G:\Test\"
	'CopyAndOverrideToRemoteServer "G:\Backup\db20210510112301.bak", "\\10.13.0.120\psp\20210510\"

End Sub	


Function GetRecentFileInFolder(tarFolder)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(tarFolder)
	Set files = folder.Files
	
	Set r = New RecentFile
	 
	For each fileIdx In files
		If  LCase(Right(fileIdx.Path,4)) = ".bak" Then
			'Wscript.Echo fileIdx.Path
			If r.Path = Empty Then
				r.Path = fileIdx.Path
				r.DateLastModified = fileIdx.DateLastModified
				'WScript.Echo fileIdx.Path
			Else
				'WScript.Echo fileIdx.Path
				If fileIdx.DateLastModified > r.DateLastModified Then
					r.Path = fileIdx.Path
					r.DateLastModified = fileIdx.DateLastModified
				End If
			End If	
		End If
	Next
	
	'Wscript.Echo r.Path
	'WScript.Echo r.DateLastModified
	
	GetRecentFileInFolder = r.Path
	
End Function

Function FolderExist(fldr)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FolderExists(fldr)) Then
      FolderExist = True
   Else
      FolderExist = False
   End If
End Function

Function CreateFolder(folder)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.CreateFolder(folder)
   CreateFolder = f.Path
End Function

Function DateTimeConvert(strdatetime,format_type,fill0)
	Dim YY, MM, DD, hh, mins, secs
	
	YY = Year(strdatetime)
	If fill0 = 1 Then
		If LEN(Month(strdatetime)) < 2 Then MM = "0" & Month(strdatetime) Else MM = Month(strdatetime) End If
		If LEN(Day(strdatetime)) < 2 Then DD = "0" & Day(strdatetime) Else DD = Day(strdatetime) End If
		If LEN(Hour(strdatetime)) < 2 Then hh = "0" & Hour(strdatetime) Else hh = Hour(strdatetime) End If
		If LEN(Minute(strdatetime)) < 2 Then mins = "0" & Minute(strdatetime) Else mins = Minute(strdatetime) End If
		If LEN(Second(strdatetime)) < 2 Then secs = "0" & Second(strdatetime) Else secs = Second(strdatetime) End If
	ElseIf fill0 = 0 then
		MM = Month(strdatetime)
		DD = Day(strdatetime)
		hh = Hour(strdatetime)
		mins = Minute(strdatetime)
		secs = Second(strdatetime)
	End if

	Select Case format_type
		Case 1
			DateTimeConvert = YY & MM & DD
		Case 2
			DateTimeConvert = YY & MM & DD & hh & mins & secs
		Case 3
			DateTimeConvert = YY & "-" & MM & "-" & DD
		Case 4
			DateTimeConvert = YY & "-" & MM & "-" & DD & " " & hh & ":" & mins & ":" & secs
	End Select 

End Function


Sub CopyAndOverrideToRemoteServer(SourceFile, Destinationfile)
	Dim ofile
	Set ofile=CreateObject("Scripting.FileSystemObject")
	ofile.GetFile(SourceFile)
	' Copy to file server
	
	'WScript.Echo SourceFile
	'WScript.Echo Destinationfile
	'Wscript.Sleep 5000
	WriteLog "Copy " & SourceFile & " to " & Destinationfile, LogFile
	
	On Error Resume Next 
	ofile.CopyFile SourceFile , Destinationfile , True
	If Err.Number <> 0 Then 
		WriteLog "Err.Number is " & Err.Number, LogFile
		WriteLog "Err.Description is " & Err.Description, LogFile
		'Wscript.Echo "Err.Number is " & Err.Number
		'Wscript.Echo "Err.Description is " & Err.Description
		'Wscript.Sleep 5000
		'WScript.Quit 
	End If 
	Set ofile= Nothing

End Sub


Sub WriteLog(msg, FileName)
	'FileName = "c:\Brad\testfile.txt"
	Dim fso, MyFile
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set MyFile = fso.OpenTextFile(FileName, ForAppending, true)
	
	' Write to the file.
	MyFile.WriteLine DateTimeConvert(NOW,4,1) & " " & msg
	MyFile.Close
End Sub