' Program: CopyRecentBakFile2RemoteFileServer
' Create by: Brad Chen
' Create Date: 2021/05/21


' 01.Get recentFile from folder
	Dim SourceFolder, SourceFile
	SourceFolder = "G:\0_Backup\folder1"
	
	Class RecentFile
	   Public Path
	   Public DateLastModified
	End Class
	
	SourceFile = GetRecentFileInFolder(SourceFolder)

	'WScript.Echo SourceFile 
	'Wscript.Sleep 5000

' 02.Create today folder if the folder not exist
	Dim DestFldr, todayFldr
	todayFldr= DateTimeConvert(NOW,1,1)
	' for example: todayFldr= "20210510"
	
	DestFldr = "\\10.13.0.120\System1\" & todayFldr
	
	If Not(FolderExist(DestFldr)) Then
		'WScript.Echo "need to create the folder"
		folder = CreateFolder(DestFldr)
	End If
	
	'WScript.Echo DestFldr
	'WScript.Echo todayFldr
	'Wscript.Sleep 3000


' 03.Copy Sourcefile to Destination folder

	
	WScript.Echo SourceFile
	WScript.Echo DestFldr

	CopyAndOverrideToRemoteServer SourceFile, DestFldr & "\"
	'CopyAndOverrideToRemoteServer SourceFile, "G:\Test\"
	

	Wscript.Sleep 3000





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
	If fill0 = 1 then ' ­n¸É0
		If LEN(Month(strdatetime)) < 2 Then MM = "0" & Month(strdatetime) Else MM = Month(strdatetime) End If
		If LEN(Day(strdatetime)) < 2 Then DD = "0" & Day(strdatetime) Else DD = Day(strdatetime) End If
		If LEN(Hour(strdatetime)) < 2 Then hh = "0" & Hour(strdatetime) Else hh = Hour(strdatetime) End If
		If LEN(Minute(strdatetime)) < 2 Then mins = "0" & Minute(strdatetime) Else mins = Minute(strdatetime) End If
		If LEN(Second(strdatetime)) < 2 Then secs = "0" & Second(strdatetime) Else secs = Second(strdatetime) End If
	ElseIf fill0 = 0 then ' ¤£¸É0
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
	
	On Error Resume Next 
	ofile.CopyFile SourceFile , Destinationfile , True
		If Err.Number <> 0 Then 
		Wscript.Echo "Err.Number is " & Err.Number
		Wscript.Echo "Err.Description is " & Err.Description
		Wscript.Sleep 5000
		Wscript.Quit 
	End If 
	Set ofile= Nothing

End Sub