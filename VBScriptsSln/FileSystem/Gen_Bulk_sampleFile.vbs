strPath = "E:\STORE"

For k = 1 To 400
	If Len(Cstr(k)) = 1 Then
		strFolder = strPath & "\McdStore00" & Cstr(k)
	ElseIf Len(Cstr(k)) = 2 Then
		strFolder = strPath & "\McdStore0" & Cstr(k)
	ElseIf Len(Cstr(k)) = 3 Then
		strFolder = strPath & "\McdStore" & Cstr(k)
	End If
	Wscript.Echo strFolder
	CreateFolderDemo(strFolder)
	GenerateFile strFolder
Next




Sub GenerateFile(strFolder)
	Dim filesys, tempname, tempfolder, tempfolderPath,tempfile, fileName, textContent
	tempfolderPath = strFolder
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set tempfolder = filesys.GetFolder(tempfolderPath)

' about 1KB each file
For i = 1 To 150
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname
	WriteToFile fileName, textContent
	WScript.Echo "1K  " & fileName
Next

' large file
For i = 1 To 100
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname & "," & Cstr(Round(Rnd()*10000000)) & "," & Cstr(Round(Rnd()*10000000))

	' about 4KB each file
	WriteManylineToFile fileName, textContent, 1600
	WScript.Echo "4K  " & fileName
	
Next

' large file
For i = 1 To 80
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname & "," & Cstr(Round(Rnd()*10000000)) & "," & Cstr(Round(Rnd()*10000000))

	
	' about 16KB each file
	WriteManylineToFile fileName, textContent, 2200
	WScript.Echo "16K " & fileName
	
Next

' large file
For i = 1 To 75
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname & "," & Cstr(Round(Rnd()*10000000)) & "," & Cstr(Round(Rnd()*10000000))
	
	' about 16KB each file
	WriteManylineToFile fileName, textContent, 2600
	WScript.Echo "xx K " & fileName

Next

' large file
For i = 1 To 75
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname & "," & Cstr(Round(Rnd()*10000000)) & "," & Cstr(Round(Rnd()*10000000))

	
	' about xxKB each file
	WriteManylineToFile fileName, textContent, 3000
	WScript.Echo "xx K " & fileName

Next



End Sub



Function CreateFolderDemo(strFolder)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.CreateFolder(strFolder)
   CreateFolderDemo = f.Path
End Function


Function WriteToFile(TextFileWithPath, WriteContent)
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(TextFileWithPath, ForWriting, True)
   f.WriteLine WriteContent 
   'Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   Set f = Nothing
   Set fso = Nothing
   'WriteToFile =   f.ReadLine
End Function


Function WriteManylineToFile(TextFileWithPath, WriteContent, n )
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile(TextFileWithPath, ForWriting, True)
	   For j = 1 to n
	   f.WriteLine WriteContent 
	   Next
   'Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   Set f = Nothing
   Set fso = Nothing
   'WriteToFile =   f.ReadLine
End Function