
	Dim filesys, tempname, tempfolder, tempfolderPath,tempfile, fileName, textContent
	tempfolderPath = "C:\Temp\Temp2"
	Set filesys = CreateObject("Scripting.FileSystemObject")
	Set tempfolder = filesys.GetFolder(tempfolderPath)

' about 1KB each file
For i = 1 To 10
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname
	WScript.Echo fileName
	WriteToFile fileName, textContent
Next

' large file
For i = 1 To 10
	tempname = filesys.GetTempName
	tempname = Replace(tempname,".tmp",".txt")
	Set tempfile = tempfolder.CreateTextFile(tempname)
	Set tempfile = Nothing
	fileName = tempfolderPath & "\" & tempname
	textContent = Cstr(i) & "," & tempname & "," & Cstr(Round(Rnd()*10000000)) & "," & Cstr(Round(Rnd()*10000000))
	WScript.Echo fileName
	
	' about 16KB each file
	'WriteManylineToFile fileName, textContent, 500
	
	' about 10KB each file
	'WriteManylineToFile fileName, textContent, 300

	' about 4KB each file
	WriteManylineToFile fileName, textContent, 100

	' about 1KB each file
	'WriteManylineToFile fileName, textContent, 10
	
Next


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