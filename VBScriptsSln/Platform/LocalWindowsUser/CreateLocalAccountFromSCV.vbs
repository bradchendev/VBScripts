Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fso, MyFile, FileName, TextLine, strAccountName, strPassword, strFullName
FileName = "c:\Temp\test.csv"
Set fso = CreateObject("Scripting.FileSystemObject")
Set MyFile = fso.OpenTextFile(FileName, ForReading)


	Dim objShell, cmdString
	Set objShell = CreateObject("WScript.Shell")
	

' Read from the file and display the results.
Do While MyFile.AtEndOfStream <> True
    TextLine = MyFile.ReadLine
    'Wscript.Echo TextLine
    
    strAccountName = Left(TextLine,instr(TextLine,",")-1)
    strPassword = strAccountName
    TextLine = Mid(TextLine,instr(TextLine,",")+1)
    strFullName = Left(TextLine,instr(TextLine,",")-1)
	Wscript.Echo strAccountName
	cmdString = "net user " & strAccountName & " " & strPassword & " /fullname:" & strFullName & " /Add"
	Wscript.Echo cmdString

	'objShell.Run cmdString ,,True

Loop
MyFile.Close

Set MyFile = Nothing

					
	Set objShell = Nothing