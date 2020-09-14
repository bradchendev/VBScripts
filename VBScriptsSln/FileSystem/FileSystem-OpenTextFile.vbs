' http://msdn.microsoft.com/en-us/library/314cz14s(VS.85).aspx

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fso, MyFile, FileName, TextLine

Set fso = CreateObject("Scripting.FileSystemObject")

' Open the file for output.
FileName = "c:\testfile.txt"

Set MyFile = fso.OpenTextFile(FileName, ForWriting, True)

' Write to the file.
MyFile.WriteLine "Hello world!"
MyFile.WriteLine "The quick brown fox"
MyFile.Close

' Open the file for input.
Set MyFile = fso.OpenTextFile(FileName, ForReading)

' Read from the file and display the results.
Do While MyFile.AtEndOfStream <> True
    TextLine = MyFile.ReadLine
    Document.Write TextLine & "<br />"
Loop
MyFile.Close
