' http://msdn.microsoft.com/en-us/library/dhyx75w2(VS.85).aspx

Const ForReading = 1, ForWriting = 2, ForAppending = 8
Const TristateUseDefault = -2, TristateTrue = -1, TristateFalse = 0
Dim fso, MyFile, FileName, TextLine

Set fso = CreateObject("Scripting.FileSystemObject")

' Open the file for output.
FileName = "c:\testfile.txt"

' 第三個參數若為true，當檔案不存在時會自動建立檔案
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


'object.OpenTextFile (filename, [ iomode, [ create, [ format ]]])
'The OpenTextFile method has these parts:

'SYNTAX
'Part	Description
'object	Required. Always the name of a FileSystemObject.
'filename	Required. String expression that identifies the file to open.
'iomode	Optional. Indicates input/output mode. Can be one of three constants: ForReading, ForWriting, or ForAppending.
'create	Optional. Boolean value that indicates whether a new file can be created if the specified filename doesn't exist. The value is True if a new file is created; False if it isn't created. The default is False.
'format	Optional. One of three Tristate values used to indicate the format of the opened file. If omitted, the file is opened as ASCII.
'Settings
'The iomode argu
ment can have any of the following settings:

'SETTINGS
'Constant	Value	Description
'ForReading	1	Open a file for reading only. You can't write to this file.
'ForWriting	2	Open a file for writing only. Use this mode to replace an existing file with new data. You can't read from this file.
'ForAppending	8	Open a file and write to the end of the file. You can't read from this file.

'The format argument can have any of the following settings:

'SETTINGS
'Constant	Value	Description
'TristateUseDefault	-2	Opens the file by using the system default.
'TristateTrue	-1	Opens the file as Unicode.
'TristateFalse	0	Opens the file as ASCII.

