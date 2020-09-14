' http://msdn.microsoft.com/en-us/library/t5399c99(VS.85).aspx

Function WriteLineToFile
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
   f.WriteLine "Hello world!" 
   f.WriteLine "VBScript is fun!"
   Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   WriteLineToFile = f.ReadAll
End Function
