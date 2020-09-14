' http://msdn.microsoft.com/en-us/library/dhyx75w2(VS.85).aspx

Function ReadAllTextFile
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
   f.Write "Hello world!"
   Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   ReadAllTextFile =   f.ReadAll
End Function
