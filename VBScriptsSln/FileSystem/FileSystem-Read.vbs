' http://msdn.microsoft.com/en-us/library/dhyx75w2(VS.85).aspx

Function ReadTextFileTest
   Const ForReading = 1, ForWriting = 2, ForAppending = 8
   Dim fso, f, Msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
   f.Write "Hello world!"
   Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   ReadTextFileTest = f.Read(5)
End Function
