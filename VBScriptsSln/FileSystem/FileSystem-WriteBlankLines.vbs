Function WriteBlankLinesToFile
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
   f.WriteBlankLines 2 
   f.WriteLine "Hello World!"
   Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   WriteBlankLinesToFile = f.ReadAll
End Function
