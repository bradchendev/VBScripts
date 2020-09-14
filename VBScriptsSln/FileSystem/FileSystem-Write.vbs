' http://msdn.microsoft.com/en-us/library/6ee7s9w2(VS.85).aspx

Function WriteToFile
   Const ForReading = 1, ForWriting = 2
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.OpenTextFile("c:\testfile.txt", ForWriting, True)
   f.Write "Hello world!" 
   Set f = fso.OpenTextFile("c:\testfile.txt", ForReading)
   WriteToFile =   f.ReadLine
End Function
