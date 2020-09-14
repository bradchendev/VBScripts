Function CreateFolderDemo
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.CreateFolder("c:\New Folder")
   CreateFolderDemo = f.Path
End Function
