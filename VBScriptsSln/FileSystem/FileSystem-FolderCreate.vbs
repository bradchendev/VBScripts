Dim folder
folder = "C:\test"


If Not(FolderExist(folder)) Then
	WScript.Echo "need to create the folder"
	folder = CreateFolder(folder)
End If

WScript.Echo folder


Function FolderExist(fldr)
   Dim fso
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FolderExists(fldr)) Then
      FolderExist = True
   Else
      FolderExist = False
   End If
End Function

Function CreateFolder(folder)
   Dim fso, f
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.CreateFolder(folder)
   CreateFolder = f.Path
End Function
