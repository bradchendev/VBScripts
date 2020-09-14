Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("C:\Windows")
Set Subfolder = folder.SubFolders
For each SubFolderIdx In Subfolder
	Wscript.Echo SubFolderIdx.Name
Next