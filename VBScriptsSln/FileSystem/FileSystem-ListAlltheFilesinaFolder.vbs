Set fso = CreateObject("Scripting.FileSystemObject")
Set folder = fso.GetFolder("C:\Windows")
Set files = folder.Files
For each fileIdx In files
	Wscript.Echo fileIdx.Name
Next