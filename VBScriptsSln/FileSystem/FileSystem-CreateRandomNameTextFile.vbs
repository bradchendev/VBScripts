
Dim filesys, tempname, tempfolder, tempfolderPath,tempfile
tempfolderPath = "C:\Temp"
Set filesys = CreateObject("Scripting.FileSystemObject")
Set tempfolder = filesys.GetFolder(tempfolderPath)
tempname = filesys.GetTempName
tempname = Replace(tempname,".tmp",".txt")
Set tempfile = tempfolder.CreateTextFile(tempname)

WScript.Echo tempfolderPath & "\" & tempname

'WScript.Echo "The temporary file" & tempfile & " has been created"

