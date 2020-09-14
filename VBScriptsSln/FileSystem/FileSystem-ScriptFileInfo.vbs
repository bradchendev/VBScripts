WScript.Echo WScript.Name
WScript.Echo WScript.FullName
WScript.Echo WScript.Path
' Windows Script Host
' C:\WINDOWS\system32\cscript.exe
' C:\WINDOWS\system32


WScript.Echo WScript.ScriptName
WScript.Echo WScript.ScriptFullName
' FileSystem-ScriptFileInfo.vbs
' C:\Users\username\source\repos\VBScripts\VBScriptsSln\FileSystem\FileSystem-ScriptFileInfo.vbs




Dim strPath, scriptBaseName
Set objFSO = CreateObject("Scripting.FileSystemObject")  
strPath = objFSO.GetParentFolderName(Wscript.ScriptFullName)
scriptBaseName = objFSO.GetBaseName(Wscript.ScriptFullName)

WScript.Echo strPath
WScript.Echo scriptBaseName
'C:\Users\1900304\source\repos\VBScripts\VBScriptsSln\FileSystem
'FileSystem-ScriptFileInfo
