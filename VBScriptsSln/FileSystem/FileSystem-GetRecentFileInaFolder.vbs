Dim targetFolder
targetFolder = "C:\Windows"
	
Class RecentFile
   Public Path
   Public DateLastModified
End Class

WScript.Echo GetRecentFileInFolder(targetFolder)

Function GetRecentFileInFolder(tarFolder)
	Set fso = CreateObject("Scripting.FileSystemObject")
	Set folder = fso.GetFolder(tarFolder)
	Set files = folder.Files
	
	Set r = New RecentFile
	 
	For each fileIdx In files
		'Wscript.Echo fileIdx.Path
		If r.Path = Empty Then
			r.Path = fileIdx.Path
			r.DateLastModified = fileIdx.DateLastModified
			'WScript.Echo fileIdx.Path
		Else
			'WScript.Echo fileIdx.Path
			If fileIdx.DateLastModified > r.DateLastModified Then
				r.Path = fileIdx.Path
				r.DateLastModified = fileIdx.DateLastModified
			End If
		End If	
	Next
	
	'Wscript.Echo r.Path
	'WScript.Echo r.DateLastModified
	
	GetRecentFileInFolder = r.Path
	
End Function



