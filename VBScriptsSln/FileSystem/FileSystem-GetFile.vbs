WScript.Echo ShowFileAccessInfo("nstd_detail_p.pdf")

Function ShowFileAccessInfo(filespec)
   Dim fso, f, s
   Set fso = CreateObject("Scripting.FileSystemObject")
   Set f = fso.GetFile(filespec)
   s = f.Path & "<br>"
   s = s & "Created: " & f.DateCreated & "<br>"
   s = s & "Last Accessed: " & f.DateLastAccessed & "<br>"
   s = s & "Last Modified: " & f.DateLastModified   
   ShowFileAccessInfo = s
End Function