Function ReportFolderStatus(fldr)
   Dim fso, msg
   Set fso = CreateObject("Scripting.FileSystemObject")
   If (fso.FolderExists(fldr)) Then
      msg = fldr & " exists."
   Else
      msg = fldr & " doesn't exist."
   End If
   ReportFolderStatus = msg
End Function
