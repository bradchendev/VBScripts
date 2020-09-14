'==========================================================================
'
' VBScript Source File -- Created with SAPIEN Technologies PrimalScript 2007
'
' NAME: 
'
' AUTHOR: Jeffery Hicks , SAPIEN Technologies
' DATE  : 11/19/2007
'
' COMMENT: 
'
'==========================================================================
On Error Resume Next
Dim strDate

centralDumpPath = "\\lab02winxp\BackupEventLogs\"
RemoteDumpPath = "C:\Temp\"

'build an array of logs to be backed up
arrLogs=Array("Application","System","Security")

'Create a FileSystemObject
Set oFS = CreateObject("Scripting.FileSystemObject")

'Open a text file of computer names ON LOCAL MACHINE WHERE SCRIPT'S RUNNING
'with one computer name per line
Set oTS = oFS.OpenTextFile("C:\servers.txt")

'Read entire file into memory

arrComputers=Split(Trim(oTS.ReadAll),VbCrLf)

'close the input file
oTS.Close

'go through array of computer names
For Each sComputer In arrComputers

  if Len(sComputer)>0 Then 'skip any blank lines
  
      For Each strLog In arrLogs
      
          strDate = CStr(Date())
          strDate = Replace(strDate, "/", "-")
          
          'dump file on each remote machine
          remoteDumpFile =  UCase(sComputer) &_
           "-" & strLog & "-" & strDate & ".evt"   
           
          rc=BackupLog(strLog,remoteDumpPath & remoteDumpFile,sComputer)
              If rc(0)=0 Then
                'successful backup and clear so copy file to central share
               strSource="\\" & sComputer & "\" & Replace(remoteDumpPath &_
                remoteDumpFile, "C:", "C$")
               strDestination=centralDumpPath & sComputer &_
                "\" & remoteDumpFile
               MoveFile strSource,strDestination
              
              Else
                 'There was an error
               Wscript.Echo "Couldn't get log " & strLog & " from " & sComputer &_
                ".  Error code: " & rc(0) & " " & rc(1)
              End If
      Next
    End If
Next

WScript.Quit 'end of main script

Function BackupLog(sLog,sFile,sComputer)
On Error Resume Next
'strLog is the name of the event log file to backup
'strFile is the name of the backup file to create. The
'filename and path must be relative to the server you
'are backing up
'strComputer is the name of the remote computer

'The function returns an array. Element 0 is an error number
'and element 1 is a description or message

Set oFS = CreateObject("Scripting.FileSystemObject")
'delete event backup if it still exists, otherwise backup 
'method will fail

If oFS.FileExists(sFile) Then oFS.DeleteFile sFile,True

 Set oWMIService = GetObject("winmgmts:" _
 & "{impersonationLevel=impersonate,(Security,Backup)}!\\" & _
 sComputer & "\root\cimv2")
 
 'query the Security logs
 Set cLogFiles = oWMIService.ExecQuery _
 ("Select * from Win32_NTEventLogFile where " & _
 "LogFileName='" & sLog & "'")
 

If cLogFiles.Count =0 Then
    BackupLog=Array("-1","Nothing to backup for event log " & sLog)
    Exit Function
End If

 'go through the collection of logs

 For Each oLogfile in cLogFiles

  'back up the log to a file LOCAL ON THAT REMOTE MACHINE
   WScript.Echo "Creating " & sFile

    oLogFile.BackupEventLog(sFile)
   
  If Err.number=0 Then
    BackupLog=Array(0,"Successfully backed up " & sLog & " to " & sFile)
 
   'no error - safe to clear the Log
    WScript.Echo "Clearing event log " & strLog & " on " & sComputer
    
    'Uncomment the next line to actually clear the log. I have it 
    'commented out for test purposes
   'oLogFile.ClearEventLog()
  Else
  
    BackupLog=Array(Err.Number,Err.Description)
   
  End If
Next

End Function

Function MoveFile(strSource,strDestination)
On Error Resume Next

Set oFS = CreateObject("Scripting.FileSystemObject")

strParentFolder = oFS.GetParentFolderName(strDestination)

   If oFS.FolderExists(strParentFolder)=False Then
    WScript.Echo "Creating " & strParentFolder
    oFS.CreateFolder strParentFolder
        If Err.Number<>0 Then
            WScript.Echo "Failed to create " & strParentFolder
            Exit Function
        End If
   End If
   
     WScript.Echo "Copying " & strSource & " to " & strDestination
    'Any existing files with same name will be overwritten
    oFS.CopyFile strSource,strDestination,True
    'if successful copy then delete source file
    If Err.Number=0 Then oFS.DeleteFile strSource

End Function