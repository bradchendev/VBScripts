'----------------------------------------------------------------------------------------------------------------------------  
'Script Name : EnumerateURLStatus.vbs     
'Author      : Matthew Beattie     
'Created     : 16/02/09   
'Description : This script enumerates the status code of a URL.  
'----------------------------------------------------------------------------------------------------------------------------  
'Initialization  Section     
'----------------------------------------------------------------------------------------------------------------------------  
Option Explicit  
Dim objFSO, scriptBaseName  
'----------------------------------------------------------------------------------------------------------------------------  
'Main Processing Section  
'----------------------------------------------------------------------------------------------------------------------------  
On Error Resume Next 
   Set objFSO     = CreateObject("Scripting.FileSystemObject")  
   scriptBaseName = objFSO.GetBaseName(Wscript.ScriptFullName)  
   ProcessScript  
   'Wscript.Echo "a test message line"
   Wscript.Echo "Err.Number is " & Err.Number
   
   If Err.Number <> 0 Then 
      Wscript.Quit  
   End If 
On Error Goto 0  
'----------------------------------------------------------------------------------------------------------------------------  
'Name       : ProcessScript -> Primary Function that controls all other script processing.     
'Parameters : None          ->      
'Return     : None          ->      
'----------------------------------------------------------------------------------------------------------------------------  
Function ProcessScript  
   Dim url, urlStatus  
   url = "https://www.google.com" 
   If Not EnumerateURLStatus(url, urlStatus) Then 
      Exit Function 
   End If 

   Select Case urlStatus   
      Case "404" 
         Wscript.Echo "The status of the URL " & url & " is " & urlStatus, vbCritical, scriptBaseName  
      Case Else 
         Wscript.Echo "The status of the URL " & url & " is " & urlStatus, vbInformation, scriptBaseName  
   End Select 
End Function 
'----------------------------------------------------------------------------------------------------------------------------  
'Name       : EnumerateURLStatus -> Enumerates the status code of a URL.    
'Parameters : url                -> Input/Output : String containing the URL of the web page to enumerate.  
'           : urlStatus          -> Input/Output : Integer containing the url status code number.  
'Return     : EnumerateURLStatus -> Returns True and the status code of the URL otherwise returns False.  
'----------------------------------------------------------------------------------------------------------------------------  
Function EnumerateURLStatus(url, urlStatus)  
   Dim objXML  
   EnumerateURLStatus = False 
   On Error Resume Next 
      Set objXML = CreateObject("MSXML2.XMLHTTP.3.0")  
      If Err.Number <> 0 Then 
         Exit Function 
      End If 
      objXML.open "GET", url, False 
      objXML.send  
      urlStatus = CInt(objXML.Status)  
      If Err.Number <> 0 Then 
         Exit Function 
      End If 
   On Error Goto 0  
   EnumerateURLStatus = True 
End Function 