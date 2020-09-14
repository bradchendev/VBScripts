Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputerName = WshNetwork.ComputerName
Set WshNetwork = Nothing

Set objComputer = GetObject("WinNT://" & strComputerName )
objComputer.Filter = Array("User")
    WScript.Echo "_______________________________________"
    WScript.Echo "Computer: " & strComputerName
    WScript.Echo "_______________________________________" 
For Each objUser in objComputer
    Wscript.Echo objUser.Name & "(" & objUser.FullName & ")" 
Next

