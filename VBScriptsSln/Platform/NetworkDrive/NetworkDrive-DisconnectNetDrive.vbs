Option Explicit

' Function: disconnect network drive
' input: strDriveLetter: (For example "Z:")
' For Example:
'
' DisconnectNetDrive "R:"
'
Sub DisconnectNetDrive(strDriveLetter)
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	objNetwork.RemoveNetworkDrive strDriveLetter
	Set objNetwork = Nothing
End Sub