Option Explicit

' Function : Check Drive Lebal Exist
' Input: (1) strDriveLetter: Drive letter For example "Z:"
' Output: True or False
'
'If CheckDriveExist("R:") = True Then
'	WScript.Echo "R: Exist"
'Elseif CheckDriveExist("R:") = False Then
'	WScript.Echo "R: not exist"
'End If
'
Function CheckDriveExist(strDriveLetter)
	Dim objShell, objNetwork, CheckDrive, AlreadyConnected, intDrive
	Set objShell = CreateObject("WScript.Shell")
	Set objNetwork = CreateObject("WScript.Network")
	Set CheckDrive = objNetwork.EnumNetworkDrives()
	'On Error Resume Next
	AlreadyConnected = False
	For intDrive = 0 To CheckDrive.Count - 1 Step 2
		'Wscript.echo CheckDrive.Item(intDrive)
		'Wscript.echo strDriveLetter
		If LCase(CheckDrive.Item(intDrive)) = LCase(strDriveLetter) Then 
			AlreadyConnected =True
		End If
		'Wscript.echo CheckDrive.Item(intDrive)
		
	Next
	Set CheckDrive = Nothing
	Set objNetwork = Nothing
	Set objShell = Nothing
	CheckDriveExist = AlreadyConnected
End Function
