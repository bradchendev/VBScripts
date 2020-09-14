Option Explicit

' Fucntion: Mapping Network Drive
' Input: 
' (1)strDriveLetter: Drive letter (For example "Z:")
' (2)strRemotePath:(For example "\\10.1.255.208\META\backup\MetaDB")
' (3)strProfile: optional (keep empty)
' (4)strUser:(file server login account)
' (5)strPassword:(login password)
'
' Output: none
'
' For Example
'MappingNetDrive "R:", _
'					"\\10.1.255.208\test", _
'					, _
'					"useraccount", _
'					"password"
'
Sub MappingNetDrive(strDriveLetter, strRemotePath, _
							strProfile, strUser, strPassword)
	Dim objNetwork
	Set objNetwork = CreateObject("WScript.Network")
	objNetwork.MapNetworkDrive strDriveLetter, strRemotePath, _
	strProfile, strUser, strPassword
	Set objNetwork = nothing
End Sub


