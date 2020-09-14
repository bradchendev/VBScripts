
' Mothed 1
' Function: Enable AD User
' Input: DN
'
' For Example:' 
' EnableUSer "cn=0670267,ou=test,dc=systex,dc=tw"
'
Sub EnableUser(strDN)
	Const ADS_UF_ACCOUNTDISABLE = 2
	Dim objUser
	Set objUser = GetObject("LDAP://" & strDN)
	intUAC = objUser.Get("userAccountControl")
	If intUAC AND ADS_UF_ACCOUNTDISABLE Then
		objUser.Put "userAccountControl", intUAC XOR ADS_UF_ACCOUNTDISABLE
		objUser.SetInfo
	End If
	Set objUser = Nothing
End Sub



' Mothed 2
' Function: Enable AD User
' Input: User object
'
' For Example:
' Set objUser = GetObject("LDAP://cn=0670267,ou=test,dc=systex,dc=tw")
'
' EnableUSer objUser
' Set objUser = Nothing
Sub EnableUser(objUser)
	Const ADS_UF_ACCOUNTDISABLE = 2
	intUAC = objUser.Get("userAccountControl")
	If intUAC AND ADS_UF_ACCOUNTDISABLE Then
		objUser.Put "userAccountControl", intUAC XOR ADS_UF_ACCOUNTDISABLE
		objUser.SetInfo
	End If
End Sub
