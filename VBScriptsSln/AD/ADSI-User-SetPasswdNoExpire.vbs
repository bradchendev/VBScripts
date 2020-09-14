
' Mothed 1
' Function: Enable Password No Expire
' Input: DN
'
' For Example:' 
' EnablePasswdNoExpire "cn=0670267,ou=test,dc=systex,dc=tw"
'
Sub EnablePasswdNoExpire(strDN)
	Const UF_DONT_EXPIRE_PASSWD = &H10000
	Dim objUser
	Set objUser = GetObject("LDAP://" & strDN)
	intUAC = objUser.Get("userAccountControl")

	If intUAC AND Not UF_DONT_EXPIRE_PASSWD Then
		objUser.Put "userAccountControl", intUAC OR UF_DONT_EXPIRE_PASSWD
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
' EnablePasswdNoExpire objUser
' Set objUser = Nothing
Sub EnablePasswdNoExpire(objUser)
	Const UF_DONT_EXPIRE_PASSWD = &H10000
	intUAC = objUser.Get("userAccountControl")

	If intUAC AND Not UF_DONT_EXPIRE_PASSWD Then
		objUser.Put "userAccountControl", intUAC OR UF_DONT_EXPIRE_PASSWD
		objUser.SetInfo
	End If

End Sub
