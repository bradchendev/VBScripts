
' Mothed 1
' Function: Enable User Dial In
' Input: DN
'
' For Example:' 
' EnableUserDialIn "cn=0670267,ou=test,dc=systex,dc=tw"
'
Sub EnableUserDialIn(strDN)
	Dim objUser
	Set objUser = GetObject("LDAP://" & strDN)
	objUser.Put "msNPAllowDialin", True
	objUser.SetInfo
	Set objUser = Nothing
End Sub


' Mothed 2
' Function: Enable User Dial In
' Input: User object
'
' For Example:
' Set objUser = GetObject("LDAP://cn=0670267,ou=test,dc=systex,dc=tw")
'
' EnableUserDialIn objUser
' Set objUser = Nothing
Sub EnableUserDialIn(objUser)

	objUser.Put "msNPAllowDialin", True
	objUser.SetInfo

End Sub