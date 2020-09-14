
' Mothed 1
' Function: Set Exchange User 傳送與接收限制KB
' Input: User object
'
' For Example:
' Set objUser = GetObject("LDAP://cn=0670267,ou=test,dc=systex,dc=tw")
'
' SetContLength objUser
' Set objUser = Nothing
Sub SetContLength(objUser)
	objUser.Put "msExchRequireAuthToSendTo", False
	objUser.Put "delivContLength", 20480
	objUser.Put "submissionContLength", 20480
	objUser.SetInfo
End Sub
