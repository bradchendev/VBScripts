


' Mothed 1
' Function: 設定轉遞(Set AltRecipient 屬性)
' Input: User object
'
' For Example:
' Set objUser = GetObject("LDAP://cn=0670267,ou=test,dc=systex,dc=tw")
'
' SetAltRecipient objUser, "cn=0670267-1,ou=test,dc=systex,dc=tw"
' Set objUser = Nothing
Sub SetAltRecipient(objUser,RecipientUserDN)
	objUser.Put "altRecipient", RecipientUserDN
	objUser.SetInfo
End Sub