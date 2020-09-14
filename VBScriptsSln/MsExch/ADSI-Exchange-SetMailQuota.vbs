'Global variables
Const ADS_PROPERTY_CLEAR = 1 
Const ADS_PROPERTY_UPDATE = 2 
Const ADS_PROPERTY_APPEND = 3
Const ADS_PROPERTY_DELETE = 4

Const INT100M_Warn  = 81920
Const INT100M_StopS = 102400
Const INT100M_StopR = 122880




' Mothed 1
' Function: Set Mail Quota 
' Input: User object
'
' For Example:
' Set objUser = GetObject("LDAP://cn=0670267,ou=test,dc=systex,dc=tw")
'
' SetMailQuota objUser
' Set objUser = Nothing
' 
Sub SetMailQuota (objUser)
	objUser.Put "mDBUseDefaults", False
	objUser.Put "mDBStorageQuota", INT100M_Warn
	objUser.Put "mDBOverQuotaLimit", INT100M_StopS
	objUser.Put "mDBOverHardQuotaLimit", INT100M_StopR
	objUser.SetInfo
End Sub

