' Function: EnumerateObjects
' Input: organizational Unit object
'
' For Example:
' Set oCont = GetObject("LDAP://ou=test,dc=systex,dc=tw")
'
' EnumerateObjects oCont
'
' Set oCont = Nothing

Sub EnumerateObjects(oCont)
	For Each oUser In oCont
		Select Case LCase(oUser.Class)
			Case "user"
				
				'If Not IsEmpty(oUser.userPrincipalName) Then
					'strMail = oUser.userPrincipalName
				'	strCN = LEFT(oUser.userPrincipalName,Instr(oUser.userPrincipalName,"@")-1)
				'End If
				
				'If Not IsEmpty(oUser.Name) Then
				'	strName = oUser.Get("name")
				'End If
				
				'If Not IsEmpty(oUser.DistinguishedName) Then
				'	strDN = oUser.Get("DistinguishedName")
				'End If
				
			Case "contact"

			Case "organizationalunit", "container"
				EnumerateObjects oUser
		End Select
	Next
End Sub