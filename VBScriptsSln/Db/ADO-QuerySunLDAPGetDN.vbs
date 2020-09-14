' Function:Query Sun LDAP 取得帳戶的DN值
' Input: (1)CN:員工帳戶
' Output: 帳戶DN
' For Example:
' 
' QuerySunLDAPGetDN "0670267", "10.1.255.222"
'
Function QuerySunLDAPGetDN(CN, SunLDAPServer)
		Dim conn,DomainContainer,searchLDAPclass,ldapStr,strDN
		
			Set conn = CreateObject("ADODB.Connection")
			conn.Provider = "ADSDSOObject"
			conn.Open "ADs Provider", "cn=directory manager", "admin1234"
			
			'SunLDAPServer = "10.1.255.222"
			DomainContainer = "ou=000000,ou=People,o=systex.com.tw,dc=systex,dc=com,dc=tw"
			searchLDAPclass = "(objectClass=inetOrgPerson)(objectClass=inetMailUser)"
			
			ldapStr = "<LDAP://" & SunLDAPServer & "/" & DomainContainer & ">;(&" & searchLDAPclass & "(uid=" & CN & "));adspath;subtree"

			Set rs = conn.Execute(ldapStr)
			
			If rs.EOF And rs.BOF Then
				strDN = ""
				Exit Function
			Else
				While Not rs.EOF
					strDN = rs.Fields(0)
					rs.MoveNext
				Wend
			End If
		Set Conn = Nothing
		
		strDN=replace(strDN,"LDAP://" & SunLDAPServer & "/","")
		
	QuerySunLDAPGetDN = strDN

End Function