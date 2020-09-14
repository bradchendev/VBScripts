

' Function: Create AD User Mailbox
' input:
' (1)strDN: User DN
' (2)strmail: User Mail address
' (3)homeMDB: Exchange Mail Database
' For Example:
'
' homeMDB = "LDAP://CN=�H�c�x�s�� (BE),CN=�w�]�x�s�s��,CN=InformationStore,CN=BE,CN=Servers,CN=�w�]�t�κ޲z�s��,CN=Administrative Groups,CN=SYSTEX,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=systex,DC=tw"
' CreateMailBox "cn=0670267,ou=test,dc=systex,dc=tw", _
'				"test@systex.com.tw", _
'				homeMDB
'
Sub CreateMailBox(strDN, strmail, homeMDB)
	Dim objUser,strNickName,strMailDomain
	
	Set objUser = GetObject("LDAP://" & strDN)
	objUser.CreateMailbox homeMDB
	objUser.SetInfo
	
	' ���wmail address
	strNickName = Left(strmail,Instr(strmail,"@")-1)
	strMailDomain = Mid(strmail,Instr(strmail,"@")+1)
	objUser.MailNickName = strNickName
	objUser.Mail = strmail
	objUser.SetInfo

	' ���w�D�nSMTP: Proxy Addresses
	objUser.PutEx ADS_PROPERTY_APPEND, "proxyAddresses", Array("SMTP:" & strmail)
	objUser.SetInfo
	' ���w���n��SMTP: Proxy Addresses
	objUser.PutEx ADS_PROPERTY_APPEND, "proxyAddresses", Array("smtp:" & strNickName & "@systex.tw")
	objUser.SetInfo
	
	Set objUser = Nothing

End Sub