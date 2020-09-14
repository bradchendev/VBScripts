

' Function: Create AD User Mailbox
' input:
' (1)strDN: User DN
' (2)strmail: User Mail address
' (3)homeMDB: Exchange Mail Database
' For Example:
'
' homeMDB = "LDAP://CN=信箱儲存區 (BE),CN=預設儲存群組,CN=InformationStore,CN=BE,CN=Servers,CN=預設系統管理群組,CN=Administrative Groups,CN=SYSTEX,CN=Microsoft Exchange,CN=Services,CN=Configuration,DC=systex,DC=tw"
' CreateMailBox "cn=0670267,ou=test,dc=systex,dc=tw", _
'				"test@systex.com.tw", _
'				homeMDB
'
Sub CreateMailBox(strDN, strmail, homeMDB)
	Dim objUser,strNickName,strMailDomain
	
	Set objUser = GetObject("LDAP://" & strDN)
	objUser.CreateMailbox homeMDB
	objUser.SetInfo
	
	' 指定mail address
	strNickName = Left(strmail,Instr(strmail,"@")-1)
	strMailDomain = Mid(strmail,Instr(strmail,"@")+1)
	objUser.MailNickName = strNickName
	objUser.Mail = strmail
	objUser.SetInfo

	' 指定主要SMTP: Proxy Addresses
	objUser.PutEx ADS_PROPERTY_APPEND, "proxyAddresses", Array("SMTP:" & strmail)
	objUser.SetInfo
	' 指定次要的SMTP: Proxy Addresses
	objUser.PutEx ADS_PROPERTY_APPEND, "proxyAddresses", Array("smtp:" & strNickName & "@systex.tw")
	objUser.SetInfo
	
	Set objUser = Nothing

End Sub