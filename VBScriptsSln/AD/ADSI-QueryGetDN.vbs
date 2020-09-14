' Function:Query AD ���o�b�᪺DN��
' Input: (1)CN:���u�b��
' (2)dc_server: For example "10.1.255.200" or "dc14"
' Output: �b��DN or �䤣��h�^�ǪŦr��
Function GetDNFromADQuery(dc_server,CN)
	Dim strDN
	Dim oConnection 'As ADODB.Connection
	Dim oRecordset 'As ADODB.Recordset
	Dim strQuery 'As String
	Dim strUPN 'As String
	Dim strADsPath 'As String

	strUPN = CN & "@systex.com.tw"
	  strADsPath = "LDAP://" & dc_server & "/dc=systex,dc=tw"

	Set oConnection = CreateObject("ADODB.Connection")
	Set oRecordset = CreateObject("ADODB.Recordset")
	oConnection.Provider = "ADsDSOObject"  'The ADSI OLE-DB provider

	oConnection.Open "ADs Provider"
	strQuery = "<" & strADsPath & ">;(&(objectClass=user)(objectCategory=person)(userprincipalName=" & strUPN & "));userPrincipalName,cn,distinguishedName;subtree"
	Set oRecordset = oConnection.Execute(strQuery)

	If oRecordset.EOF And oRecordset.BOF Then
		'WScript.Echo "No duplicate UPN found"
		'WarningMessages = WarningMessages & "�bAD�䤣�� " & CN & " �b��<br>"
		'End If
		GetDNFromADQuery = ""
		Exit Function
	Else
		While Not oRecordset.EOF
			'WScript.Echo oRecordset.Fields("userPrincipalName") & " found!" & vbLf & oRecordset.Fields("cn") & " located at " & oRecordset.Fields("distinguishedName")
			strDN = oRecordset.Fields("distinguishedName")
			oRecordset.MoveNext
		Wend
	End If

	Set oRecordset = Nothing
	Set oConnection = Nothing

	GetDNFromADQuery = strDN

End Function