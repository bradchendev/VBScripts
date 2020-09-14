' Function:送出通知eMail
' Input: (1)Subject:主旨
'		(2)Sender: 寄件者
'		(2)Recipient:收件者
'		(3)MailServer:電子郵件伺服器
'		(4)htmlbody:郵件HTML內容
' Output: 無
' For Example
'
' SendMail "Test Mail", _
'			"test@systex.com.tw", _
'			"bradchen@systex.com.tw", _
'			"msmail.systex.com.tw", _
'			"<font color=red>test html message</font><br>"
'
Sub SendMail(Subject,Sender,Recipient,MailServer, htmlbody)
	Dim objMessage
	Set objMessage = CreateObject("CDO.Message")
	objMessage.Subject = Subject
	objMessage.From = Sender
	objMessage.To = Recipient

	'The line below shows how to send using HTML included directly in your script
	objMessage.HTMLBody = htmlbody
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
	'Name or IP of Remote SMTP Server
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserver") = MailServer
	'Server port (typically 25)
	objMessage.Configuration.Fields.Item _
	("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	objMessage.Configuration.Fields.Update
	objMessage.Send
End Sub
