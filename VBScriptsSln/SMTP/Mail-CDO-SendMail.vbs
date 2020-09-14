' Function:�e�X�q��eMail
' Input: (1)Subject:�D��
'		(2)Sender: �H���
'		(2)Recipient:�����
'		(3)MailServer:�q�l�l����A��
'		(4)htmlbody:�l��HTML���e
' Output: �L
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
