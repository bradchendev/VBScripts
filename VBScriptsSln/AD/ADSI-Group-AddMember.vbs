' Function:�W�[�@�Ӹs�զ���
' Input: (1)objGroup:�n�M���������s�ժ���
' Output: �L
Sub AddGroupMember(objGroup, strUserDN)
	objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(strUserDN)
	'Wscript.Echo strUserDN
	objGroup.SetInfo
End Sub