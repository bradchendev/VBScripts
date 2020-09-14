' Function:增加一個群組成員
' Input: (1)objGroup:要清除成員的群組物件
' Output: 無
Sub AddGroupMember(objGroup, strUserDN)
	objGroup.PutEx ADS_PROPERTY_APPEND, "member", Array(strUserDN)
	'Wscript.Echo strUserDN
	objGroup.SetInfo
End Sub