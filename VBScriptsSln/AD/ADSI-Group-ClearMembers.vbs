' Function:清空群組的成員
' Input: (1)objGroup:要清除成員的群組物件
' Output: 無
Sub ClearGroupMembers(objGroup)
	objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
	objGroup.SetInfo
End Sub