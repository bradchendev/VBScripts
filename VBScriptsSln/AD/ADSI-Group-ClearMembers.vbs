' Function:�M�Ÿs�ժ�����
' Input: (1)objGroup:�n�M���������s�ժ���
' Output: �L
Sub ClearGroupMembers(objGroup)
	objGroup.PutEx ADS_PROPERTY_CLEAR, "member", 0
	objGroup.SetInfo
End Sub