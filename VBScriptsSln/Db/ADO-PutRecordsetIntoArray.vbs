' Function:把recordset資料放入Array
' Input: (1)rs: recordset
' Output: recordset的資料array
Function GetDataIntoArray(rs,ColumnName)
	Dim DataArray()
	i =0
	If rs.BOF and rs.EOF Then
		reDim Preserve DataArray(1)
		DataArray(0) = "NoData"
	Else
		rs.MoveFirst
		Do Until rs.EOF
			reDim Preserve DataArray(i+1)
			DataArray(i) = rs(ColumnName)
			i = i + 1
   		rs.MoveNext
		Loop
	End If
	GetDataIntoArray = DataArray
End Function