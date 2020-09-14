
msgbox ConvertStringToDatetimeFormat("20081025112500")

' Input String For Example: 20081025112500
' OutPut String For Example: 2008-10-25 11:25:00
Function ConvertStringToDatetimeFormat(strdatetime)
	Dim strYear, strMonth, strDate, strHour, strMin, strSec
	strYear = Left(strdatetime,4)
	strMonth = Mid(strdatetime,5,2)
	strDate = Mid(strdatetime,7,2)
	
	strHour = Mid(strdatetime,9,2)
	strMin = Mid(strdatetime,11,2)
	strSec = Mid(strdatetime,13,2)

	ConvertStringToDatetimeFormat = strYear & "-" & strMonth & "-" & strDate & " " & strHour & ":" & strMin & ":" & strSec
End Function