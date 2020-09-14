
'msgbox DateTimeConvert(NOW,"yyyy-mm-dd",1)
WScript.Echo DateTimeConvert(NOW,1,1)
WScript.Echo DateTimeConvert(NOW,2,1)
WScript.Echo DateTimeConvert(NOW,3,1)
WScript.Echo DateTimeConvert(NOW,4,1)
WScript.Echo DateTimeConvert(NOW,1,0)
WScript.Echo DateTimeConvert(NOW,2,0)
WScript.Echo DateTimeConvert(NOW,3,0)
WScript.Echo DateTimeConvert(NOW,4,0)

' Function: Convert Datetime into Formated String
' Input:
' (1)strdatetime: 丁 ┪ 丁ㄧ计
' (2)format_type: 锣Α
' format_type = 1 "20080301"
' format_type = 2 "20080301121021"
' format_type = 3 "2008-03-01"
' format_type = 4 "2008-03-01 12:10:21"
' (3)fill0: 琌干0
' fill0 = 1(璶干0)
' fill0 = 0(ぃ干0)
Function DateTimeConvert(strdatetime,format_type,fill0)
	Dim YY, MM, DD, hh, mins, secs
	
	YY = Year(strdatetime)
	If fill0 = 1 then ' 璶干0
		If LEN(Month(strdatetime)) < 2 Then MM = "0" & Month(strdatetime) Else MM = Month(strdatetime) End If
		If LEN(Day(strdatetime)) < 2 Then DD = "0" & Day(strdatetime) Else DD = Day(strdatetime) End If
		If LEN(Hour(strdatetime)) < 2 Then hh = "0" & Hour(strdatetime) Else hh = Hour(strdatetime) End If
		If LEN(Minute(strdatetime)) < 2 Then mins = "0" & Minute(strdatetime) Else mins = Minute(strdatetime) End If
		If LEN(Second(strdatetime)) < 2 Then secs = "0" & Second(strdatetime) Else secs = Second(strdatetime) End If
	ElseIf fill0 = 0 then ' ぃ干0
		MM = Month(strdatetime)
		DD = Day(strdatetime)
		hh = Hour(strdatetime)
		mins = Minute(strdatetime)
		secs = Second(strdatetime)
	End if

	Select Case format_type
		Case 1
			DateTimeConvert = YY & MM & DD
		Case 2
			DateTimeConvert = YY & MM & DD & hh & mins & secs
		Case 3
			DateTimeConvert = YY & "-" & MM & "-" & DD
		Case 4
			DateTimeConvert = YY & "-" & MM & "-" & DD & " " & hh & ":" & mins & ":" & secs
	End Select 

End Function