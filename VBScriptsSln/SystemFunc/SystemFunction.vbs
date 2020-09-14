' ===== Data type ==================================
Dim strMyName
strMyName = "Brad Chen"
WScript.Echo strMyName
WScript.Echo TypeName(strMyName)

strMyName = 2008
WScript.Echo strMyName
WScript.Echo TypeName(strMyName)

strMyName = 20081234567890
WScript.Echo strMyName
WScript.Echo TypeName(strMyName)

strMyName = #2008-12-09#
WScript.Echo strMyName
WScript.Echo TypeName(strMyName)



' ===== Date Time ==================================
WScript.Echo "���Ѫ��ɶ��O" & NOW()
WScript.Echo "���Ѫ�����O" & DATE()
WScript.Echo "���ѬO�@�g���� " & WEEKDAY(DATE()) & " ��"
WScript.Echo "���� " & WEEKDAYNAME(WEEKDAY(DATE()))

WScript.Echo WEEKDAY(DATE()) 
' 1 �P����
' 2 �P���@
' 3 �P���G
' 4 �P���T
' 5 �P���|
' 6 �P����
' 7 �P����





WScript.Echo DateAdd("m",1,"31-Jan-01")
' Output is 1931/2/1

'*  yyyy - Year
'* q - Quarter
'* m - Month
'* y - Day of year
'* d - Day
'* w - Weekday
'* ww - Week of year
'* h - Hour
'* n - Minute
'* s - Second


WScript.Echo Dateadd("d",1,Date())
WScript.Echo Dateadd("d",-1,Date())

WScript.Echo Dateadd("m",1,Date())
WScript.Echo Dateadd("m",-1,Date())

WScript.Echo Dateadd("yyyy",1,Date())
WScript.Echo Dateadd("yyyy",-1,Date())


'
' �H�����Ͷü� Rnd()
For i = 0 to 100
	Wscript.echo Cstr(Round(Rnd*1000000))
next
