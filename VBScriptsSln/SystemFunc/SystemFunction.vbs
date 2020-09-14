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
WScript.Echo "今天的時間是" & NOW()
WScript.Echo "今天的日期是" & DATE()
WScript.Echo "今天是一週的第 " & WEEKDAY(DATE()) & " 天"
WScript.Echo "今天 " & WEEKDAYNAME(WEEKDAY(DATE()))

WScript.Echo WEEKDAY(DATE()) 
' 1 星期日
' 2 星期一
' 3 星期二
' 4 星期三
' 5 星期四
' 6 星期五
' 7 星期六





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
' 隨機產生亂數 Rnd()
For i = 0 to 100
	Wscript.echo Cstr(Round(Rnd*1000000))
next
