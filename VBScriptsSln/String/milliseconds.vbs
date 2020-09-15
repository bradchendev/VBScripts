'capture the date and timer together so if the date changes while 
'the other code runs the values you are using don't change
t = Timer
dateStr = Date()
temp = Int(t)

milliseconds = Int((t-temp) * 1000)

seconds = temp mod 60
temp    = Int(temp/60)
minutes = temp mod 60
hours   = Int(temp/60)
label = "AM"

If hours > 12 Then
    label = "PM"
    hours = hours-12
End If

'format it and add the date
strTime = LeftPad(hours, "0", 2) & ":"
strTime = strTime & LeftPad(minutes, "0", 2) & ":"
strTime = strTime & LeftPad(seconds, "0", 2) & "."
strTime = strTime & LeftPad(milliseconds, "0", 3)


WScript.Echo dateStr & " " & strTime & " " & label

'this function adds characters to a string to meet the desired length
Function LeftPad(str, addThis, howMany)
    LeftPad = String(howMany - Len(str), addThis) & str
End Function





' Method 2
Dim t, temp,milliseconds,seconds,minutes,hours,strTime,strTime2
t = Timer
temp = Int(t)
milliseconds = Int((t-temp) * 1000)
seconds = temp mod 60
temp    = Int(temp/60)	
minutes = temp mod 60
hours   = Int(temp/60)
		  
'format it and add the date
strTime = LeftPad(hours, "0", 2) & ":"
strTime = strTime & LeftPad(minutes, "0", 2) & ":"
strTime = strTime & LeftPad(seconds, "0", 2) & "."
strTime = strTime & LeftPad(milliseconds, "0", 3)

'this function adds characters to a string to meet the desired length
Function LeftPad(str, addThis, howMany)
    LeftPad = String(howMany - Len(str), addThis) & str
End Function
