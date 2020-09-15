Dim t,t2, temp,temp2,milliseconds,milliseconds2,seconds,minutes,hours,strTime,strTime2
t = Timer
temp = Int(t)
milliseconds = Int((t-temp) * 1000)

Wscript.Echo "milliseconds: " & milliseconds


Wscript.Sleep 3000 

t2 = Timer
temp2 = Int(t2)
milliseconds2 = Int((t2-temp2) * 1000)


Wscript.Echo "milliseconds2:" & milliseconds2
Wscript.Echo milliseconds2 - milliseconds

Wscript.Echo "t: "&t
Wscript.Echo "t2:"&t2	  
'Wscript.Echo t2-t & "second"
Wscript.Echo cstr((t2-t)* 1000) & "ms"