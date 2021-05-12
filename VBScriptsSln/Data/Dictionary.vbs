Set d = CreateObject("Scripting.Dictionary")
d.Add "0", "Athens"   'Add some keys and items
d.Add "1", "Belgrade"
d.Add "2", "Cairo"


For Each I in d
  WScript.Echo "D.Item(" & I & ") : " & d.Item(I)
Next

'output
'D.Item(0) : Athens
'D.Item(1) : Belgrade
'D.Item(2) : Cairo
