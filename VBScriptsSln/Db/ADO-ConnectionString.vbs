
' MDAC ªº ODBC
'Conn.Open "Driver={SQL SERVER};server=" & Serverhost & ";uid=" & uid & ";pwd=" & pwd & ";database=" & dbName

' MDAC ªº OLD DB
'Conn.Open "Provider=SQLOLEDB; Data Source=" & ServerHost & "; Initial Catalog=" & DBName & ";Integrated Security=SSPI;"

' SQL Native Client OLE DB
'Conn.ConnectionString = "Provider=SQLNCLI;" _
'         & "Server=(local);" _
'         & "Database=META;" _ 
'         & "Integrated Security=SSPI;" _
'         & "DataTypeCompatibility=80;" _
'         & "MARS Connection=True;"
'	
'	";Uid=" & uid & _
'	";Pwd=" & pwd & ";"
'

Dim Conn, rs
Set Conn = CreateObject("ADODB.Connection")
Conn.ConnectionString = "Provider=SQLNCLI;" _
         & "Server=(local);" _
         & "Database=META;" _ 
         & "Integrated Security=SSPI;" _
         & "DataTypeCompatibility=80;" _
         & "MARS Connection=True;"
Conn.Open

Set rs = Conn.EXECUTE("SELECT * FROM myTable")

Set rs = Nothing
Set Conn = Nothing
