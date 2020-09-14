
' Support Policies for SQL Server Native Client
' https://docs.microsoft.com/en-us/sql/relational-databases/native-client/applications/support-policies-for-sql-server-native-client?view=sql-server-ver15
' SQL Server Native Client 11.0 supports connections to, SQL Server 2008, SQL Server 2008 R2, SQL Server 2012 (11.x), SQL Server 2014 (12.x), and Azure SQL Database.



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


'Conn.ConnectionString = "Provider=SQLNCLI;" _
'         & "Server=(local);" _
'         & "Database=META;" _ 
'         & "Integrated Security=SSPI;" _
'         & "DataTypeCompatibility=80;" _
'         & "MARS Connection=True;"


' Drivers Tab on ODBC Data Source Administrator (32-bit)
' SQL Server Native Client 10.0
'Conn.ConnectionString = "Provider=SQLNCLI10;" _
'         & "Server=(local);" _
'         & "Database=META;" _ 
'         & "Integrated Security=SSPI;" _
'         & "DataTypeCompatibility=80;" _
'         & "MARS Connection=True;"


' Drivers Tab on ODBC Data Source Administrator (32-bit)
' SQL Server Native Client 11.0
'Conn.ConnectionString = "Provider=SQLNCLI11;" _
'         & "Server=(local);" _
'         & "Database=META;" _ 
'         & "Integrated Security=SSPI;" _
'         & "DataTypeCompatibility=80;" _
'         & "MARS Connection=True;"


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
