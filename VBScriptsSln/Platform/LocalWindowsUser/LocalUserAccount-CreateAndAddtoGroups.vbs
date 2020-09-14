Set WshNetwork = WScript.CreateObject("WScript.Network")
strComputerName = WshNetwork.ComputerName
Set WshNetwork = Nothing

	Dim colAccounts
	Set colAccounts = GetObject("WinNT://" & strComputerName & "")
	
	CreateReplicationAccount colAccounts ,"repl_snapshot" ,"1qaz@WSX" ,"SQL Server Replication Snapshot Agent" ,"SQL Server Replication Snapshot Agent" ,strComputerName
	CreateReplicationAccount colAccounts ,"repl_logreader" ,"1qaz@WSX" ,"SQL Server Replication Logreader Agent" ,"SQL Server Replication Logreader Agent" ,strComputerName
	CreateReplicationAccount colAccounts ,"repl_distribution" ,"1qaz@WSX" ,"SQL Server Replication Distribution Agent" ,"SQL Server Replication Distribution Agent" ,strComputerName
	CreateReplicationAccount colAccounts ,"repl_merge" ,"1qaz@WSX" ,"SQL Server Replication Merge Agent" ,"SQL Server Replication Merge Agent" ,strComputerName

	Set colAccounts = Nothing
	
Sub CreateReplicationAccount( colAccounts ,strUser ,strPassword ,strFullName ,strDescription ,strComputerName)
	Dim objUser, objLocalAdmGroup

	Set objUser = colAccounts.Create("user", strUser)
	objUser.SetPassword strPassword
	objUser.FullName = strFullName
	objUser.Description = strDescription
	objUser.SetInfo
	Flags = objUser.Get("UserFlags")
	objUser.put "Userflags", flags OR &H10000
	objUser.setinfo


	Set objLocalAdmGroup = GetObject("WinNT://" & strComputerName & "/Users,group")
	objLocalAdmGroup.Add(objUser.AdsPath)
	Wscript.Echo "Create user " & "repl_snapshot" & " and Added user to " & strComputerName & "'s local Users group"

	Set objUser = Nothing

End Sub




