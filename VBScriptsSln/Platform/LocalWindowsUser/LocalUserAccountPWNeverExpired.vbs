strComputer = "."
strUser = "repl_snapshot"
Set User = Getobject("WinNT://" & strComputer & "/" & strUser)
Flags = User.Get("UserFlags")

User.put "Userflags", flags OR &H10000
user.setinfo
Set User = nothing
