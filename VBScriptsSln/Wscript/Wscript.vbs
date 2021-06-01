

Wscript.Echo "Script name:" & Wscript.ScriptName
Wscript.Echo "Script path:" & Wscript.ScriptFullName

'Script name: test.vbs
'Script path: C:\scripts\test.vbs


Wscript.Echo "Build number: "& _
    ScriptEngineMajorVersion & ¡§.¡¨ & ScriptEngineMinorVersion & ¡§.¡¨ & ScriptEngineBuildVersion

'Build number: 5.6.8820
Wscript.Echo "Version: " & Wscript.Version

strScriptHost = LCase(Wscript.FullName)

If Right(strScriptHost, 11) = "wscript.exe" Then
    Wscript.Echo "This script is running under WScript."
Else
    Wscript.Echo "This script is running under CScript."
End If

-- https://devblogs.microsoft.com/scripting/how-can-i-determine-the-name-of-a-script-while-that-script-is-running/
