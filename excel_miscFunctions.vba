Option Explicit

Sub FormatShortDate()

    Dim c As Range
    Set c = Selection
    
    c.NumberFormat = "m/d/yyyy"

End Sub

Sub GetLogonScriptStuff()
    Dim objNet, domain, username
    Set objNet = CreateObject("WScript.NetWork")
    domain = objNet.UserDomain
    username = objNet.username
    
    Dim wshProcess, dc, wshShell
    Set wshShell = CreateObject("Wscript.Shell")
    Set wshProcess = wshShell.Environment("Process")
    dc = wshProcess("LogonServer")
    
    Dim UserObj, loginscript
    Set UserObj = GetObject("WinNT://" & domain & "/" & username)
    loginscript = dc & "\netlogon\" & UserObj.loginscript
    
    Debug.Print loginscript
End Sub
