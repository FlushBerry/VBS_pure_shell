

Option Explicit

' =========================================================
'  VBS PURE SHELL v1.0
'  Aucun appel à cmd.exe / powershell.exe
'  Lancement : cscript.exe //NoLogo pureshell.vbs
'          ou  wscript.exe pureshell.vbs (mode GUI InputBox)
' =========================================================

Dim g_Shell, g_FSO, g_Net, g_WMI, g_GUI
Set g_Shell = CreateObject("WScript.Shell")
Set g_FSO   = CreateObject("Scripting.FileSystemObject")
Set g_Net   = CreateObject("WScript.Network")
Set g_WMI   = GetObject("winmgmts:\\.\root\cimv2")

' Détection du host : cscript (console) ou wscript (GUI)
g_GUI = (InStr(LCase(WScript.FullName), "wscript.exe") > 0)

' =========================================================
'  COMMANDES NATIVES (COM uniquement)
' =========================================================

Function Cmd_Whoami()
    Cmd_Whoami = g_Net.UserDomain & "\" & g_Net.UserName & _
                 " @ " & g_Net.ComputerName
End Function

Function Cmd_Hostname()
    Cmd_Hostname = g_Net.ComputerName
End Function

Function Cmd_Pwd()
    Cmd_Pwd = g_Shell.CurrentDirectory
End Function

Function Cmd_Cd(path)
    On Error Resume Next
    If path = "" Then
        Cmd_Cd = g_Shell.CurrentDirectory
        Exit Function
    End If
    g_Shell.CurrentDirectory = path
    If Err.Number <> 0 Then
        Cmd_Cd = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_Cd = "OK -> " & g_Shell.CurrentDirectory
    End If
End Function

Function Cmd_Dir(path)
    Dim folder, f, out
    If path = "" Then path = g_Shell.CurrentDirectory
    On Error Resume Next
    Set folder = g_FSO.GetFolder(path)
    If Err.Number <> 0 Then
        Cmd_Dir = "[ERR] Chemin invalide: " & path
        Err.Clear
        Exit Function
    End If
    out = "Contenu de " & folder.Path & vbCrLf & String(60, "-") & vbCrLf
    For Each f In folder.SubFolders
        out = out & "<DIR>       " & f.Name & vbCrLf
    Next
    For Each f In folder.Files
        out = out & Right("            " & f.Size, 12) & "  " & f.Name & vbCrLf
    Next
    Cmd_Dir = out
End Function

Function Cmd_Type(file)
    Dim ts, out
    On Error Resume Next
    Set ts = g_FSO.OpenTextFile(file, 1)
    If Err.Number <> 0 Then
        Cmd_Type = "[ERR] " & Err.Description
        Err.Clear
        Exit Function
    End If
    out = ts.ReadAll()
    ts.Close
    Cmd_Type = out
End Function

Function Cmd_Write(file, content)
    Dim ts
    On Error Resume Next
    Set ts = g_FSO.OpenTextFile(file, 2, True) ' 2 = Write, True = Create
    If Err.Number <> 0 Then
        Cmd_Write = "[ERR] " & Err.Description
        Err.Clear
        Exit Function
    End If
    ts.Write content
    ts.Close
    Cmd_Write = "OK -> écrit " & Len(content) & " octets dans " & file
End Function

Function Cmd_Del(file)
    On Error Resume Next
    g_FSO.DeleteFile file, True
    If Err.Number <> 0 Then
        Cmd_Del = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_Del = "OK -> supprimé: " & file
    End If
End Function

Function Cmd_Copy(src, dst)
    On Error Resume Next
    g_FSO.CopyFile src, dst, True
    If Err.Number <> 0 Then
        Cmd_Copy = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_Copy = "OK -> " & src & " => " & dst
    End If
End Function

Function Cmd_Ipconfig()
    Dim col, a, out, ip
    Set col = g_WMI.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled = True")
    out = ""
    For Each a In col
        out = out & "=== " & a.Description & " ===" & vbCrLf
        out = out & "  MAC     : " & a.MACAddress & vbCrLf
        If Not IsNull(a.IPAddress) Then
            For Each ip In a.IPAddress
                out = out & "  IP      : " & ip & vbCrLf
            Next
        End If
        If Not IsNull(a.DefaultIPGateway) Then
            For Each ip In a.DefaultIPGateway
                out = out & "  Gateway : " & ip & vbCrLf
            Next
        End If
        If Not IsNull(a.DNSServerSearchOrder) Then
            For Each ip In a.DNSServerSearchOrder
                out = out & "  DNS     : " & ip & vbCrLf
            Next
        End If
        out = out & "  DHCP    : " & a.DHCPEnabled & vbCrLf & vbCrLf
    Next
    Cmd_Ipconfig = out
End Function

Function Cmd_Ps()
    Dim col, p, out
    Set col = g_WMI.ExecQuery("SELECT ProcessId, Name, ExecutablePath FROM Win32_Process")
    out = "PID       Name                          Path" & vbCrLf & String(80, "-") & vbCrLf
    For Each p In col
        out = out & Right("        " & p.ProcessId, 8) & "  " & _
              Left(p.Name & String(30, " "), 30) & "  " & _
              CStr(p.ExecutablePath) & vbCrLf
    Next
    Cmd_Ps = out
End Function

Function Cmd_Kill(pid)
    Dim col, p
    On Error Resume Next
    Set col = g_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId = " & pid)
    For Each p In col
        p.Terminate()
        Cmd_Kill = "OK -> PID " & pid & " terminé"
        Exit Function
    Next
    Cmd_Kill = "[ERR] PID introuvable: " & pid
End Function

Function Cmd_Netstat()
    Dim wmiNet, col, c, out
    On Error Resume Next
    Set wmiNet = GetObject("winmgmts:\\.\root\StandardCimv2")
    If Err.Number <> 0 Then
        Cmd_Netstat = "[ERR] StandardCimv2 indisponible: " & Err.Description
        Err.Clear
        Exit Function
    End If
    Set col = wmiNet.ExecQuery("SELECT * FROM MSFT_NetTCPConnection")
    out = "Proto  Local                       Remote                      State" & vbCrLf
    out = out & String(85, "-") & vbCrLf
    For Each c In col
        out = out & "TCP    " & _
              Left(c.LocalAddress & ":" & c.LocalPort & String(26, " "), 26) & "  " & _
              Left(c.RemoteAddress & ":" & c.RemotePort & String(26, " "), 26) & "  " & _
              c.State & vbCrLf
    Next
    Cmd_Netstat = out
End Function

Function Cmd_Services()
    Dim col, s, out
    Set col = g_WMI.ExecQuery("SELECT Name, State, StartMode FROM Win32_Service")
    out = "State       StartMode    Name" & vbCrLf & String(70, "-") & vbCrLf
    For Each s In col
        out = out & Left(s.State & String(12, " "), 12) & _
              Left(s.StartMode & String(13, " "), 13) & s.Name & vbCrLf
    Next
    Cmd_Services = out
End Function

Function Cmd_Users()
    Dim col, u, out
    Set col = g_WMI.ExecQuery("SELECT * FROM Win32_UserAccount WHERE LocalAccount = True")
    out = "Name                 Disabled  Lockout  PasswordRequired" & vbCrLf
    out = out & String(70, "-") & vbCrLf
    For Each u In col
        out = out & Left(u.Name & String(20, " "), 20) & "  " & _
              Left(CStr(u.Disabled) & String(8, " "), 8) & "  " & _
              Left(CStr(u.Lockout) & String(7, " "), 7) & "  " & _
              CStr(u.PasswordRequired) & vbCrLf
    Next
    Cmd_Users = out
End Function

Function Cmd_RegRead(keypath)
    On Error Resume Next
    Dim val : val = g_Shell.RegRead(keypath)
    If Err.Number <> 0 Then
        Cmd_RegRead = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_RegRead = keypath & " = " & val
    End If
End Function

Function Cmd_RegWrite(keypath, value, regtype)
    On Error Resume Next
    If regtype = "" Then regtype = "REG_SZ"
    g_Shell.RegWrite keypath, value, regtype
    If Err.Number <> 0 Then
        Cmd_RegWrite = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_RegWrite = "OK -> " & keypath & " = " & value
    End If
End Function

Function Cmd_Env(name)
    Dim vars, v, out
    If name = "" Then
        vars = Array("USERNAME","USERDOMAIN","COMPUTERNAME","USERPROFILE", _
                     "TEMP","APPDATA","LOCALAPPDATA","PATH","OS", _
                     "PROCESSOR_ARCHITECTURE","LOGONSERVER","HOMEPATH")
        out = ""
        For Each v In vars
            out = out & v & " = " & g_Shell.ExpandEnvironmentStrings("%" & v & "%") & vbCrLf
        Next
        Cmd_Env = out
    Else
        Cmd_Env = name & " = " & g_Shell.ExpandEnvironmentStrings("%" & name & "%")
    End If
End Function

Function Cmd_Download(url, dest)
    Dim http, stream
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    If Err.Number <> 0 Then
        Cmd_Download = "[ERR] HTTP: " & Err.Description
        Err.Clear
        Exit Function
    End If
    If http.Status <> 200 Then
        Cmd_Download = "[ERR] HTTP " & http.Status
        Exit Function
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1
    stream.Open
    stream.Write http.responseBody
    stream.SaveToFile dest, 2
    stream.Close
    Cmd_Download = "OK -> " & dest & " (" & LenB(http.responseBody) & " octets)"
End Function

Function Cmd_HttpGet(url)
    Dim http
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    If Err.Number <> 0 Then
        Cmd_HttpGet = "[ERR] " & Err.Description
        Err.Clear
        Exit Function
    End If
    Cmd_HttpGet = "HTTP " & http.Status & vbCrLf & vbCrLf & http.responseText
End Function

Function Cmd_Run(binary, args)
    ' Exec via WMI — parent devient wmiprvse.exe (évasion parent-child)
    Dim startup, process, pid, cmd
    On Error Resume Next
    Set startup = g_WMI.Get("Win32_ProcessStartup").SpawnInstance_
    startup.ShowWindow = 0
    Set process = g_WMI.Get("Win32_Process")
    
    cmd = binary
    If args <> "" Then cmd = cmd & " " & args
    
    process.Create cmd, Null, startup, pid
    If Err.Number <> 0 Then
        Cmd_Run = "[ERR] " & Err.Description
        Err.Clear
    Else
        Cmd_Run = "OK -> PID " & pid & " (via WMI)"
    End If
End Function

Function Cmd_Sleep(ms)
    If Not IsNumeric(ms) Then
        Cmd_Sleep = "[ERR] Argument numérique attendu"
        Exit Function
    End If
    WScript.Sleep CLng(ms)
    Cmd_Sleep = "OK -> sleep " & ms & "ms"
End Function

Function Cmd_Help()
    Cmd_Help = _
        "=== VBS Pure Shell — Commandes ===" & vbCrLf & _
        " [Système]" & vbCrLf & _
        "   whoami / hostname / pwd / env [var]" & vbCrLf & _
        " [Fichiers]" & vbCrLf & _
        "   dir [path] | ls        type <file> | cat" & vbCrLf & _
        "   cd <path>              write <file> <content>" & vbCrLf & _
        "   del <file>             copy <src> <dst>" & vbCrLf & _
        " [Réseau]" & vbCrLf & _
        "   ipconfig               netstat" & vbCrLf & _
        "   download <url> <dst>   httpget <url>" & vbCrLf & _
        " [Processus]" & vbCrLf & _
        "   ps / tasklist          kill <pid>" & vbCrLf & _
        "   run <bin> [args]       (via WMI)" & vbCrLf & _
        " [Système avancé]" & vbCrLf & _
        "   services               users" & vbCrLf & _
        "   regread <key>          regwrite <key> <val> [type]" & vbCrLf & _
        " [Divers]" & vbCrLf & _
        "   sleep <ms>             clear / cls" & vbCrLf & _
        "   help                   exit / quit"
End Function

' =========================================================
'  PARSER
' =========================================================

Function ParseAndExec(input)
    Dim cmd, rest, spacePos, args, a1, a2, a3
    input = Trim(input)
    If input = "" Then
        ParseAndExec = ""
        Exit Function
    End If
    
    spacePos = InStr(input, " ")
    If spacePos > 0 Then
        cmd  = LCase(Left(input, spacePos - 1))
        rest = Trim(Mid(input, spacePos + 1))
    Else
        cmd  = LCase(input)
        rest = ""
    End If
    
    ' Split en 3 args max
    args = Split(rest, " ", 3)
    a1 = "" : a2 = "" : a3 = ""
    If UBound(args) >= 0 Then a1 = args(0)
    If UBound(args) >= 1 Then a2 = args(1)
    If UBound(args) >= 2 Then a3 = args(2)
    
    Select Case cmd
        Case "whoami"              : ParseAndExec = Cmd_Whoami()
        Case "hostname"            : ParseAndExec = Cmd_Hostname()
        Case "pwd"                 : ParseAndExec = Cmd_Pwd()
        Case "cd"                  : ParseAndExec = Cmd_Cd(rest)
        Case "dir", "ls"           : ParseAndExec = Cmd_Dir(rest)
        Case "type", "cat"         : ParseAndExec = Cmd_Type(rest)
        Case "write"               : ParseAndExec = Cmd_Write(a1, Mid(rest, Len(a1) + 2))
        Case "del", "rm"           : ParseAndExec = Cmd_Del(rest)
        Case "copy", "cp"          : ParseAndExec = Cmd_Copy(a1, a2)
        Case "ipconfig", "ifconfig": ParseAndExec = Cmd_Ipconfig()
        Case "ps", "tasklist"      : ParseAndExec = Cmd_Ps()
        Case "kill"                : ParseAndExec = Cmd_Kill(rest)
        Case "netstat"             : ParseAndExec = Cmd_Netstat()
        Case "services"            : ParseAndExec = Cmd_Services()
        Case "users"               : ParseAndExec = Cmd_Users()
        Case "regread"             : ParseAndExec = Cmd_RegRead(rest)
        Case "regwrite"            : ParseAndExec = Cmd_RegWrite(a1, a2, a3)
        Case "env"                 : ParseAndExec = Cmd_Env(rest)
        Case "download"            : ParseAndExec = Cmd_Download(a1, a2)
        Case "httpget"             : ParseAndExec = Cmd_HttpGet(rest)
        Case "run"                 : ParseAndExec = Cmd_Run(a1, Mid(rest, Len(a1) + 2))
        Case "sleep"               : ParseAndExec = Cmd_Sleep(rest)
        Case "clear", "cls"        : ParseAndExec = "__CLEAR__"
        Case "help", "?"           : ParseAndExec = Cmd_Help()
        Case "exit", "quit"        : ParseAndExec = "__EXIT__"
        Case Else                  : ParseAndExec = "[ERR] Commande inconnue: " & cmd & " (tapez 'help')"
    End Select
End Function

' =========================================================
'  REPL — mode CONSOLE (cscript) ou GUI (wscript)
' =========================================================

Sub REPL_Console()
    Dim input, output
    WScript.StdOut.WriteLine "=========================================="
    WScript.StdOut.WriteLine " VBS Pure Shell — tapez 'help' ou 'exit'"
    WScript.StdOut.WriteLine "=========================================="
    
    Do
        WScript.StdOut.Write vbCrLf & "[" & g_Shell.CurrentDirectory & "]> "
        If WScript.StdIn.AtEndOfStream Then Exit Do
        input = WScript.StdIn.ReadLine()
        
        output = ParseAndExec(input)
        
        If output = "__EXIT__" Then
            WScript.StdOut.WriteLine "Bye."
            Exit Do
        ElseIf output = "__CLEAR__" Then
            ' Pas de vrai clear en cscript — on simule
            WScript.StdOut.WriteLine String(50, vbCrLf)
        ElseIf output <> "" Then
            WScript.StdOut.WriteLine output
        End If
        
        WScript.Sleep 50
    Loop
End Sub

Sub REPL_GUI()
    Dim input, output, history
    history = ""
    
    Do
        input = InputBox( _
            "Répertoire: " & g_Shell.CurrentDirectory & vbCrLf & vbCrLf & _
            "Dernière sortie:" & vbCrLf & _
            Left(history, 500) & vbCrLf & vbCrLf & _
            "Commande (help/exit):", _
            "VBS Pure Shell", "")
        
        ' Annulation = exit
        If input = "" Then Exit Do
        
        output = ParseAndExec(input)
        
        If output = "__EXIT__" Then Exit Do
        If output = "__CLEAR__" Then
            history = ""
        Else
            history = output
            ' Affichage dans MsgBox si sortie non vide
            If Len(output) > 0 Then
                MsgBox output, vbInformation, "Résultat: " & input
            End If
        End If
        
        WScript.Sleep 100
    Loop
    
    MsgBox "Fermeture du shell dans 3s...", vbInformation
    WScript.Sleep 3000
End Sub

' =========================================================
'  ENTRY POINT
' =========================================================

If g_GUI Then
    REPL_GUI()
Else
    REPL_Console()
End If

' Cleanup
Set g_Shell = Nothing
Set g_FSO   = Nothing
Set g_Net   = Nothing
Set g_WMI   = Nothing
