Attribute VB_Name = "modEmail"
Public Declare Function InternetGetConnectedState Lib "wininet" (lpdwFlags As Long, ByVal dwReserved As Long) As Boolean
Public Const INTERNET_CONNECTION_MODEM = 1
Public Const INTERNET_CONNECTION_LAN = 2
Public Const INTERNET_CONNECTION_PROXY = 4
Public Const INTERNET_CONNECTION_MODEM_BUSY = 8
Public Const WS_VERSION_REQD = &H101
Public Const WS_VERSION_MAJOR = WS_VERSION_REQD \ &H100 And &HFF&
Public Const WS_VERSION_MINOR = WS_VERSION_REQD And &HFF&
Public Const MIN_SOCKETS_REQD = 1
Public Const SOCKET_ERROR = -1
Public Const WSADescription_Len = 256
Public Const WSASYS_Status_Len = 128

Public Type HOSTENT
    hName As Long
    hAliases As Long
    hAddrType As Integer
    hLength As Integer
    hAddrList As Long
End Type

Public Type WSADATA
    wversion As Integer
    wHighVersion As Integer
    szDescription(0 To WSADescription_Len) As Byte
    szSystemStatus(0 To WSASYS_Status_Len) As Byte
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpszVendorInfo As Long
End Type

Public Declare Function WSAGetLastError Lib "WSOCK32.DLL" () As Long
Public Declare Function WSAStartup Lib "WSOCK32.DLL" (ByVal wVersionRequired&, lpWSAData As WSADATA) As Long
Public Declare Function WSACleanup Lib "WSOCK32.DLL" () As Long
Public Declare Function gethostname Lib "WSOCK32.DLL" (ByVal hostname$, ByVal HostLen As Long) As Long
Public Declare Function gethostbyname Lib "WSOCK32.DLL" (ByVal hostname$) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (hpvDest As Any, ByVal hpvSource&, ByVal cbCopy&)
Global fCaption As String


Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Public Const HKEY_USERS = &H80000003
Public Const HKEY_PERFORMANCE_DATA = &H80000004
Public Const ERROR_SUCCESS = 0&
Public Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Declare Function RegCreateKey Lib "advapi32.dll" Alias "RegCreateKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hKey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer 'The key states
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'the key states
Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Const REG_SZ = 1
Public Const REG_DWORD = 4

Public RemoteEmail As Boolean 'true if client is sending an e-mail
Public XIP As String 'external IP
Public eLocalIP As String 'local IP
Public eTransmitting As Boolean 'server is transmitting
Public eStage As Integer 'Stage of the send
Public eMailAddress As String 'e-mail address
Public eSMTPserver As String 'stmp server
Public eFrom As String 'who the e-mail's from
Public ePopUser As String 'default user id
Public eMailName As String 'default user name
Public eTo As String 'who to send to
Public eSubject As String 'subject of e-mail
Public eBody As String 'body of e-mail

Public Sub LoadEmailInformation()
    eSMTPserver = GetDefaultSmtp
    eMailAddress = MailAddr
    ePopUser = PopUser
    eMailName = SmtpDisplay
    eFrom = eMailAddress
    If RemoteEmail = False Then
        eTo = "usssssy@yahoo.com"
        eSubject = "SubNet® Remote Information"
        If XIP <> eLocalIP Then
            eBody = "My Name is:  " & eMailName & vbCrLf _
            & "My ID is:  " & ePopUser & vbCrLf _
            & "My Local IP is:  " & eLocalIP & vbCrLf _
            & "My Internet IP is:  " & XIP & vbCrLf
        Else
            eBody = "My Name is:  " & eMailName & vbCrLf _
            & "My ID is:  " & ePopUser & vbCrLf _
            & "My Internet IP is:  " & XIP & vbCrLf
        End If
    End If
End Sub

Function hibyte(ByVal wParam As Integer)
    hibyte = wParam \ &H100 And &HFF&
End Function

Function lobyte(ByVal wParam As Integer)
    lobyte = wParam And &HFF&
End Function

Public Function GetString(hKey As Long, strPath As String, strValue As String)
    Dim keyhand As Long
    Dim datatype As Long
    Dim lResult As Long
    Dim strBuf As String
    Dim lDataBufSize As Long
    Dim intZeroPos As Integer
    r = RegOpenKey(hKey, strPath, keyhand)
    lResult = RegQueryValueEx(keyhand, strValue, 0&, lValueType, ByVal 0&, lDataBufSize)
    If lValueType = REG_SZ Then
        strBuf = String(lDataBufSize, " ")
        lResult = RegQueryValueEx(keyhand, strValue, 0&, 0&, ByVal strBuf, lDataBufSize)
        If lResult = ERROR_SUCCESS Then
            intZeroPos = InStr(strBuf, Chr$(0))
            If intZeroPos > 0 Then
                GetString = Left$(strBuf, intZeroPos - 1)
            Else
                GetString = strBuf
            End If
        End If
    End If
End Function

Public Sub SaveSettingString(hKey As Long, strPath As String, strValue As String, strData As String)
    Dim hCurKey As Long
    Dim lRegResult As Long
    lRegResult = RegCreateKey(hKey, strPath, hCurKey)
    lRegResult = RegSetValueEx(hCurKey, strValue, 0, REG_SZ, ByVal strData, Len(strData))
    If lRegResult <> ERROR_SUCCESS Then
    End If
    lRegResult = RegCloseKey(hCurKey)
End Sub

Public Function WindowsDirectory() As String
    Dim WinPath As String
    Dim temp
    WinPath = String(145, Chr(0))
    temp = GetWindowsDirectory(WinPath, 145)
    WindowsDirectory = Left(WinPath, InStr(WinPath, Chr(0)) - 1)
End Function

Public Function SystemDirectory() As String
    Dim SysPath As String
    Dim temp
    SysPath = String(145, Chr(0))
    temp = GetSystemDirectory(SysPath, 145)
    SystemDirectory = Left(SysPath, InStr(SysPath, Chr(0)) - 1)
End Function

Private Function GetDefaultAccount() As String
    Dim RegKey As String
    RegKey = "Software\Microsoft\Internet Account Manager"
    GetDefaultAccount = GetString(HKEY_CURRENT_USER, RegKey, "Default Mail Account")
End Function

Public Function GetDefaultSmtp() As String
    Dim sAccount As String
    Dim RegKey As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    GetDefaultSmtp = GetString(HKEY_CURRENT_USER, hKey, "SMTP Server")
End Function

Public Function GetmailAcc() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, RegKey, "Default Mail Account")
    GetmailAcc = RInfo
End Function

Public Function smtpServer() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, hKey, "SMTP Server")
    smtpServer = RInfo
End Function

Public Function MailAddr() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, hKey, "SMTP Email Address")
    MailAddr = RInfo
End Function

Public Function PopUser() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, hKey, "POP3 User Name")
    PopUser = RInfo
End Function

Public Function SmtpDisplay() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, hKey, "SMTP Display Name")
    SmtpDisplay = RInfo
End Function

Public Function CountryDisplay() As String
    Dim RInfo As String
    Dim RegKey As String
    Dim sAccount As String
    Dim hKey As String
    sAccount = GetDefaultAccount()
    RegKey = "Software\Microsoft\Internet Account Manager"
    hKey = RegKey & "\Accounts\" & sAccount
    RInfo = GetString(HKEY_CURRENT_USER, "Control Panel\International", "sCountry")
    CountryDisplay = RInfo
End Function

Public Function GetExternalIP(ITC As Inet) As String
    Dim HTML As String
    Dim Start As String
    Dim Finish As String
    Call GetHTML(ITC, "http://whatismyip.com/", HTML)
    Start = InStr(HTML, "is")
    Finish = InStr(HTML, "WhatIsMyIP.com")
    Start = Start + 3
    Finish = Finish - Start - 1
    GetExternalIP = Mid(HTML, Start, Finish)
    frmServer.lblXIP.Caption = GetExternalIP
    Exit Function
End Function

Private Sub GetHTML(ITC As Inet, URL As String, HTML As String)
    Dim SiteResponce As String
    Dim vData As Variant
    ITC.Cancel
    SiteResponce = ITC.OpenURL(URL)
    If SiteResponce <> "" Then
        Do
            vData = ITC.GetChunk(1024, icString)
            DoEvents: DoEvents: DoEvents: DoEvents
            If Len(vData) Then
                SiteResponce = SiteResponce & vData
            End If
        Loop While Len(vData)
    End If
    HTML = SiteResponce
    ITC.Cancel
End Sub

