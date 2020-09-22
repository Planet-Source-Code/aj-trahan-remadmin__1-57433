VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmChat 
   BorderStyle     =   0  'None
   Caption         =   "Subnet Chat"
   ClientHeight    =   5610
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6960
   LinkTopic       =   "Form1"
   Picture         =   "frmChat.frx":0000
   ScaleHeight     =   5610
   ScaleWidth      =   6960
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtServer 
      Height          =   375
      Left            =   1560
      TabIndex        =   7
      Text            =   "127.0.0.1"
      Top             =   5760
      Width           =   1335
   End
   Begin VB.ListBox ConnectionList 
      Height          =   1035
      ItemData        =   "frmChat.frx":7F752
      Left            =   3000
      List            =   "frmChat.frx":7F754
      TabIndex        =   6
      Top             =   5760
      Width           =   3825
   End
   Begin VB.CommandButton cmdTalk 
      Caption         =   "Talk"
      Height          =   375
      Left            =   600
      TabIndex        =   5
      Top             =   5760
      Width           =   855
   End
   Begin MSWinsockLib.Winsock sockVoice 
      Index           =   0
      Left            =   120
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6802
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00404040&
      Height          =   200
      Left            =   1560
      TabIndex        =   3
      Top             =   5100
      Width           =   200
   End
   Begin VB.TextBox txtChat 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
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
      Height          =   3875
      Left            =   210
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   415
      Width           =   6300
   End
   Begin VB.TextBox txtSend 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
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
      Height          =   285
      Left            =   240
      TabIndex        =   0
      Top             =   4535
      Width           =   6255
   End
   Begin VB.Label lblTyping 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(Slave) Is Typing A Message."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0C0&
      Height          =   210
      Left            =   1800
      TabIndex        =   4
      Top             =   4300
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Enable Voice"
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
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   5100
      Width           =   1335
   End
   Begin VB.Image imgChatSel 
      Height          =   390
      Index           =   1
      Left            =   5280
      Picture         =   "frmChat.frx":7F756
      Top             =   5040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgChatSel 
      Height          =   390
      Index           =   0
      Left            =   240
      Picture         =   "frmChat.frx":80F90
      Top             =   5040
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgChat 
      Height          =   390
      Index           =   1
      Left            =   5280
      Picture         =   "frmChat.frx":83B0C
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Image imgChat 
      Height          =   390
      Index           =   0
      Left            =   240
      Picture         =   "frmChat.frx":85346
      Top             =   5040
      Width           =   1170
   End
End
Attribute VB_Name = "frmChat"
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
Public CLOSINGAPPLICATION As Boolean
Public wStream As Object

Private Sub Check1_Click()
Dim mmsg As String
If Check1.Value = 1 Then
    mmsg = "|CHAT|(**** " & User & " has enabled Voice ****)"
    mmsg = Encrypt(mmsg)
    If frmMain.sockMain.State = sckConnected Then
        frmMain.sockChat.SendData mmsg
        txtChat.Text = txtChat.Text & "(**** " & User & " has enabled Voice ****)" & vbCrLf
        frmMain.lblInfo(8).Caption = "Connected"
        Enable_Voice
        
    End If
Else
    mmsg = "|CHAT|(**** " & User & " has disabled Voice ****)"
    mmsg = Encrypt(mmsg)
    If frmMain.sockMain.State = sckConnected Then
        frmMain.sockChat.SendData mmsg
        txtChat.Text = txtChat.Text & "(**** " & User & " has disabled Voice ****)" & vbCrLf
        frmMain.lblInfo(8).Caption = "Pending.."
        Disable_Voice
    End If
End If

End Sub
Private Sub Disable_Voice()
        ConnectionList_DblClick
        sockVoice_Close 0
End Sub
Private Sub Enable_Voice()
    Dim rc As Long
    Dim Idx As Long
    Dim LocalPort As Long
    Dim RemotePort As Long
        Idx = InstanceTCP(sockVoice)
        
        If (Idx > 0) Then
            ConnectionList.Enabled = False
            
            On Error Resume Next
            If Not Connect(sockVoice(Idx), txtServer.Text, VOICEPORT) Then
                Unload sockVoice(Idx)
            End If
            
            ConnectionList.Enabled = True
        End If

End Sub

Private Sub cmdTalk_Click()
    Dim rc As Long
    Dim iPort As Integer
    Dim itm As Integer
    If (Not wStream.Playing And wStream.PlayDeviceFree And _
        Not wStream.Recording And wStream.RecDeviceFree) Then
        wStream.Playing = True
        cmdTalk.Caption = "&Playing"
        Screen.MousePointer = vbHourglass
        
        iPort = wStream.StreamInQueue
        Do While (iPort <> NULLPORTID)
            For itm = 0 To ConnectionList.ListCount - 1
                If (ConnectionList.ItemData(itm) = iPort) Then
                    ConnectionList.TopIndex = itm
                    ConnectionList.Selected(itm) = True
                    Exit For
                End If
            Next
            
            rc = wStream.PlayWave(Me.hwnd, iPort)
            Call wStream.RemoveStreamFromQueue(iPort)
            iPort = wStream.StreamInQueue
        Loop
        
        ConnectionList.TopIndex = 0
        If (ConnectionList.ListCount > 0) Then
            ConnectionList.Selected(0) = True
            ConnectionList.Selected(0) = False
        End If
        Screen.MousePointer = vbDefault
        cmdTalk.Caption = "&Talk"
        wStream.Playing = False
    End If
End Sub

Private Sub ConnectionList_DblClick()
    Dim MemberID As String
    Dim Idx As Long
    If (ConnectionList.Text = "") Then Exit Sub
    MemberID = ConnectionList.List(ConnectionList.ListIndex)
    
    Call GetIdxFromMemberID(sockVoice, MemberID, Idx)
    Call RemoveConnectionFromList(sockVoice(Idx), ConnectionList)
    Call Disconnect(sockVoice(Idx))
    Unload sockVoice(Idx)

    cmdTalk.Enabled = (ConnectionList.ListCount > 0)
End Sub

Private Sub Form_Load()
    Dim rc As Long
    Dim Idx As Long
    Dim TCPidx As Long
    CLOSINGAPPLICATION = False
    txtServer.Text = RemHst
    
    Set wStream = CreateObject("WaveStreaming.WaveStream")
    Call wStream.InitACMCodec(WAVE_FORMAT_GSM610, TIMESLICE)

    cmdTalk.Enabled = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.lblInfo(8).Caption = "Pending.."
    Dim Idx As Long
    Dim Socket As Winsock
    CLOSINGAPPLICATION = True
    For Each Socket In sockVoice
        Call Disconnect(Socket)
    Next
    Set wStream = Nothing
End Sub

Private Sub imgChat_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgChatSel(Index).Visible = True
If Index = 0 Then
    If Check1.Value = 0 Then
        Exit Sub
    Else
    Dim rc As Long
    If (Not wStream.Playing And _
        Not wStream.Recording And _
            wStream.RecDeviceFree And _
            wStream.PlayDeviceFree) Then
        wStream.Recording = True
        cmdTalk.Caption = "&Talking"
        Screen.MousePointer = vbHourglass
        rc = wStream.RecordWave(Me.hwnd, sockVoice)
        Screen.MousePointer = vbDefault
        cmdTalk.Caption = "&Talk"
        
        If Not wStream.Playing And _
               wStream.PlayDeviceFree And _
               wStream.RecDeviceFree Then
            Call cmdTalk_Click
        End If
    End If
    End If
End If

End Sub

Private Sub imgChat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgChatSel(Index).Visible = False
imgChat(Index).Visible = True
If Index = 0 Then
    wStream.Recording = False
    Exit Sub
End If
If Index = 1 Then
    frmMain.sockChat.Close
    frmMain.sockMain.SendData Encrypt("|CLOSECHAT|")
    frmMain.lblStatus.Caption = " Status:  Closing Chat Client"
    frmMain.lblInfo(9).Caption = "Pending.."
    frmMain.tmrMessage.Enabled = True
    Unload Me
End If
End Sub

Private Sub sockVoice_Close(Index As Integer)
    Call RemoveConnectionFromList(sockVoice(Index), ConnectionList)
    Call Disconnect(sockVoice(Index))
    cmdTalk.Enabled = (ConnectionList.ListCount > 0)
    frmMain.lblInfo(8).Caption = "Pending.."
End Sub

Private Sub sockVoice_Connect(Index As Integer)
    Call AddConnectionToList(sockVoice(Index), ConnectionList)
    'Call ResPlaySound(RingOutId)
    cmdTalk.Enabled = True
    frmMain.lblInfo(8).Caption = "Connected"
End Sub

Private Sub sockVoice_ConnectionRequest(Index As Integer, ByVal requestID As Long)

    Dim rc As Long
    Dim Idx As Long
    Dim RmHost As String
    If (sockVoice(Index).RemoteHost <> "") Then
        RmHost = UCase(sockVoice(Index).RemoteHost)
    Else
        RmHost = UCase(sockVoice(Index).RemoteHostIP)
    End If
    rc = vbYes
                    
    If (rc = vbYes) Then
        Idx = InstanceTCP(sockVoice)
        If (Idx > 0) Then
            sockVoice(Idx).LocalPort = 0
            Call sockVoice(Idx).Accept(requestID)
            Call AddConnectionToList(sockVoice(Idx), ConnectionList)
            Call ResPlaySound(RingInId)
            cmdTalk.Enabled = True
        End If
    End If
End Sub

Private Sub sockVoice_DataArrival(Index As Integer, ByVal bytesTotal As Long)
    Dim rc As Long
    Dim WaveData() As Byte
    Static ExBytes(MAXTCP) As Long
    Static ExData(MAXTCP) As Variant
With wStream
    If (sockVoice(Index).BytesReceived > 0) Then
        Do While (sockVoice(Index).BytesReceived > 0)
            If (ExBytes(Index) = 0) Then
                If (.waveChunkSize <= sockVoice(Index).BytesReceived) Then
                    Call sockVoice(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize)
                    Call .SaveStreamBuffer(Index, WaveData)
                    Call .AddStreamToQueue(Index)
                Else
                    ExBytes(Index) = sockVoice(Index).BytesReceived
                    Call sockVoice(Index).GetData(ExData(Index), vbByte + vbArray, ExBytes(Index))
                End If
            Else
                Call sockVoice(Index).GetData(WaveData, vbByte + vbArray, .waveChunkSize - ExBytes(Index))
                ExData(Index) = MidB(ExData(Index), 1) & MidB(WaveData, 1)
                Call .SaveStreamBuffer(Index, ExData(Index))
                Call .AddStreamToQueue(Index)
                ExBytes(Index) = 0
                ExData(Index) = ""
            End If
        Loop
        
        If (Not .Playing And .PlayDeviceFree And _
            Not .Recording And .RecDeviceFree) Then
            Call cmdTalk_Click
        End If
    End If
End With
End Sub

Private Sub txtChat_Change()
txtChat.SelStart = Len(txtChat.Text)
End Sub

Private Sub txtSend_Change()
Dim tType As String
If Len(txtSend.Text) > 0 Then
    tTyping = True
Else
    tTyping = False
    tType = Encrypt("|NOTTYPING|")
    If frmMain.sockMain.State = sckConnected Then
        frmMain.sockMain.SendData tType
    End If
End If
If Len(txtSend.Text) > 1 Then
    Exit Sub
Else
    If tTyping = True Then
        tType = Encrypt("|TYPING|")
        If frmMain.sockMain.State = sckConnected Then
            frmMain.sockMain.SendData tType
        End If
    End If
End If
End Sub

Private Sub txtSend_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    If frmMain.sockMain.State = sckConnected Then
        If txtSend.Text = "" Then Exit Sub
        If txtSend.Text = " " Then
            Exit Sub
        Else
            frmMain.sockChat.SendData Encrypt("|CHAT|" & "{" & User & ":} " & txtSend.Text)
            txtChat.Text = txtChat.Text & "{" & User & ":} " & txtSend.Text & vbCrLf
            txtSend.Text = ""
            txtSend.SetFocus
        End If
    End If
End If

End Sub

