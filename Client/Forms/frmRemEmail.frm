VERSION 5.00
Begin VB.Form frmRemEmail 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   Picture         =   "frmRemEmail.frx":0000
   ScaleHeight     =   7110
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtSubject 
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
      Left            =   1920
      TabIndex        =   1
      Top             =   1135
      Width           =   4635
   End
   Begin VB.TextBox txtSendTo 
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
      Left            =   1920
      TabIndex        =   0
      Top             =   480
      Width           =   4635
   End
   Begin VB.TextBox txtBody 
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
      Height          =   4520
      Left            =   310
      MultiLine       =   -1  'True
      TabIndex        =   2
      Top             =   1640
      Width           =   6240
   End
   Begin VB.Image imgClickSel 
      Height          =   390
      Index           =   1
      Left            =   5400
      Picture         =   "frmRemEmail.frx":A11A2
      Top             =   6480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgClickSel 
      Height          =   390
      Index           =   0
      Left            =   2880
      Picture         =   "frmRemEmail.frx":A29DC
      Top             =   6480
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgClick 
      Height          =   390
      Index           =   1
      Left            =   5400
      Picture         =   "frmRemEmail.frx":A4216
      Top             =   6480
      Width           =   1170
   End
   Begin VB.Image imgClick 
      Height          =   390
      Index           =   0
      Left            =   2880
      Picture         =   "frmRemEmail.frx":A5A50
      Top             =   6480
      Width           =   1170
   End
End
Attribute VB_Name = "frmRemEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SendTo As Boolean
Dim HasBody As Boolean

Private Sub Form_Load()
imgClick(0).Visible = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub imgClick_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClickSel(Index).Visible = True
End Sub

Private Sub imgClick_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    imgClickSel(Index).Visible = False
    If Index = 0 Then SendRemoteEmail
    If Index = 1 Then Unload Me
End Sub

Private Sub SendRemoteEmail()
    If SendTo = False Then Exit Sub
    Dim tTo As String
    Dim sSub As String
    Dim bBod As String
    tTo = txtSendTo.Text & "|"
    sSub = txtSubject.Text & "|"
    bBod = txtBody.Text
    RemEmail = "|EMAIL|" & tTo & sSub & bBod
    RemEmail = Encrypt(RemEmail)
    If frmMain.sockMain.State = sckConnected Then
        frmMain.sockMain.SendData RemEmail
        Pause 200
    End If
    Unload Me
End Sub

Private Sub txtSendTo_Change()
If InStr(1, txtSendTo.Text, "@") <> 0 Then
    SendTo = True
    imgClick(0).Visible = True
Else
    SendTo = False
    imgClick(0).Visible = False
End If
End Sub

