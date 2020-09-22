VERSION 5.00
Begin VB.Form frmLogin 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   0  'None
   Caption         =   "SubNet Login"
   ClientHeight    =   2865
   ClientLeft      =   5550
   ClientTop       =   4740
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmLogin.frx":0000
   ScaleHeight     =   2865
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer AlphaTrans 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3960
      Top             =   2280
   End
   Begin VB.TextBox txtPass 
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
      IMEMode         =   3  'DISABLE
      Left            =   2160
      TabIndex        =   3
      Top             =   1875
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.TextBox txtUser 
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
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Text            =   "Usssssy"
      Top             =   1560
      Width           =   2175
   End
   Begin VB.TextBox txtServer 
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
      ForeColor       =   &H80000009&
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Text            =   "127.0.0.1"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "You Must Supply A Password"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   15
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Image imgCancelSel 
      Height          =   390
      Left            =   2640
      Picture         =   "frmLogin.frx":A522
      Top             =   2280
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgConnectSel 
      Height          =   390
      Left            =   600
      Picture         =   "frmLogin.frx":D2EE
      Top             =   2280
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   390
      Left            =   2640
      Picture         =   "frmLogin.frx":10194
      Top             =   2280
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   600
      Picture         =   "frmLogin.frx":12D4E
      Top             =   2280
      Width           =   1170
   End
End
Attribute VB_Name = "frmLogin"
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

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwflags As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const GWL_EXSTYLE = (-20)
Private Const LWA_ALPHA = &H2
Private Const WS_EX_LAYERED = &H80000

Dim Current As Integer ' current alpha transparency 0 = transparent 255 = opaque
Dim Max As Integer
Private Sub Form_Load()
Load frmVerify
txtPass.PasswordChar = Chr(164)
txtServer.Text = RemHst
txtUser.Text = User
txtPass.Text = Passy
Current = 255
Max = 0
Transparent frmLogin.hwnd, Current
Me.Height = 2865
Me.Width = 4680
Load_Login
End Sub
Private Sub Load_Login()
On Error GoTo Err
Open App.Path & "\Client.ini" For Input As #1
    Input #1, RemHst
    Input #1, User
Close
RemHst = Decrypt(RemHst)
User = Decrypt(User)
txtServer.Text = RemHst
txtUser.Text = User
Exit Sub
Err:
Close
Open App.Path & "\Client.ini" For Output As #1
    Print #1, Encrypt(RemHst)
    Print #1, Encrypt(User)
Close
Load_Login
End Sub
Private Sub AlphaTrans_Timer()

Current = Current - 20 '+5 is smooth from the 150 start, you can change
                      'experiment to find the one you like
If Current + 1 <= Max Then
    AlphaTrans.Enabled = False
    Transparent frmLogin.hwnd, 0
    Reset_Images frmMain
    Disable_Images frmMain
    Enable_imgMain frmMain
    Unload Me
    Exit Sub
End If

Transparent frmLogin.hwnd, Current

End Sub
Private Function Transparent(ByVal hwnd As Long, Perc As Integer) As Long
    Dim Msg As Long
    On Error Resume Next
    If Perc < 0 Or Perc > 255 Then
      Transparent = 1
    Else
      Msg = GetWindowLong(hwnd, GWL_EXSTYLE)
      Msg = Msg Or WS_EX_LAYERED
      SetWindowLong hwnd, GWL_EXSTYLE, Msg
      SetLayeredWindowAttributes hwnd, 0, Perc, LWA_ALPHA
      Transparent = 0
    End If
    If Err Then
      Transparent = 2
    End If
End Function

Private Sub Form_LostFocus()
Me.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgConnectSel.Visible = True
Image1.Visible = False
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image1.Visible = True
imgConnectSel.Visible = False
If Label1.Visible = True Then GoTo Finish_Connecting
If txtUser.Text = "" Then
    Load frmWarning
    frmMain.StayOnTop frmWarning
    frmWarning.Height = 1410
    frmWarning.Width = 4680
    frmWarning.Left = (Screen.Width - frmWarning.Width) / 2
    frmWarning.Top = (Screen.Height - frmWarning.Height) / 2
    frmWarning.Show
    frmLogin.Visible = False
    Beep
    Exit Sub
End If
Load frmVerify
StayOnTop frmVerify
    frmVerify.Height = 1410
    frmVerify.Width = 4665
    frmVerify.Top = (Screen.Height - frmVerify.Height) / 2
    frmVerify.Left = (Screen.Width - frmVerify.Width) / 2
frmVerify.Show
User = txtUser.Text
Passy = txtPass.Text
RemHst = txtServer.Text
Open App.Path & "\Client.ini" For Output As #1
    Print #1, Encrypt(RemHst)
    Print #1, Encrypt(User)
Close
Reset_Images frmMain
Connect_Main frmMain, frmMain.sockMain
frmMain.tmrConnect.Enabled = True
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
Unload Me
Exit Sub

Finish_Connecting:

Load frmVerify
StayOnTop frmVerify
    frmVerify.Height = 1410
    frmVerify.Width = 4665
    frmVerify.Top = (Screen.Height - frmVerify.Height) / 2
    frmVerify.Left = (Screen.Width - frmVerify.Width) / 2
    frmVerify.Show
Dim mmssgg As String
User = txtUser.Text
Passy = txtPass.Text
RemHst = txtServer.Text
Open App.Path & "\Client.ini" For Output As #1
    Print #1, Encrypt(RemHst)
    Print #1, Encrypt(User)
Close
Reset_Images frmMain
'frmMain.tmrConnect.Enabled = True
mmssgg = "|SLOGIN|" & User & ":" & Passy
mmssgg = Encrypt(mmssgg)
If frmMain.sockMain.State = sckConnected Then
    frmMain.sockMain.SendData mmssgg
End If
frmMain.tmrConnect.Enabled = True
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
Unload Me
End Sub
Public Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCancelSel.Visible = True
Image2.Visible = False

End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
imgCancelSel.Visible = False
Login = False
AlphaTrans.Enabled = True
Reset_Images frmMain
frmMain.Login_Cancel
Unload frmVerify
End Sub


Private Sub txtPass_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then ConnectOnEnter
End Sub
Private Sub ConnectOnEnter()
Image1.Visible = True
imgConnectSel.Visible = False
If Label1.Visible = True Then GoTo Finish_Connecting
If txtUser.Text = "" Then
    Load frmWarning
    frmMain.StayOnTop frmWarning
    frmWarning.Height = 1410
    frmWarning.Width = 4680
    frmWarning.Left = (Screen.Width - frmWarning.Width) / 2
    frmWarning.Top = (Screen.Height - frmWarning.Height) / 2
    frmWarning.Show
    frmLogin.Visible = False
    Beep
    Exit Sub
End If
Load frmVerify
StayOnTop frmVerify
    frmVerify.Height = 1410
    frmVerify.Width = 4665
    frmVerify.Top = (Screen.Height - frmVerify.Height) / 2
    frmVerify.Left = (Screen.Width - frmVerify.Width) / 2
frmVerify.Show
User = txtUser.Text
Passy = txtPass.Text
RemHst = txtServer.Text
Open App.Path & "\Client.ini" For Output As #1
    Print #1, Encrypt(RemHst)
    Print #1, Encrypt(User)
Close
Reset_Images frmMain
Connect_Main frmMain, frmMain.sockMain
'AlphaTrans.Enabled = True
frmMain.tmrConnect.Enabled = True
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
Unload Me
Exit Sub

Finish_Connecting:

Load frmVerify
StayOnTop frmVerify
    frmVerify.Height = 1410
    frmVerify.Width = 4665
    frmVerify.Top = (Screen.Height - frmVerify.Height) / 2
    frmVerify.Left = (Screen.Width - frmVerify.Width) / 2
    frmVerify.Show
Dim mmssgg As String
User = txtUser.Text
Passy = txtPass.Text
RemHst = txtServer.Text
Open App.Path & "\Client.ini" For Output As #1
    Print #1, Encrypt(RemHst)
    Print #1, Encrypt(User)
Close
Reset_Images frmMain
frmMain.tmrConnect.Enabled = True
mmssgg = "|SLOGIN|" & User & ":" & Passy
mmssgg = Encrypt(mmssgg)
If frmMain.sockMain.State = sckConnected Then
    frmMain.sockMain.SendData mmssgg
End If
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
Unload Me
End Sub
