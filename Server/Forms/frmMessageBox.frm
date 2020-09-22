VERSION 5.00
Begin VB.Form frmMessageBox 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmMessageBox.frx":0000
   ScaleHeight     =   2040
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgIgnoreSel 
      Height          =   390
      Left            =   3240
      Picture         =   "frmMessageBox.frx":1EDDA
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgIgnore 
      Height          =   390
      Left            =   3240
      Picture         =   "frmMessageBox.frx":20614
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRetrySel 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":21E4E
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRetry 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":23688
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgAbortSel 
      Height          =   390
      Left            =   360
      Picture         =   "frmMessageBox.frx":24EC2
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgAbort 
      Height          =   390
      Left            =   360
      Picture         =   "frmMessageBox.frx":266FC
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblMsgTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Microsoft Windows"
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
      Height          =   135
      Left            =   1200
      TabIndex        =   1
      Top             =   0
      Width           =   3375
   End
   Begin VB.Image imgNoSel 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":27F36
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgNo 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":29770
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgYesSel 
      Height          =   390
      Left            =   360
      Picture         =   "frmMessageBox.frx":2AFAA
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgYes 
      Height          =   390
      Left            =   360
      Picture         =   "frmMessageBox.frx":2C7E4
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgCancelSel 
      Height          =   390
      Left            =   3240
      Picture         =   "frmMessageBox.frx":2E01E
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgCancel 
      Height          =   390
      Left            =   3240
      Picture         =   "frmMessageBox.frx":2F858
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgOKSel 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":31092
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgOk 
      Height          =   390
      Left            =   1800
      Picture         =   "frmMessageBox.frx":328CC
      Top             =   1440
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      Height          =   825
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   3765
   End
   Begin VB.Image imgExclamation 
      Height          =   780
      Left            =   120
      Picture         =   "frmMessageBox.frx":34106
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image imgCritical 
      Height          =   780
      Left            =   120
      Picture         =   "frmMessageBox.frx":35F58
      Top             =   240
      Visible         =   0   'False
      Width           =   735
   End
End
Attribute VB_Name = "frmMessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    frmServer.tmrStatus.Enabled = True
End Sub

Private Sub Form_Terminate()
    frmServer.tmrStatus.Enabled = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmServer.tmrStatus.Enabled = True
End Sub

Private Sub imgAbort_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgAbortSel.Visible = True
End Sub

Private Sub imgAbort_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|ABORT"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub imgCancel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgCancelSel.Visible = True
End Sub

Private Sub imgCancel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|CANCEL"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub imgIgnore_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgIgnoreSel.Visible = True
End Sub

Private Sub imgIgnore_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|IGNORE"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub imgNo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgNoSel.Visible = True
End Sub

Private Sub imgNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|NO"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub
Private Sub imgOk_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgOKSel.Visible = True
End Sub

Private Sub imgOk_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|OK"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub imgRetry_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgRetrySel.Visible = True
End Sub

Private Sub imgRetry_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|RETRY"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub imgYes_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    imgYesSel.Visible = True
End Sub

Private Sub imgYes_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    mMessageButton = "|MSGBUTTON|YES"
    SendMsgBtn frmServer.W1
    Unload Me
End Sub

Private Sub lblMsgTitle_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub
