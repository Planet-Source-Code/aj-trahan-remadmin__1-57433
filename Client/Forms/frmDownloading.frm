VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDownloading 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   Picture         =   "frmDownloading.frx":0000
   ScaleHeight     =   3480
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer Timer4 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1800
      Top             =   720
   End
   Begin VB.Timer tmrSpeed 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   1320
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   75
      Left            =   840
      Top             =   720
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   240
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   294
      ImageHeight     =   80
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":7188
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":1861A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":29AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":3AF3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":4C3D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":5D862
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":6ECF4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":80186
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":91618
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":A2AAA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":B3F3C
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":C53CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":D6860
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDownloading.frx":E7CF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar objProg 
      Height          =   135
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   238
      _Version        =   393216
      Appearance      =   0
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Server"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   1
      Left            =   3090
      TabIndex        =   6
      Top             =   1480
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Client"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   1480
      Width           =   1335
   End
   Begin VB.Label lblSpeed 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Download Speed = 0 KB(s)/Sec"
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
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   4335
   End
   Begin VB.Label lblFileName 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1440
      TabIndex        =   3
      Top             =   2110
      Width           =   3015
   End
   Begin VB.Label lblBytes 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   2160
      TabIndex        =   2
      Top             =   2440
      Width           =   2295
   End
   Begin VB.Label lblGotten 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   1200
      TabIndex        =   1
      Top             =   2760
      Width           =   3255
   End
   Begin VB.Image Image1 
      Height          =   1230
      Left            =   120
      Top             =   360
      Width           =   4440
   End
End
Attribute VB_Name = "frmDownloading"
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
Dim iImage As Long
Dim rReceived As Long
Dim dCount As Long
Dim kKB As String

Private Sub Form_Load()
If bFileTransfer = False Then
    Unload Me
    Exit Sub
End If
Me.Top = 0
Me.Left = 0
If lblFileName.Caption = "" Then
    lblFileName.Caption = "Remote Snapshot"
End If
Image1.Picture = ImageList1.ListImages(1).Picture
iImage = 2
Change_Pictures
Me.Refresh
rReceived = 0
dCount = 0
If uUpLoad = True Then
    Ready_Upload
Else
    tmrSpeed.Enabled = True
End If
End Sub
Private Sub Ready_Upload()
Timer4.Enabled = True
tmrSpeed.Enabled = False
lblSpeed.Caption = "Upload Speed = 0 KB(s)/Sec"
End Sub
Private Sub Change_Pictures()
    iImage = iImage - 1
    If iImage <= 0 Then iImage = 14
    Image1.Picture = ImageList1.ListImages(iImage).Picture
End Sub
Private Sub Change4Upload()
    iImage = iImage + 1
    If iImage >= 14 Then iImage = 1
    Image1.Picture = ImageList1.ListImages(iImage).Picture
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Timer1_Timer()
If uUpLoad = True Then
    Change4Upload
Else
    Change_Pictures
End If
End Sub

Private Sub Timer4_Timer()
lblGotten.Caption = sSent
If lblGotten.Caption = bbIt Then Exit Sub
dCount = dCount + 1
rReceived = lblGotten.Caption
rReceived = ((bbIt - rReceived) / 1024) / dCount
lblSpeed.Caption = "Upload Speed = " & rReceived & " KB(s)/Sec"
End Sub

Private Sub tmrSpeed_Timer()
lblGotten.Caption = bBytes
If lblGotten.Caption = "0" Then Exit Sub
dCount = dCount + 1
rReceived = lblGotten.Caption
rReceived = (rReceived / 1024) / dCount
lblSpeed.Caption = "Download Speed = " & rReceived & " KB(s)/Sec"
End Sub
