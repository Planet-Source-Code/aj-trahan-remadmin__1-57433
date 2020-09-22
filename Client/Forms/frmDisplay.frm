VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmDisplay 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Remote Imaging"
   ClientHeight    =   7065
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9330
   LinkTopic       =   "Form1"
   Picture         =   "frmDisplay.frx":0000
   ScaleHeight     =   7065
   ScaleWidth      =   9330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Visible         =   0   'False
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   720
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplay.frx":41EB42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplay.frx":41F41C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplay.frx":41FCF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDisplay.frx":4205D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image imgMouse 
      Height          =   480
      Left            =   4200
      Picture         =   "frmDisplay.frx":420EAA
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image imgSubNet 
      Height          =   240
      Left            =   0
      Picture         =   "frmDisplay.frx":421774
      Top             =   0
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.Image imgMin 
      Height          =   120
      Left            =   4680
      Picture         =   "frmDisplay.frx":422A76
      Top             =   120
      Width           =   120
   End
   Begin VB.Image imgMinMax 
      Height          =   120
      Left            =   4800
      Picture         =   "frmDisplay.frx":422B78
      Top             =   120
      Width           =   120
   End
   Begin VB.Image imgClose 
      Height          =   120
      Left            =   4920
      Picture         =   "frmDisplay.frx":422C7A
      Top             =   120
      Width           =   120
   End
   Begin VB.Image imgBottomRight 
      Height          =   75
      Left            =   4800
      Picture         =   "frmDisplay.frx":422D7C
      Top             =   1320
      Width           =   90
   End
   Begin VB.Image imgBottomLeft 
      Height          =   90
      Left            =   2160
      Picture         =   "frmDisplay.frx":422E22
      Top             =   1320
      Width           =   90
   End
   Begin VB.Image imgBottom 
      Height          =   120
      Left            =   1320
      Picture         =   "frmDisplay.frx":422EDC
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   4410
   End
   Begin VB.Image imgTopRight 
      Height          =   210
      Left            =   4800
      Picture         =   "frmDisplay.frx":423BEE
      Top             =   120
      Width           =   90
   End
   Begin VB.Image imgTopLeft 
      Height          =   210
      Left            =   1920
      Picture         =   "frmDisplay.frx":423D48
      Top             =   120
      Width           =   1185
   End
   Begin VB.Image imgRight 
      Height          =   1635
      Left            =   4800
      Picture         =   "frmDisplay.frx":424AAA
      Stretch         =   -1  'True
      Top             =   0
      Width           =   120
   End
   Begin VB.Image imgLeft 
      Height          =   1605
      Left            =   2160
      Picture         =   "frmDisplay.frx":425384
      Stretch         =   -1  'True
      Top             =   0
      Width           =   90
   End
   Begin VB.Image imgTop 
      Height          =   210
      Left            =   1920
      Picture         =   "frmDisplay.frx":425C72
      Stretch         =   -1  'True
      Top             =   120
      Width           =   3135
   End
   Begin VB.Image imgBlank 
      Height          =   4305
      Left            =   960
      Picture         =   "frmDisplay.frx":428174
      Top             =   2280
      Width           =   6045
   End
   Begin VB.Image imgDisplay 
      Height          =   7500
      Left            =   480
      Picture         =   "frmDisplay.frx":47D07A
      Stretch         =   -1  'True
      Top             =   1800
      Width           =   7500
   End
End
Attribute VB_Name = "frmDisplay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hHeight As Long
Dim wWidth As Long

Private Sub Form_Load()
'StayOnTop Me
Me.Width = (Screen.Width / 3) * 2
Me.Height = (Screen.Height / 3) * 2
hHeight = Me.Height
wWidth = Me.Width
ChangeView
nNewScreen = False
End Sub
Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub ChangeView()
imgTop.Top = 0
imgTop.Left = 0
imgTop.Width = Me.Width
imgTopLeft.Top = 0
imgTopLeft.Left = 0
imgTopRight.Top = 0
imgTopRight.Left = Me.Width - imgTopRight.Width
imgLeft.Left = 0
imgLeft.Height = Me.Height
imgRight.Left = Me.Width - imgRight.Width
imgRight.Height = Me.Height
imgBottom.Top = Me.Height - imgBottom.Height
imgBottom.Left = 0
imgBottom.Width = Me.Width
imgBottomLeft.Top = Me.Height - imgBottomLeft.Height
imgBottomLeft.Left = 0
imgBottomRight.Left = Me.Width - imgBottomRight.Width
imgBottomRight.Top = Me.Height - imgBottomRight.Height
imgDisplay.Width = Me.Width - imgLeft.Width - imgRight.Width
imgDisplay.Height = Me.Height - imgTop.Height - imgBottom.Height
imgDisplay.Top = imgTop.Top + imgTop.Height
imgDisplay.Left = imgLeft.Left + imgLeft.Width
imgBlank.Top = (imgDisplay.Height - imgBlank.Height) / 2
imgBlank.Left = (imgDisplay.Width - imgBlank.Width) / 2
imgMin.Top = 50
imgMinMax.Top = 50
imgClose.Top = 50
imgClose.Left = Me.Width - 200
imgMinMax.Left = imgClose.Left - imgClose.Width - 20
imgMin.Left = imgMinMax.Left - imgMinMax.Width - 20
Me.Left = (Screen.Height - Me.Height) / 2
If Me.Width = Screen.Width Then
    Me.Top = 0
    Me.Left = 0
Else
    Me.Top = (Screen.Height - Me.Height) / 2
End If
Me.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\Screen.jpg"
End Sub

Private Sub imgClose_Click()
    nNewScreen = True
    frmMain.sockScreen.SendData "|STOPTIMER|"
    Unload Me
End Sub

Private Sub imgDisplay_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
'if right click(button=2), refresh image
If nNewScreen = True Then Exit Sub
If Button = 2 Then
    Kill App.Path & "\Screen.jpg"
    Open App.Path & "\Screen.jpg" For Binary As #50
    nNewScreen = True
    frmMain.sockScreen.SendData "|STOPTIMER|"
    Pause 20
    frmMain.sockScreen.SendData "|SENDNEW|"
End If
End Sub

Private Sub imgMin_Click()
    imgTopLeft.Visible = False
    imgTop.Visible = False
    imgLeft.Visible = False
    imgBottomLeft.Visible = False
    imgBottom.Visible = False
    imgSubNet.Visible = True
    Me.Width = imgSubNet.Width
    Me.Height = imgSubNet.Height
    Me.Left = 0
    Me.Top = Screen.Height - 450 - Me.Height

End Sub

Private Sub imgMinMax_Click()
Me.Visible = False
If Me.Width = wWidth Then
    Me.Width = Screen.Width
    Me.Height = Screen.Height - 450
Else
    Me.Width = wWidth
    Me.Height = hHeight
End If
ChangeView
End Sub

Private Sub imgSubNet_Click()
Me.Visible = False
imgSubNet.Visible = False
imgTop.Visible = True
imgTopLeft.Visible = True
imgLeft.Visible = True
imgBottom.Visible = True
imgBottomLeft.Visible = True
Me.Height = hHeight
Me.Width = wWidth
ChangeView
End Sub

Private Sub imgTop_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub


