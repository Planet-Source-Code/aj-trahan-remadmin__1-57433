VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   3285
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1200
      Left            =   3120
      Top             =   360
   End
   Begin VB.Timer AlphaTrans 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2640
      Top             =   360
   End
   Begin VB.Image imgOKSel 
      Height          =   390
      Left            =   240
      Picture         =   "frmAbout.frx":BD0F
      Top             =   2400
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgWebSel 
      Height          =   390
      Left            =   1440
      Picture         =   "frmAbout.frx":E707
      Top             =   2400
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image imgSysInfoSel 
      Height          =   390
      Left            =   3000
      Picture         =   "frmAbout.frx":11BBD
      Top             =   2400
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.Image imgWeb 
      Height          =   390
      Left            =   1440
      Picture         =   "frmAbout.frx":14F4F
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Image imgSysInfo 
      Height          =   390
      Left            =   3000
      Picture         =   "frmAbout.frx":1806A
      Top             =   2400
      Width           =   1515
   End
   Begin VB.Image imgOK 
      Height          =   390
      Left            =   240
      Picture         =   "frmAbout.frx":1B08F
      Top             =   2400
      Width           =   1170
   End
End
Attribute VB_Name = "frmAbout"
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


Dim Current As Integer
Dim Max As Integer
Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image4_Click()
Unload Me
End Sub

Private Sub Label1_Click()
Unload Me
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

Private Sub AlphaTrans_Timer()
Current = Current - 20
If Current - 1 >= Max Then
    AlphaTrans.Enabled = False
    Transparent frmAbout.hwnd, 255
    Exit Sub
End If

Transparent frmAbout.hwnd, Current
End Sub

Private Sub Form_Load()
About = True
Current = 255
Max = 255
Transparent frmAbout.hwnd, Current
End Sub

Public Sub DoTheStuff()
Timer1.Enabled = True
AlphaTrans.Enabled = True
End Sub

Private Sub imgOK_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOKSel.Visible = True
imgOK.Visible = False
End Sub

Private Sub imgOK_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgOK.Visible = True
imgOKSel.Visible = False
About = False
frmMain.imgAbout.Visible = True
frmMain.imgAboutSel.Visible = False
frmMain.imgAboutSel.Top = 1800
DoTheStuff
End Sub

Private Sub imgSysInfo_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSysInfoSel.Visible = True
imgSysInfo.Visible = False

End Sub

Private Sub imgSysInfo_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgSysInfo.Visible = True
imgSysInfoSel.Visible = False
End Sub

Private Sub imgWeb_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgWebSel.Visible = True
imgWeb.Visible = False

End Sub

Private Sub imgWeb_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgWeb.Visible = True
imgWebSel.Visible = False
End Sub

Private Sub Timer1_Timer()
Unload Me
End Sub
