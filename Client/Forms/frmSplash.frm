VERSION 5.00
Begin VB.Form frmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Splash Screen"
   ClientHeight    =   3000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   Picture         =   "frmSplash.frx":0000
   ScaleHeight     =   3000
   ScaleWidth      =   5400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Interval        =   1
      Left            =   1080
      Top             =   2400
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2400
   End
   Begin VB.Timer Trans 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   600
      Top             =   2400
   End
End
Attribute VB_Name = "frmSplash"
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
Private Sub Trans_Timer()

Current = Current - 5
If Current + 1 <= Max Then
    Trans.Enabled = False
    Transparent frmSplash.hwnd, 0
    frmMain.Show
    frmLogin.Show modal, frmMain
    Unload Me
    Exit Sub
End If

Transparent frmSplash.hwnd, Current

End Sub

Private Sub Form_Load()
Trans.Interval = 1
Current = 0
Max = 255
Transparent frmSplash.hwnd, Current
FRM = 255
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

Private Sub Timer1_Timer()
Timer1.Enabled = False
Current = 255
Max = 0
Transparent frmSplash.hwnd, Current
Trans.Enabled = True
End Sub


Private Sub Timer3_Timer()
Current = Current + 5
If Current - 1 >= Max Then
    Timer3.Enabled = False
    Transparent frmSplash.hwnd, 255
    Timer1.Enabled = True
    Load frmMain
    Exit Sub
End If

Transparent frmSplash.hwnd, Current
End Sub

