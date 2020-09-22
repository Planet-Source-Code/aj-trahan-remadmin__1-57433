VERSION 5.00
Begin VB.Form frmConnectError 
   BorderStyle     =   0  'None
   Caption         =   "Unable To Connect"
   ClientHeight    =   1785
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmConnectError.frx":0000
   ScaleHeight     =   1785
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Height          =   390
      Left            =   1745
      Picture         =   "frmConnectError.frx":1B35A
      Top             =   1200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   1745
      Picture         =   "frmConnectError.frx":1CB94
      Top             =   1200
      Width           =   1170
   End
End
Attribute VB_Name = "frmConnectError"
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
Private Sub Form_Load()

End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Load frmMain
frmLogin.Show
Unload Me
End Sub
