VERSION 5.00
Begin VB.Form frmRejection 
   BorderStyle     =   0  'None
   Caption         =   "Connection Rejected By Server"
   ClientHeight    =   1410
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4665
   LinkTopic       =   "Form1"
   Picture         =   "frmRejection.frx":0000
   ScaleHeight     =   1410
   ScaleWidth      =   4665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Height          =   390
      Left            =   1800
      Picture         =   "frmRejection.frx":157F2
      Top             =   840
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Left            =   1800
      Picture         =   "frmRejection.frx":1702C
      Top             =   840
      Width           =   1170
   End
End
Attribute VB_Name = "frmRejection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
lLoggedIn = False
Reset_Images frmMain
frmMain.Login_Cancel
frmMain.tmrConnect.Enabled = False
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = True
End Sub

Private Sub Image1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2.Visible = False
Unload Me
End Sub
