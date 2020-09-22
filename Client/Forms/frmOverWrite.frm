VERSION 5.00
Begin VB.Form frmOverWrite 
   BorderStyle     =   0  'None
   Caption         =   "OverWrite A File"
   ClientHeight    =   2280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5535
   LinkTopic       =   "Form1"
   Picture         =   "frmOverWrite.frx":0000
   ScaleHeight     =   2280
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Image Image2 
      Height          =   390
      Index           =   1
      Left            =   3000
      Picture         =   "frmOverWrite.frx":29222
      Top             =   1680
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   0
      Left            =   1320
      Picture         =   "frmOverWrite.frx":2AA5C
      Top             =   1680
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   1
      Left            =   3000
      Picture         =   "frmOverWrite.frx":2C296
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   0
      Left            =   1320
      Picture         =   "frmOverWrite.frx":2DAD0
      Top             =   1680
      Width           =   1170
   End
   Begin VB.Label lblFileName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File Name"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   870
      Width           =   5055
   End
End
Attribute VB_Name = "frmOverWrite"
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

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).Visible = True
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).Visible = False
If Index = 0 Then
    frmCustComDlg.Transfer_File
    GoTo Ending
End If
If Index = 1 Then
    Unload Me
End If
Ending:
Unload Me
End Sub

