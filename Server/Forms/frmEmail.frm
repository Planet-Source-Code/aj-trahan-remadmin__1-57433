VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmEmail 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6945
   LinkTopic       =   "Form1"
   Picture         =   "frmEmail.frx":0000
   ScaleHeight     =   7110
   ScaleWidth      =   6945
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock sockMail 
      Left            =   360
      Top             =   6480
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrEmail 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   840
      Top             =   6480
   End
   Begin VB.TextBox txtTo 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1920
      TabIndex        =   2
      Top             =   480
      Width           =   4655
   End
   Begin VB.TextBox txtSubject 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   1920
      TabIndex        =   1
      Top             =   1140
      Width           =   4655
   End
   Begin VB.TextBox txtBody 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   4495
      Left            =   310
      TabIndex        =   0
      Top             =   1650
      Width           =   6235
   End
   Begin VB.Image imgSendSel 
      Height          =   390
      Index           =   0
      Left            =   2880
      Picture         =   "frmEmail.frx":A11A2
      Top             =   6360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgSend 
      Height          =   390
      Index           =   0
      Left            =   2880
      Picture         =   "frmEmail.frx":A29DC
      Top             =   6360
      Width           =   1170
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSend_Click()

End Sub

Private Sub Form_Load()
    If RemoteEmail = True Then
        frmServer.lblStatus.Caption = " Status:  Sending Remote E-Mail..."
    End If
    tmrEmail.Enabled = True
End Sub

Private Sub imgSend_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imgSendSel(Index).Visible = True
End Sub

Private Sub imgSend_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    imgSendSel(Index).Visible = False
    eStage = 0
    Transmit eStage
End Sub

Private Sub sockMail_DataArrival(ByVal bytesTotal As Long)
    Dim eData As String
    sockMail.GetData eData, vbString
    If eTransmitting = True Then
        eStage = eStage + 1
        Transmit eStage
    Else
        If sockMail.State <> sckClosed Then sockMail.Close
    End If
End Sub

Private Sub tmrEmail_Timer()
    Dim Flags As Long
    Dim Result As Boolean
    Result = InternetGetConnectedState(Flags, 0)
    If Result Then
        If RemoteEmail = False Then
            frmServer.lblStatus.Caption = " Status:  Sending Local Information To, " & eMailAddress
        End If
        tmrEmail.Enabled = False
        sockMail.LocalPort = 0
        sockMail.Protocol = sckTCPProtocol
        sockMail.Connect eSMTPserver, "25"
        eTransmitting = True
        eStage = 0
    End If
End Sub

Private Sub txtBody_Change()
    eBody = txtBody.Text
End Sub

Private Sub txtSubject_Change()
    eSubject = txtSubject.Text
End Sub

Private Sub txtTo_Change()
    eTo = txtTo.Text
End Sub

Private Sub Transmit(eStage As Integer)
    Dim Helo As String, temp As String
    Dim pos As Integer
Select Case eStage
    Case 1:
        Helo = eFrom
        pos = Len(Helo) - InStr(Helo, "@")
        temp = "HELO " & Right$(Helo, pos) & vbCrLf
        sockMail.SendData temp
    Case 2:
        temp = "MAIL FROM: <" & Trim(eFrom) & ">" & vbCrLf
        sockMail.SendData temp
    Case 3:
        temp = "RCPT TO: <" & Trim(eTo) & ">" & vbCrLf
        sockMail.SendData temp
    Case 4:
        temp = "DATA" & vbCrLf
        sockMail.SendData "DATA" & vbCrLf
    Case 5:
        temp = temp & "From: " & eFrom & vbNewLine
        temp = temp & "To: " & eTo & vbNewLine
        temp = temp & "Subject: " & eSubject & vbNewLine
        temp = temp & vbCrLf & eBody
        temp = temp & vbCrLf & Now
        sockMail.SendData temp
        sockMail.SendData vbCrLf & "." & vbCrLf
        eStage = 0
        eTransmitting = False
        Pause 10
        sockMail.Close
        RemoteEmail = False
        frmServer.lblStatus.Caption = " Status:  E-mail Sent"
        frmServer.tmrStatus.Enabled = True
        Unload Me
End Select
End Sub
