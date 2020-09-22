VERSION 5.00
Begin VB.Form frmCustComDlg 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6045
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8520
   LinkTopic       =   "Form1"
   Picture         =   "frmCustComDlg.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   8520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.FileListBox fileList 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2010
      Left            =   1800
      TabIndex        =   3
      Top             =   3120
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox txtFileName 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1800
      TabIndex        =   2
      Top             =   5520
      Width           =   3375
   End
   Begin VB.DirListBox dirPath 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   1800
      TabIndex        =   1
      Top             =   720
      Width           =   6495
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1800
      TabIndex        =   0
      Top             =   240
      Width           =   3375
   End
   Begin VB.Label lblUpLoad 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Select A File To Upload"
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
      Left            =   1800
      TabIndex        =   4
      Top             =   2930
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.Image imgUploadSel 
      Height          =   390
      Left            =   5520
      Picture         =   "frmCustComDlg.frx":A7ABA
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgUpload 
      Height          =   390
      Left            =   5520
      Picture         =   "frmCustComDlg.frx":A92F4
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image3 
      Height          =   1215
      Index           =   3
      Left            =   293
      Picture         =   "frmCustComDlg.frx":AAB2E
      Top             =   4535
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Index           =   2
      Left            =   278
      Picture         =   "frmCustComDlg.frx":AF8A4
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Index           =   1
      Left            =   280
      Picture         =   "frmCustComDlg.frx":B461A
      Top             =   1950
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1215
      Index           =   0
      Left            =   270
      Picture         =   "frmCustComDlg.frx":B9390
      Top             =   670
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   1
      Left            =   6960
      Picture         =   "frmCustComDlg.frx":BE106
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image2 
      Height          =   390
      Index           =   0
      Left            =   5520
      Picture         =   "frmCustComDlg.frx":BF940
      Top             =   5520
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   1
      Left            =   6960
      Picture         =   "frmCustComDlg.frx":C117A
      Top             =   5520
      Width           =   1170
   End
   Begin VB.Image Image1 
      Height          =   390
      Index           =   0
      Left            =   5520
      Picture         =   "frmCustComDlg.frx":C29B4
      Top             =   5520
      Width           =   1170
   End
End
Attribute VB_Name = "frmCustComDlg"
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

Dim fFileName As String
Dim pPath As String
Dim dDuplicate As Boolean


Private Sub dirPath_Change()
dDuplicate = False
fileList.Path = dirPath.Path
End Sub

Private Sub Drive1_Change()
dirPath.Path = Drive1.Drive
End Sub

Private Sub fileList_Click()
'need to come up w/ something to put the file name in txtfilename.text
Dim X As Integer
X = fileList.ListIndex
txtFileName.Text = fileList.List(X)

End Sub

Private Sub Form_Load()
StayOnTop frmCustComDlg
Me.Height = 6045
Me.Width = 8520
Me.Left = (Screen.Width - Me.Width) / 2
Me.Top = (Screen.Height - Me.Height) / 2
dirPath.Path = Drive1.Drive
fileList.Path = dirPath.Path
If uUpLoad = True Then
    lblUpLoad.Visible = True
    UpLoad_File
End If
End Sub
Private Sub UpLoad_File()
    dirPath.Height = 2190
    fileList.Visible = True
    Image1(0).Visible = False
    Image2(0).Visible = False
    imgUpload.Visible = True
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).Visible = True
End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image2(Index).Visible = False

If Index = 0 Then
    pPath = dirPath.Path
    fFileName = Mid(txtFileName.Text, 1, Len(txtFileName.Text) - 1)
    Check4Duplicate
    Exit Sub
End If
If Index = 1 Then
    uUpLoad = False
    Unload Me
End If
End Sub
Private Sub Check4Duplicate()
Dim I As Integer
For I = 0 To fileList.ListCount - 1
    If LCase(fFileName) = LCase(fileList.List(I)) Then dDuplicate = True
Next I

If dDuplicate = True Then
    Load frmOverWrite
    StayOnTop frmOverWrite
    With frmOverWrite
        .Height = 2280
        .Width = 5535
        .Top = (Screen.Height - .Height) / 2
        .Left = (Screen.Width - .Width) / 2
        .lblFileName.Caption = fFileName
        .Show
    End With
    Beep
    Exit Sub
Else
    Transfer_File
    Exit Sub
End If
End Sub
Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Public Sub Transfer_File()
    bBytes = 0
    Open pPath & "\" & fFileName For Binary As #1
    bFileTransfer = True
    frmDownloading.lblFileName = frmMain.lvFiles.SelectedItem.Text
    frmDownloading.lblGotten.Caption = "0"
    frmDownloading.Show modal, frmMain
    frmMain.sockExplorer.SendData "|GETFILE|" & frmMain.lvFiles.SelectedItem.Key
    Pause 50
    Unload Me
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(Index).Visible = False
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Image3(Index).Visible = True
If Index = 0 Then
    dirPath.Path = Environ("homedrive") & Environ("homepath") & "\Desktop"
End If
If Index = 1 Then
    dirPath.Path = Environ("homedrive") & Environ("homepath") & "\My Documents"
End If
If Index = 2 Then
    dirPath.Path = Environ("homedrive") & "\"
End If
If Index = 3 Then
    dirPath.Path = Environ("homedrive") & Environ("homepath") & "\NetHood"
End If
End Sub


Private Sub imgUpload_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgUploadSel.Visible = True
End Sub

Private Sub imgUpload_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error GoTo NoFile

    imgUploadSel.Visible = False
    FileToSend = txtFileName.Text
    Me.Visible = False
    If Len(FileToSend) <= 0 Then Exit Sub
    bFileTransfer = True
    Load frmDownloading
    frmDownloading.Show modal, frmMain
    frmDownloading.lblFileName.Caption = FileToSend
    Msg = "|UPLOAD|" & FolderClick & "\" & FileToSend
    frmMain.sockExplorer.SendData Msg
    Pause 10
    FileToSend = dirPath.Path & "\" & FileToSend
    Call SendFile(FileToSend, frmMain.sockExplorer)
    frmMain.sockExplorer.SendData "|DONEUPLOAD|"
NoFile:
Unload frmDownloading
bFileTransfer = False
uUpLoad = False
Unload Me
End Sub

