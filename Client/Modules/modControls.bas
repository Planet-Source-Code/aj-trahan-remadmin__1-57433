Attribute VB_Name = "modControls"
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
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public lLoggedIn As Boolean
Public fFilenName As String
Public pProcess As Boolean
Public iInverse As Boolean
Public rReady As Boolean
Public RemHst As String
Public User As String
Public Passy As String
Public tTyping As Boolean
Public dDesk As Boolean
Public uUpLoad As Boolean
Public FileToSend As String
Public bbIt As Long
Public sSent As Long
Public bBytes As Long
Public FRM As Integer
Public RemEmail As String
Public nNewScreen As Boolean
'***** For Changing the Screen Resolution ****
Private Declare Function EnumDisplaySettings Lib "user32" Alias "EnumDisplaySettingsA" (ByVal lpszDeviceName As Long, ByVal iModeNum As Long, lpDevMode As Any) As Boolean
Private Declare Function ChangeDisplaySettings Lib "user32" Alias "ChangeDisplaySettingsA" (lpDevMode As Any, ByVal dwflags As Long) As Long
    Private Const CCDEVICENAME = 32
    Private Const CCFORMNAME = 32
    Private Const DM_BITSPERPEL = &H60000
    Private Const DM_PELSWIDTH = &H80000
    Private Const DM_PELSHEIGHT = &H100000
Private Type DEVMODE
    dmDeviceName As String * CCDEVICENAME
    dmSpecVersion As Integer
    dmDriverVersion As Integer
    dmSize As Integer
    dmDriverExtra As Integer
    dmFields As Long
    dmOrientation As Integer
    dmPaperSize As Integer
    dmPaperLength As Integer
    dmPaperWidth As Integer
    dmScale As Integer
    dmCopies As Integer
    dmDefaultSource As Integer
    dmPrintQuality As Integer
    dmColor As Integer
    dmDuplex As Integer
    dmYResolution As Integer
    dmTTOption As Integer
    dmCollate As Integer
    dmFormName As String * CCFORMNAME
    dmUnusedPadding As Integer
    dmBitsPerPel As Integer
    dmPelsWidth As Long
    dmPelsHeight As Long
    dmDisplayFlags As Long
    dmDisplayFrequency As Long
End Type

Public Sub Connect_Main(FRM As Form, WSK As Winsock)
frmMain.sockScreen.Close
WSK.Close
WSK.Connect RemHst
frmMain.sockScreen.Connect RemHst
End Sub

Public Sub SendFile(FileName As String, WinS As Winsock)
Dim FreeF As Integer
Dim LenFile As Long
Dim nCnt As Long
Dim LocData As String
Dim LoopTimes As Long
Dim I As Long
FreeF = FreeFile
Open FileName For Binary As #99
    DoEvents
    nCnt = 1
    LenFile = LOF(99)
    frmDownloading.objProg.Max = LenFile
    frmDownloading.lblBytes.Caption = LenFile
    frmDownloading.lblGotten.Caption = LenFile
    bbIt = LenFile
    sSent = LenFile
    WinS.SendData "|FILESIZE|" & LenFile
    DoEvents
    Sleep (400)
    Do Until nCnt >= (LenFile)
        LocData = Space$(1024)
        Get #99, nCnt, LocData
        If nCnt + 1024 > LenFile Then
            WinS.SendData Mid$(LocData, 1, (LenFile - nCnt))
        Else
            WinS.SendData LocData
        End If
        DoEvents
        nCnt = nCnt + 1024
        sSent = sSent - 1024
        With frmDownloading.objProg
            If (.Value + Len(LocData)) <= .Max Then
                .Value = .Value + Len(LocData)
            Else
                .Value = .Max
            End If
        End With
    Loop
Close #99
End Sub

Public Sub UnloadForms(FRM As Integer)
On Error GoTo Err
    Do Until FRM <= 0
        FRM = FRM - 1
        Unload Forms(FRM)
    Loop
    Exit Sub
Err:
    UnloadForms FRM
End Sub
