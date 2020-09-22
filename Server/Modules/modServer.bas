Attribute VB_Name = "modServer"
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

' for user password verification
Private Declare Function LogonUser Lib "Advapi32" Alias "LogonUserA" (ByVal lpszUsername As String, ByVal lpszDomain As Any, ByVal lpszPassword As String, ByVal dwLogonType As Long, ByVal dwLogonProvider As Long, phToken As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long
Private Declare Function NetUserChangePassword Lib "netapi32.dll" (ByVal sDomain As String, ByVal sUserName As String, ByVal sOldPassword As String, ByVal sNewPassword As String) As Long
'for entire desktop capture
Option Base 0

Private Type PALETTEENTRY
   peRed As Byte
   peGreen As Byte
   peBlue As Byte
   peFlags As Byte
End Type

Private Type LOGPALETTE
   palVersion As Integer
   palNumEntries As Integer
   palPalEntry(255) As PALETTEENTRY  ' Enough for 256 colors.
End Type

Private Type GUID
   Data1 As Long
   Data2 As Integer
   Data3 As Integer
   Data4(7) As Byte
End Type

Private Const RASTERCAPS As Long = 38
Private Const RC_PALETTE As Long = &H100
Private Const SIZEPALETTE As Long = 104

Private Type RECT
   Left As Long
   Top As Long
   Right As Long
   Bottom As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDc As Long, ByVal iCapabilitiy As Long) As Long
Private Declare Function GetSystemPaletteEntries Lib "gdi32" (ByVal hDc As Long, ByVal wStartIndex As Long, ByVal wNumEntries As Long, lpPaletteEntries As PALETTEENTRY) As Long
Private Declare Function CreatePalette Lib "gdi32" (lpLogPalette As LOGPALETTE) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDCDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hDCSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SelectPalette Lib "gdi32" (ByVal hDc As Long, ByVal hPalette As Long, ByVal bForceBackground As Long) As Long
Private Declare Function RealizePalette Lib "gdi32" (ByVal hDc As Long) As Long
Private Declare Function GetWindowDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long

Private Type PicBmp
   Size As Long
   Type As Long
   hBmp As Long
   hPal As Long
   Reserved As Long
End Type

Private Declare Function OleCreatePictureIndirect Lib "olepro32.dll" (PicDesc As PicBmp, RefIID As GUID, ByVal fPictureOwnsHandle As Long, IPic As IPicture) As Long
' for "pausing"
Public Declare Function GetTickCount Lib "kernel32" () As Long
' for form on top
#If Win32 Then
Public Const HWND_TOPMOST& = -1
#Else
Public Const HWND_TOPMOST& = -1
#End If 'WIN32
#If Win32 Then
    Public Const SWP_NOMOVE& = &H2
    Public Const SWP_NOSIZE& = &H1
#Else
    Public Const SWP_NOMOVE& = &H2
    Public Const SWP_NOSIZE& = &H1
#End If 'WIN32
#If Win32 Then
 Declare Function SetWindowPos& Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
#Else
 Declare Sub SetWindowPos Lib "User" (ByVal hwnd As Integer, ByVal hWndInsertAfter As Integer, ByVal x As Integer, ByVal y As Integer, ByVal cx As Integer, ByVal cy As Integer, ByVal wFlags As Integer)
#End If 'WIN32
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2
'for shut-down,logoff,reboot of PC
Private Type LUID
    UsedPart As Long
    IgnoredForNowHigh32BitPart As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount As Long
    TheLuid As LUID
    Attributes As Long
End Type
Private Const EWX_SHUTDOWN As Long = 1
Private Const EWX_FORCE As Long = 4
Private Const EWX_REBOOT = 2
Private Declare Function ExitWindowsEx Lib "user32" (ByVal _
    dwOptions As Long, ByVal dwReserved As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function OpenProcessToken Lib "Advapi32" (ByVal _
    ProcessHandle As Long, _
    ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "Advapi32" _
    Alias "LookupPrivilegeValueA" _
    (ByVal lpSystemName As String, ByVal lpName As String, lpLuid _
    As LUID) As Long
Private Declare Function AdjustTokenPrivileges Lib "Advapi32" _
    (ByVal TokenHandle As Long, _
    ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES _
    , ByVal BufferLength As Long, _
    PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
' strings used by the server
Public nFo As String
Public StrtUp As Long
Public DownLoad As Long
Public AutoLogin As Long
Public verify As Long
Public Email As Long
Public Hidden As Long
Public Addy As String
Public lLocked As Boolean
Public tTyping As Boolean
Public FRMz As Integer
Public cConnected As Boolean
Public cRemCons As Long
Public mMessageButton As String
Public ScreenSending As Boolean
Public Sub SendMsgBtn(Sock As Winsock)
    Dim SMB As String
    SMB = Encrypt(mMessageButton)
    Sock.SendData SMB
End Sub

Sub Pause(HowLong As Long)
    Dim u%, tick As Long
    tick = GetTickCount()
    Do
        u% = DoEvents
    Loop Until tick + HowLong < GetTickCount
End Sub
Public Sub I_AM_UNLOADING()
    Dim Wdir As String
    Dim nName As String
    nName = App.Path
    nName = nName & "\" & App.EXEName & ".exe"
    Wdir = Environ("windir")
    Wdir = Wdir & "\System\" & App.EXEName & ".exe"
    'FileCopy nName, Wdir
'need to finish the sub and add it to start-up in the registry or
'short-cut in the Start-Up folder

End Sub
Private Sub AdjustToken()
    Const TOKEN_ADJUST_PRIVILEGES = &H20
    Const TOKEN_QUERY = &H8
    Const SE_PRIVILEGE_ENABLED = &H2
    Dim hdlProcessHandle As Long
    Dim hdlTokenHandle As Long
    Dim tmpLuid As LUID
    Dim tkp As TOKEN_PRIVILEGES
    Dim tkpNewButIgnored As TOKEN_PRIVILEGES
    Dim lBufferNeeded As Long
    hdlProcessHandle = GetCurrentProcess()
    OpenProcessToken hdlProcessHandle, (TOKEN_ADJUST_PRIVILEGES Or _
    TOKEN_QUERY), hdlTokenHandle
    ' Get the LUID for shutdown privilege.
    LookupPrivilegeValue "", "SeShutdownPrivilege", tmpLuid
    tkp.PrivilegeCount = 1    ' One privilege to set
    tkp.TheLuid = tmpLuid
    tkp.Attributes = SE_PRIVILEGE_ENABLED
    'Enable the shutdown privilege in the access token of this process.
    AdjustTokenPrivileges hdlTokenHandle, False, _
        tkp, Len(tkpNewButIgnored), tkpNewButIgnored, lBufferNeeded
End Sub
Public Sub ReBooT()
    AdjustToken
    ExitWindowsEx (EWX_REBOOT), &HFFFF
End Sub
Public Sub ShutDown()
    AdjustToken
    ExitWindowsEx (EWX_SHUTDOWN), &HFFFF
End Sub
Public Sub LogOff()
    ExitWindowsEx EWX_FORCE Or EWX_LOGOFF, 0&
End Sub
' The next two functions are for user password verification
Public Function UserValidate(sUserName As String, sPassword As String, Optional sDomain As String) As Boolean
    Dim lReturn As Long
    Const NERR_BASE = 2100
    Const NERR_PasswordCantChange = NERR_BASE + 143
    Const NERR_PasswordHistConflict = NERR_BASE + 144
    Const NERR_PasswordTooShort = NERR_BASE + 145
    Const NERR_PasswordTooRecent = NERR_BASE + 146
    
    If Len(sDomain) = 0 Then
        sDomain = Environ$("USERDOMAIN")
    End If
    
    'Call API to check password.
    lReturn = NetUserChangePassword(StrConv(sDomain, vbUnicode), StrConv(sUserName, vbUnicode), StrConv(sPassword, vbUnicode), StrConv(sPassword, vbUnicode))
    
    'Test return value.
    Select Case lReturn
    Case 0, NERR_PasswordCantChange, NERR_PasswordHistConflict, NERR_PasswordTooShort, NERR_PasswordTooRecent
        UserValidate = True
    Case Else
        UserValidate = False
    End Select
End Function
Public Function DLLErrorText(ByVal lLastDLLError As Long) As String
    Dim sBuff As String * 256
    Dim lCount As Long
    Const FORMAT_MESSAGE_ALLOCATE_BUFFER = &H100, FORMAT_MESSAGE_ARGUMENT_ARRAY = &H2000
    Const FORMAT_MESSAGE_FROM_HMODULE = &H800, FORMAT_MESSAGE_FROM_STRING = &H400
    Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000, FORMAT_MESSAGE_IGNORE_INSERTS = &H200
    Const FORMAT_MESSAGE_MAX_WIDTH_MASK = &HFF
    
    lCount = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM Or FORMAT_MESSAGE_IGNORE_INSERTS, 0, lLastDLLError, 0&, sBuff, Len(sBuff), ByVal 0)
    If lCount Then
        DLLErrorText = Left$(sBuff, lCount - 2)    'Remove line feeds
    End If
End Function
Public Function CaptureScreen() As Picture
    Dim hWndScreen As Long
    hWndScreen = GetDesktopWindow()
    Set CaptureScreen = CaptureWindow(hWndScreen, False, 0, 0, Screen.Width \ Screen.TwipsPerPixelX, Screen.Height \ Screen.TwipsPerPixelY)
End Function
Public Function CaptureWindow(ByVal hWndSrc As Long, ByVal Client As Boolean, ByVal LeftSrc As Long, ByVal TopSrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long) As Picture
    Dim hDCMemory As Long
    Dim hBmp As Long
    Dim hBmpPrev As Long
    Dim r As Long
    Dim hDCSrc As Long
    Dim hPal As Long
    Dim hPalPrev As Long
    Dim RasterCapsScrn As Long
    Dim HasPaletteScrn As Long
    Dim PaletteSizeScrn As Long
    Dim LogPal As LOGPALETTE
  
    If Client Then
        hDCSrc = GetDC(hWndSrc)
    Else
        hDCSrc = GetWindowDC(hWndSrc)
    End If

    hDCMemory = CreateCompatibleDC(hDCSrc)
    hBmp = CreateCompatibleBitmap(hDCSrc, WidthSrc, HeightSrc)
    hBmpPrev = SelectObject(hDCMemory, hBmp)

    RasterCapsScrn = GetDeviceCaps(hDCSrc, RASTERCAPS)
    HasPaletteScrn = RasterCapsScrn And RC_PALETTE
    PaletteSizeScrn = GetDeviceCaps(hDCSrc, SIZEPALETTE)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        LogPal.palVersion = &H300
        LogPal.palNumEntries = 256
        r = GetSystemPaletteEntries(hDCSrc, 0, 256, LogPal.palPalEntry(0))
        hPal = CreatePalette(LogPal)
        hPalPrev = SelectPalette(hDCMemory, hPal, 0)
        r = RealizePalette(hDCMemory)
    End If
    r = BitBlt(hDCMemory, 0, 0, WidthSrc, HeightSrc, hDCSrc, LeftSrc, TopSrc, vbSrcCopy)
    hBmp = SelectObject(hDCMemory, hBmpPrev)
    If HasPaletteScrn And (PaletteSizeScrn = 256) Then
        hPal = SelectPalette(hDCMemory, hPalPrev, 0)
    End If
    r = DeleteDC(hDCMemory)
    r = ReleaseDC(hWndSrc, hDCSrc)
    Set CaptureWindow = CreateBitmapPicture(hBmp, hPal)
End Function
Public Function CreateBitmapPicture(ByVal hBmp As Long, ByVal hPal As Long) As Picture
    Dim r As Long
    Dim Pic As PicBmp
    Dim IPic As IPicture
    Dim IID_IDispatch As GUID
    
    With IID_IDispatch
        .Data1 = &H20400
        .Data4(0) = &HC0
        .Data4(7) = &H46
    End With
    
    With Pic
        .Size = Len(Pic)
        .Type = vbPicTypeBitmap
        .hBmp = hBmp
        .hPal = hPal
    End With

    r = OleCreatePictureIndirect(Pic, IID_IDispatch, 1, IPic)
    Set CreateBitmapPicture = IPic
End Function
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
