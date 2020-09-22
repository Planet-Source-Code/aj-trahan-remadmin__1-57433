Attribute VB_Name = "modVoice"
Option Explicit

' Application User Defined Types...
' Sound Format
Public Const WAVE_FORMAT_PCM = &H1                  ' Microsoft Windows PCM Wave Format
Public Const WAVE_FORMAT_ADPCM = &H11               ' ADPCM Wave Format
Public Const WAVE_FORMAT_IMA_ADPCM = &H11           ' IMA ADPCM Wave Format
Public Const WAVE_FORMAT_DVI_ADPCM = &H11           ' DVI ADPCM Wave Format
Public Const WAVE_FORMAT_DSPGROUP_TRUESPEECH = &H22 ' DSP Group Wave Format
Public Const WAVE_FORMAT_GSM610 = &H31              ' GSM610 Wave Format
Public Const WAVE_FORMAT_MSN_AUDIO = &H32           ' MSN Audio Wave Format

Public Const TIMESLICE = 0.2            ' Time Slicing 1/5 Second

' Application Constants...
Public Const NoOfRings = 1                  ' Number Of Times In/Out Bound Calls Ring...

Public Const phoneHungUp = 3                ' Hangup Status Icon...
Public Const phoneRingIng = 2               ' Ringing Status Icon...
Public Const phoneAnswered = 1              ' Answered Status Icon...
Public Const mikeNO = 6
Public Const mikeOFF = 7
Public Const mikeON = 8
Public Const speakNO = 9
Public Const speakOFF = 10
Public Const speakON = 11

Public Const RingInId = 101                 ' Ringing InBound Sound...
Public Const RingOutId = 102                ' Ringing OutBound Sound...

' Toolbar constants...
Public Const tbCALL = 2
Public Const tbHANGUP = 3
Public Const tbAUTOANSWER = 5

'== flag values for wFlags parameter ==================================
Public Const SND_SYNC = &H0                 '  play synchronously (default)
Public Const SND_ASYNC = &H1                '  play asynchronously
Public Const SND_NODEFAULT = &H2            '  don't use default sound
Public Const SND_MEMORY = &H4               '  lpszSoundName points to a memory file
Public Const SND_LOOP = &H8                 '  loop the sound until next sndPlaySound
Public Const SND_NOSTOP = &H10              '  don't stop any currently playing sound

'== MCI Wave API Declarations ================================================
Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal SoundData As Any, ByVal uFlags As Long) As Long
'== TCP Port Array Processing Const.s ===================================
Public Const MINTCP = 1                     ' Minimum index for tcpsocket control instance
Public Const MAXTCP = 32                    ' Maximum index for tcpsocket control instance
Public Const VOICEPORT = 6802               ' Voice chat Port To Listen On...
Public Const NULLPORTID = 0                 ' Null Port ID - A Port ID That Will Never Be Used...

Public Sub DebugSocket(TCPSocket As Winsock)
' Prints Information In A TCP Socket, For Debugging TCP Events...
    Debug.Print "TCPSocket.RemoteHost", TCPSocket.RemoteHost
    Debug.Print "TCPSocket.RemoteHostIP", TCPSocket.RemoteHostIP
    Debug.Print "TCPSocket.RemotePort", TCPSocket.RemotePort
    Debug.Print "TCPSocket.LocalHostName", TCPSocket.LocalHostName
    Debug.Print "TCPSocket.LocalIP", TCPSocket.LocalIP
    Debug.Print "TCPSocket.LocalPort", TCPSocket.LocalPort
    Debug.Print "TCPSocket.State", TCPSocket.State
    Debug.Print "====================================================="
End Sub

Public Sub ResPlaySound(ResourceId As Long)
' Uses Sound Play Sound To Play Back PreRecorded WaveFiles
    Dim sndBuff As String
    sndBuff = StrConv(LoadResData(ResourceId, "WAVE"), vbUnicode)
    Call sndPlaySound(sndBuff, SND_SYNC Or SND_MEMORY)
End Sub

Public Sub AddConnectionToList(Socket As Winsock, ConnList As ListBox)
' Adds A Connection Reference To A ListBox - [(Server)(LocalPort)(RemotePort)]
    Dim MemberID As String                      ' Connection Reference Variable
    ' Create MemberID From HostName and RemoteIP
    MemberID = Socket.RemoteHostIP & "  [" & _
               Format(Socket.RemotePort, "0") & "] - [" & _
               Format(Socket.LocalPort, "0") & "]"

    ConnList.AddItem MemberID                   ' Add New Member To List
    ConnList.ItemData(ConnList.NewIndex) = Socket.Index
End Sub

Public Sub RemoveConnectionFromList(Socket As Winsock, ConnList As ListBox)
' Removes A Connection Reference From A ListBox
    Dim Conn As Long                                ' Connection Array Element Variable
    Dim MemberID As String                          ' Connection Reference Variable
    ' Create MemberID From HostName and RemoteIP
    MemberID = Socket.RemoteHostIP & "  [" & _
               Format(Socket.RemotePort, "0") & "] - [" & _
               Format(Socket.LocalPort, "0") & "]"
    
    For Conn = 0 To ConnList.ListCount - 1          ' Search Each Member In List
        If (ConnList.List(Conn) = MemberID) Then    ' Look For MemberID In List
            ConnList.RemoveItem Conn                ' Remove MemberID From List
        End If
    Next                                            ' Next Connection
End Sub

Public Sub GetIdxFromMemberID(Sockets As Variant, MemberID As String, Index As Long)
    Dim Idx As Long                                 ' Socket cntl index
    Dim LocPortID As Long                           ' Local Port ID
    Dim RemPortID As Long                           ' Remote Port ID
    Dim RemoteIP As String                          ' Remote Host IP address
    Dim sStart As Long                              ' Substring begin position
    Dim sEnd As Long                                ' Substring end postition
    Dim Socket As Winsock                               ' Winsock socket
    sStart = 1
    sEnd = InStr(1, MemberID, " ") - 1              ' Get end of remote ip address
    If (sEnd > 1) Then
        RemoteIP = Mid(MemberID, sStart, sEnd)      ' Get remote host ip address
        sStart = InStr(sEnd, MemberID, "[") + 1     ' Get start of remote port
        If (sStart > 1) Then                        ' If Start found
            sEnd = InStr(sStart, MemberID, "]") - 1 ' Get end of remote port
            If (sEnd > 2) Then                      ' If end found
                RemPortID = Val(Mid(MemberID, sStart, sEnd)) ' Get RemotePort
                sStart = InStr(sEnd, MemberID, "[") + 1      ' Get start of local port
                If (sStart > 1) Then
                    sEnd = InStr(sStart, MemberID, "]") - 1  ' Get end of local port
                    If (sEnd > 2) Then                       ' If End Found
                        LocPortID = Val(Mid(MemberID, sStart, sEnd)) ' Extract local port
                        For Each Socket In Sockets
                            If ((Socket.RemoteHostIP = RemoteIP) And _
                                (Socket.RemotePort = RemPortID) And _
                                (Socket.LocalPort = LocPortID) And _
                                (Socket.Index > 0)) Then    ' Was a match found???
                                Index = Socket.Index        ' Save and return index
                                Exit Sub                    ' Done... exit
                            End If
                        Next
                    End If
                End If
            End If
        End If
    End If
End Sub

Public Function InstanceTCP(TCPArray As Variant) As Long
    Dim Ind As Long                                 ' Array Index Var...
    InstanceTCP = -1                                ' Set Default Value
    On Error GoTo InitControl                       ' IF Error Then Control Is Available
    
    For Ind = MINTCP To MAXTCP                      ' For Each Member In TCPArray() + 1
        If (TCPArray(Ind).Name = "") Then           ' If Control Is Not Valid Then..
        End If                                      ' ..A Runtime Error Will Occure
    Next                                            ' Search Next Item In Array
InitControl:                                        ' Initialize New Control
    On Error GoTo ErrorHandler                      ' Enable Error Handling...
            
    If ((Ind >= MINTCP) And (Ind <= MAXTCP)) Then   ' Check to make sure index value is with in range
        Load TCPArray(Ind)                          ' Create New Member In TCPArray
        InstanceTCP = Ind                           ' Return New TCPctl Index
    End If
    
    Exit Function                                   ' Exit
ErrorHandler:                                       ' Handler
    Debug.Print Err.Number, Err.Description         ' Debug Errors
    Resume Next                                     ' Ignore Error And Continue
End Function

Public Function Connect(Socket As Winsock, RemHost As String, RemPort As Long) As Boolean
' Attempts To Open A Socket Connection
    Connect = False                             ' Set default return code
    Call CloseListen(Socket)                    ' Stop Listening On LocalPort
    Socket.LocalPort = 0                        ' Not necessary, but done just in case
    Call Socket.Connect(RemHost, RemPort)       ' Connect To Server
    
    Do While ((Socket.State = sckConnecting) Or _
              (Socket.State = sckConnectionPending) Or _
              (Socket.State = sckResolvingHost) Or _
              (Socket.State = sckHostResolved) Or _
              (Socket.State = sckOpen))         ' Attempting To Connect...
        DoEvents                                ' Post Events
    Loop                                        ' Keep Waiting

    Connect = (Socket.State = sckConnected)     ' Did Socket Connect On Port...
End Function

Public Function Listen(Socket As Winsock) As Long
' Starts Listening For A Remote Connection Request On The LocalPort ID
    If (Socket.State <> sckListening) Then      ' Is Socket Already Listening
        If (Socket.LocalPort = 0) Then          ' If local port is not initialized then...
            Socket.LocalPort = VOICEPORT        ' Set standard application port
        End If
        Call Socket.Listen                      ' Listen On Local Port...
    End If
End Function

Public Function CloseListen(Socket As Winsock) As Long
' Stops Listening On A LocalPort For A Remote Connection Request...
    If (Socket.State = sckListening) Then       ' Is Socket Listening?
        Socket.Close                            ' Close Listen
    End If
End Function

Public Sub Disconnect(Socket As Winsock)
' Disconnects A Remote WinSock TCP/IP Connection If Connected
    If (Socket.State <> sckClosed) Then         ' Is Socket Already Closed?
        Socket.Close                            ' Close Socket
    End If
End Sub
