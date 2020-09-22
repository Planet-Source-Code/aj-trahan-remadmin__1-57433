Attribute VB_Name = "modImages"
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

Public FolderClick As String
Public Login As Boolean
Public About As Boolean
Public Sub Reset_Images(FRM As Form)
FRM.imgConnect.Visible = True
FRM.imgRemAdmin.Visible = True
FRM.imgServOpt.Visible = True
FRM.imgRemExpl.Visible = True
FRM.imgRemDesk.Visible = True
FRM.imgAbout.Visible = True
FRM.imgConnectSel.Visible = False
FRM.imgConnectSel.Top = 1800
FRM.imgRemAdminSel.Visible = False
FRM.imgRemAdminSel.Top = 1800
FRM.imgServOptSel.Visible = False
FRM.imgServOptSel.Top = 1800
FRM.imgRemExplSel.Visible = False
FRM.imgRemExplSel.Top = 1800
FRM.imgRemDeskSel.Visible = False
FRM.imgRemDeskSel.Top = 1800
FRM.imgAboutSel.Visible = False
FRM.imgAboutSel.Top = 1800

End Sub
Public Sub Disable_Images(FRM As Form)
Disable_imgMain FRM
Disable_imgRemoteAdmin FRM
Disable_RemoteExplorer FRM
Disable_ServerOptions FRM
Disable_RemDesk FRM
End Sub
Public Sub Disable_RemDesk(FRM As Form)
FRM.lblRemDesk.Visible = False
FRM.imgRemScreen.Visible = False
FRM.VScroll1.Visible = False
FRM.HScroll1.Visible = False
FRM.imgRE(6).Visible = False
FRM.imgRE(7).Visible = False
FRM.imgbar.Visible = False
End Sub
Public Sub Enable_RemDesk(FRM As Form)
If lLoggedIn = False Then
    Enable_imgMain frmMain
    Exit Sub
End If
If FRM.sockMain.State = sckConnected Then
    If dDesk = True Then
        FRM.imgRemScreen.Visible = True
        FRM.VScroll1.Visible = True
        FRM.HScroll1.Visible = True
    End If
End If
FRM.imgRE(6).Visible = True
FRM.imgRE(7).Visible = True
FRM.imgbar.Visible = True
End Sub
Public Sub Disable_imgMain(FRM As Form)
FRM.imgMain.Visible = False
For I = 0 To 15
    FRM.lblInfo(I).Visible = False
    If I <= 4 Then
        FRM.imgMainInfo(I).Visible = False
    End If
Next I
FRM.lstProcesses.Visible = False
End Sub
Public Sub Enable_imgMain(FRM As Form)
FRM.imgMain.Visible = True
For I = 0 To 15
    FRM.lblInfo(I).Visible = True
    If I <= 4 Then
        FRM.imgMainInfo(I).Visible = True
    End If
Next I
FRM.lstProcesses.Visible = True
End Sub
Public Sub Disable_ServerOptions(FRM As Form)
FRM.imgServerOptions.Visible = False
For I = 0 To 5
    FRM.imgSO(I).Visible = False
    FRM.chkSO(I).Visible = False
Next I
FRM.txtEmail.Visible = False
FRM.lstUsers.Visible = False
'FRM.SendNewSettings
End Sub
Public Sub Enable_ServerOptions(FRM As Form)
If lLoggedIn = False Then
    Enable_imgMain frmMain
    Exit Sub
End If
FRM.imgServerOptions.Visible = True
For I = 0 To 5
    FRM.imgSO(I).Visible = True
    FRM.chkSO(I).Visible = True
Next I
FRM.txtEmail.Visible = True
FRM.lstUsers.Visible = True
End Sub
Public Sub Disable_RemoteExplorer(FRM As Form)
FRM.imgRemoteExplorer.Visible = False
For I = 0 To 5
    FRM.imgRE(I).Visible = False
    If I <= 2 Then
        FRM.lblRE(I).Visible = False
    End If
Next I
FRM.TVTreeView.Visible = False
FRM.lvFiles.Visible = False
End Sub
Public Sub Enable_RemoteExplorer(FRM As Form)
Unload frmDisplay
If lLoggedIn = False Then
    Enable_imgMain frmMain
    Exit Sub
End If
FRM.imgRemoteExplorer.Visible = True
For I = 0 To 5
    FRM.imgRE(I).Visible = True
    If I <= 2 Then
        FRM.lblRE(I).Visible = True
    End If
Next I
FRM.TVTreeView.Visible = True
FRM.lvFiles.Visible = True
End Sub
Public Sub Disable_imgRemoteAdmin(FRM As Form)
FRM.imgRemoteAdmin.Visible = False
For I = 0 To 12
    FRM.imgRA(I).Visible = False
    If I <= 8 Then
        FRM.optRA(I).Visible = False
    End If
Next I
FRM.lblButton.Visible = False
FRM.Label1.Visible = False
FRM.imgRA(14).Visible = False
FRM.imgRAS(14).Visible = False
If iInverse = True Then FRM.imgRA(13).Visible = False
FRM.txtMSGTitle.Visible = False
FRM.txtMessage.Visible = False
FRM.txtURL.Visible = False
End Sub
Public Sub Enable_imgRemoteAdmin(FRM As Form)
If lLoggedIn = False Then
    Enable_imgMain frmMain
    Exit Sub
End If
FRM.imgRemoteAdmin.Visible = True
For I = 0 To 12
    FRM.imgRA(I).Visible = True
    If I <= 8 Then
        FRM.optRA(I).Visible = True
    End If
Next I
If iInverse = True Then
    FRM.imgRA(13).Top = 3840
    FRM.imgRAS(13).Top = 3840
    FRM.imgRA(13).Visible = True
    FRM.imgRA(3).Visible = False
End If
FRM.lblButton.Visible = True
FRM.Label1.Visible = True
FRM.imgRA(14).Visible = True
FRM.txtMSGTitle.Visible = True
FRM.txtMessage.Visible = True
FRM.txtURL.Visible = True
End Sub
