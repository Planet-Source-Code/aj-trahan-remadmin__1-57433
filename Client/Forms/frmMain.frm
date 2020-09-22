VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "SubNet®"
   ClientHeight    =   9165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0CCE
   ScaleHeight     =   9165
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSWinsockLib.Winsock sockScreen 
      Left            =   1560
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6804
   End
   Begin VB.VScrollBar VScroll1 
      Height          =   6255
      Left            =   10720
      TabIndex        =   48
      Top             =   2280
      Width           =   255
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   120
      TabIndex        =   47
      Top             =   8550
      Width           =   10575
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   3240
      Top             =   7440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14AC7C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14B556
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picRemDesk2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   1200
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   46
      Top             =   9240
      Width           =   375
   End
   Begin VB.PictureBox picRemDesk 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   375
      Left            =   360
      ScaleHeight     =   315
      ScaleWidth      =   435
      TabIndex        =   45
      Top             =   9240
      Width           =   495
   End
   Begin VB.Timer tmrMessage 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   2520
      Top             =   240
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   4440
      Top             =   240
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   5
      Left            =   6480
      TabIndex        =   43
      Top             =   3360
      Width           =   255
   End
   Begin VB.Timer tmrRefreshProcesses 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   3000
      Top             =   240
   End
   Begin VB.Timer tmrConnect 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   3480
      Top             =   240
   End
   Begin MSWinsockLib.Winsock sockProcesses 
      Left            =   2040
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6805
   End
   Begin MSWinsockLib.Winsock sockChat 
      Left            =   1080
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6803
   End
   Begin MSWinsockLib.Winsock sockExplorer 
      Left            =   600
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6801
   End
   Begin MSWinsockLib.Winsock sockMain 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemotePort      =   6800
   End
   Begin VB.ListBox lstProcesses 
      Appearance      =   0  'Flat
      BackColor       =   &H005F8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5070
      Left            =   10680
      TabIndex        =   42
      Top             =   3240
      Width           =   4935
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   4
      Left            =   8040
      TabIndex        =   25
      Top             =   4365
      Width           =   255
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   3
      Left            =   2400
      TabIndex        =   24
      Top             =   3240
      Width           =   255
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   2
      Left            =   2760
      TabIndex        =   23
      Top             =   3000
      Width           =   255
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   1
      Left            =   3120
      TabIndex        =   22
      Top             =   2700
      Width           =   255
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6000
      TabIndex        =   21
      Text            =   "usssssy@yahoo.com"
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CheckBox chkSO 
      BackColor       =   &H00404040&
      Caption         =   "Check1"
      Height          =   255
      Index           =   0
      Left            =   4320
      TabIndex        =   20
      Top             =   2450
      Width           =   255
   End
   Begin VB.ListBox lstUsers 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   2010
      Left            =   10680
      TabIndex        =   19
      Top             =   2400
      Width           =   4815
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   8
      Left            =   9720
      TabIndex        =   12
      Top             =   4440
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   7
      Left            =   10200
      TabIndex        =   11
      Top             =   4200
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   6
      Left            =   9360
      TabIndex        =   10
      Top             =   3960
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   5
      Left            =   9960
      TabIndex        =   9
      Top             =   3720
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   4
      Left            =   7320
      TabIndex        =   8
      Top             =   4550
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   3
      Left            =   8195
      TabIndex        =   7
      Top             =   4300
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   2
      Left            =   7320
      TabIndex        =   6
      Top             =   4030
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   1
      Left            =   7320
      TabIndex        =   5
      Top             =   3775
      Width           =   255
   End
   Begin VB.OptionButton optRA 
      BackColor       =   &H00404040&
      Height          =   255
      Index           =   0
      Left            =   7200
      TabIndex        =   4
      Top             =   3480
      Value           =   -1  'True
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Height          =   375
      Left            =   3960
      ScaleHeight     =   315
      ScaleWidth      =   315
      TabIndex        =   1
      Top             =   240
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3000
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14BE30
            Key             =   "CD"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C182
            Key             =   "HD"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C4D4
            Key             =   "ND"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14C7F5
            Key             =   "RC2"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CB36
            Key             =   "CLOSED"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14CE88
            Key             =   "OPEN"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D1DA
            Key             =   "FD"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   3600
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   62
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D52C
            Key             =   "FILE"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14D87E
            Key             =   "MDB"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DBD0
            Key             =   "HDI"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14DF22
            Key             =   "UDL"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E274
            Key             =   "ARX"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E5C6
            Key             =   "XMX"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14E918
            Key             =   "DWG"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14EC6A
            Key             =   "M3U"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14EFBC
            Key             =   "SEU"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F30E
            Key             =   "VB"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F660
            Key             =   "FRM"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14F9B2
            Key             =   "CTL"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":14FD04
            Key             =   "BAS"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150056
            Key             =   "GID"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1503A8
            Key             =   "XFM"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1506FA
            Key             =   "CRT"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150A4C
            Key             =   "URL"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":150D9E
            Key             =   "ASP"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1510F0
            Key             =   "SWF"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151442
            Key             =   "CAT"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151794
            Key             =   "SCR"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151AE6
            Key             =   "CHM"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":151E38
            Key             =   "POT"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15218A
            Key             =   "XLA"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1524DC
            Key             =   "DOT"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15282E
            Key             =   "CLS"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152B80
            Key             =   "JS"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":152ED2
            Key             =   "XLS"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153224
            Key             =   "DBX"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153576
            Key             =   "PPT"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1538C8
            Key             =   "PPA"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153C1A
            Key             =   "REG"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":153F6C
            Key             =   "HTT"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1542BE
            Key             =   "FNT"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":154610
            Key             =   "XML"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":154962
            Key             =   "SC"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":154CB4
            Key             =   "EXE"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155006
            Key             =   "HLP"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155358
            Key             =   "BMP"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1556AA
            Key             =   "PNT"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1559FC
            Key             =   "ANI"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":155D4E
            Key             =   "DLL"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1560A0
            Key             =   "IE"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1563F2
            Key             =   "RAR"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15673A
            Key             =   "RTF"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":156A8C
            Key             =   "DOC"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":156DDE
            Key             =   "TXT"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157130
            Key             =   "INI"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157482
            Key             =   "WAV"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1577D4
            Key             =   "ZIP"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157B26
            Key             =   "PDF"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":157F15
            Key             =   "FOLDER"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158267
            Key             =   "JBF"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1585B9
            Key             =   "PRC"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15890B
            Key             =   "FRX"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158C5D
            Key             =   "DSP"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":158FAF
            Key             =   "C++"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":159301
            Key             =   "H++"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":159653
            Key             =   "RC"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1599A5
            Key             =   "CPP"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":159CF7
            Key             =   "APP"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":15A05E
            Key             =   "ICO"
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtURL 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   6840
      TabIndex        =   13
      Text            =   "http://planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=57433&lngWId=1"
      Top             =   6240
      Width           =   3855
   End
   Begin VB.TextBox txtMessage 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   735
      Left            =   7680
      MaxLength       =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Text            =   "frmMain.frx":15A3C1
      Top             =   2760
      Width           =   3015
   End
   Begin VB.TextBox txtMSGTitle 
      Appearance      =   0  'Flat
      BackColor       =   &H005D8071&
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   7680
      MaxLength       =   28
      TabIndex        =   2
      Text            =   "SubNet Administrator"
      Top             =   2385
      Width           =   3015
   End
   Begin MSComctlLib.TreeView TVTreeView 
      Height          =   5475
      Left            =   10560
      TabIndex        =   14
      Top             =   2160
      Width           =   2970
      _ExtentX        =   5239
      _ExtentY        =   9657
      _Version        =   393217
      Indentation     =   88
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "BankGothic Lt BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ListView lvFiles 
      Height          =   5520
      Left            =   10680
      TabIndex        =   15
      Top             =   1800
      Width           =   7230
      _ExtentX        =   12753
      _ExtentY        =   9737
      View            =   2
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList2"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "BankGothic Lt BT"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Image imgBar 
      Height          =   75
      Left            =   0
      Picture         =   "frmMain.frx":15A3F5
      Top             =   2185
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.Label lblButton 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   255
      Left            =   7920
      TabIndex        =   50
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Remote Button Pushed:"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   5760
      TabIndex        =   49
      Top             =   4920
      Width           =   2175
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   14
      Left            =   4080
      Picture         =   "frmMain.frx":15CF6B
      Top             =   7900
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   14
      Left            =   4080
      Picture         =   "frmMain.frx":15E7A5
      Top             =   7900
      Width           =   1170
   End
   Begin VB.Image imgRemScreen 
      Height          =   6210
      Left            =   120
      MouseIcon       =   "frmMain.frx":15FFDF
      Stretch         =   -1  'True
      Top             =   2280
      Width           =   10575
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   5
      Left            =   6000
      Picture         =   "frmMain.frx":1608A9
      Top             =   2760
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   5
      Left            =   6000
      Picture         =   "frmMain.frx":1620E3
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   13
      Left            =   4080
      Picture         =   "frmMain.frx":16391D
      Top             =   4200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgMainInfoSel 
      Height          =   285
      Index           =   0
      Left            =   8665
      Picture         =   "frmMain.frx":165157
      Top             =   7530
      Visible         =   0   'False
      Width           =   1950
   End
   Begin VB.Image imgDisconnectSel 
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":168741
      Top             =   2160
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgDisconnect 
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":16DD5D
      Top             =   560
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgConnectSel 
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":172AA5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgMainInfo 
      Height          =   495
      Index           =   4
      Left            =   7670
      Picture         =   "frmMain.frx":178343
      Top             =   7860
      Width           =   495
   End
   Begin VB.Image imgMainInfo 
      Height          =   495
      Index           =   3
      Left            =   6720
      Picture         =   "frmMain.frx":17B1F8
      Top             =   7860
      Width           =   495
   End
   Begin VB.Image imgMainInfo 
      Height          =   495
      Index           =   2
      Left            =   5760
      Picture         =   "frmMain.frx":17DFBE
      Top             =   7860
      Width           =   495
   End
   Begin VB.Image imgMainInfo 
      Height          =   285
      Index           =   0
      Left            =   8680
      Picture         =   "frmMain.frx":180C2E
      Top             =   7515
      Width           =   1935
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6804"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   15
      Left            =   3680
      TabIndex        =   41
      Top             =   8040
      Width           =   975
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6803"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   14
      Left            =   3600
      TabIndex        =   40
      Top             =   7720
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6802"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   13
      Left            =   3600
      TabIndex        =   39
      Top             =   7395
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6801"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   12
      Left            =   3600
      TabIndex        =   38
      Top             =   7080
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "6800"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   11
      Left            =   3600
      TabIndex        =   37
      Top             =   6800
      Width           =   1095
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   10
      Left            =   2040
      TabIndex        =   36
      Top             =   8040
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   9
      Left            =   2040
      TabIndex        =   35
      Top             =   7720
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   8
      Left            =   2040
      TabIndex        =   34
      Top             =   7395
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   7
      Left            =   2040
      TabIndex        =   33
      Top             =   7080
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   6
      Left            =   2040
      TabIndex        =   32
      Top             =   6800
      Width           =   1335
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   4
      Left            =   1080
      TabIndex        =   30
      Top             =   5180
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   3
      Left            =   1080
      TabIndex        =   29
      Top             =   4605
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   1080
      TabIndex        =   28
      Top             =   4080
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H005D8071&
      Height          =   255
      Index           =   1
      Left            =   1080
      TabIndex        =   27
      Top             =   3525
      Width           =   3255
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   26
      Top             =   3015
      Width           =   3255
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   4
      Left            =   7320
      Picture         =   "frmMain.frx":184201
      Top             =   3240
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":187C2C
      Top             =   2760
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   2
      Left            =   7320
      Picture         =   "frmMain.frx":18B820
      Top             =   2280
      Visible         =   0   'False
      Width           =   2085
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   1
      Left            =   4080
      Picture         =   "frmMain.frx":18F43F
      Top             =   6960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgSOS 
      Height          =   390
      Index           =   0
      Left            =   2760
      Picture         =   "frmMain.frx":192216
      Top             =   6960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   4
      Left            =   7320
      Picture         =   "frmMain.frx":1951A3
      Top             =   3240
      Width           =   2085
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   3
      Left            =   7320
      Picture         =   "frmMain.frx":19888B
      Top             =   2760
      Width           =   2085
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   2
      Left            =   7320
      Picture         =   "frmMain.frx":19C0CF
      Top             =   2280
      Width           =   2085
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   1
      Left            =   4080
      Picture         =   "frmMain.frx":19F8EF
      Top             =   6960
      Width           =   1170
   End
   Begin VB.Image imgSO 
      Height          =   390
      Index           =   0
      Left            =   2760
      Picture         =   "frmMain.frx":1A24D3
      Top             =   6960
      Width           =   1170
   End
   Begin VB.Label lblRE 
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   2
      Left            =   5160
      TabIndex        =   18
      Top             =   2025
      Width           =   2775
   End
   Begin VB.Label lblRE 
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   1
      Left            =   2280
      TabIndex        =   17
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label lblRE 
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   16
      Top             =   1810
      Width           =   6855
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   7
      Left            =   5880
      Picture         =   "frmMain.frx":1A5211
      Top             =   1780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   6
      Left            =   3960
      Picture         =   "frmMain.frx":1A7DB5
      Top             =   1780
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   5
      Left            =   5520
      Picture         =   "frmMain.frx":1AAA90
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   4
      Left            =   4200
      Picture         =   "frmMain.frx":1ADA49
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   3
      Left            =   6840
      Picture         =   "frmMain.frx":1B085C
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   2
      Left            =   8160
      Picture         =   "frmMain.frx":1B373C
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   1
      Left            =   1800
      Picture         =   "frmMain.frx":1B650E
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRES 
      Height          =   390
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":1B92E0
      Top             =   8160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   7
      Left            =   5880
      Picture         =   "frmMain.frx":1BC0C7
      Top             =   1780
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   6
      Left            =   3960
      Picture         =   "frmMain.frx":1BEB39
      Top             =   1780
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   5
      Left            =   5520
      Picture         =   "frmMain.frx":1C16B2
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   4
      Left            =   4200
      Picture         =   "frmMain.frx":1C43E1
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   3
      Left            =   6840
      Picture         =   "frmMain.frx":1C700F
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   2
      Left            =   8160
      Picture         =   "frmMain.frx":1C9CD8
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   1
      Left            =   1800
      Picture         =   "frmMain.frx":1CC8A5
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgRE 
      Height          =   390
      Index           =   0
      Left            =   480
      Picture         =   "frmMain.frx":1CF472
      Top             =   8160
      Width           =   1170
   End
   Begin VB.Image imgEXIT 
      Height          =   120
      Left            =   10800
      Picture         =   "frmMain.frx":1D204C
      Top             =   45
      Width           =   120
   End
   Begin VB.Image imgAbout 
      Height          =   1515
      Left            =   9360
      Picture         =   "frmMain.frx":1D22F3
      Top             =   180
      Width           =   1425
   End
   Begin VB.Image imgRemDesk 
      Height          =   1545
      Left            =   7440
      Picture         =   "frmMain.frx":1D94D5
      Top             =   165
      Width           =   1755
   End
   Begin VB.Image imgServOpt 
      Height          =   1515
      Left            =   5640
      Picture         =   "frmMain.frx":1E22B7
      Top             =   180
      Width           =   1680
   End
   Begin VB.Image imgRemExpl 
      Height          =   1515
      Left            =   3675
      Picture         =   "frmMain.frx":1EA789
      Top             =   180
      Width           =   1860
   End
   Begin VB.Image imgRemAdmin 
      Height          =   1515
      Left            =   2040
      Picture         =   "frmMain.frx":1F3A8F
      Top             =   180
      Width           =   1500
   End
   Begin VB.Image imgConnect 
      Height          =   975
      Left            =   120
      Picture         =   "frmMain.frx":1FB12D
      Top             =   560
      Width           =   1740
   End
   Begin VB.Label lblStatus 
      BackStyle       =   0  'Transparent
      Caption         =   " Status:  Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   8880
      Width           =   6015
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   7
      Left            =   9360
      Picture         =   "frmMain.frx":2009CB
      Top             =   4800
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   7
      Left            =   9360
      Picture         =   "frmMain.frx":2038AB
      Top             =   4800
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   8
      Left            =   9360
      Picture         =   "frmMain.frx":206574
      Top             =   5725
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   8
      Left            =   9360
      Picture         =   "frmMain.frx":209110
      Top             =   5725
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   5
      Left            =   4080
      Picture         =   "frmMain.frx":20BB85
      Top             =   5880
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   5
      Left            =   4080
      Picture         =   "frmMain.frx":20EA26
      Top             =   5880
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   4
      Left            =   4080
      Picture         =   "frmMain.frx":211737
      Top             =   4885
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   4
      Left            =   4080
      Picture         =   "frmMain.frx":2143D7
      Top             =   4885
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   3
      Left            =   4080
      Picture         =   "frmMain.frx":216F58
      Top             =   3880
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   3
      Left            =   4080
      Picture         =   "frmMain.frx":219E0B
      Top             =   3880
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   11
      Left            =   2880
      Picture         =   "frmMain.frx":21CAB1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   11
      Left            =   2880
      Picture         =   "frmMain.frx":21F882
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   10
      Left            =   3600
      Picture         =   "frmMain.frx":222446
      Top             =   2160
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   10
      Left            =   3600
      Picture         =   "frmMain.frx":224FEA
      Top             =   2160
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":227A47
      Top             =   3960
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   2
      Left            =   600
      Picture         =   "frmMain.frx":22A8EE
      Top             =   3960
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   1
      Left            =   600
      Picture         =   "frmMain.frx":22D572
      Top             =   3360
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   1
      Left            =   600
      Picture         =   "frmMain.frx":230106
      Top             =   3360
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   0
      Left            =   600
      Picture         =   "frmMain.frx":232B80
      Top             =   2760
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   0
      Left            =   600
      Picture         =   "frmMain.frx":235752
      Top             =   2760
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   6
      Left            =   4080
      Picture         =   "frmMain.frx":2381C6
      Top             =   6880
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   6
      Left            =   4080
      Picture         =   "frmMain.frx":23AEA1
      Top             =   6880
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   390
      Index           =   9
      Left            =   9360
      Picture         =   "frmMain.frx":23DA1A
      Top             =   7200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   9
      Left            =   9360
      Picture         =   "frmMain.frx":2406F5
      Top             =   7200
      Width           =   1170
   End
   Begin VB.Image imgRAS 
      Height          =   615
      Index           =   12
      Left            =   9815
      Picture         =   "frmMain.frx":24326E
      Top             =   7890
      Visible         =   0   'False
      Width           =   630
   End
   Begin VB.Image imgRA 
      Height          =   480
      Index           =   12
      Left            =   9920
      Picture         =   "frmMain.frx":245D98
      Top             =   7985
      Width           =   465
   End
   Begin VB.Image imgRemAdminSel 
      Height          =   1425
      Left            =   1875
      Picture         =   "frmMain.frx":24845F
      Top             =   1800
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Image imgRemExplSel 
      Height          =   1365
      Left            =   3675
      Picture         =   "frmMain.frx":250A39
      Top             =   1800
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Image imgServOptSel 
      Height          =   1410
      Left            =   5520
      Picture         =   "frmMain.frx":258BDF
      Top             =   1800
      Visible         =   0   'False
      Width           =   1830
   End
   Begin VB.Image imgRemDeskSel 
      Height          =   1395
      Left            =   7440
      Picture         =   "frmMain.frx":261341
      Top             =   1800
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.Image imgAboutSel 
      Height          =   1290
      Left            =   9240
      Picture         =   "frmMain.frx":26964B
      Top             =   1800
      Visible         =   0   'False
      Width           =   1740
   End
   Begin VB.Image imgRA 
      Height          =   390
      Index           =   13
      Left            =   4080
      Picture         =   "frmMain.frx":270B75
      Top             =   4200
      Visible         =   0   'False
      Width           =   1170
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Pending.."
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Index           =   5
      Left            =   1080
      TabIndex        =   31
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Image imgMainInfo 
      Height          =   480
      Index           =   1
      Left            =   9910
      Picture         =   "frmMain.frx":2723AF
      Top             =   8040
      Width           =   465
   End
   Begin VB.Label lblRemDesk 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "This Option Is only Available With Subnet-Professional®"
      BeginProperty Font 
         Name            =   "BankGothic Lt BT"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   1095
      Left            =   3600
      TabIndex        =   44
      Top             =   4560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.Image imgRemoteAdmin 
      Height          =   6975
      Left            =   0
      Picture         =   "frmMain.frx":274A76
      Top             =   1755
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.Image imgRemoteExplorer 
      Height          =   7065
      Left            =   0
      Picture         =   "frmMain.frx":36FC9C
      Top             =   1745
      Visible         =   0   'False
      Width           =   11055
   End
   Begin VB.Image imgMain 
      Height          =   7440
      Left            =   0
      Picture         =   "frmMain.frx":46E29A
      Top             =   1740
      Width           =   11055
   End
   Begin VB.Image imgServerOptions 
      Height          =   6135
      Left            =   120
      Picture         =   "frmMain.frx":57A09C
      Top             =   1800
      Visible         =   0   'False
      Width           =   9780
   End
End
Attribute VB_Name = "frmMain"
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

Dim StartUp As Long
Dim FileDL As Long
Dim AutoLogin As Long
Dim VerifyLogin As Long
Dim AutoEmail As Long
Dim Addy As String
Dim Hidden As Long
Dim FCount As Integer
Dim LVFileCount As Long
Dim FList As String
Dim Data As String
Dim Msg As String
Dim bGettingdesktop As Boolean

' **** for icon in sys tray ****
Private Type NOTIFYICONDATA
    cbSize As Long
    hwnd As Long
    uId As Long
    uFlags As Long
    ucallbackMessage As Long
    hIcon As Long
    szTip As String * 64
End Type
Private Const NIM_ADD = &H0
Private Const NIM_MODIFY = &H1
Private Const NIM_DELETE = &H2
Private Const WM_MOUSEMOVE = &H200
Private Const NIF_MESSAGE = &H1
Private Const NIF_ICON = &H2
Private Const NIF_TIP = &H4
Private Const WM_LBUTTONDBLCLK = &H203
Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_RBUTTONDBLCLK = &H206
Private Const WM_RBUTTONDOWN = &H204
Private Const WM_RBUTTONUP = &H205
Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Dim T As NOTIFYICONDATA


Private Sub Command2_Click()
tmrRefrScreen.Enabled = False
End Sub

Private Sub Form_Load()
TVTreeView.Top = 2345
TVTreeView.Left = 290
lvFiles.Top = 2313
lvFiles.Left = 3450
If txtMSGTitle.Enabled = True Then
    txtMSGTitle.BackColor = &H5D8071
End If
FList = "|FILES|"
RemHst = "127.0.0.1"
User = "Administrator"
imgConnectSel.Top = imgConnect.Top
imgConnectSel.Visible = True
Disable_Images frmMain
lstUsers.Left = 480
lstUsers.Top = 4680
lstProcesses.Top = 2280
lstProcesses.Left = 5520
lstProcesses.AddItem "Pending.."
lstUsers.AddItem "Pending.."
For I = 0 To 10
    lblInfo(I).ForeColor = lstProcesses.BackColor
Next I
AddIcon2Tray
Load frmLogin
rReady = True
End Sub
Public Sub AddIcon2Tray()
    T.cbSize = Len(T)
    T.hwnd = Picture1.hwnd
    T.uId = 1&
    T.uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
    T.ucallbackMessage = WM_MOUSEMOVE
          T.hIcon = Me.Icon
          T.szTip = "SubNet® - ©All Rights Reserved 2004-2005" & Chr$(0)
    Shell_NotifyIcon NIM_ADD, T
End Sub
Public Sub RemoveIconFromTray()
On Error Resume Next
    T.cbSize = Len(T)
    T.hwnd = Picture1.hwnd
    T.uId = 1&
    Shell_NotifyIcon NIM_DELETE, T
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    SendMessage Me.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Kill App.Path & "\Desktop.bmp"
Kill App.Path & "\Screen.bmp"
RemoveIconFromTray
UnloadForms FRM
End
End Sub

Private Sub Form_Terminate()
On Error Resume Next
Kill App.Path & "\Desktop.bmp"
Kill App.Path & "\Screen.bmp"
RemoveIconFromTray
UnloadForms FRM
End
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Kill App.Path & "\Desktop.bmp"
Kill App.Path & "\Screen.bmp"
RemoveIconFromTray
UnloadForms FRM
End
End Sub

Private Sub HScroll1_Change()
    picRemDesk2.PaintPicture picRemDesk.Picture, 0, 0, _
    picRemDesk2.Width, picRemDesk2.Height, _
    HScroll1.Value, VScroll1.Value, _
    picRemDesk2.Width, picRemDesk2.Height
    Refresh_Desk
End Sub

Private Sub HScroll1_Scroll()
    picRemDesk2.PaintPicture picRemDesk.Picture, 0, 0, _
    picRemDesk2.Width, picRemDesk2.Height, _
    HScroll1.Value, VScroll1.Value, _
    picRemDesk2.Width, picRemDesk2.Height
    Refresh_Desk
End Sub

Private Sub Refresh_Desk()
imgRemScreen.Picture = picRemDesk2.Image
End Sub
Private Sub imgAbout_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgConnectSel.Visible = True Then
    frmLogin.Show
    Exit Sub
End If
imgAboutSel.Top = imgAbout.Top + 30
imgAboutSel.Visible = True
imgAbout.Visible = False
frmAbout.Show modal, frmMain
End Sub

Private Sub imgCloseApp_Click()
Unload Me
'End
End Sub


Private Sub imgCloseMe_Click()
Unload Me
'End
End Sub

Private Sub imgCloseMe2_Click()
'Unload frmAbout
'Unload frmLogin
Unload Me
'End
End Sub

Private Sub imgAboutSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmMain.imgAbout.Visible = True
frmMain.imgAboutSel.Visible = False
frmMain.imgAboutSel.Top = 1800
frmAbout.DoTheStuff
frmAbout.Show modal, frmMain
End Sub

Private Sub imgConnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
imgConnectSel.Top = imgConnect.Top
imgConnectSel.Visible = True
Load frmLogin
End Sub

Private Sub imgConnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
frmLogin.Show modal, frmMain
End Sub

Private Sub imgConnectSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
frmLogin.Show modal, frmMain
frmLogin.AlphaTrans.Enabled = True
End Sub

Private Sub imgDisconnect_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgDisconnectSel.Top = imgDisconnect.Top
imgDisconnectSel.Visible = True
imgDisconnect.Visible = False
lLoggedIn = False
End Sub
Public Sub Login_Cancel()
TVTreeView.Nodes.Clear
lvFiles.ListItems.Clear
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
imgDisconnectSel.Visible = False
imgConnect.Visible = True
imgDisconnect.Visible = False
Close_Connections
For I = 0 To 10
    lblInfo(I).Caption = "Pending.."
Next I
lstProcesses.Clear
lstProcesses.AddItem "Pending.."
lblRE(0).Caption = "Pending.."
lblRE(2).Caption = "Pending.."
lblStatus.Caption = " Status:  Disconnected"
tmrRefreshProcesses.Enabled = False
Unload frmChat
sockChat.Close
End Sub
Private Sub imgDisconnect_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
tmrConnect.Enabled = False
TVTreeView.Nodes.Clear
lvFiles.ListItems.Clear
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
imgDisconnectSel.Visible = False
imgConnect.Visible = True
imgDisconnect.Visible = False
Close_Connections
For I = 0 To 10
    lblInfo(I).Caption = "Pending.."
Next I
lstProcesses.Clear
lstProcesses.AddItem "Pending.."
lblRE(0).Caption = "Pending.."
lblRE(2).Caption = "Pending.."
lblStatus.Caption = " Status:  Disconnected"
tmrRefreshProcesses.Enabled = False
Unload frmChat
Unload frmDisplay
sockChat.Close
dDesk = False
End Sub
Private Sub Close_Connections()
sockMain.Close
sockExplorer.Close
sockProcesses.Close
sockChat.Close
lblStatus.Caption = " Status:  Disconnected From Remote Server"
End Sub
Private Sub imgEXIT_Click()
'Unload frmAbout
'Unload frmLogin
'Unload frmChat
'Unload frmWarning
Unload Me
'End
End Sub

Private Sub imgMainInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
On Error Resume Next
If sockMain.State <> sckConnected Then
    GoTo Closer
Else
    If Index = 0 Then
        If pProcess = False Then
            Exit Sub
        End If
    End If
    imgMainInfoSel(Index).Visible = True
    imgMainInfo(Index).Visible = False
End If
Exit Sub
Closer:
If Index = 1 Then
    imgMainInfoSel(Index).Visible = True
    imgMainInfo(Index).Visible = False
End If
End Sub

Private Sub imgMainInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Index = 1 Then GoTo CloseButton
If Index = 2 Then GoTo SsDd
If Index = 3 Then GoTo lLoO
If Index = 4 Then GoTo rRbB
If sockMain.State <> sckConnected Then Exit Sub
imgMainInfo(Index).Visible = True
imgMainInfoSel(Index).Visible = False

If Index = 0 Then
    If pProcess = False Then Exit Sub
    Dim Proc As String
    Dim P As Integer
    P = lstProcesses.ListIndex
    Proc = lstProcesses.List(P)
    If Proc = "" Then Exit Sub
    If sockMain.State = sckConnected Then
        Msg = Encrypt("|STOPPROCESS|" & Proc)
        sockMain.SendData Msg
        Pause 10
        tmrRefreshProcesses.Enabled = False
        sockProcesses.Close
        sockProcesses.Connect RemHst
        lblInfo(10).Caption = "Connected"
        tmrRefreshProcesses.Enabled = True
        Exit Sub
    End If
    Exit Sub
End If
SsDd:
If Index = 2 Then
    Msg = Encrypt("|SHUTDOWN|")
    If sockMain.State = sckConnected Then
        sockMain.SendData Msg
    End If
    Exit Sub
End If
lLoO:
If Index = 3 Then
    Msg = Encrypt("|LOGOFF|")
    If sockMain.State = sckConnected Then
        sockMain.SendData Msg
    End If
    Exit Sub
End If
rRbB:
If Index = 4 Then
    Msg = Encrypt("|REBOOT|")
    If sockMain.State = sckConnected Then
        sockMain.SendData Msg
    End If
    Exit Sub
End If
Exit Sub
CloseButton:
    'Unload frmAbout
    'Unload frmLogin
    'Unload frmChat
    'Unload frmWarning
    Unload Me
    'End
End Sub

Private Sub imgRA_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If sockMain.State <> sckConnected Then Exit Sub
imgRAS(Index).Visible = True
End Sub

Private Sub imgRA_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If rReady = False Then Exit Sub
imgRAS(Index).Visible = False
If sockMain.State = sckConnected Then
    If Index = 0 Then
        Msg = Encrypt("|SHOW|")
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 1 Then
        Msg = Encrypt("|HIDE|")
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 2 Then
        Reset_Images frmMain
        Disable_Images frmMain
        imgServOptSel.Top = imgServOpt.Top
        imgServOptSel.Visible = True
        imgServOpt.Visible = False
        Enable_ServerOptions frmMain
        Exit Sub
    End If
    If Index = 3 Then
        imgRA(13).Top = 3840
        imgRAS(13).Top = 3840
        imgRA(13).Visible = True
        imgRA(3).Visible = False
        Msg = Encrypt("|INVERSEMOUSE|")
        sockMain.SendData Msg
        iInverse = True
        Exit Sub
    End If
    If Index = 13 Then
        imgRA(3).Visible = True
        imgRA(13).Visible = False
        Msg = Encrypt("|NORMALMOUSE|")
        sockMain.SendData Msg
        iInverse = False
        Exit Sub
    End If
    If Index = 4 Then
        Msg = Encrypt("|EMPTYRECYCLEBIN|")
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 5 Then
        If bFileTransfer = True Then Exit Sub
        Close
        On Error Resume Next
        Kill App.Path & "\Desktop.jpg"
        Open App.Path & "\Desktop.jpg" For Binary As #1
        bGettingdesktop = True
        bFileTransfer = True
        Load frmDownloading
        Msg = "|DESKTOP|"
        sockExplorer.SendData Msg
        Exit Sub
    End If
    If Index = 7 Then
        Send_PopUp
    End If
    If Index = 8 Then
        Msg = Encrypt("|IE|" & txtURL.Text)
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 9 Then
        'Shell "regsvr32 WaveStream.dll /s"
        Load frmChat
        StayOnTop frmChat
        frmChat.Show
        frmChat.Height = 5610
        frmChat.Width = 6945
        frmChat.Top = (Screen.Height - frmChat.Height) / 2
        frmChat.Left = (Screen.Width - frmChat.Width) / 2
        Msg = "|CHAT|"
        Msg = Encrypt(Msg)
        If sockMain.State = sckConnected Then
            sockMain.SendData Msg
        End If
        sockChat.RemoteHost = RemHst
        sockChat.Connect
        lblInfo(9).Caption = "Connected"
        lblStatus.Caption = " Status:  Connecting Chat Client"
        tmrMessage.Enabled = True
        Exit Sub
    End If
    If Index = 10 Then
        Msg = Encrypt("|LOCK|")
        If sockMain.State = sckConnected Then
            sockMain.SendData Msg
        End If
        Exit Sub
    End If
    If Index = 11 Then
        Msg = Encrypt("|UNLOCK|")
        If sockMain.State = sckConnected Then
            sockMain.SendData Msg
        End If
        Exit Sub
    End If
    If Index = 12 Then
        'Unload frmAbout
        'Unload frmChat
        'Unload frmLogin
        'Unload frmSplash
        'Unload frmWarning
        Unload Me
        'End
        Exit Sub
    End If
    If Index = 14 Then
        Load frmRemEmail
        frmRemEmail.Show modal, frmMain
        Exit Sub
    End If
End If
End Sub
Private Sub Send_PopUp()
Dim sStyle As Long
For I = 0 To 8
    If optRA(I).Value = True Then sStyle = I
Next I
Msg = "|POPUP|" & sStyle & txtMSGTitle.Text & "|" & txtMessage.Text
Msg = Encrypt(Msg)
If sockMain.State = sckConnected Then
    sockMain.SendData Msg
End If
End Sub
Public Function StayOnTop(Form As Form)
Dim lFlags As Long
Dim lStay As Long
lFlags = SWP_NOSIZE Or SWP_NOMOVE
lStay = SetWindowPos(Form.hwnd, HWND_TOPMOST, 0, 0, 0, 0, lFlags)
End Function

Private Sub imgRE_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If sockMain.State <> sckConnected Then Exit Sub
imgRES(Index).Visible = True
End Sub

Private Sub imgRE_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
imgRES(Index).Visible = False
If sockExplorer.State = sckConnected Then

    If Index = 3 Then
        Msg = "|EXECUTE|" & lvFiles.SelectedItem.Key
        Msg = Encrypt(Msg)
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 4 Then
        If lblRE(0).Caption = "Pending.." Then
            TVTreeView_Collapse TVTreeView.Nodes(3)
            TVTreeView_NodeClick TVTreeView.Nodes(3)
            TVTreeView.Nodes.Item(3).Selected = True
            Me.MousePointer = vbDefault
            Exit Sub
        End If
        uUpLoad = True
        Load frmCustComDlg
        frmCustComDlg.Show
        Exit Sub
    End If
    If Index = 5 Then
        If chkSO(1).Value = 0 Then Exit Sub
        bBytes = 0
        Dim YTR As String
        Dim rty As Integer
        YTR = Me.lvFiles.SelectedItem.Text
        YTR = Mid(YTR, 1, Len(YTR) - 1)
        rty = Len(YTR)
        For I = rty To 1 Step -1
            If Mid(YTR, I, 1) = "(" Then
                YTR = Mid(YTR, 1, I - 1)
                Exit For
            End If
        Next I
        If sockMain.State <> sckConnected Then
            Load frmConnectError
            frmConnectError.Show
            StayOnTop frmConnectError
            frmConnectError.Height = 1785
            frmConnectError.Width = 4680
            frmConnectError.Top = (Screen.Height - frmConnectError.Height) / 2
            frmConnectError.Left = (Screen.Width - frmConnectError.Width) / 2
            Beep
            Reset_Images frmMain
            Disable_Images frmMain
            imgConnectSel.Top = imgConnect.Top
            imgConnectSel.Visible = True
            imgConnect.Visible = False
            tmrRefreshProcesses.Enabled = True
            Exit Sub
        End If
        Load frmCustComDlg
        frmCustComDlg.txtFileName.Text = YTR
        frmCustComDlg.Show
        Exit Sub
    End If
    If Index = 6 Then
        Load frmDisplay
        frmDisplay.Show
        DoEvents
        Open App.Path & "\Screen.jpg" For Binary As #50
        nNewScreen = True
        sockScreen.SendData "|SENDNEW|"
        Exit Sub
    End If
    If Index = 7 Then
        Unload frmDisplay
        Exit Sub
    End If
End If
NoFile:
End Sub

Private Sub imgRemAdmin_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
If imgConnectSel.Visible = True Then
    frmLogin.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
imgRemAdminSel.Top = imgRemAdmin.Top
imgRemAdmin.Visible = False
imgRemAdminSel.Visible = True
Disable_Images frmMain
Enable_imgRemoteAdmin frmMain
End Sub

Private Sub imgRemAdminSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If

Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
End Sub

Private Sub imgRemDesk_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
If imgConnectSel.Visible = True Then
    frmLogin.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
imgRemDeskSel.Top = imgRemDesk.Top
imgRemDeskSel.Visible = True
imgRemDesk.Visible = False
Disable_Images frmMain
Enable_RemDesk frmMain
End Sub

Private Sub imgRemDeskSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
End Sub

Private Sub imgRemExpl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
If imgConnectSel.Visible = True Then
    frmLogin.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
imgRemExplSel.Top = imgRemExpl.Top
imgRemExplSel.Visible = True
imgRemExpl.Visible = False
Disable_Images frmMain
Enable_RemoteExplorer frmMain
End Sub

Private Sub imgRemExplSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
End Sub

Private Sub imgServOpt_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
If imgConnectSel.Visible = True Then
    frmLogin.Show modal, frmMain
    Exit Sub
End If
Reset_Images frmMain
imgServOptSel.Top = imgServOpt.Top
imgServOptSel.Visible = True
imgServOpt.Visible = False
Disable_Images frmMain
Enable_ServerOptions frmMain
End Sub

Private Sub imgServOptSel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If imgAboutSel.Visible = True Then
    frmAbout.Show modal, frmMain
    Exit Sub
End If
Disable_Images frmMain
Enable_imgMain frmMain
Reset_Images frmMain
End Sub
Public Sub SendNewSettings()
If sockMain.State = sckConnected Then
    Msg = "|SETTINGS|"
    For I = 0 To 5
        Msg = Msg & chkSO(I).Value & ":"
    Next I
    Msg = Msg & txtEmail.Text
    Msg = Encrypt(Msg)
    sockMain.SendData Msg
Else
    Exit Sub
End If
End Sub
Private Sub imgSO_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

If sockMain.State <> sckConnected Then Exit Sub
imgSOS(Index).Visible = True
End Sub

Private Sub imgSO_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

imgSOS(Index).Visible = False
If sockMain.State = sckConnected Then
    If Index = 2 Then
        Msg = Encrypt("|GETSETTINGS|")
        sockMain.SendData Msg
        Exit Sub
    End If
    If Index = 3 Then
        Reset_Images frmMain
        Disable_Images frmMain
        Enable_imgMain frmMain
        Exit Sub
    End If
    If Index = 4 Then
        If chkSO(5).Value = 0 Then
            chkSO(5).Value = 1
        Else
            chkSO(5).Value = 0
        End If
        Exit Sub
    End If
    If Index = 5 Then
        SendNewSettings
        Exit Sub
    End If
End If
End Sub

Private Sub lvFiles_Click()
On Error GoTo Err
Dim FuckinFile As String
FuckinFile = lvFiles.SelectedItem.Key
Err:
End Sub

Private Sub picture1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static rec As Boolean, Msg As Long
Dim Retval As String
Dim returnstring
Dim RetValue
Msg = X / Screen.TwipsPerPixelX
If rec = False Then
    rec = True
    Select Case Msg
    Case WM_LBUTTONDOWN:
    Case WM_LBUTTONDBLCLK
        Me.Height = 9180
        'Restore
        'show the program
        Me.Show
        
        If About = True Then frmAbout.Show modal, frmMain
    Case WM_LBUTTONUP:
    Case WM_RBUTTONDBLCLK: 'not used in this program
    Case WM_RBUTTONDOWN:   'not used in this program
    Case WM_RBUTTONUP:
        'frmMeOptions.PopupMenu frmMeOptions.mnuOptions (Also not used)
        'need to use custom form
    End Select
    rec = False
End If
End Sub

Private Sub sockChat_DataArrival(ByVal bytesTotal As Long)
Dim cData As String
sockChat.GetData cData
cData = Decrypt(cData)

If InStr(1, cData, "|CHAT|") <> 0 Then
    frmChat.txtChat.Text = frmChat.txtChat.Text & Mid(cData, 7, Len(cData)) & vbCrLf
    Exit Sub
End If

End Sub

Private Sub sockExplorer_Close()
TVTreeView.Nodes.Clear
lvFiles.ListItems.Clear
sockExplorer.Close
End Sub

Private Sub sockExplorer_Connect()
TVTreeView.Nodes.Add , , "xxxROOTxxx", lblInfo(0).Caption, "RC2", "RC2"
sockExplorer.SendData "|ENUMDRVS|"
End Sub

Private Sub sockExplorer_DataArrival(ByVal bytesTotal As Long)
Dim Strdata As String
sockExplorer.GetData Strdata, vbString

If InStr(1, Strdata, "|NOT|") <> 0 Then
    Load frmSecure
    frmSecure.Show
    StayOnTop frmSecure
    frmSecure.Height = 1415
    frmSecure.Width = 4670
    frmSecure.Left = (Screen.Width - frmSecure.Width) / 2
    frmSecure.Top = (Screen.Height - frmSecure.Height) / 2
    Beep
    TVTreeView_Collapse TVTreeView.Nodes(3)
    TVTreeView_NodeClick TVTreeView.Nodes(3)
    TVTreeView.Nodes.Item(3).Selected = True
    Me.MousePointer = vbDefault
    Exit Sub
End If
If InStr(1, Strdata, "|CANT|") <> 0 Then
    Close #1
    bFileTransfer = False
    If bGettingdesktop = True Then
        bGettingdesktop = False
    End If
    MsgBox "The Server Is Not Allowing File Download.", vbInformation, "NO DOWNLOADING ALLOWED"
    Unload frmDownloading
    Exit Sub
End If
If InStr(1, Strdata, "|COMPLEET|") <> 0 Then
    'Dim SDLEN2 As Long
    'SDLEN2 = Len(Strdata) - 10
    'Strdata = Mid(Strdata, 1, SDLEN2)
    frmDownloading.objProg.Value = frmDownloading.objProg.Max
    bFileTransfer = False
    Put #1, , Strdata
    Close #1
    DoEvents
    If bGettingdesktop = True Then
        bGettingdesktop = False
    End If
    Unload frmDownloading
    Exit Sub
End If
If InStr(1, Strdata, "|COMPLETE|") <> 0 Then
    'Debug.Print Strdata
    'Dim SDLEN As Long
    'SDLEN = Len(Strdata)
    'SDLEN = SDLEN - 10
    'Strdata = Mid(Strdata, 1, SDLEN)
    frmDownloading.objProg.Value = frmDownloading.objProg.Max
    frmDownloading.Caption = "FILE RECEIVED"
    frmDownloading.lblBytes.Caption = "DOWNLOAD COMPLETE"
    Put #1, , Strdata
    Close #1
    bFileTransfer = False
    DoEvents
    If bGettingdesktop = True Then
        bGettingdesktop = False
    End If
    Unload frmDownloading
    Reset_Images frmMain
    imgRemDeskSel.Top = imgRemDesk.Top
    imgRemDeskSel.Visible = True
    imgRemDesk.Visible = False
    Disable_Images frmMain
    Enable_RemDesk frmMain
    Screen_Shot
    Exit Sub
End If
If InStr(1, Strdata, "|SOME|") <> 0 Then
    lblStatus.Caption = " Status:  Receiving Remote File Information"
    FList = FList & Mid(Strdata, 7, Len(Strdata))
    tmrMessage.Enabled = True
    Exit Sub
End If
If InStr(1, Strdata, "|DRVS|") <> 0 Then
    lblStatus.Caption = " Status:  Retreiving Remote Drives"
    Populate_Tree_With_Drives Strdata, TVTreeView
    tmrMessage.Enabled = True
    Exit Sub
End If
If InStr(1, Strdata, "|FOLDERS|") <> 0 Then
    lblStatus.Caption = " Status:  Receiving Remote Folder Information"
    Populate_Folders Strdata, TVTreeView
    tmrMessage.Enabled = True
    Exit Sub
End If
If InStr(1, Strdata, "|FILES|") <> 0 Then
    lblStatus.Caption = " Status:  Populating Remote Files"
    Call Populate_Files(FList, lvFiles)
    frmMain.MousePointer = vbDefault
    FList = "|FILES|"
    tmrMessage.Enabled = True
    Exit Sub
End If
If bFileTransfer = True Then
    If InStr(1, Strdata, "|FILESIZE|") <> 0 Then
        frmDownloading.Show modal, frmMain
        frmDownloading.lblBytes.Caption = CLng(Mid$(Strdata, 11, Len(Strdata)))
        frmDownloading.objProg.Max = CLng(Mid$(Strdata, 11, Len(Strdata)))
        Exit Sub
    End If
    Put #1, , Strdata
    bBytes = bBytes + Len(Strdata)
    With frmDownloading.objProg
        If (.Value + Len(Strdata)) <= .Max Then
            .Value = .Value + Len(Strdata)
        Else
            .Value = .Max
        End If
    End With
End If
sockExplorer_DataArrival_Exit:
Exit Sub
sockExplorer_DataArrival_Error:
bGettingdesktop = False
MsgBox Err.Description, vbCritical, "SUBNET® DOWNLOADER"
Exit Sub
End Sub
Private Sub Screen_Shot()
    dDesk = True
    Reset_Images frmMain
    imgRemDeskSel.Top = imgRemDesk.Top
    imgRemDeskSel.Visible = True
    imgRemDesk.Visible = False
    Disable_Images frmMain
    Enable_RemDesk frmMain
    picRemDesk.Picture = LoadPicture(App.Path & "\Desktop.jpg")
    picRemDesk2.Height = imgRemScreen.Height
    picRemDesk2.Width = imgRemScreen.Width
    imgRemScreen.Visible = True
    HScroll1.Visible = True
    VScroll1.Visible = True
    HScroll1.Min = 0
    HScroll1.Max = ScaleX(picRemDesk.Picture.Width, 8, vbTwips) - picRemDesk2.Width
    HScroll1.LargeChange = 50 * Screen.TwipsPerPixelX
    HScroll1.SmallChange = Screen.TwipsPerPixelX
    VScroll1.Min = 0
    VScroll1.Max = ScaleX(picRemDesk.Picture.Height, 8, vbTwips) - picRemDesk2.Height
    VScroll1.LargeChange = 50 * Screen.TwipsPerPixelY
    VScroll1.SmallChange = Screen.TwipsPerPixelY
    HScroll1_Change
End Sub
Private Sub sockMain_Close()
lblInfo(6).Caption = "Pending.."
lblInfo(7).Caption = "Pending.."
tmrRefreshProcesses.Enabled = False

TVTreeView.Nodes.Clear
lvFiles.ListItems.Clear
Reset_Images frmMain
Disable_Images frmMain
Enable_imgMain frmMain
imgDisconnectSel.Visible = False
imgConnect.Visible = True
imgDisconnect.Visible = False
Close_Connections
For I = 0 To 10
    lblInfo(I).Caption = "Pending.."
Next I
lstUsers.Clear
lstUsers.AddItem "Pending.."
lstProcesses.Clear
lstProcesses.AddItem "Pending.."
lblRE(0).Caption = "Pending.."
lblRE(2).Caption = "Pending.."
lblStatus.Caption = " Status:  Disconnected"
tmrRefreshProcesses.Enabled = False
End Sub

Private Sub sockMain_Connect()
lblStatus.Caption = " Status:  Connected to, " & RemHst & " as, " & User
imgDisconnect.Visible = True
imgConnect.Visible = False
lblInfo(6).Caption = "Connected"
End Sub
Private Sub Open_Explorer()
If sockExplorer.State <> sckClosed Then
    sockExplorer.Close
    TVTreeView.Nodes.Clear
    lvFiles.ListItems.Clear
End If
If sockExplorer.State <> sckConnected Then
    With sockExplorer
        .RemoteHost = RemHst
        .Connect
    End With
Else
    TVTreeView.Nodes.Clear
    lvFiles.ListItems.Clear
    sockExplorer.Close
    With sockExplorer
        .RemoteHost = RemHst
        .Connect
    End With
End If
lblInfo(7).Caption = "Connected"
End Sub
Private Sub sockMain_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
sockMain.GetData Data
Data = Decrypt(Data)

If InStr(1, Data, "|OPTIONS|") <> 0 Then
    Unload frmVerify
    lLoggedIn = True
    Adjust_Settings Mid(Data, 10, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|INFO|") <> 0 Then
    ShowInfo Mid(Data, 7, Len(Data))
    Open_Explorer
    Exit Sub
End If
If InStr(1, Data, "|MSGBUTTON|") <> 0 Then
    lblButton.Caption = Mid(Data, 12, Len(Data))
    Exit Sub
End If
If InStr(1, Data, "|TYPING|") <> 0 Then
    frmChat.lblTyping.Visible = True
    Exit Sub
End If
If InStr(1, Data, "|NOTTYPING|") <> 0 Then
    frmChat.lblTyping.Visible = False
    Exit Sub
End If
If InStr(1, Data, "|BANNED|") <> 0 Then
    Unload frmVerify
    Temp_Ban
    Exit Sub
End If
If InStr(1, Data, "|SECURELOGIN|") <> 0 Then
    Unload frmVerify
    tmrConnect.Enabled = False
    Load frmLogin
    frmLogin.txtPass.Visible = True
    frmLogin.txtPass.Text = Passy
    frmLogin.Label1.Visible = True
    frmLogin.Show modal, frmMain
    Exit Sub
End If

End Sub
Private Sub Temp_Ban()
Close_Connections
Load frmRejection
StayOnTop frmRejection
With frmRejection
    .Height = 1410
    .Width = 4665
    .Top = (Screen.Height - .Height) / 2
    .Left = (Screen.Width - .Width) / 2
    .Show
End With
End Sub
Private Sub Adjust_Settings(sString As String)
StartUp = Mid(sString, 1, 1)
FileDL = Mid(sString, 3, 1)
AutoLogin = Mid(sString, 5, 1)
VerifyLogin = Mid(sString, 7, 1)
AutoEmail = Mid(sString, 9, 1)
Hidden = Mid(sString, 11, 1)
Addy = Mid(sString, 13, Len(sString))
chkSO(0).Value = StartUp
chkSO(1).Value = FileDL
chkSO(2).Value = AutoLogin
chkSO(3).Value = VerifyLogin
chkSO(4).Value = AutoEmail
chkSO(5).Value = Hidden
txtEmail.Text = Addy
End Sub
Private Sub ShowInfo(Info As String)
Dim z As Integer
z = InStr(1, Info, ":")
lblInfo(0).Caption = Mid(Info, 1, z - 1)
Info = Mid(Info, z + 1, Len(Info))
z = InStr(1, Info, ":")
lblInfo(1).Caption = Mid(Info, 1, z - 1)
Info = Mid(Info, z + 1, Len(Info))
z = InStr(1, Info, ":")
lblInfo(2).Caption = Format(Mid(Info, 1, z - 1), "###.##") & " Hours"
Info = Mid(Info, z + 1, Len(Info))
z = InStr(1, Info, ":")
lblInfo(3).Caption = Mid(Info, 1, z - 1)
Info = Mid(Info, z + 1, Len(Info))
z = InStr(1, Info, ":")
lblInfo(4).Caption = Mid(Info, 1, z - 1)
Info = Mid(Info, z + 1, Len(Info))
lblInfo(5).Caption = Mid(Info, 1, Len(Info))
lstProcesses.Clear
sockProcesses.Close
sockProcesses.Connect RemHst
If lblInfo(3).Caption = "Windows NT" Then
    pProcess = False
Else
    pProcess = True
End If
End Sub

Private Sub sockProcesses_Close()
lblInfo(10).Caption = "Pending.."
End Sub

Private Sub sockProcesses_DataArrival(ByVal bytesTotal As Long)
sockProcesses.GetData Data
Data = Decrypt(Data)
NewProcess Data
End Sub

Private Sub NewProcess(Process As String)
Dim WFM As Integer
For I = 1 To Len(Process)
    If Mid(Process, I, 1) = ":" Then
        WFM = WFM + 1
    End If
Next I
AddTheProcesses Process, WFM
End Sub
Private Sub AddTheProcesses(Processes As String, HowManyTimes As Integer)
lstProcesses.Clear
On Error Resume Next
Dim Strt As Integer
Dim Cnt As Integer
AGAIN:
Strt = InStr(2, Processes, ":")
lstProcesses.AddItem Mid(Processes, 2, Strt - 2)
Cnt = Cnt + 1
Processes = Mid(Processes, Strt, Len(Processes))
If Cnt < HowManyTimes Then GoTo AGAIN
tmrRefreshProcesses.Enabled = True
End Sub

Private Sub sockScreen_DataArrival(ByVal bytesTotal As Long)
Dim sScreen As String
sockScreen.GetData sScreen, vbString

If InStr(1, sScreen, "|MOUSEPOS|") <> 0 Then
    Place_Mouse Mid(sScreen, 11, Len(sScreen)), 4
    Exit Sub
End If
If InStr(1, sScreen, "|MCLICK|") <> 0 Then
    mMouseClicked Mid(sScreen, 9, Len(sScreen))
    Exit Sub
End If
If InStr(1, sScreen, "|FINISHED|") <> 0 Then
    'Dim lLenScreen As Long
    'lLenScreen = Len(sScreen)
    'lLenScreen = lLenScreen - 10
    'sScreen = Mid(sScreen, 1, lLenScreen)
    Put #50, , sScreen
    Close #50
    nNewScreen = False
    DoEvents
    frmDisplay.imgDisplay.Picture = LoadPicture(App.Path & "\Screen.jpg")
    Picture1.Picture = LoadPicture(App.Path & "\Screen.jpg")
    frmDisplay.imgBlank.Visible = False
    frmDisplay.Show
    Exit Sub
End If
If nNewScreen = True Then
    If InStr(1, sScreen, "|FILESIZE|") <> 0 Then
        Exit Sub
    End If
    Put #50, , sScreen
End If
End Sub
Private Sub mMouseClicked(PosAndButton As String)
Dim bBbutton As String
Dim Seper1 As Long
Dim Seper2 As Long
Seper1 = InStr(1, PosAndButton, "|")
Seper2 = InStr(Seper1 + 1, PosAndButton, "|")
bBbutton = Mid(PosAndButton, Seper2 + 1, Len(PosAndButton))
Dim WB As Integer
If bBbutton = "RIGHT-BUTTON" Then WB = 2
If bBbutton = "LEFT-BUTTON" Then WB = 1
If bBbutton = "MIDDLE-BUTTON" Then WB = 3
Place_Mouse Mid(PosAndButton, 1, Seper2 - 1), WB
End Sub

Private Sub Place_Mouse(Possitions As String, WB As Integer)
    frmDisplay.imgMouse.Visible = False
    Dim Xplace As Single
    Dim Yplace As Single
    Dim Seppp As Long
    Seppp = InStr(1, Possitions, "|")
    Xplace = Mid(Possitions, 1, Seppp - 1)
    Yplace = Mid(Possitions, Seppp + 1, Len(Possitions))
    frmDisplay.imgMouse.Left = frmDisplay.imgDisplay.Left + (frmDisplay.imgDisplay.Width * Xplace)
    frmDisplay.imgMouse.Top = frmDisplay.imgDisplay.Top + (frmDisplay.imgDisplay.Height * Yplace)
    Set frmDisplay.imgMouse.Picture = frmDisplay.ImageList1.ListImages(WB).Picture
    frmDisplay.imgMouse.Visible = True
End Sub

Public Sub tmrConnect_Timer()
tmrConnect.Enabled = False
If lstProcesses.ListCount < 2 Then
    Load frmConnectError
    frmConnectError.Show
    StayOnTop frmConnectError
    frmConnectError.Height = 1785
    frmConnectError.Width = 4670
    frmConnectError.Top = (Screen.Height - frmConnectError.Height) / 2
    frmConnectError.Left = (Screen.Width - frmConnectError.Width) / 2
    Beep
    Reset_Images frmMain
    Disable_Images frmMain
    imgConnectSel.Top = imgConnect.Top
    imgConnectSel.Visible = True
    imgConnect.Visible = False
    tmrRefreshProcesses.Enabled = True
    Unload frmVerify
End If

End Sub

Private Sub tmrMessage_Timer()
tmrMessage.Enabled = False
lblStatus.Caption = " Status:  Connected to, " & RemHst & " as, " & User
End Sub

Private Sub tmrRefreshProcesses_Timer()
sockProcesses.Close
If sockMain.State = sckConnected Then
    sockProcesses.Connect RemHst
End If
End Sub

Private Sub TVTreeView_Collapse(ByVal Node As MSComctlLib.Node)
On Error GoTo tvTreeView_Collapse_Error
If Node.Key = "xxxROOTxxx" Then
    Exit Sub
End If
Delete_Child_Nodes Me.TVTreeView, Node
tvTreeView_Collapse_Exit:
    Exit Sub
tvTreeView_Collapse_Error:
    MsgBox Err.Description, vbCritical, "Explorer Collapse"
    Exit Sub
End Sub

Private Sub TVTreeView_NodeClick(ByVal Node As MSComctlLib.Node)
Dim FLDR As String
FLDR = Node.Key
FolderClick = FLDR
For I = Len(FLDR) To 1 Step -1
    If Mid(FLDR, I, 1) = "\" Then
        FLDR = Mid(FLDR, I + 1, Len(FLDR))
        Call SendFolderName(FLDR)
    End If
Next I
On Error GoTo tvTreeView_NodeClick_Error
Dim sData As String
LVFileCount = 0
Me.MousePointer = vbHourglass
sData = "|FOLDERS|" & Node.Key
sockExplorer.SendData (sData)
tvTreeView_NodeClick_Exit:
    Exit Sub
tvTreeView_NodeClick_Error:
    Me.MousePointer = vbDefault
    If Err.Number = 40006 Then
        MsgBox "Remote connection lost!", vbExclamation, "Explorer Click"
        Exit Sub
    End If
    MsgBox Err.Description, vbCritical, "Explorer Click"
    Exit Sub
End Sub
Private Sub SendFolderName(FldName As String)
If FldName = "" Then FldName = "C:"
lblRE(0).Caption = FldName
sockExplorer.SendData "|FN|" & FldName
Pause 10
End Sub

Private Sub VScroll1_Change()
    picRemDesk2.PaintPicture picRemDesk.Picture, 0, 0, _
    picRemDesk2.Width, picRemDesk2.Height, _
    HScroll1.Value, VScroll1.Value, _
    picRemDesk2.Width, picRemDesk2.Height
    Refresh_Desk
End Sub

Private Sub VScroll1_Scroll()
    picRemDesk2.PaintPicture picRemDesk.Picture, 0, 0, _
    picRemDesk2.Width, picRemDesk2.Height, _
    HScroll1.Value, VScroll1.Value, _
    picRemDesk2.Width, picRemDesk2.Height
    Refresh_Desk
End Sub
