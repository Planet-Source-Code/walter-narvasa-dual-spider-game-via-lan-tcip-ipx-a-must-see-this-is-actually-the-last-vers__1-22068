VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmFighter 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "The Spider on the Stick Version 2 by Walter A. Narvasa"
   ClientHeight    =   7470
   ClientLeft      =   2040
   ClientTop       =   1860
   ClientWidth     =   9585
   Icon            =   "frmFighter.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "frmFighter.frx":0CCA
   ScaleHeight     =   7470
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame Connectivity 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   7215
      Left            =   0
      TabIndex        =   38
      Top             =   120
      Visible         =   0   'False
      Width           =   9615
      Begin VB.CommandButton cmdSendMessage 
         Caption         =   "&Send Message"
         Enabled         =   0   'False
         Height          =   375
         Left            =   7560
         TabIndex        =   51
         Top             =   6315
         Width           =   1455
      End
      Begin VB.CommandButton cmdClose 
         Caption         =   "&Close"
         Height          =   375
         Left            =   480
         TabIndex        =   50
         Top             =   5040
         Width           =   2535
      End
      Begin VB.CommandButton cmdPlay 
         Caption         =   "&Start Play"
         Height          =   375
         Left            =   480
         TabIndex        =   44
         Top             =   4560
         Width           =   2535
      End
      Begin VB.CommandButton cmdDisconnect 
         Caption         =   "&Disconnect Game"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   40
         Top             =   4080
         Width           =   2535
      End
      Begin VB.CommandButton cmdConnect 
         Caption         =   "&Connect to Game Host"
         Height          =   375
         Left            =   480
         TabIndex        =   39
         Top             =   3600
         Width           =   2535
      End
      Begin VB.TextBox txtPlayerName 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5160
         TabIndex        =   48
         Text            =   " "
         Top             =   2520
         Width           =   2295
      End
      Begin VB.TextBox txtDataReceived 
         Enabled         =   0   'False
         Height          =   1815
         Left            =   3840
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   47
         Text            =   "frmFighter.frx":E1D0C
         Top             =   4320
         Width           =   5175
      End
      Begin VB.TextBox txtDataSent 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3840
         TabIndex        =   46
         Text            =   " "
         Top             =   6360
         Width           =   3615
      End
      Begin VB.TextBox txtIPAddress 
         Enabled         =   0   'False
         Height          =   285
         Left            =   5640
         TabIndex        =   42
         Text            =   " "
         Top             =   1560
         Width           =   1335
      End
      Begin VB.CommandButton cmdListen 
         Caption         =   "Game Host &Listen"
         Height          =   375
         Left            =   480
         TabIndex        =   41
         Top             =   3120
         Width           =   2535
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Multiplayer (LAN IPX/Internet-TCPIP) Connection Options"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1215
         Left            =   480
         TabIndex        =   53
         Top             =   1680
         Width           =   2415
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Host-Remote Chat Box"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   300
         Left            =   3840
         TabIndex        =   52
         Top             =   3960
         Width           =   2790
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3360
         X2              =   9480
         Y1              =   3720
         Y2              =   3720
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         X1              =   3360
         X2              =   3360
         Y1              =   120
         Y2              =   7080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Player's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   4680
         TabIndex        =   49
         Top             =   2040
         Width           =   3330
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Please Enter Game Host IP Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   300
         Left            =   4080
         TabIndex        =   43
         Top             =   1080
         Width           =   4395
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00000000&
         BorderColor     =   &H00FF0000&
         BorderWidth     =   3
         Height          =   6975
         Left            =   120
         Shape           =   4  'Rounded Rectangle
         Top             =   120
         Width           =   9375
      End
   End
   Begin VB.PictureBox picPlayerPunchNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6840
      Picture         =   "frmFighter.frx":E1D0E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   90
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6960
      Picture         =   "frmFighter.frx":E249B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   92
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7080
      Picture         =   "frmFighter.frx":E2DD6
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   89
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7200
      Picture         =   "frmFighter.frx":E368C
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   91
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7320
      Picture         =   "frmFighter.frx":E3E9D
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   93
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7440
      Picture         =   "frmFighter.frx":E469B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   94
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardNormalRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7560
      Picture         =   "frmFighter.frx":E4EDA
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   95
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   840
      Picture         =   "frmFighter.frx":E570E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   82
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   720
      Picture         =   "frmFighter.frx":E5E07
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   83
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "frmFighter.frx":E6657
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   84
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   480
      Picture         =   "frmFighter.frx":E6E34
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   85
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   360
      Picture         =   "frmFighter.frx":E758B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   86
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   240
      Picture         =   "frmFighter.frx":E7CE1
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   87
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardNormalLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "frmFighter.frx":E841E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   88
      Top             =   1560
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6840
      Picture         =   "frmFighter.frx":E8B44
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   81
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6960
      Picture         =   "frmFighter.frx":E92D4
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   80
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7080
      Picture         =   "frmFighter.frx":E9C14
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   79
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7200
      Picture         =   "frmFighter.frx":EA4D6
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   78
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7320
      Picture         =   "frmFighter.frx":EACE8
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   77
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7440
      Picture         =   "frmFighter.frx":EB4E9
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   76
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardJungleRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7560
      Picture         =   "frmFighter.frx":EBCCF
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   75
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   840
      Picture         =   "frmFighter.frx":EC4B4
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   74
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   720
      Picture         =   "frmFighter.frx":ECC44
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   73
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "frmFighter.frx":ED583
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   72
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   480
      Picture         =   "frmFighter.frx":EDE3B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   71
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   360
      Picture         =   "frmFighter.frx":EE644
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   70
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   240
      Picture         =   "frmFighter.frx":EEE3E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   69
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardJungleLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "frmFighter.frx":EF62A
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   68
      Top             =   6240
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6840
      Picture         =   "frmFighter.frx":EFE1B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   67
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6960
      Picture         =   "frmFighter.frx":F05A9
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   66
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7080
      Picture         =   "frmFighter.frx":F0EE4
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   65
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7200
      Picture         =   "frmFighter.frx":F179F
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   64
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7320
      Picture         =   "frmFighter.frx":F1FAB
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   63
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7440
      Picture         =   "frmFighter.frx":F279F
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   62
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardCaveRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7560
      Picture         =   "frmFighter.frx":F2F96
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   61
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   840
      Picture         =   "frmFighter.frx":F377C
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   60
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   720
      Picture         =   "frmFighter.frx":F3F0A
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   59
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "frmFighter.frx":F483E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   58
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   480
      Picture         =   "frmFighter.frx":F50F2
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   57
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   360
      Picture         =   "frmFighter.frx":F5904
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   56
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   240
      Picture         =   "frmFighter.frx":F6106
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   55
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardCaveLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "frmFighter.frx":F6907
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   54
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrMultiplayerLanInternet 
      Interval        =   1
      Left            =   2880
      Top             =   0
   End
   Begin VB.PictureBox picPlayerPunchDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6840
      Picture         =   "frmFighter.frx":F70F2
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   37
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   6960
      Picture         =   "frmFighter.frx":F787B
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   36
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7080
      Picture         =   "frmFighter.frx":F81B0
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   35
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7200
      Picture         =   "frmFighter.frx":F8A5D
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   34
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7320
      Picture         =   "frmFighter.frx":F926A
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   33
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7440
      Picture         =   "frmFighter.frx":F9A5A
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   32
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardDomesticRight 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   7560
      Picture         =   "frmFighter.frx":FA23C
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   31
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunchDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   840
      Picture         =   "frmFighter.frx":FAA1E
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   30
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHitDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   720
      Picture         =   "frmFighter.frx":FB1AA
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   29
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLossDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   600
      Picture         =   "frmFighter.frx":FBADE
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   28
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKickDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   480
      Picture         =   "frmFighter.frx":FC382
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   27
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBackDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   360
      Picture         =   "frmFighter.frx":FCB89
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   26
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerNoneDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   240
      Picture         =   "frmFighter.frx":FD37C
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   25
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompLoss 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4200
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   9
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompPunch 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4320
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   11
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompKick 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4440
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   8
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompHit 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4560
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   7
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompBack 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4680
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   5
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerPunch 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3480
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   18
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerHit 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3360
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   14
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerLoss 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3240
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   15
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerKick 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3120
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   17
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerBack 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   3000
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   12
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrRegeneration 
      Enabled         =   0   'False
      Interval        =   700
      Left            =   4800
      Top             =   0
   End
   Begin VB.Timer tmrPlayerRecover 
      Interval        =   500
      Left            =   3480
      Top             =   0
   End
   Begin VB.Timer tmrComputerAI 
      Interval        =   500
      Left            =   5280
      Top             =   0
   End
   Begin VB.PictureBox picCompNone 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4800
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   10
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picCompForward 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   4920
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   6
      Top             =   4680
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.Timer tmrCompRecover 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4080
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar w 
      Height          =   255
      Left            =   1200
      TabIndex        =   2
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar l 
      Height          =   255
      Left            =   6120
      TabIndex        =   3
      Top             =   1560
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picPlayerNone 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   2880
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   16
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForward 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   2760
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   13
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin VB.PictureBox picPlayerForwardDomesticLeft 
      AutoSize        =   -1  'True
      Height          =   1530
      Left            =   120
      Picture         =   "frmFighter.frx":FDB5A
      ScaleHeight     =   1470
      ScaleWidth      =   1815
      TabIndex        =   24
      Top             =   3120
      Visible         =   0   'False
      Width           =   1875
   End
   Begin MSWinsockLib.Winsock ws 
      Left            =   5760
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Points: $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7560
      TabIndex        =   99
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblCompPoints 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   8790
      TabIndex        =   98
      Top             =   360
      Width           =   135
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Points: $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   120
      TabIndex        =   97
      Top             =   360
      Width           =   945
   End
   Begin VB.Label lblPlayerPoints 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1350
      TabIndex        =   96
      Top             =   360
      Width           =   135
   End
   Begin VB.Image m 
      Height          =   1470
      Left            =   7560
      Picture         =   "frmFighter.frx":FE345
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Image f 
      Height          =   1470
      Left            =   240
      Picture         =   "frmFighter.frx":FEB84
      Top             =   5280
      Width           =   1815
   End
   Begin VB.Label GameStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Status:"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   45
      Top             =   7200
      Width           =   9375
   End
   Begin VB.Label lblCompMoney 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   7950
      TabIndex        =   23
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblPlayerMoney 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   3240
      TabIndex        =   22
      Top             =   840
      Width           =   375
   End
   Begin VB.Label lblCompCaption 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Current Earnings: $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   5760
      TabIndex        =   21
      Top             =   840
      Width           =   2070
   End
   Begin VB.Label lblPlayerCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Current Earnings: $"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   270
      Left            =   1020
      TabIndex        =   20
      Top             =   840
      Width           =   2145
   End
   Begin VB.Shape Shape3 
      BackColor       =   &H00000000&
      BorderColor     =   &H00004080&
      BorderWidth     =   3
      DrawMode        =   15  'Merge Pen Not
      FillColor       =   &H00004080&
      FillStyle       =   0  'Solid
      Height          =   135
      Left            =   120
      Top             =   6800
      Width           =   9375
   End
   Begin VB.Label lblFighter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "The Spider on the Stick"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   555
      Left            =   2040
      TabIndex        =   19
      Top             =   120
      Width           =   5340
   End
   Begin VB.Label lblWinner 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1335
      Left            =   2880
      TabIndex        =   4
      Top             =   3600
      Width           =   4095
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   1215
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   720
      Width           =   9375
   End
   Begin VB.Label lblCompName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Computer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   6120
      TabIndex        =   1
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblPlayerName 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Player"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00000000&
      BorderColor     =   &H00FF0000&
      BorderWidth     =   3
      Height          =   5055
      Left            =   120
      Top             =   2040
      Width           =   9375
   End
   Begin VB.Menu mnuGame 
      Caption         =   "&Game"
      Begin VB.Menu mnuNewGameItem 
         Caption         =   "&New Game"
         Begin VB.Menu mnuSinglePlayerItem 
            Caption         =   "&Single Player vs Computer"
            Checked         =   -1  'True
            Shortcut        =   ^S
         End
         Begin VB.Menu mnuTwoPlayerItem 
            Caption         =   "Multiplayer (&Keyboard)"
            Shortcut        =   ^K
         End
         Begin VB.Menu mnuTwoPlayer2Item 
            Caption         =   "Multiplayer (&LAN/Internet)"
            Shortcut        =   ^L
         End
      End
      Begin VB.Menu mnuFightItem 
         Caption         =   "&Fight!"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuVarietyItem 
         Caption         =   "Variety of Different Spider Selection"
         Begin VB.Menu mnuVarietyItemPlayer1 
            Caption         =   "Variety of Spiders Selection for Player 1"
            Begin VB.Menu mnuSpiderNormalItem 
               Caption         =   "&Normal Spider - Default"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuSpiderDomesticItem 
               Caption         =   "&Domestic Spider -Worth $250.00"
            End
            Begin VB.Menu mnuSpiderCaveItem 
               Caption         =   "&Cave Spider - Worth $500.00"
            End
            Begin VB.Menu mnuSpiderJungleItem 
               Caption         =   "&Jungle Spider - Worth $750.00"
            End
         End
         Begin VB.Menu mnuVarietyItemPlayer2 
            Caption         =   "Variety of Spiders Selection for Player 2"
            Begin VB.Menu mnuSpiderNormalItem2 
               Caption         =   "&Normal Spider - Default"
               Checked         =   -1  'True
            End
            Begin VB.Menu mnuSpiderDomesticItem2 
               Caption         =   "&Domestic Spider -Worth $250.00"
            End
            Begin VB.Menu mnuSpiderCaveItem2 
               Caption         =   "&Cave Spider - Worth $500.00"
            End
            Begin VB.Menu mnuSpiderJungleItem2 
               Caption         =   "&Jungle Spider - Worth $750.00"
            End
         End
      End
      Begin VB.Menu mnuDifficultyItem 
         Caption         =   "&Type of Stages"
         Begin VB.Menu mnuNormalItem 
            Caption         =   "&Normal - Default"
            Checked         =   -1  'True
            Shortcut        =   ^N
         End
         Begin VB.Menu mnuEasyItem 
            Caption         =   "&Easy - House"
            Shortcut        =   ^E
         End
         Begin VB.Menu mnuMediumItem 
            Caption         =   "&Medium - Basketball Court"
            Shortcut        =   ^M
         End
         Begin VB.Menu mnuHardItem 
            Caption         =   "&Hard - Playground"
            Shortcut        =   ^H
         End
      End
      Begin VB.Menu separator2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuControlsItem 
         Caption         =   "&Control Setting"
         Shortcut        =   ^C
      End
      Begin VB.Menu mnuHowtoPlayItem 
         Caption         =   "H&ow to Play"
         Shortcut        =   ^O
      End
      Begin VB.Menu mnuSpiderInformation 
         Caption         =   "Spider &Information"
         Shortcut        =   ^I
      End
      Begin VB.Menu mnuCreditsItem 
         Caption         =   "C&redits"
         Shortcut        =   ^R
      End
      Begin VB.Menu separator 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExitItem 
         Caption         =   "E&xit"
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmFighter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================================================================================================
'
' Developed by Walter A. Narvasa
' jawoltze@edsamail.com.ph
'
' Walter A. Narvasa of
' WANCOM SYSTEMS
'
' READ THIS BEFORE USING THE CODE:
'
' You can study and view the source code for creating your
' own apps, but do not reproduce/release The Spider on the Stick fully
' or partially for any commercial and/or personal purposes. All
' rights of this product is related to it's author. Any violation
' of above conditions will be treated seriously and is punishable.
'
' I do not have full time to add complete explanation, read the help
' file (click Help->Contents) in The Spider on the Stick. Contact me for
' additional help/suggestions
'
'
' VISIT MY WEBSITE : http://jawoltze.gq.nu/
'
'=============================================================================================================================
Option Explicit

Private a As Boolean, aa As Boolean
Private bb As Integer, b As Integer
Private i As Integer, i2 As Integer
Dim wsdata As String 'winsock data that will be send
Dim SingleSelected As Boolean
Dim KeyboardSelected As Boolean

Private Sub Form_Load()
    Randomize
    SingleSelected = False
    KeyboardSelected = False
    aaa = lblPlayerName.Caption
    aaaa = lblCompName.Caption
    f.Left = 240
    m.Left = 7560
    f.Top = 5280
    m.Top = 5280
    bb = 100
    b = 100
    l.Value = b
    w.Value = bb
    f.Picture = picPlayerForwardNormalLeft.Picture
    m.Picture = picPlayerForwardNormalRight.Picture
    lblWinner.Caption = ""
    lblFighter.ForeColor = vbRed
    Call mnuNormalItem_Click
    Call mnuSpiderNormalItem_Click
    Call mnuSpiderNormalItem2_Click
    bb = 100
    b = 100
    w.Value = bb
    l.Value = b
    tmrComputerAI.Enabled = False
    tmrCompRecover.Enabled = False
    i2 = 5
    i = 5
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If lblFighter.ForeColor = vbGreen Then
        'constant moves for left player
        If KeyCode = vbKeyRight Then
            f.Picture = picPlayerForward.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            If m.Left - f.Left < 1000 Then
                Exit Sub
            Else
                f.ZOrder (1)
                f.Left = f.Left + 200
            End If
        ElseIf KeyCode = vbKeyLeft Then
                If f.Left - 300 < 100 Then
                    Exit Sub
                End If
            i2 = 1
            f.ZOrder (1)
            f.Picture = picPlayerBack.Picture
            f.Left = f.Left - 200
        ElseIf KeyCode = vbKeyControl Then
            a = True
            f.ZOrder (1)
            f.Picture = picPlayerPunch.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
        ElseIf KeyCode = vbKeyShift Then
            f.ZOrder (1)
            a = True
            f.Picture = picPlayerKick.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
        End If
        'assign a multiplayer on keyboard right player
        If mnuTwoPlayerItem.Checked = True Then
            mnuSinglePlayerItem.Checked = False
            mnuTwoPlayer2Item.Checked = False
            If KeyCode = vbKeyA Then
                m.Picture = picCompForward.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
                If m.Left - f.Left < 1000 Then
                    Exit Sub
                Else
                    m.ZOrder (1)
                    m.Left = m.Left - 200
                End If
            ElseIf KeyCode = vbKeyS Then
                If m.Left + 300 >= 5280 Then
                    Exit Sub
                End If
                i = 1
                m.ZOrder (1)
                m.Picture = picCompBack.Picture
                m.Left = m.Left + 200
            ElseIf KeyCode = vbKeyG Then
                aa = True
                m.ZOrder (1)
                m.Picture = picCompPunch.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
            ElseIf KeyCode = vbKeyH Then
                m.ZOrder (1)
                aa = True
                m.Picture = picCompKick.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
            End If
        'assign a multiplayer lan/internet connectivity left player left moves to remote right player
        ElseIf mnuTwoPlayer2Item.Checked = True Then
            mnuSinglePlayerItem.Checked = False
            mnuTwoPlayerItem.Checked = False
            If CurrentGameConnection = "Server" Then
                If KeyCode = vbKeyRight Then
                    ws.SendData "KeyRight-" & (7800 - f.Left)
                ElseIf KeyCode = vbKeyLeft Then
                    ws.SendData "KeyLeft-" & (7800 - f.Left)
                ElseIf KeyCode = vbKeyControl Then
                    ws.SendData "KeyControl"
                ElseIf KeyCode = vbKeyShift Then
                    ws.SendData "KeyShift"
                End If
            ElseIf CurrentGameConnection = "Client" Then
                If KeyCode = vbKeyRight Then
                    ws.SendData "KeyRight2-" & (7800 - f.Left)
                ElseIf KeyCode = vbKeyLeft Then
                    ws.SendData "KeyLeft2-" & (7800 - f.Left)
                ElseIf KeyCode = vbKeyControl Then
                    ws.SendData "KeyControl2"
                ElseIf KeyCode = vbKeyShift Then
                    ws.SendData "KeyShift2"
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    i = 5
    i2 = 5
    If mnuTwoPlayerItem.Checked = True Then
        If lblFighter.ForeColor = vbGreen Then
            If aa = True Then
                If m.Left - f.Left < 1000 Then
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrCompRecover.Enabled = True
                End If
            End If
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
                On Error GoTo health3:
                    w.Value = bb
health3:
    If Err.Number = 380 Then
            bb = 0
            w.Value = b
            tmrPlayerRecover.Enabled = False
            w.Value = bb
            f.Picture = picPlayerLoss.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            m.Left = f.Left + f.Width
            lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
            lblWinner.Caption = aaaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If

            End If
        m.Picture = picCompNone.Picture
        aa = False
    End If
End If
'///////////////////////////////////////////////////////////////////
    If lblFighter.ForeColor = vbGreen Then
        If a = True Then
            If m.Left - f.Left < 1000 Then
                b = b - i
                m.Picture = picCompHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If l.Value - 5 = 0 Then
                tmrCompRecover.Enabled = False
                l.Value = b
                m.Picture = picCompLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                f.Left = m.Left - 2000
                b = 0
                l.Value = b
                lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) + Val(lblPlayerMoney.Caption)
                lblWinner.Caption = aaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health2:
                l.Value = b
health2:
        If Err.Number = 380 Then
            b = 0
            l.Value = b
            tmrCompRecover.Enabled = False
            l.Value = b
            m.Picture = picCompLoss.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            f.Left = m.Left - 2000
            lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) + Val(lblPlayerMoney.Caption)
            lblWinner.Caption = aaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If
            End If
        
        f.Picture = picPlayerNone.Picture
        a = False
    End If
End Sub

Private Sub mnuControlsItem_Click()
    frmControls.Show vbModal
End Sub

Private Sub mnuCreditsItem_Click()
    frmCredits.Show vbModal
End Sub

Private Sub mnuExitItem_Click()
    Unload frmControls
    Unload frmCredits
    Unload Me
End Sub

Private Sub mnuFightItem_Click()
    lblFighter.ForeColor = vbGreen
    If mnuSinglePlayerItem.Checked = True Then
        mnuDifficultyItem.Enabled = False
        tmrComputerAI.Enabled = True
        tmrCompRecover.Enabled = False
        Randomize
        f.Left = 240
        m.Left = 7560
        f.Top = 5280
        m.Top = 5280
        bb = 100
        b = 100
        l.Value = b
        w.Value = bb
        m.Picture = picCompNone.Picture
        f.Picture = picPlayerNone.Picture
        lblWinner.Caption = ""
        tmrRegeneration.Enabled = True
        bb = 100
        b = 100
        w.Value = bb
        l.Value = b
        tmrComputerAI.Enabled = True
        tmrCompRecover.Enabled = False
        i2 = 5
        i = 5
        mnuFightItem.Enabled = True
        mnuDifficultyItem.Enabled = True
    ElseIf mnuTwoPlayerItem.Checked = True Or mnuTwoPlayer2Item.Checked = True Then
        mnuFightItem.Enabled = False
        mnuDifficultyItem.Enabled = False
        tmrComputerAI.Enabled = False
        tmrCompRecover.Enabled = False
        Randomize
        f.Left = 240
        m.Left = 7560
        f.Top = 5280
        m.Top = 5280
        bb = 100
        b = 100
        l.Value = b
        w.Value = bb
        m.Picture = picCompNone.Picture
        f.Picture = picPlayerNone.Picture
        lblWinner.Caption = ""
        bb = 100
        b = 100
        tmrRegeneration.Enabled = True
        w.Value = bb
        l.Value = b
        tmrComputerAI.Enabled = False
        tmrCompRecover.Enabled = False
        i2 = 5
        i = 5
        mnuFightItem.Enabled = True
    End If
End Sub

Private Sub mnuHowtoPlayItem_Click()
    MsgBox "The object of the game is to eliminate your opponnent." & vbCrLf & _
            "And from there you can earn money and can be added to each" & vbCrLf & _
            "player and used the earned money as points that can be used" & vbCrLf & _
            "to buy yourself a new fighting spider according to its price." & vbCrLf & _
            "There are threee (3) set of games in which you can choose" & vbCrLf & _
            "and each of them have different set of game options." & vbCrLf & _
            "Just click on the Game Menu and you may select a game play." & vbCrLf & _
            "There are shortcut keys that you could use for your own convenience." & vbCrLf & _
            "Have fun with this game!!", vbOKOnly + vbInformation, "How to Play:"
End Sub

Private Sub mnuSpiderInformation_Click()
    frmSpiderInfo.Show
End Sub
 
Private Sub mnuNormalItem_Click()
    On Error GoTo ErrorLoad
    CurrentSpiderStage = "Normal"
    Me.Picture = LoadPicture(App.Path + "\Graphics\Default Background.jpg")
    mnuNormalItem.Checked = True
    mnuEasyItem.Checked = False
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = False
    Exit Sub
ErrorLoad:
    MsgBox "Cannot find Default Background!", vbOKOnly + vbCritical, "Warning:"
End Sub

Private Sub mnuEasyItem_Click()
    On Error GoTo ErrorLoad
    CurrentSpiderStage = "Easy"
    Me.Picture = LoadPicture(App.Path + "\Graphics\House Background.jpg")
    mnuNormalItem.Checked = False
    mnuEasyItem.Checked = True
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = False
    Exit Sub
ErrorLoad:
    MsgBox "Cannot find House Background!", vbOKOnly + vbCritical, "Warning:"
End Sub

Private Sub mnuMediumItem_Click()
    On Error GoTo ErrorLoad
    CurrentSpiderStage = "Medium"
    Me.Picture = LoadPicture(App.Path + "\Graphics\Basketball Background.jpg")
    mnuNormalItem.Checked = False
    mnuEasyItem.Checked = False
    mnuMediumItem.Checked = True
    mnuHardItem.Checked = False
    Exit Sub
ErrorLoad:
    MsgBox "Cannot find Basketball Court Background!", vbOKOnly + vbCritical, "Warning:"
End Sub

Private Sub mnuHardItem_Click()
    On Error GoTo ErrorLoad
    CurrentSpiderStage = "Hard"
    Me.Picture = LoadPicture(App.Path + "\Graphics\Playground Background.jpg")
    mnuNormalItem.Checked = False
    mnuEasyItem.Checked = False
    mnuMediumItem.Checked = False
    mnuHardItem.Checked = True
    Exit Sub
ErrorLoad:
    MsgBox "Cannot find Playground Background!", vbOKOnly + vbCritical, "Warning:"
End Sub

Private Sub mnuSinglePlayerItem_Click()
    CurrentGameType = "Singleplayer"
    mnuSinglePlayerItem.Checked = True
    mnuTwoPlayerItem.Checked = False
    mnuTwoPlayer2Item.Checked = False
    mnuDifficultyItem.Enabled = True
    If SingleSelected = False Then
        Do
            aaa = InputBox("Enter the player's name", "Fighter")
            If aaa = "" Then
                MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
            End If
        Loop While aaa = ""
        Do
            aaaa = InputBox("Enter the computer's name", "Fighter")
            If aaaa = "" Then
                MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
            End If
        Loop While aaaa = ""
        SingleSelected = True
    End If
    lblPlayerName.Caption = aaa
    lblCompName.Caption = aaaa
    Me.Caption = "The Spider on the Stick - Single Player vs Computer Mode"
    Call mnuFightItem_Click
End Sub

Private Sub mnuSpiderNormalItem_Click()
    CurrentSpiderType = "Normal"
    mnuSpiderNormalItem.Checked = True
    mnuSpiderDomesticItem.Checked = False
    mnuSpiderCaveItem.Checked = False
    mnuSpiderJungleItem.Checked = False
    picPlayerForward.Picture = picPlayerForwardNormalLeft.Picture
    picPlayerNone.Picture = picPlayerNoneNormalLeft.Picture
    picPlayerBack.Picture = picPlayerBackNormalLeft.Picture
    picPlayerKick.Picture = picPlayerKickNormalLeft.Picture
    picPlayerLoss.Picture = picPlayerLossNormalLeft.Picture
    picPlayerHit.Picture = picPlayerHitNormalLeft.Picture
    picPlayerPunch.Picture = picPlayerPunchNormalLeft.Picture
End Sub

Private Sub mnuSpiderDomesticItem_Click()
    If Val(lblPlayerPoints.Caption) >= 250 Then '500
        lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) - 250 '500
        CurrentSpiderType = "Domestic"
        mnuSpiderNormalItem.Checked = False
        mnuSpiderDomesticItem.Checked = True
        mnuSpiderCaveItem.Checked = False
        mnuSpiderJungleItem.Checked = False
        picPlayerForward.Picture = picPlayerForwardDomesticLeft.Picture
        picPlayerNone.Picture = picPlayerNoneDomesticLeft.Picture
        picPlayerBack.Picture = picPlayerBackDomesticLeft.Picture
        picPlayerKick.Picture = picPlayerKickDomesticLeft.Picture
        picPlayerLoss.Picture = picPlayerLossDomesticLeft.Picture
        picPlayerHit.Picture = picPlayerHitDomesticLeft.Picture
        picPlayerPunch.Picture = picPlayerPunchDomesticLeft.Picture
    End If
End Sub

Private Sub mnuSpiderCaveItem_Click()
    If Val(lblPlayerPoints.Caption) >= 500 Then '750
        lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) - 500 '750
        CurrentSpiderType = "Cave"
        mnuSpiderNormalItem.Checked = False
        mnuSpiderDomesticItem.Checked = False
        mnuSpiderCaveItem.Checked = True
        mnuSpiderJungleItem.Checked = False
        picPlayerForward.Picture = picPlayerForwardCaveLeft.Picture
        picPlayerNone.Picture = picPlayerNoneCaveLeft.Picture
        picPlayerBack.Picture = picPlayerBackCaveLeft.Picture
        picPlayerKick.Picture = picPlayerKickCaveLeft.Picture
        picPlayerLoss.Picture = picPlayerLossCaveLeft.Picture
        picPlayerHit.Picture = picPlayerHitCaveLeft.Picture
        picPlayerPunch.Picture = picPlayerPunchCaveLeft.Picture
    End If
End Sub

Private Sub mnuSpiderJungleItem_Click()
    If Val(lblPlayerPoints.Caption) >= 750 Then '1000
        lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) - 750 '1000
        CurrentSpiderType = "Jungle"
        mnuSpiderNormalItem.Checked = False
        mnuSpiderDomesticItem.Checked = False
        mnuSpiderCaveItem.Checked = False
        mnuSpiderJungleItem.Checked = True
        picPlayerForward.Picture = picPlayerForwardJungleLeft.Picture
        picPlayerNone.Picture = picPlayerNoneJungleLeft.Picture
        picPlayerBack.Picture = picPlayerBackJungleLeft.Picture
        picPlayerKick.Picture = picPlayerKickJungleLeft.Picture
        picPlayerLoss.Picture = picPlayerLossJungleLeft.Picture
        picPlayerHit.Picture = picPlayerHitJungleLeft.Picture
        picPlayerPunch.Picture = picPlayerPunchJungleLeft.Picture
    End If
End Sub

'-------
Private Sub mnuSpiderNormalItem2_Click()
    CurrentSpiderType = "Normal"
    mnuSpiderNormalItem2.Checked = True
    mnuSpiderDomesticItem2.Checked = False
    mnuSpiderCaveItem2.Checked = False
    mnuSpiderJungleItem2.Checked = False
    picCompForward.Picture = picPlayerForwardNormalRight.Picture
    picCompNone.Picture = picPlayerNoneNormalRight.Picture
    picCompBack.Picture = picPlayerBackNormalRight.Picture
    picCompKick.Picture = picPlayerKickNormalRight.Picture
    picCompLoss.Picture = picPlayerLossNormalRight.Picture
    picCompHit.Picture = picPlayerHitNormalRight.Picture
    picCompPunch.Picture = picPlayerPunchNormalRight.Picture
End Sub

Private Sub mnuSpiderDomesticItem2_Click()
    If Val(lblCompPoints.Caption) >= 250 Then '500
        lblCompPoints.Caption = Val(lblCompPoints.Caption) - 250 '500
        CurrentSpiderType = "Domestic"
        mnuSpiderNormalItem2.Checked = False
        mnuSpiderDomesticItem2.Checked = True
        mnuSpiderCaveItem2.Checked = False
        mnuSpiderJungleItem2.Checked = False
        picCompForward.Picture = picPlayerForwardDomesticRight.Picture
        picCompNone.Picture = picPlayerNoneDomesticRight.Picture
        picCompBack.Picture = picPlayerBackDomesticRight.Picture
        picCompKick.Picture = picPlayerKickDomesticRight.Picture
        picCompLoss.Picture = picPlayerLossDomesticRight.Picture
        picCompHit.Picture = picPlayerHitDomesticRight.Picture
        picCompPunch.Picture = picPlayerPunchDomesticRight.Picture
    End If
End Sub

Private Sub mnuSpiderCaveItem2_Click()
    If Val(lblCompPoints.Caption) >= 500 Then '750
        lblCompPoints.Caption = Val(lblCompPoints.Caption) - 500 '750
        CurrentSpiderType = "Cave"
        mnuSpiderNormalItem2.Checked = False
        mnuSpiderDomesticItem2.Checked = False
        mnuSpiderCaveItem2.Checked = True
        mnuSpiderJungleItem2.Checked = False
        picCompForward.Picture = picPlayerForwardCaveRight.Picture
        picCompNone.Picture = picPlayerNoneCaveRight.Picture
        picCompBack.Picture = picPlayerBackCaveRight.Picture
        picCompKick.Picture = picPlayerKickCaveRight.Picture
        picCompLoss.Picture = picPlayerLossCaveRight.Picture
        picCompHit.Picture = picPlayerHitCaveRight.Picture
        picCompPunch.Picture = picPlayerPunchCaveRight.Picture
    End If
End Sub

Private Sub mnuSpiderJungleItem2_Click()
    If Val(lblCompPoints.Caption) >= 750 Then '1000
        lblCompPoints.Caption = Val(lblCompPoints.Caption) - 750 '1000
        CurrentSpiderType = "Jungle"
        mnuSpiderNormalItem2.Checked = False
        mnuSpiderDomesticItem2.Checked = False
        mnuSpiderCaveItem2.Checked = False
        mnuSpiderJungleItem2.Checked = True
        picCompForward.Picture = picPlayerForwardJungleRight.Picture
        picCompNone.Picture = picPlayerNoneJungleRight.Picture
        picCompBack.Picture = picPlayerBackJungleRight.Picture
        picCompKick.Picture = picPlayerKickJungleRight.Picture
        picCompLoss.Picture = picPlayerLossJungleRight.Picture
        picCompHit.Picture = picPlayerHitJungleRight.Picture
        picCompPunch.Picture = picPlayerPunchJungleRight.Picture
    End If
End Sub

Private Sub mnuTwoPlayerItem_Click()
    CurrentGameType = "Multiplayer Keyboard"
    mnuSinglePlayerItem.Checked = False
    mnuTwoPlayerItem.Checked = True
    mnuTwoPlayer2Item.Checked = False
    mnuDifficultyItem.Enabled = False
    If KeyboardSelected = False Then
        Do
        aaa = InputBox("Enter the player's name", "Fighter")
            If aaa = "" Then
                MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
            End If
        Loop While aaa = ""
        Do
        aaaa = InputBox("Enter the second player's name", "Fighter")
            If aaaa = "" Then
                MsgBox ("Sorry, but that was an invalid name" & vbCrLf & "Please enter another name"), vbExclamation, "Fighter"
            End If
        Loop While aaaa = ""
        KeyboardSelected = True
    End If
    lblPlayerName.Caption = aaa
    lblCompName.Caption = aaaa
    Me.Caption = "The Spider on the Stick - Multiplayer Keyboard Mode"
    Call mnuFightItem_Click
End Sub


Private Sub tmrCompRecover_Timer()
    m.Picture = picCompNone.Picture
    tmrCompRecover.Enabled = False
End Sub

Private Sub tmrComputerAI_Timer()
    If lblFighter.ForeColor = vbGreen Then
    'EASY MODE
    If mnuEasyItem.Checked = True Or mnuNormalItem.Checked = True Then
        Dim intAction As Integer
        intAction = Int(7 * Rnd) + 1
        tmrComputerAI.Interval = 500
        i = 5
        'Go Back
        If intAction = 1 Or intAction = 2 Then
            i = 1
            m.Picture = picCompBack.Picture
            If m.Left + 200 > 5000 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left + 200
        'Go Forward
        ElseIf intAction = 3 Or intAction = 4 Then
            m.Picture = picCompForward.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            If m.Left - f.Left < 800 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left - 200
        End If
        'Punch
        If intAction = 5 Or 6 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompPunch.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
            End If
        End If
        'Kick
        If intAction = 7 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompKick.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
            End If
        End If
    End If
    'MEDIUM MODE
        If mnuMediumItem.Checked = True Then
            intAction = Int(6 * Rnd) + 1
            tmrComputerAI.Interval = 100
            i = 5
            'Go Back
            If intAction = 1 Or intAction = 2 Then
                i = 1
                m.Picture = picCompBack.Picture
                    If m.Left + 200 > 5000 Then
                        Exit Sub
                    End If
                f.ZOrder (1)
                m.Picture = picCompForward.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
                m.Left = m.Left + 200
            'Go Forward
            m.Picture = picCompForward.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            ElseIf intAction = 3 Or intAction = 4 Then
                If m.Left - f.Left < 800 Then
                    Exit Sub
                End If
                f.ZOrder (1)
                m.Picture = picCompForward.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
                m.Left = m.Left - 200
            End If
            'Punch
            If intAction = 5 Then
                If m.Left - f.Left < 900 Then
                    m.ZOrder (1)
                    m.Picture = picCompPunch.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrPlayerRecover.Enabled = True
                End If
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                    m.Left = f.Left + f.Width
                    bb = 0
                    w.Value = bb
                    lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                    lblWinner.Caption = aaaa & " Wins!"
                    tmrRegeneration.Enabled = False
                    tmrComputerAI.Enabled = False
                    lblFighter.ForeColor = vbRed
                Else
                On Error GoTo health
                    w.Value = bb
                End If
            End If
            'Kick
            If intAction = 6 Then
                If m.Left - f.Left < 900 Then
                    m.ZOrder (1)
                    m.Picture = picCompKick.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrPlayerRecover.Enabled = True
                End If
                
                If w.Value - 5 = 0 Then
                    tmrPlayerRecover.Enabled = False
                    w.Value = bb
                    f.Picture = picPlayerLoss.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                    m.Left = f.Left + f.Width
                    bb = 0
                    w.Value = bb
                    lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                    lblWinner.Caption = aaaa & " Wins!"
                    tmrRegeneration.Enabled = False
                    tmrComputerAI.Enabled = False
                    lblFighter.ForeColor = vbRed
                Else
                On Error GoTo health:
                    w.Value = bb
                End If
            End If
        End If
    'HARD MODE
        If mnuHardItem.Checked = True Then
            intAction = Int(8 * Rnd) + 1
            tmrComputerAI.Interval = 1
            i = 5
        'Go Back
        If intAction = 1 Then
            i = 1
            m.Picture = picCompBack.Picture
            If m.Left + 200 > 5000 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Picture = picCompForward.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            m.Left = m.Left + 200
        'Go Forward
        ElseIf intAction = 2 Or intAction = 3 Then
            m.Picture = picCompForward.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
            If m.Left - f.Left < 800 Then
                Exit Sub
            End If
            f.ZOrder (1)
            m.Left = m.Left - 200
        End If
        'Punch
        If intAction = 4 Or intAction = 5 Or intAction = 6 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompPunch.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
health:
    If Err.Number = 380 Then
        bb = 0
        w.Value = bb
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
        Resume Next
    End If
            End If
        End If
        'Kick
        If intAction = 7 Or intAction = 8 Then
            If m.Left - f.Left < 900 Then
                m.ZOrder (1)
                m.Picture = picCompKick.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
                bb = bb - i2
                f.Picture = picPlayerHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrPlayerRecover.Enabled = True
            End If
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health:
                w.Value = bb
            End If
        End If
    End If
End If
End Sub

Private Sub tmrMultiplayerLanInternet_Timer()
    'activate multiplayer lan/internet Virtual Key_Up timer
    i = 5
    i2 = 5
    If mnuTwoPlayer2Item.Checked = True Then
        If lblFighter.ForeColor = vbGreen Then
            If aa = True Then
                If m.Left - f.Left < 1000 Then
                    bb = bb - i2
                    f.Picture = picPlayerHit.Picture
                    'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                    tmrCompRecover.Enabled = True
                End If
            End If
            If w.Value - 5 = 0 Then
                tmrPlayerRecover.Enabled = False
                w.Value = bb
                f.Picture = picPlayerLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                m.Left = f.Left + f.Width
                bb = 0
                w.Value = bb
                lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
                lblWinner.Caption = aaaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
                On Error GoTo health3:
                    w.Value = bb
health3:
    If Err.Number = 380 Then
            bb = 0
            w.Value = b
            tmrPlayerRecover.Enabled = False
            w.Value = bb
            f.Picture = picPlayerLoss.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            m.Left = f.Left + f.Width
            lblCompPoints.Caption = Val(lblCompPoints.Caption) + Val(lblCompMoney.Caption)
            lblWinner.Caption = aaaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If
            End If
        m.Picture = picCompNone.Picture
        aa = False
    End If
End If
    If lblFighter.ForeColor = vbGreen Then
        If a = True Then
            If m.Left - f.Left < 1000 Then
                b = b - i
                m.Picture = picCompHit.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Hit.wav", 1)
                tmrCompRecover.Enabled = True
            End If
        End If
    
            If l.Value - 5 = 0 Then
                tmrCompRecover.Enabled = False
                l.Value = b
                m.Picture = picCompLoss.Picture
                'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
                f.Left = m.Left - 2000
                b = 0
                l.Value = b
                lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) + Val(lblPlayerMoney.Caption)
                lblWinner.Caption = aaa & " Wins!"
                tmrRegeneration.Enabled = False
                tmrComputerAI.Enabled = False
                lblFighter.ForeColor = vbRed
            Else
            On Error GoTo health2:
                l.Value = b
health2:
        If Err.Number = 380 Then
            b = 0
            l.Value = b
            tmrCompRecover.Enabled = False
            l.Value = b
            m.Picture = picCompLoss.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Win.wav", 1)
            f.Left = m.Left - 2000
            lblPlayerPoints.Caption = Val(lblPlayerPoints.Caption) + Val(lblPlayerMoney.Caption)
            lblWinner.Caption = aaa & " Wins!"
            tmrRegeneration.Enabled = False
            tmrComputerAI.Enabled = False
            lblFighter.ForeColor = vbRed
        End If
            End If
        f.Picture = picPlayerNone.Picture
        a = False
    End If
End Sub

Private Sub tmrPlayerRecover_Timer()
    f.Picture = picPlayerNone.Picture
    tmrPlayerRecover.Enabled = False
End Sub

Private Sub tmrRegeneration_Timer()
    If w.Value < 100 Then
        bb = bb + 1
        w.Value = bb
        lblPlayerMoney.Caption = Format(bb, "####0")
    End If
    If l.Value < 100 Then
        b = b + 1
        l.Value = b
        lblCompMoney.Caption = Format(b, "####0")
    End If
End Sub

'************************************* Multiplayer-Network Start Code **********************************************

Private Sub mnuTwoPlayer2Item_Click()
    mnuSinglePlayerItem.Checked = False
    mnuTwoPlayerItem.Checked = False
    mnuTwoPlayer2Item.Checked = True
    CurrentGameType = "Multiplayer LAN/Internet"
    Me.Caption = "The Spider on the Stick - Multiplayer LAN/Internet Mode"
    Connectivity.Visible = True
    txtIPAddress.Enabled = True
    txtIPAddress.Text = ws.LocalIP 'makes winsock give you your computer's IP
    txtIPAddress.SetFocus
    wsdata = txtDataSent.Text 'tells winsock that the data you will send will be that of txtDataSent.Text
    txtPlayerName.Enabled = True
    txtDataReceived.Enabled = True
    txtDataSent.Enabled = True
    cmdSendMessage.Enabled = True
End Sub

Private Sub cmdPlay_Click()
    If Trim(txtPlayerName) = "" Then
        MsgBox "Cannot start game without assigning" & _
        vbCrLf & "your own Player's Name!", vbOKOnly + vbCritical, _
        IIf(CurrentGameConnection = "Server", "Warning: Host Player", "Warning: Remote Player")
        txtPlayerName.SetFocus
    Else
        If CurrentGameConnection = "Server" Then
            If GameStatus.Caption = "User Status: Listening..." Then
                MsgBox "Cannot start game without remote opponent!", vbOKOnly + vbCritical, "Warning: Host Player"
                Exit Sub
            Else
                Call PlaySettings
                ws.SendData "Server-" & txtPlayerName
            End If
        ElseIf CurrentGameConnection = "Client" Then
            If GameStatus.Caption = "User Status: Connecting..." Then
                MsgBox "Cannot start game without host opponent!", vbOKOnly + vbCritical, "Warning: Remote Player"
                Exit Sub
            Else
                Call PlaySettings
                ws.SendData "Client-" & txtPlayerName
            End If
        End If
    End If
End Sub

Function PlaySettings()
    Connectivity.Visible = False
    txtIPAddress.Enabled = False
    txtPlayerName.Enabled = False
    txtDataReceived.Enabled = False
    txtDataSent.Enabled = False
    cmdSendMessage.Enabled = False
    StartMultiplayer = True
    Call mnuFightItem_Click
End Function

Private Sub cmdListen_Click()
    CurrentGameConnection = "Server"
    ws.LocalPort = "3999" 'port that your computer will listen to if you press listen
    ws.Listen
    GameStatus.Caption = "User Status: Listening..."
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = False
End Sub

Private Sub cmdConnect_Click()
    CurrentGameConnection = "Client"
    ws.RemoteHost = txtIPAddress.Text 'assigns txtIPAddress.Text to the host
    ws.RemotePort = "3999" 'port to connect to. I have deliberately made it this number to stop possible interference from other apps
    ws.Connect
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmdListen.Enabled = False
    GameStatus.Caption = "User Status: Connecting..."
    txtPlayerName.SetFocus
End Sub

Private Sub cmdDisconnect_Click()
    GameStatus.Caption = "User Status: Disconnected..."
    ws.SendData "DIS"
    ws.Close 'close winsock (disconnect)
    cmdDisconnect.Enabled = False
End Sub

Private Sub cmdClose_Click()
    Connectivity.Visible = False
    txtIPAddress.Enabled = False
    txtPlayerName.Enabled = False
    txtDataReceived.Enabled = False
    txtDataSent.Enabled = False
    cmdSendMessage.Enabled = False
End Sub

Private Sub ws_Connect()
    GameStatus.Caption = "User Status: Connected to Host Player..."
    ws.SendData "CON"
    cmdConnect.Enabled = False
    cmdDisconnect.Enabled = True
    cmdListen.Enabled = False
End Sub

Private Sub ws_ConnectionRequest(ByVal requestID As Long)
    ws.Close
    ws.Accept requestID 'accepts the other person and allows the two programs to connect to each other ready for game
End Sub

'DATA DECIPHERING
Private Sub ws_DataArrival(ByVal bytesTotal As Long)
    ws.GetData wsdata 'get the data (i named this wsdata...see top)
    If Left(wsdata, 4) = "TALK" Then
        Dim str1
        str1 = Right(wsdata, Len(wsdata) - 4) 'get rid of the 'chat' part
        txtDataReceived.Text = txtDataReceived.Text & "OTHER PLAYER" & ": " & str1 & vbCrLf 'displays all the necessary info in the txtDataReceived textbox
        txtDataReceived.SelStart = Len(txtDataReceived.Text)
    End If
    If Left(wsdata, 3) = "DIS" Then
        GameStatus.Caption = "User Status: The Remote Player has closed their Connection"
        cmdDisconnect.Enabled = False
        ws.Close
    End If
    If Left(wsdata, 3) = "CON" Then
        GameStatus.Caption = "User Status: The Remote Player is Connected"
        cmdConnect.Enabled = False
        cmdDisconnect.Enabled = True
        cmdListen.Enabled = False
    End If
    'host moves via winsock transfered to right player
    If Left(wsdata, 8) = "KeyRight" Then
        m.Picture = picCompForward.Picture
        'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
        If m.Left - f.Left < 1000 Then
            Exit Sub
        Else
            m.ZOrder (1)
            m.Left = Val(ExtractArgument(2, wsdata, "-"))
        End If
    End If
    If Left(wsdata, 7) = "KeyLeft" Then
        If m.Left + 300 >= 7800 Then
            Exit Sub
        End If
            i = 1
            m.ZOrder (1)
            m.Picture = picCompBack.Picture
            m.Left = Val(ExtractArgument(2, wsdata, "-"))
    End If
    If Left(wsdata, 10) = "KeyControl" Then
            aa = True
            m.ZOrder (1)
            m.Picture = picCompPunch.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
    End If
    If Left(wsdata, 8) = "KeyShift" Then
        m.ZOrder (1)
        aa = True
        m.Picture = picCompKick.Picture
        'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
    End If
    'remote moves via winsock transfered to left player
    If Left(wsdata, 9) = "KeyRight2" Then
        m.Picture = picCompForward.Picture
        'Call sndPlaySound(App.Path & "\Sounds\Forward.wav", 1)
        If m.Left - f.Left < 1000 Then
            Exit Sub
        Else
            m.ZOrder (1)
            m.Left = Val(ExtractArgument(2, wsdata, "-"))
        End If
    End If
    If Left(wsdata, 8) = "KeyLeft2" Then
        If m.Left + 300 >= 7800 Then
            Exit Sub
        End If
            i = 1
            m.ZOrder (1)
            m.Picture = picCompBack.Picture
            m.Left = Val(ExtractArgument(2, wsdata, "-"))
    End If
    If Left(wsdata, 11) = "KeyControl2" Then
            aa = True
            m.ZOrder (1)
            m.Picture = picCompPunch.Picture
            'Call sndPlaySound(App.Path & "\Sounds\Punch.wav", 1)
    End If
    If Left(wsdata, 9) = "KeyShift2" Then
        m.ZOrder (1)
        aa = True
        m.Picture = picCompKick.Picture
        'Call sndPlaySound(App.Path & "\Sounds\Kick.wav", 1)
    End If
    If Left(wsdata, 6) = "Server" Then
        lblPlayerName.Caption = txtPlayerName
        lblCompName.Caption = ExtractArgument(2, wsdata, "-")
        aaa = txtPlayerName
        aaaa = ExtractArgument(2, wsdata, "-")
    End If
    If Left(wsdata, 6) = "Client" Then
        lblPlayerName.Caption = txtPlayerName
        lblCompName.Caption = ExtractArgument(2, wsdata, "-")
        aaa = txtPlayerName
        aaaa = ExtractArgument(2, wsdata, "-")
    End If
End Sub

Private Sub txtDataSent_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then 'if enter is pressed
        Call cmdSendMessage_Click 'go to the cmdSendMessage event code
    End If 'this makes it easier to send the text as you don't have to keep pressing the button
End Sub

Private Sub cmdSendMessage_Click()
    ws.SendData "TALK" & txtDataSent.Text 'send some text to alert the data arrival of the prescence of text, followed by the actual text
    txtDataReceived.Text = txtDataReceived.Text & "ME" & ": " & txtDataSent.Text & vbCrLf 'displays your text in your textbox
    txtDataReceived.SelStart = Len(txtDataReceived.Text)
    txtDataSent.Text = "" 'clears the speech box ready for the next speech
    txtDataSent.SetFocus
End Sub

Function ExtractArgument(ArgNum As Integer, srchstr As String, Delim As String) As String
    'Extract an argument or token from a string based on its position and a delimiter.
    On Error GoTo Err_ExtractArgument
    Dim ArgCount As Integer
    Dim LastPos As Integer
    Dim Pos As Integer
    Dim Arg As String
    Arg = ""
    LastPos = 1
    If ArgNum = 1 Then Arg = srchstr
        Do While InStr(srchstr, Delim) > 0
            Pos = InStr(LastPos, srchstr, Delim)
        If Pos = 0 Then
            'No More Args found
            If ArgCount = ArgNum - 1 Then Arg = Mid(srchstr, LastPos)
            Exit Do
        Else
            ArgCount = ArgCount + 1
            If ArgCount = ArgNum Then
                Arg = Mid(srchstr, LastPos, Pos - LastPos)
                Exit Do
            End If
        End If
        LastPos = Pos + 1
    Loop
    '---------
    ExtractArgument = Arg
    Exit Function
Err_ExtractArgument:
    MsgBox "Error " & Err & ": " & Error
    Resume Next
End Function

 
'************************************* Multiplayer-Network End Code **********************************************
