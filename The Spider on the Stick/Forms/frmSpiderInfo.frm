VERSION 5.00
Begin VB.Form frmSpiderInfo 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Spider Information"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   Icon            =   "frmSpiderInfo.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   8205
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Timer timMain 
      Interval        =   100
      Left            =   0
      Top             =   5640
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   3480
      TabIndex        =   7
      Top             =   5640
      Width           =   1215
   End
   Begin VB.PictureBox picOut 
      BackColor       =   &H00000000&
      Height          =   5490
      Left            =   50
      ScaleHeight     =   5430
      ScaleWidth      =   8055
      TabIndex        =   0
      Top             =   50
      Width           =   8115
      Begin VB.PictureBox picUp 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1035
         Left            =   0
         ScaleHeight     =   1035
         ScaleWidth      =   8010
         TabIndex        =   4
         Top             =   -30
         Width           =   8010
         Begin VB.Line Line3 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   120
            X2              =   7920
            Y1              =   960
            Y2              =   960
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "A brief Spider Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   300
            Left            =   0
            TabIndex        =   6
            Top             =   660
            Width           =   7905
         End
         Begin VB.Label lblMain 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The Spider on the Stick"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FFFF&
            Height          =   345
            Left            =   120
            TabIndex        =   5
            Top             =   120
            Width           =   7890
         End
         Begin VB.Line Line1 
            BorderColor     =   &H0000FFFF&
            BorderWidth     =   2
            X1              =   90
            X2              =   7920
            Y1              =   600
            Y2              =   600
         End
      End
      Begin VB.PictureBox picIn 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8355
         Left            =   120
         ScaleHeight     =   8355
         ScaleWidth      =   7815
         TabIndex        =   1
         Top             =   1080
         Width           =   7815
         Begin VB.PictureBox picPlayerForwardJungleRight 
            AutoSize        =   -1  'True
            Height          =   1530
            Left            =   120
            Picture         =   "frmSpiderInfo.frx":000C
            ScaleHeight     =   1470
            ScaleWidth      =   1815
            TabIndex        =   17
            Top             =   6600
            Width           =   1875
         End
         Begin VB.PictureBox picPlayerForwardCaveRight 
            AutoSize        =   -1  'True
            Height          =   1530
            Left            =   120
            Picture         =   "frmSpiderInfo.frx":07F1
            ScaleHeight     =   1470
            ScaleWidth      =   1815
            TabIndex        =   13
            Top             =   4560
            Width           =   1875
         End
         Begin VB.PictureBox picPlayerForwardDomesticRight 
            AutoSize        =   -1  'True
            Height          =   1530
            Left            =   120
            Picture         =   "frmSpiderInfo.frx":0FD7
            ScaleHeight     =   1470
            ScaleWidth      =   1815
            TabIndex        =   9
            Top             =   2520
            Width           =   1875
         End
         Begin VB.PictureBox picPlayerForwardNormalRight 
            AutoSize        =   -1  'True
            Height          =   1530
            Left            =   120
            Picture         =   "frmSpiderInfo.frx":17B9
            ScaleHeight     =   1470
            ScaleWidth      =   1815
            TabIndex        =   8
            Top             =   480
            Width           =   1875
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSpiderInfo.frx":1FED
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   1170
            Index           =   3
            Left            =   2160
            TabIndex        =   16
            Top             =   6600
            Width           =   5610
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The Jungle Spider"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00E0E0E0&
            Height          =   240
            Index           =   3
            Left            =   120
            TabIndex        =   15
            Top             =   6240
            Width           =   1755
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSpiderInfo.frx":20FB
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   1170
            Index           =   2
            Left            =   2160
            TabIndex        =   14
            Top             =   4560
            Width           =   5610
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The Cave Spider"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000C000&
            Height          =   240
            Index           =   2
            Left            =   120
            TabIndex        =   12
            Top             =   4200
            Width           =   1590
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSpiderInfo.frx":2220
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   1170
            Index           =   0
            Left            =   2160
            TabIndex        =   11
            Top             =   2520
            Width           =   5610
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The Domestic Spider"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   1
            Left            =   120
            TabIndex        =   10
            Top             =   2160
            Width           =   1980
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmSpiderInfo.frx":2313
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   1170
            Index           =   1
            Left            =   2160
            TabIndex        =   3
            Top             =   480
            Width           =   5610
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "The Default Spider"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   -1  'True
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   2
            Top             =   120
            Width           =   1785
         End
      End
   End
End
Attribute VB_Name = "frmSpiderInfo"
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
Dim CurScheme As Integer
Dim EasterFlag As Boolean

Private Sub cmdOK_Click()
Unload Me
End Sub

Private Sub Form_Load()
CurScheme = 2
EasterFlag = False
timMain.Interval = 30
picIn.Top = picOut.ScaleHeight + 20
End Sub

Private Sub Image1_DblClick()
If EasterFlag = False Then
    MsgBox "OK, you cracked one easter egg, one more is there to crack", vbInformation + vbOKOnly, "Icon Hunter"
    Image1.ToolTipText = "Don't right click me please"
    
End If

EasterFlag = True
    CurScheme = CurScheme + 1
    If CurScheme = 5 Then CurScheme = 1
    ChangeState CurScheme
End Sub

Private Sub Image1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then
    If EasterFlag = True Then
        frmAuthor.Show vbModal
    End If
End If
End Sub

Private Sub Label1_Click(Index As Integer)
picIn_Click
End Sub

Private Sub Label2_Click()
picIn_Click
End Sub


Private Sub lblCap_Click(Index As Integer)
picIn_Click
End Sub


Private Sub lblMain_Click()
picIn_Click
End Sub


Private Sub picIn_Click()
timMain.Enabled = Not timMain.Enabled
End Sub

Private Sub timMain_Timer()

picIn.Top = picIn.Top - 10
If picIn.Top + picIn.Height < picUp.Height + picUp.Top Then picIn.Top = picOut.ScaleHeight + 20


If EasterFlag = False Then Exit Sub

If picIn.Top = picOut.ScaleHeight + 20 Then
    ChangeState CurScheme
    CurScheme = CurScheme + 1
    If CurScheme = 5 Then CurScheme = 1
End If

End Sub


Sub ChangeState(State As Integer)

Select Case State
    Case 1
        Dim myC As Control
        
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbGreen
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbRed
            End If
        Next myC
     Case 2
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbWhite
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbBlack
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbBlack
            End If
        Next myC

     Case 3
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbRed
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbGreen
            End If
        Next myC

     Case 4
        For Each myC In Me.Controls
            If TypeOf myC Is PictureBox Then
                myC.BackColor = vbBlack
            ElseIf TypeOf myC Is Label Then
                myC.ForeColor = vbWhite
            ElseIf TypeOf myC Is Line Then
                myC.BorderColor = vbWhite
            End If
        Next myC

End Select

End Sub
