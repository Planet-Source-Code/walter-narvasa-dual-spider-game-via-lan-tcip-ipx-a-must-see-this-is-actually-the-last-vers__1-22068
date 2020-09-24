VERSION 5.00
Begin VB.Form frmCredits 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "About The Spider on the Stick"
   ClientHeight    =   6030
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   8205
   Icon            =   "frmCredits.frx":0000
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
      Left            =   1200
      Top             =   2700
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&Ok"
      Height          =   330
      Left            =   3480
      TabIndex        =   14
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
         TabIndex        =   11
         Top             =   0
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
            Caption         =   "A Spider Dual Fighting Game"
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
            Height          =   420
            Left            =   120
            TabIndex        =   13
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
            TabIndex        =   12
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
         Height          =   7275
         Left            =   120
         ScaleHeight     =   7275
         ScaleWidth      =   7830
         TabIndex        =   1
         Top             =   1080
         Width           =   7830
         Begin VB.Line Line2 
            BorderColor     =   &H000000FF&
            Index           =   4
            X1              =   60
            X2              =   1725
            Y1              =   4335
            Y2              =   4335
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Disclaimer"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Index           =   4
            Left            =   165
            TabIndex        =   16
            Top             =   4065
            Width           =   1005
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCredits.frx":000C
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
            Height          =   525
            Index           =   5
            Left            =   360
            TabIndex        =   15
            Top             =   4440
            Width           =   7215
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FFFF00&
            Index           =   3
            X1              =   75
            X2              =   1740
            Y1              =   3075
            Y2              =   3075
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCredits.frx":00CB
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   480
            Index           =   4
            Left            =   360
            TabIndex        =   10
            Top             =   3240
            Width           =   7215
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Information"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFF00&
            Height          =   240
            Index           =   3
            Left            =   210
            TabIndex        =   9
            Top             =   2820
            Width           =   1095
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00FF00FF&
            Index           =   2
            X1              =   0
            X2              =   1665
            Y1              =   5490
            Y2              =   5490
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCredits.frx":0164
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   690
            Index           =   3
            Left            =   330
            TabIndex        =   8
            Top             =   5640
            Width           =   6975
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Redistribution"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF00FF&
            Height          =   240
            Index           =   2
            Left            =   105
            TabIndex        =   7
            Top             =   5220
            Width           =   1320
         End
         Begin VB.Line Line2 
            BorderColor     =   &H0000FF00&
            Index           =   1
            X1              =   105
            X2              =   1770
            Y1              =   1590
            Y2              =   1590
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   $"frmCredits.frx":01ED
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   825
            Index           =   2
            Left            =   390
            TabIndex        =   6
            Top             =   1755
            Width           =   7335
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Comments"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H0000FF00&
            Height          =   240
            Index           =   1
            Left            =   180
            TabIndex        =   5
            Top             =   1320
            Width           =   1005
         End
         Begin VB.Line Line2 
            BorderColor     =   &H000080FF&
            Index           =   0
            X1              =   60
            X2              =   1725
            Y1              =   330
            Y2              =   330
         End
         Begin VB.Label Label1 
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Email: jawoltze@edsamail.com.ph / Website: http://jawoltze.gq.nu/"
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
            Height          =   240
            Index           =   1
            Left            =   360
            TabIndex        =   4
            Top             =   720
            Width           =   6495
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Walter A. Narvasa"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   3
            Top             =   480
            Width           =   1425
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Developed By "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000080FF&
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   2
            Top             =   60
            Width           =   1380
         End
      End
   End
End
Attribute VB_Name = "frmCredits"
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
