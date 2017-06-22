VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Automatic curing"
   ClientHeight    =   3810
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   9480
   LinkTopic       =   "Form3"
   ScaleHeight     =   3810
   ScaleWidth      =   9480
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   4560
      Top             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   735
      Left            =   5880
      TabIndex        =   7
      Top             =   2760
      Width           =   1935
   End
   Begin VB.CheckBox Check2 
      Caption         =   "msgbox"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   1320
      TabIndex        =   6
      Top             =   2760
      Width           =   1935
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1065
      Left            =   5400
      TabIndex        =   5
      Top             =   480
      Width           =   2535
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "3"
      Top             =   480
      Width           =   615
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "-"
      Top             =   1080
      Width           =   615
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   3720
      Top             =   2880
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   2040
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label8 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   9000
      TabIndex        =   9
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label5 
      Caption         =   "Next one sending-keys:"
      Height          =   375
      Left            =   2880
      TabIndex        =   12
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   8160
      TabIndex        =   11
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   10
      Top             =   2160
      Width           =   855
   End
   Begin VB.Label Label3 
      Caption         =   "hr"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   600
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Period:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   480
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Key:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   2
      Top             =   1080
      Width           =   855
   End
   Begin VB.Menu reset 
      Caption         =   "Reset(&R)"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit(&E)"
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim b As String
Dim a, c, t As Double

Private Sub Command1_Click()
b = Text2.Text
If Command1.Caption = "start" Then
        Timer1.Enabled = True
        Command1.Caption = "pause"
        Text4.Enabled = False
        Text1.Enabled = False
            If Text4.Text = "" Then
            Text4.Text = c
            Else
            a = c - Val(Text4.Text)
            End If
Else
Timer1.Enabled = False
Command1.Caption = "start"
Text4.Enabled = True
Text1.Enabled = True
End If
ProgressBar1.Max = c
End Sub

Private Sub exit_Click()
Unload Form3
End Sub

Private Sub Form_Activate()
z = Val(Text1.Text)
c = z * 3600
End Sub


Private Sub reset_Click()
a = 0
Text4.Text = ""
ProgressBar1.Value = 0
End Sub

Private Sub Text1_Change()
aa = Val(Text1.Text)
c = aa * 3600
End Sub

Private Sub Text2_Change()
b = Text2.Text
End Sub

Private Sub Timer1_Timer()
a = a + 1
Text4.Text = c - a
If a <= c And a >= 0 Then
ProgressBar1.Value = a
Else
ProgressBar1.Value = 0
End If
d = a / c * 10000 \ 1
d = d / 100
If d < 1 Then
Label7.Caption = "0" & d
Else
Label7.Caption = d
End If
If a Mod c = 0 Then
a = 0
SendKeys b
    If Check2.Value = 1 Then
    zzz = MsgBox("please check and you get 5 seconds", vbInformation + vbOKOnly, "Message")
    Timer1.Enabled = False
    Timer3.Enabled = True
    End If
End If
End Sub

Private Sub Timer3_Timer()
a = a + 1
Text4.Text = 5 - a
d = a / 5 * 10000 \ 1
d = d / 100
Label7.Caption = d
If a = 5 Then
Timer3.Enabled = False
Timer1.Enabled = True
a = 0
End If

End Sub
