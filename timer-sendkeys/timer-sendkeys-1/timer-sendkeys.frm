VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "timer-sendkeys"
   ClientHeight    =   7125
   ClientLeft      =   225
   ClientTop       =   855
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleMode       =   0  'User
   ScaleWidth      =   7110
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   735
      Left            =   4800
      TabIndex        =   12
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CheckBox Check1 
      Caption         =   "shutdown"
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
      Left            =   960
      TabIndex        =   11
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   6360
   End
   Begin VB.CommandButton Command2 
      Caption         =   "shutdown"
      Height          =   615
      Left            =   4920
      TabIndex        =   10
      Top             =   6000
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   8
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   5040
      Width           =   855
   End
   Begin VB.TextBox Text1 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1200
      TabIndex        =   4
      Top             =   5040
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   3480
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   345
      TabIndex        =   3
      Top             =   2880
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Scrolling       =   1
   End
   Begin VB.Label Label10 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   16
      Top             =   120
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "period is "
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   0
      TabIndex        =   15
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label8 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   14
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label7 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   13
      Top             =   3000
      Width           =   615
   End
   Begin VB.Label Label6 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4920
      TabIndex        =   9
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label5 
      Caption         =   "min"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Top             =   5160
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "h"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   5160
      Width           =   375
   End
   Begin VB.Label Label3 
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
      Left            =   4440
      TabIndex        =   2
      Top             =   1920
      Width           =   615
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   48
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1440
      TabIndex        =   1
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Next one sending-keys:"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
   Begin VB.Menu setting 
      Caption         =   "setting(&s)"
      Index           =   1
      Begin VB.Menu reset 
         Caption         =   "reset"
      End
      Begin VB.Menu key 
         Caption         =   "key sending"
      End
      Begin VB.Menu period 
         Caption         =   "period"
      End
   End
   Begin VB.Menu exit 
      Caption         =   "Exit(&E)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Double
Dim b As Double
Dim c As Double
Dim hour, min, sec, time As Integer

Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Height = 8000
Else
Form1.Height = 5430
End If
End Sub

Private Sub Command1_Click()
If Command1.Caption = "start" Then
y = MsgBox("Your setting is sending " & b & " per " & c & "s", vbInformation + vbOKCancel, "Cheking box")
    If y = vbOK Then
    Timer1.Enabled = True
    Command1.Caption = "pause"
    period.Enabled = False
    key.Enabled = False
    reset.Enabled = False
    Label10.Caption = c
    End If
Else
Timer1.Enabled = False
Command1.Caption = "start"
Timer1.Enabled = False
period.Enabled = True
key.Enabled = True
reset.Enabled = True
End If
ProgressBar1.Max = c
End Sub

Private Sub Command2_Click()
Text1.Enabled = Not Text1.Enabled
Text2.Enabled = Not Text2.Enabled
Text3.Enabled = Not Text3.Enabled
Timer2.Enabled = Not Timer2.Enabled
Text1.Text = Val(Text1.Text)
hour = Val(Text1.Text)
Text2.Text = Val(Text2.Text)
min = Val(Text2.Text)
Text3.Text = Val(Text3.Text)
sec = Val(Text3.Text)
time = hour * 3600 + min * 60 + sec
If Command2.Caption = "shutdown" Then
Command2.Caption = "cancel shutdown"
Else
Command2.Caption = "shutdown"
End If
End Sub

Private Sub exit_Click()
Unload Form1
End Sub

Private Sub Form_Load()
a = 0
b = 9
c = 1803
Form1.Height = 5430
End Sub

Private Sub key_Click()
b = InputBox("Please input a key", "Inputbox")
End Sub

Private Sub period_Click()
c = Val(InputBox("Please input the period", "Inputbox"))
Do Until c > 0
z = MsgBox("please input a number", vbCritical + vbOKCancel, "error")
    If z = vbOK Then
    c = Val(InputBox("Please input the period", "Inputbox"))
    End If
Loop
End Sub

Private Sub reset_Click()
a = 0
End Sub

Private Sub Timer1_Timer()
a = a + 1
Label2.Caption = c - a
ProgressBar1.Value = a
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
End If
For i = 0 To 5
If a Mod c = 6 * i Then
SendKeys b
End If
If a Mod c = 6 * i + 3 Then
SendKeys 0
End If
Next
End Sub

Private Sub Timer2_Timer()
time = time - 1
sec = sec - 1
If sec <= 0 Then
    If min > 0 Then
    sec = sec + 60
    min = min - 1
    Else
        If hour > 0 Then
        min = min + 60
        hour = hour - 1
        Else
        min = 0
        sec = 0
        End If
    End If
End If
Text1.Text = hour
Text2.Text = min
Text3.Text = sec
If time = 0 Then
Shell ("shutdown -s -t 5")
End If
End Sub
