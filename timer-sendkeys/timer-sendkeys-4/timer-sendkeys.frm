VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "timer-sendkeys"
   ClientHeight    =   7125
   ClientLeft      =   1530
   ClientTop       =   1635
   ClientWidth     =   7110
   LinkTopic       =   "Form1"
   ScaleHeight     =   7125
   ScaleMode       =   0  'User
   ScaleWidth      =   7110
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   600
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
      Left            =   1680
      TabIndex        =   17
      Top             =   1560
      Width           =   2535
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
      Left            =   1920
      TabIndex        =   16
      Top             =   3480
      Width           =   1935
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   735
      Left            =   4800
      TabIndex        =   11
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
      Left            =   360
      TabIndex        =   10
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
      TabIndex        =   9
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
      TabIndex        =   7
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
      TabIndex        =   5
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
      TabIndex        =   3
      Top             =   5040
      Width           =   855
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   6480
      Top             =   0
   End
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   375
      Left            =   345
      TabIndex        =   2
      Top             =   2880
      Width           =   5295
      _ExtentX        =   9340
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
      TabIndex        =   15
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
      TabIndex        =   14
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
      Left            =   6600
      TabIndex        =   13
      Top             =   3000
      Width           =   255
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
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
      Left            =   5640
      TabIndex        =   12
      Top             =   3000
      Width           =   855
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
      TabIndex        =   8
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
      TabIndex        =   6
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
      TabIndex        =   4
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
      TabIndex        =   1
      Top             =   1920
      Width           =   615
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
      Caption         =   "Setting(&s)"
      Begin VB.Menu gs 
         Caption         =   "General setting"
      End
      Begin VB.Menu ac 
         Caption         =   "Automatic curing"
      End
   End
   Begin VB.Menu reset 
      Caption         =   "Reset(&R)"
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
Dim a, b, c, e, f As Double
Dim hour, min, sec, time As Double

Private Sub ac_Click()
Form3.Visible = True
End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then
Form1.Height = 8000
Else
Form1.Height = 5430
End If
End Sub

Private Sub Command1_Click()
c = Val(form2.Text1.Text)
If Command1.Caption = "start" Then
    If a = 0 Then
    y = MsgBox("Your setting is sending " & b & " per " & c & "s", vbInformation + vbOKCancel, "Cheking box")
        If y = vbOK Then
        Timer1.Enabled = True
        Command1.Caption = "pause"
        reset.Enabled = False
        form2.Text1.Enabled = False
        Label10.Caption = c
        Text4.Enabled = False
            If Text4.Text = "" Then
            Text4.Text = c
            Else
            a = c - Abs(Val(Text4.Text))
            End If
        End If
    Else
        Timer1.Enabled = True
        Command1.Caption = "pause"
        reset.Enabled = False
        form2.Text1.Enabled = False
        Label10.Caption = c
        Text4.Enabled = False
        a = c - Val(Text4.Text)
    End If
Else
Timer1.Enabled = False
Command1.Caption = "start"
reset.Enabled = True
form2.Text1.Enabled = True
Text4.Enabled = True
End If
If Form3.Visible = True Then
Form3.Timer2.Enabled = True
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
Unload Form3
Unload form2
Unload Form1
End Sub

Private Sub Form_Load()
a = 0
b = 9
c = 1803
Form1.Height = 5430
End Sub


Private Sub Form_Unload(Cancel As Integer)
Unload Form3
Unload form2
Unload Form1
End Sub

Private Sub gs_Click()
form2.Visible = True
End Sub

Private Sub reset_Click()
a = 0
Text4.Text = ""
Label7.Caption = ""
ProgressBar1.Value = 0
End Sub



Private Sub Text4_Change()
s1 = c - a
h1 = s1 \ 3600
s2 = s1 Mod 3600
m1 = s2 \ 60
s3 = s2 Mod 60
Label3.ToolTipText = h1 & "hr " & m1 & "min " & s3 & "sec"
End Sub

Private Sub Timer1_Timer()
a = a + 1
Text4.Text = c - a
f = Val(form2.Text4.Text)
If a <= c And a >= 0 Then
ProgressBar1.Value = a
Else
ProgressBar1.Value = 0
End If
d = a / c * 10000 \ 1
d = d / 100
If d < 1 And d >= 0 Then
Label7.Caption = "0" & d
Else
Label7.Caption = d
End If
If a = c Then
a = 0
b = form2.Text2.Text
e = form2.Text3.Text
SendKeys b
    If Check2.Value = 1 Then
    zzz = MsgBox("please check and you get 5 seconds", vbInformation + vbOKOnly, "Message")
    Timer1.Enabled = False
    Timer3.Enabled = True
    End If
End If
For i = 0 To f
If a Mod c = 6 * i Then
b = form2.Text2.Text
SendKeys b
End If
If a Mod c = 6 * i + 3 Then
e = form2.Text3.Text
SendKeys e
End If
Next
End Sub

Private Sub Timer2_Timer()
time = time - 1
sec = sec - 1
If sec < 0 Then
sec = sec + 60
min = min - 1
    If min < 0 Then
    min = min + 60
    hour = hour - 1
        If hour < 0 Then
        min = 0
        sec = 0
        hour = 0
        End If
    End If
End If
Text1.Text = hour
Text2.Text = min
Text3.Text = sec
If time = 0 And Check1.Value = 1 Then
Shell ("shutdown -s -t 5")
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
