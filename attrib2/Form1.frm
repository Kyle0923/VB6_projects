VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "杀毒"
   ClientHeight    =   4440
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9585
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4440
   ScaleWidth      =   9585
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Left            =   5640
      Top             =   2520
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      Height          =   405
      Left            =   2880
      TabIndex        =   19
      Text            =   "00"
      Top             =   2520
      Width           =   1095
   End
   Begin VB.CheckBox Check1 
      Caption         =   "use timer to run"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   18
      Top             =   2400
      Width           =   2415
   End
   Begin VB.OptionButton o1 
      Caption         =   "完成后打开"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   17
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
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
      Left            =   240
      TabIndex        =   16
      Text            =   "请输入盘符"
      Top             =   360
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "确定盘符"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   960
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar p2 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3840
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   3000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "反选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "选项2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Width           =   2655
      Begin VB.CheckBox cre 
         Caption         =   "rd recycled"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox ca 
         Caption         =   "del autorun.inf"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox ce 
         Caption         =   "del *.exe"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "反选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全选"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox cr 
      Caption         =   "-r"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox cs 
      Caption         =   "-s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox ch 
      Caption         =   "-h"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "选项1"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2055
      Left            =   2280
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   3240
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   4000
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4095
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4080
      TabIndex        =   20
      Top             =   2640
      Width           =   375
   End
   Begin VB.Menu look 
      Caption         =   "视图"
      Begin VB.Menu s 
         Caption         =   "最小化"
         Shortcut        =   ^S
      End
      Begin VB.Menu h 
         Caption         =   "隐藏"
         Shortcut        =   ^H
      End
   End
   Begin VB.Menu esc 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Single
Dim t, t0, t1 As Integer
Dim cmd, cmd1, cmd2 As String
Dim p As String

Private Sub ce_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
msg = MsgBox("可能会误删文件(*.exe)是否继续?", vbInformation + vbYesNo, "确认")
If msg = vbNo Then
ce.Value = 0
Else
ce.Value = 1
End If
End Sub


Private Sub cmdstart_Click()
a = 0
If Check1.Value = 1 Then
Timer1.Interval = 1000
Timer1.Enabled = True
Else
Timer1.Interval = 0
Timer1.Enabled = False
End If
If cmd1 = "" Then
cmd1 = "attrib "
End If
If cs.Value = 1 Then
cmd1 = cmd1 & "-s "
End If
For i = 1 To 1000
a = a + 1
p1.Value = a
Next
If cr.Value = 1 Then
cmd1 = cmd1 & "-r "
End If
For i = 1 To 1000
a = a + 1
p1.Value = a
Next
If ch.Value = 1 Then
cmd1 = cmd1 & "-h "
End If
For i = 1 To 1000
a = a + 1
p1.Value = a
Next
cmd1 = cmd1 & "/s /d"
run = Shell("cmd /c " & cmd1)
For i = 1 To 1000
a = a + 1
p1.Value = a
Next
b = 0
If ce.Value = 1 Then
run = Shell("cmd /c del " & p & "*.exe")
End If
For i = 1 To 1000
b = b + 1
p2.Value = b
Next
If ca.Value = 1 Then
run = Shell("cmd /c del " & p & "autorun.inf")
End If
For i = 1 To 1000
b = b + 1
p2.Value = b
Next
If cre.Value = 1 Then
run = Shell("cmd /c Rd " & p & "recycled /s /q")
End If
For i = 1 To 1000
b = b + 1
p2.Value = b
Next
If Check1.Value = 0 Then
msg = MsgBox("杀毒完成！", vbInformation + vbOKOnly, "OK")
End If
If o1.Value = True Then
t1 = t1 + 1
If t1 = 1 Then
run = Shell("cmd /c start " & p)
End If
End If
End Sub

Private Sub Command1_Click()
cs.Value = 1
cr.Value = 1
ch.Value = 1
End Sub

Private Sub Command2_Click()
cs.Value = Abs(cs.Value - 1)
cr.Value = Abs(cr.Value - 1)
ch.Value = Abs(ch.Value - 1)
End Sub

Private Sub Command3_Click()
ca.Value = 1
cre.Value = 1
msg = MsgBox("可能会误删文件(*.exe)是否继续?", vbInformation + vbYesNo, "确认")
If msg = vbNo Then
ce.Value = 0
Else
ce.Value = 1
End If
End Sub

Private Sub Command4_Click()
ce.Value = Abs(ce.Value - 1)
ca.Value = Abs(ca.Value - 1)
cre.Value = Abs(cre.Value - 1)
End Sub

Private Sub Command5_Click()
Text1.Text = Trim(Text1.Text)
p = Text1.Text
If Len(p) = 2 And Right(p, 1) = ":" Then
cmd1 = "attrib " & p & " "
o1.Visible = True
o1.Caption = "完成后打开" & p
Else
msg = MsgBox("请输入盘符", vbCritical, "error")
Text1.Text = ""
End If
End Sub


Private Sub esc_Click()
Unload Form1
End Sub



Private Sub h_Click()
Me.Visible = False
End Sub

Private Sub s_Click()
Me.WindowState = 1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5_Click
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text1.Text = "请输入盘符" Then
Text1.Text = ""
End If
End Sub

Private Sub Text2_LostFocus()
t0 = Val(Text2.Text)
End Sub

Private Sub Timer1_Timer()
t = Val(Text2.Text)
t = t - 1
If t <= 0 Then
cmdstart_Click
t = t0
End If
Text2.Text = t
End Sub
