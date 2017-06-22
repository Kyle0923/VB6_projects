VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "杀毒"
   ClientHeight    =   4260
   ClientLeft      =   150
   ClientTop       =   840
   ClientWidth     =   9585
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   9585
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command5 
      Caption         =   "确定盘符"
      Height          =   495
      Left            =   240
      TabIndex        =   20
      Top             =   1080
      Width           =   1695
   End
   Begin VB.Frame Frame3 
      Caption         =   "完成后打开"
      Height          =   735
      Left            =   240
      TabIndex        =   16
      Top             =   2280
      Width           =   7575
      Begin VB.CheckBox c3 
         Caption         =   "j:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5640
         TabIndex        =   19
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox c2 
         Caption         =   "h:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   3480
         TabIndex        =   18
         Top             =   240
         Width           =   975
      End
      Begin VB.CheckBox c1 
         Caption         =   "g:"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "Form1.frx":0442
      Left            =   240
      List            =   "Form1.frx":044F
      TabIndex        =   15
      Text            =   "请选择盘符"
      Top             =   600
      Width           =   1695
   End
   Begin MSComctlLib.ProgressBar p2 
      Height          =   375
      Left            =   240
      TabIndex        =   14
      Top             =   3600
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   3000
   End
   Begin VB.CommandButton Command4 
      Caption         =   "反选"
      Height          =   495
      Left            =   6600
      TabIndex        =   13
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "全选"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   1680
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Caption         =   "选项2"
      Height          =   2055
      Left            =   5160
      TabIndex        =   8
      Top             =   240
      Width           =   2655
      Begin VB.CheckBox cre 
         Caption         =   "rd recycled"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   1080
         Width           =   1695
      End
      Begin VB.CheckBox ca 
         Caption         =   "del autorun.inf"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1935
      End
      Begin VB.CheckBox ce 
         Caption         =   "del *.exe"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "反选"
      Height          =   495
      Left            =   3720
      TabIndex        =   7
      Top             =   1680
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全选"
      Height          =   495
      Left            =   2640
      TabIndex        =   6
      Top             =   1680
      Width           =   855
   End
   Begin VB.CheckBox cr 
      Caption         =   "-r"
      Height          =   255
      Left            =   2760
      TabIndex        =   5
      Top             =   960
      Width           =   975
   End
   Begin VB.CheckBox cs 
      Caption         =   "-s"
      Height          =   255
      Left            =   2760
      TabIndex        =   4
      Top             =   600
      Width           =   975
   End
   Begin VB.CheckBox ch 
      Caption         =   "-h"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "选项1"
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
      Top             =   3120
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
      Height          =   3855
      Left            =   8280
      TabIndex        =   0
      Top             =   120
      Width           =   975
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
Dim cmd1, cmd2 As String
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
msg = MsgBox("杀毒完成！", vbInformation + vbOKOnly, "OK")
If c1.Value = 1 Then
run = Shell("cmd /c start g:")
End If
If c2.Value = 1 Then
run = Shell("cmd /c start h:")
End If
If c3.Value = 1 Then
run = Shell("cmd /c start j:")
End If
End Sub


Private Sub Combo1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command5_Click
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
Select Case Combo1.Text
Case "g:"
c1.Enabled = True
c2.Enabled = False
c3.Enabled = False
cmd1 = "attrib g:* "
p = "g:"
Case "h:"
c2.Enabled = True
c1.Enabled = False
c3.Enabled = False
cmd1 = "attrib h:* "
p = "h:"
Case "j:"
c3.Enabled = True
c1.Enabled = False
c2.Enabled = False
cmd1 = "attrib j:* "
p = "j:"
Case Else
msg = MsgBox("请选择盘符", vbCritical, "error")
End Select
End Sub

Private Sub esc_Click()
Unload Form1
End Sub
