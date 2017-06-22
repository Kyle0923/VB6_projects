VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "timer runer"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "visible"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   3
      Top             =   1800
      Width           =   2775
   End
   Begin VB.CommandButton cmdstart 
      Caption         =   "start"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2280
      Width           =   3135
   End
   Begin VB.Timer Timer1 
      Left            =   4200
      Top             =   1680
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
      Left            =   1560
      TabIndex        =   2
      Text            =   "input file name"
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   840
      TabIndex        =   1
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "run:"
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
      Left            =   840
      TabIndex        =   5
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "s"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3240
      TabIndex        =   4
      Top             =   480
      Width           =   615
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b As Integer
Dim c As String
Private Sub cmdstart_Click()
If Check1.Value = 0 Then
Me.Visible = False
End If
If cmdstart.Caption = "end" Then
Timer1.Enabled = False
Timer1.Interval = 0
b = a
Text1.Text = a
cmdstart.Caption = "start"
Else
Timer1.Enabled = True
Timer1.Interval = 1000
a = Val(Text1.Text)
c = Trim(Text2.Text)
cmdstart.Caption = "end"
End If
End Sub

Private Sub Form_Load()
Check1.Value = 1
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdstart_Click
End If
End Sub


Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
cmdstart_Click
End If
End Sub

Private Sub Text2_LostFocus()
If Text2.Text = "" Then
Text2.Text = "input file name"
End If
End Sub

Private Sub Text2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Text2.Text = "input file name" Then
Text2.Text = ""
End If
End Sub

Private Sub Timer1_Timer()
b = Val(Text1.Text)
b = b - 1
If b <= 0 Then
run = Shell("cmd /c start " & c)
b = a
End If
Text1.Text = b
End Sub
