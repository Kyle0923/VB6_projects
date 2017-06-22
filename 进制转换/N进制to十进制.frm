VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "进制转换"
   ClientHeight    =   3885
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   3885
   ScaleWidth      =   11895
   StartUpPosition =   3  '窗口缺省
   Begin VB.CheckBox Check1 
      Caption         =   "显示十进制数值"
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
      TabIndex        =   10
      Top             =   120
      Width           =   2295
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   4200
      TabIndex        =   9
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   840
      TabIndex        =   3
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   24
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3180
      TabIndex        =   7
      Top             =   4050
      Width           =   5655
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   3480
      TabIndex        =   5
      Top             =   1920
      Visible         =   0   'False
      Width           =   6975
   End
   Begin VB.TextBox Text2 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   26.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   645
      Left            =   8520
      TabIndex        =   1
      Top             =   840
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   5655
   End
   Begin VB.Label Label4 
      Caption         =   "进制数值:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1920
      TabIndex        =   8
      Top             =   3000
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "十进制数值："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   1920
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "原始数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   4
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "进制"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.Menu reset 
      Caption         =   "重设"
   End
   Begin VB.Menu exit 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim n, o1, o2 As Long
Private Sub Check1_Click()
If Check1.Value = 1 Then
Text4.Top = 3000
Label4.Top = 3000
Text5.Top = 3000
Command1.Top = 4050
Form1.Height = 5835
Label3.Visible = True
Text3.Visible = True
Else
Label3.Visible = False
Text3.Visible = False
Text4.Top = 1920
Label4.Top = 1920
Text5.Top = 1920
Command1.Top = 2970
Form1.Height = 4755
End If
End Sub

Private Sub Command1_Click()
Text3.Text = ""
Text5.Text = ""
If o1 > 10 Then
If Right(Text1.Text, 1) <> "." Then
Text1.Text = Text1.Text & "."
End If
T1 = Text1.Text
n = 0
a = 1
Do
a = InStr(a, T1, ".")
If a > 0 Then
n = n + 1
a = a + Len(".")
End If
Loop Until a = 0
b = 1
T1 = "." & T1
For i = 1 To n
b1 = InStr(b, T1, ".")
b = 1 + b1
b2 = InStr(b, T1, ".")
b3 = b2 - b1 - 1
s1 = Val(Mid(T1, b1 + 1, b3))
If s1 > o1 - 1 Then
m = MsgBox("原始数非" & o & "进制数值", vbCritical, "错误")
Text2.Text = ""
Text2.SetFocus
Text3.Text = ""
Else
c1 = Val(s1)
d = d + c1 * o1 ^ (n - i)
Text3.Text = d
End If
Next
Else
For i = 1 To n
s1 = Right(Text1.Text, i)
s2 = Val(Left(s1, 1))
a = s2 * (o1 ^ (i - 1))
b = b + a
Next
Text3.Text = b
End If
oo = Val(Text3.Text)
n1 = Log(oo) / Log(o2)
n1 = (n1 + 1) \ 1
ss = Val(Text3.Text)
If o2 <= 10 Then
For iii = 1 To n1
s3 = ss Mod o2
Text5.Text = s3 & Text5.Text
ss = ss \ o2
Next
Text5.Text = Val(Text5.Text)
Else
For iii = 1 To n1
s3 = ss Mod o2
Text5.Text = s3 & "." & Text5.Text
ss = ss \ o2
Next
If Left(Text5.Text, 1) = "0" Then
Text5.Text = Mid(Text5.Text, 3, Len(Text5.Text) - 2)
End If
End If
End Sub

Private Sub exit_Click()
Unload Form1
End Sub

Private Sub Form_Load()
Text4.Top = 1920
Label4.Top = 1920
Text5.Top = 1920
Command1.Top = 2970
End Sub

Private Sub reset_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
End Sub




Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text2.SetFocus
End If
End Sub


Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Text4.SetFocus
End If
End Sub

Private Sub Text2_LostFocus()
n = Len(Text1.Text)
o1 = Val(Text2.Text)
For ii = 1 To n
s = Mid(Text1.Text, ii, 1)
s1 = Val(s)
If s1 > o1 - 1 Then
m = MsgBox("原始数非" & o & "进制数值", vbCritical, "错误")
Text2.Text = ""
ii = n
Text2.SetFocus
End If
Next
End Sub



Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
Command1.SetFocus
End If

End Sub

Private Sub Text4_LostFocus()
o2 = Val(Text4.Text)
End Sub
