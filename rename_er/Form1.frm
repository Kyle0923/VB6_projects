VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "重命名工具"
   ClientHeight    =   4110
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   8280
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   7.5
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   4110
   ScaleWidth      =   8280
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar p1 
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox hz 
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
      Left            =   7080
      TabIndex        =   2
      Text            =   "后缀名"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox f 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5520
      TabIndex        =   1
      Text            =   "首文件的num"
      Top             =   480
      Width           =   1215
   End
   Begin VB.TextBox fnum 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   4
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox fn 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   2520
      TabIndex        =   3
      Top             =   1200
      Width           =   5535
   End
   Begin VB.TextBox pn 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   1800
      TabIndex        =   0
      Top             =   480
      Width           =   3615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   2760
      Width           =   7815
   End
   Begin VB.Label Label4 
      Caption         =   "."
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6840
      TabIndex        =   9
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "文件数："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "new filename："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   7
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Label Label1 
      Caption         =   "pathname:"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   480
      Width           =   1455
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
Private Sub Command1_Click()
p1.Max = num + 1
a = 0
pn.Text = Trim(pn.Text)
fn.Text = Trim(fn.Text)
num = Val(fnum.Text)
f.Text = Trim(f.Text)
hz.Text = Trim(hz.Text)
hzm = "." & hz.Text
filenum = Val(f.Text)
p1.Value = a + 1
For i = 1 To num
If i > 9 And i < 100 Then
namei = "0" & Trim(Str(i))
End If
If i < 10 Then
namei = "00" & Trim(Str(i))
End If
FileName = pn.Text & filenum & hzm
newname = fn.Text & namei & hzm
run = Shell("cmd /c rename " & FileName & " " & newname)
filenum = filenum + 1
a = a + 1
If a > p1.Max Then
p1.Value = p1.Max
Else
p1.Value = a
End If
Next
msg = MsgBox("finish", vbInformation, "ok")
End Sub

Private Sub exit_Click()
Unload Form1
End Sub

Private Sub f_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If f.Text = "首文件的num" Then
f.Text = ""
f.FontSize = 12
End If
End Sub

Private Sub fnum_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub hz_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If hz.Text = "后缀名" Then
hz.Text = ""
End If
End Sub

