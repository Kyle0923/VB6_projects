VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "start"
   ClientHeight    =   2460
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6345
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
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   2460
   ScaleWidth      =   6345
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Command3 
      Caption         =   "解密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   3
      Top             =   1560
      Width           =   3135
   End
   Begin VB.CommandButton Command2 
      Caption         =   "加密"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2760
      TabIndex        =   2
      Top             =   600
      Width           =   3135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000000&
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   480
      MaxLength       =   3
      TabIndex        =   0
      Text            =   "请输入密钥"
      Top             =   1080
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "密匙："
      BeginProperty Font 
         Name            =   "Txt"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   1335
   End
   Begin VB.Menu esc 
      Caption         =   "退出(&E)"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
a = Val(Me.Text1.Text)
Me.Text1.Text = a
If Me.Command1.Caption = "确定" Then
    If a <= -26 Or a >= 26 Or a = 0 Then
    msg1 = MsgBox("密钥必须在-25到25之间且不为0", vbCritical, "error")
    Me.Text1.Text = ""
    Else
    msg2 = MsgBox("确定密钥为" & a, vbInformation + vbOKCancel, "Are you sure?")
        If msg2 = vbOK Then
        Me.Text1.Enabled = False
        Me.Command1.Caption = "重设密钥"
        Else
        Me.Text1.Text = ""
        End If
    End If
Else
Me.Text1.Enabled = True
Me.Text1.Text = ""
Me.Command1.Caption = "确定"
End If
End Sub

Private Sub Command2_Click()
If Me.Text1.Enabled = False Then
Form1.Hide
Form2.Show
Else
msgk = MsgBox("未设置密钥", vbCritical, "error")
End If
End Sub

Private Sub Command3_Click()
If Me.Text1.Enabled = False Then
Form1.Hide
Form3.Show
Else
msgk = MsgBox("未设置密钥", vbCritical, "error")
End If
End Sub


Private Sub esc_Click()
Unload Form1
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If Me.Text1.Text = "请输入密钥" Then
Me.Text1.Text = ""
End If
Me.Text1.ForeColor = vbBlack
Me.Text1.FontSize = 16
Me.Text1.FontUnderline = False
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub Text1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Me.Text1.Text = ""
Me.Text1.ForeColor = vbBlack
Me.Text1.FontSize = 16
Me.Text1.FontUnderline = False
End Sub
