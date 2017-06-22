VERSION 5.00
Begin VB.Form password 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "黄宗皓的文件夹"
   ClientHeight    =   3045
   ClientLeft      =   11625
   ClientTop       =   7935
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "确认"
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   2160
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Height          =   360
      IMEMode         =   3  'DISABLE
      Left            =   1320
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   1320
      Width           =   1935
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
      Height          =   360
      Left            =   1320
      TabIndex        =   0
      Text            =   "User"
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "密码："
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
      TabIndex        =   2
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "用户名："
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
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.Menu exit 
      Caption         =   "退出(&E)"
   End
End
Attribute VB_Name = "password"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Key = Trim(Text2.Text)
Select Case Text1.Text
Case "User"
a = Shell("cmd /c if exist " & Key & ".inf start E:\黄宗皓的文件夹 else start keyerror.vbs")
pass = Shell("cmd /c ren E:\黄宗皓的文件夹\safer\lock.inf pass.inf")
user.Show
password.Hide
Case "administrator"
    If Key = "713520" Then
    administrator.Show
    password.Hide
    End If
End Select
End Sub


Private Sub exit_Click()
Unload password
Unload user
Unload administrator
End Sub

Private Sub Form_Unload(Cancel As Integer)
cmd = Shell("cmd /c ren E:\黄宗皓的文件夹\safer\pass.inf lock.inf")
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
