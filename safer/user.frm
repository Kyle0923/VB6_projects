VERSION 5.00
Begin VB.Form user 
   Caption         =   "user"
   ClientHeight    =   3015
   ClientLeft      =   11640
   ClientTop       =   7950
   ClientWidth     =   3585
   LinkTopic       =   "Form2"
   Moveable        =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3585
   Begin VB.CommandButton Command1 
      Caption         =   "关闭"
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label Label1 
      Caption         =   "黄宗皓的文件夹已解锁"
      BeginProperty Font 
         Name            =   "新宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   3375
   End
   Begin VB.Menu back 
      Caption         =   "返回(&B)"
   End
   Begin VB.Menu exit 
      Caption         =   "退出(&E)"
   End
End
Attribute VB_Name = "user"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub back_Click()
password.Show
user.Hide
End Sub

Private Sub Command1_Click()
Unload password
Unload user
Unload administrator
End Sub

Private Sub exit_Click()
Unload password
Unload user
Unload administrator
End Sub

