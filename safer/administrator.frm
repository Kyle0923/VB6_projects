VERSION 5.00
Begin VB.Form administrator 
   Caption         =   "administrator"
   ClientHeight    =   3015
   ClientLeft      =   11640
   ClientTop       =   7950
   ClientWidth     =   3585
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   3585
   Begin VB.Menu exit 
      Caption         =   "退出(&E)"
   End
End
Attribute VB_Name = "administrator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()

End Sub

Private Sub exit_Click()
Unload password
Unload user
Unload administrator
End Sub

Private Sub Form_Load()
msg = MsgBox("欢迎登陆管理员账号", vbInformation, "welcome")
End Sub
