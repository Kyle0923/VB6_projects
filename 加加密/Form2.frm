VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   Caption         =   "加密"
   ClientHeight    =   3240
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   7950
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3240
   ScaleWidth      =   7950
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdreset 
      Caption         =   "重设"
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
      Left            =   3480
      TabIndex        =   3
      Top             =   1800
      Width           =   2655
   End
   Begin MSComctlLib.ProgressBar p2 
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   2880
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.TextBox Text2 
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
      Left            =   240
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   7455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "返回"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   6360
      TabIndex        =   0
      Top             =   1800
      Width           =   1335
   End
   Begin MSComctlLib.ProgressBar p1 
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   2520
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   53
   End
   Begin VB.CommandButton Command1 
      Caption         =   "加密"
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
      TabIndex        =   2
      Top             =   1800
      Width           =   2655
   End
   Begin VB.TextBox Text1 
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
      Left            =   240
      TabIndex        =   1
      Top             =   360
      Width           =   7455
   End
   Begin VB.Label Label2 
      Caption         =   "加密文："
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
      TabIndex        =   8
      Top             =   960
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "原文："
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
      TabIndex        =   7
      Top             =   120
      Width           =   1815
   End
   Begin VB.Menu esc 
      Caption         =   "退出"
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a As Single
Private Sub cmdreset_Click()
msga = MsgBox("确定重设本窗口?", vbInformation + vbOKCancel, "重设?")
If msga = vbOK Then
Me.Text1.Locked = False
Me.Text1.Text = ""
Me.Text2.Text = ""
p1.Value = 0
p2.Value = 0
Me.Command1.Enabled = True
End If
End Sub

Private Sub Command1_Click()
If Text1.Text <> "" Then
    Me.Command1.Enabled = False
    Text1.Locked = True
    Me.cmdreset.Enabled = False
    Me.Text2.Text = ""
    a = 1
    p1.Value = a
    str1 = Trim(Me.Text1.Text)
    c = Len(str1)
    p2.Max = c
        For i = 1 To c
        p2.Value = i
            If p2.Value = p2.Max Then
            Me.cmdreset.Enabled = True
            End If
        str2 = Mid(str1, i, 1)
            For b = 97 To 122
            a = 0
            a = a + 1
            p1.Value = a
            str3 = str2
                If str2 = Chr(b) Then
                Key = Val(Form1.Text1.Text)
                ans = b + Key + i
                    If ans > 122 Then
                    ans = ans - 26
                    End If
                    If ans < 97 Then
                    ans = ans + 26
                    End If
                str3 = Replace(str2, Chr(b), Chr(ans))
                b = 123
                Me.Text2.Text = Me.Text2.Text & str3
                a = a + 1
                p1.Value = a
                End If
            Next b
    For b = 65 To 90
        a = a + 1
        p1.Value = a
        str4 = str2
            If str2 = Chr(b) Then
            Key = Val(Form1.Text1.Text)
            ans = b + Key + i
            If ans > 90 Then
            ans = ans - 26
            End If
            If ans < 65 Then
            ans = ans + 26
            End If
        str4 = Replace(str2, Chr(b), Chr(ans))
        b = 91
        Me.Text2.Text = Me.Text2.Text & str4
        a = a + 1
        p1.Value = a
        End If
    Next b
        If str3 = str2 And str4 = str2 Then
        Me.Text2.Text = Me.Text2.Text & str3
        End If
    Next i
    p1.Value = p1.Max
Else
msg = MsgBox("请输入原文", vbCritical, "error")
End If
End Sub

Private Sub Command2_Click()
Form2.Hide
Form1.Show
Me.Text1.Locked = False
Me.Text1.Text = ""
Me.Text2.Text = ""
p1.Value = 0
p2.Value = 0
Me.Command1.Enabled = True
End Sub

Private Sub esc_Click()
Unload Form2
End Sub

Private Sub Form_Load()
Me.Text1.Locked = False
Me.Text1.Text = ""
Me.Text2.Text = ""
p1.Value = 0
p2.Value = 0
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
Command1_Click
End If
End Sub
