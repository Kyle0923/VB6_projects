VERSION 5.00
Begin VB.Form form2 
   Caption         =   "Setting"
   ClientHeight    =   4380
   ClientLeft      =   4290
   ClientTop       =   2610
   ClientWidth     =   4365
   LinkTopic       =   "Form2"
   ScaleHeight     =   4380
   ScaleWidth      =   4365
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
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
      Left            =   2355
      TabIndex        =   9
      Top             =   3120
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
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
      Left            =   435
      TabIndex        =   8
      Top             =   3120
      Width           =   1575
   End
   Begin VB.TextBox Text4 
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
      Left            =   1800
      TabIndex        =   7
      TabStop         =   0   'False
      Text            =   "5"
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text3 
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
      Left            =   1800
      TabIndex        =   6
      TabStop         =   0   'False
      Text            =   "0"
      Top             =   1680
      Width           =   1695
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
      Left            =   1800
      TabIndex        =   5
      TabStop         =   0   'False
      Text            =   "9"
      Top             =   1080
      Width           =   1695
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
      Left            =   1800
      TabIndex        =   4
      TabStop         =   0   'False
      Text            =   "2200"
      Top             =   480
      Width           =   1695
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Times:"
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
      Left            =   720
      TabIndex        =   3
      Top             =   2280
      Width           =   855
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Key2:"
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
      Left            =   720
      TabIndex        =   2
      Top             =   1680
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Key1:"
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
      Left            =   720
      TabIndex        =   1
      Top             =   1080
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Period:"
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
      Left            =   720
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a1, a2, a3, a4 As String

Private Sub Command1_Click()
Form1.Enabled = True
form2.Visible = False
End Sub

Private Sub Command2_Click()
Text1.Text = a1
Text2.Text = a2
Text3.Text = a3
Text4.Text = a4
Form1.Enabled = True
form2.Visible = False
End Sub

Private Sub Form_Activate()
a1 = Text1.Text
a2 = Text2.Text
a3 = Text3.Text
a4 = Text4.Text
Form1.Enabled = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
Form1.Enabled = True
End Sub

Private Sub Text1_GotFocus()
If Val(Text1.Text) = 2200 Then
Text1.Text = ""
End If
End Sub

Private Sub Text1_LostFocus()
c = Val(Text1.Text)
Do Until c > 0
z = MsgBox("please input a number", vbCritical + vbOKOnly, "error")
Text1.Text = 2200
c = Val(Text1.Text)
Loop
Text1.Text = Val(Text1.Text)
End Sub

Private Sub Text2_GotFocus()
If Val(Text2.Text) = 9 Then
Text2.Text = ""
End If
End Sub


Private Sub Text3_GotFocus()
If Val(Text3.Text) = 0 Then
Text3.Text = ""
End If
End Sub

Private Sub Text4_GotFocus()
If Val(Text4.Text) = 5 Then
Text4.Text = ""
End If
End Sub

Private Sub Text4_LostFocus()
c = Val(Text4.Text)
Do Until c > 0
z = MsgBox("please input a number", vbCritical + vbOKOnly, "error")
Text4.Text = 5
c = Val(Text4.Text)
Loop
Text4.Text = Val(Text4.Text)
End Sub




