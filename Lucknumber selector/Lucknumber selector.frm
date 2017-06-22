VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   Caption         =   "Lucknumber selector"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ProgressBar p1 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   2640
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
      Max             =   3000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "开始"
      BeginProperty Font 
         Name            =   "薛文轩钢笔楷体"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   600
      TabIndex        =   0
      Top             =   1920
      Width           =   3255
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "00"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   68.25
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   2640
      TabIndex        =   2
      Top             =   240
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Lucknumber:"
      BeginProperty Font 
         Name            =   "Gill Sans MT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
p1.Value = 0
Randomize
a = Int((52 * Rnd(52)) + 1)
Do While a = 2 Or a = 3 Or a = 14 Or a = 32 Or a = 12 Or a = 10
a = Int((52 * Rnd(52)) + 1)
Loop
For x = 1 To 1000
p1.Value = p1.Value + 1
Next x
If a < 10 Then
a = "0" & a
End If
For x = 1 To 1000
p1.Value = p1.Value + 1
Next x
Label2.Caption = a
For x = 1 To 1000
p1.Value = p1.Value + 1
Next x
End Sub

