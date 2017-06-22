VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   1560
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton Command1 
      Caption         =   "start"
      Height          =   855
      Left            =   1320
      TabIndex        =   0
      Top             =   1920
      Width           =   2175
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   240
      Top             =   360
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "ËÎÌå"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Double

Private Sub Command1_Click()
If Command1.Caption = "start" Then
Command1.Caption = "pause"
Else
Command1.Caption = "start"
End If
Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Timer1_Timer()
a = a + 1
If a = 1900 Then
Timer3.Enabled = True
End If
If a = 1900 + 1000 Then
Timer2.Enabled = True
a = 0
End If
Label1.Caption = a
End Sub

Private Sub Timer2_Timer()
b = b + 1
For i = 1 To 3
If b = 6 * i Then
SendKeys "9"
End If
If b = 6 * i - 3 Then
SendKeys "0"
End If
Next
If b = 10 Then
Timer2.Enabled = False
b = 0
End If
End Sub

Private Sub Timer3_Timer()
c = c + 1
If c = 3 Then
SendKeys "-"
End If
If c = 6 Then
SendKeys "0"
End If
If c = 9 Then
Timer3.Enabled = False
c = 0
End If
End Sub
