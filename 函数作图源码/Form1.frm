VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "初等函数作图器 - Programmed by LG [Ver 1.2.0.1]"
   ClientHeight    =   11400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13725
   FillColor       =   &H000000FF&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   11400
   ScaleWidth      =   13725
   StartUpPosition =   2  '屏幕中心
   Begin VB.OptionButton Option1 
      Caption         =   "0.0001[易死机注意]"
      Height          =   255
      Index           =   5
      Left            =   11040
      TabIndex        =   21
      Tag             =   "0.0001"
      Top             =   9480
      Width           =   2295
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0.002"
      Height          =   255
      Index           =   3
      Left            =   11040
      TabIndex        =   20
      Tag             =   "0.002"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0.005"
      Height          =   255
      Index           =   4
      Left            =   11040
      TabIndex        =   19
      Tag             =   "0.005"
      Top             =   8760
      Width           =   1455
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0.01"
      Height          =   255
      Index           =   2
      Left            =   9480
      TabIndex        =   17
      Tag             =   "0.01"
      Top             =   9480
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "0.1"
      Height          =   255
      Index           =   1
      Left            =   9480
      TabIndex        =   16
      Tag             =   "0.1"
      Top             =   9120
      Width           =   1215
   End
   Begin VB.OptionButton Option1 
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   9480
      TabIndex        =   15
      Tag             =   "1"
      Top             =   8760
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   10
      Text            =   "Form1.frx":0000
      Top             =   10080
      Width           =   9015
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   7095
      Left            =   9480
      ScaleHeight     =   7035
      ScaleWidth      =   4035
      TabIndex        =   9
      Top             =   1080
      Width           =   4095
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   375
      Left            =   9480
      Max             =   50
      Min             =   1
      TabIndex        =   7
      Top             =   600
      Value           =   10
      Width           =   4095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "清空"
      Height          =   375
      Left            =   7320
      TabIndex        =   5
      Top             =   600
      Width           =   1815
   End
   Begin VB.CheckBox Check1 
      Caption         =   "自动清空"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   600
      Value           =   1  'Checked
      Width           =   1095
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   9000
      Left            =   120
      ScaleHeight     =   8940
      ScaleWidth      =   8940
      TabIndex        =   2
      Top             =   1080
      Width           =   9000
   End
   Begin VB.CommandButton Command1 
      Caption         =   "作图！"
      Height          =   375
      Left            =   7320
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   6375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "版本1.2.0.1于2010年11月20日编写并发布"
      Height          =   255
      Left            =   9600
      TabIndex        =   22
      Top             =   10200
      Width           =   3375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "绘图精度："
      Height          =   180
      Left            =   9480
      TabIndex        =   18
      Top             =   8400
      Width           =   900
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "程序仅为交流、学习之用 QQ:563259858"
      Height          =   255
      Left            =   9600
      TabIndex        =   14
      Top             =   10440
      Width           =   3615
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright LinGeng 2010~2011"
      Height          =   255
      Left            =   9600
      TabIndex        =   13
      Top             =   11040
      Width           =   2655
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "您可以自由传播本程序，但不得用于商业用途"
      Height          =   255
      Left            =   9600
      TabIndex        =   12
      Top             =   10800
      Width           =   3735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "本程序由LG于2010年8月30日编写"
      Height          =   255
      Left            =   9600
      TabIndex        =   11
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "10"
      ForeColor       =   &H000000FF&
      Height          =   180
      Left            =   11760
      TabIndex        =   8
      Top             =   240
      Width           =   180
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "坐标系大小（x * x）："
      Height          =   255
      Left            =   9600
      TabIndex        =   6
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "f(x)="
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************
'=======本程序由LG编写=======
'====仅为交流、学习之目的====
'=可以自由、免费地使用与传播=
'****************************
'*****但严禁用于商业用途*****
'****************************
'==LG的邮箱: lgiscj@163.com==
'==LG的QQ  : 563259858     ==
'===如果发现BUG,请立即通知===
'=如果对程序有建议，也请告知=
'==========谢谢合作==========
'****************************

'Ver 1.2.0.1
'2010年11月20日新增：
'  在改变精度、坐标系大小时可重绘之前的所有函数


'程序原理：使用VB的ScriptControl组件，将每个函数自变量值代入表达式分别计算
'          所得的结果即为对应函数值
    
    'i:确定初次作图
    Dim i As Integer
    'h:确定坐标系大小
    Dim h As Integer
    'counting:确定图像数量
    Dim counting As Integer
    'strC:函数表达式
    Dim strC As String
    'steping:绘图精度
    Dim steping As Single
    '定义变量x,y
    Dim x, y As Double
    '储存输入的函数式
    Dim fx(100) As String
    '临时使用
    Dim tmpNum, tmpNum2, tmpNum3, tmpNum4 As Integer
    '保留性清空
    Dim nclean As Boolean

Private Sub Command1_Click()
    '跳过全部错误
    'On Error Resume Next
    
    '若第一次作图，画坐标系
    If i = 0 Then
        Clear
        i = 1
    End If
     
    '启动自动清空时，清空绘图区
    If Me.Check1.Value = 1 Then Clear
    
    '函数数量+1
    counting = counting + 1
    
    '根据函数数目设定颜色
    Dim cor As Long
    Select Case counting
        Case 1
            cor = RGB(255, 0, 0)
        Case 2
            cor = RGB(0, 0, 255)
        Case 3
            cor = RGB(255, 0, 255)
        Case 4
            cor = RGB(0, 255, 255)
        Case 5
            cor = RGB(255, 255, 0)
        Case Else
            '随机颜色
            cor = RGB(Int(Rnd * 253) + 1, Int(Rnd * 253) + 1, Int(Rnd * 253) + 1)
    End Select
    
    '显示颜色对应的函数
    Me.Picture2.FontSize = 12
    Me.Picture2.ForeColor = cor
    Me.Picture2.Print ("f(x" & counting & ")= " & Text1.Text)

    '读入表达式
    strC = Me.Text1.Text
    
    '储存表达式
    fx(counting) = strC
    
    '大写改小写
    strC = LCase(strC)
    
    '常量
    strC = Replace(strC, "e", Exp(1))
    strC = Replace(strC, "pi", 3.14159265358979)
    
    'ax改为a * x
    For tmpNum = 0 To 9
        For tmpNum2 = 97 To 122
            strC = Replace(strC, tmpNum & Chr(tmpNum2), tmpNum & "*" & Chr(tmpNum2))
        Next tmpNum2
    Next tmpNum

    '逐个描点
    DrawFx strC, cor

End Sub

Private Sub Clear()
'清空绘图区、信息区、重定义坐标系大小、画坐标轴、函数数目归零、清空保存的函数
Me.Picture1.Cls
Me.Picture2.Cls
Me.Picture1.Scale (-h, h)-(h, -h)
Me.Picture1.Line (-h, 0)-(h, 0), RGB(0, 255, 0)
Me.Picture1.Line (0, -h)-(0, h), RGB(0, 255, 0)
If nclean = False Then
    For tmpNum = 0 To counting
        fx(tmpNum) = ""
    Next
    counting = 0
End If
End Sub

Private Sub Command2_Click()
'清空
Clear
End Sub

Private Sub Form_Load()
'窗体标题
Dim ver As String
ver = "1.0.6.1"
Me.Caption = "初等函数作图器[" & ver & "]--Programmed by LG"

'变量初始化
i = 0
h = 10
counting = 0
steping = 0.01
nclean = False
redrawing = False
End Sub

Private Sub HScroll1_Change()
'改变坐标系大小
h = Me.HScroll1.Value
Me.Label3.Caption = h
nclean = True
Clear
nclean = False
tmpNum3 = counting
counting = 0
'自动重新画函数
For tmpNum4 = 1 To tmpNum3
    Me.Text1.Text = fx(tmpNum4)
    Command1_Click
Next tmpNum4
End Sub

Private Sub Option1_Click(Index As Integer)
'改变精度
steping = Option1(Index).Tag
nclean = True
Clear
nclean = False
tmpNum3 = counting
counting = 0
'自动重新画函数
For tmpNum4 = 1 To tmpNum3
    Me.Text1.Text = fx(tmpNum4)
    Command1_Click
Next tmpNum4
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
'按下ENTER画图
If KeyAscii = 13 Then
Command1_Click
End If
End Sub

Private Sub DrawFx(strFx As String, colr As Long)
    '设置执行器
    Set s = CreateObject("ScriptControl")
    s.Language = "VBScript"
    
    '逐个描点
    For x = -h To h Step steping
        y = s.eval(Replace(strFx, "x", x))
        '画点
      If Not y > h Then Me.Picture1.PSet (x, y), colr
      y = 0
    Next x
    
    '释放执行器
    Set s = Nothing
End Sub
