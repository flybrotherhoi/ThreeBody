VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form TwoBodyMode 
   AutoRedraw      =   -1  'True
   Caption         =   "二体模式"
   ClientHeight    =   9990
   ClientLeft      =   585
   ClientTop       =   795
   ClientWidth     =   18285
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9990
   ScaleMode       =   0  'User
   ScaleWidth      =   18222.17
   Begin VB.CommandButton Command6 
      Caption         =   "退出系统"
      Height          =   315
      Left            =   255
      TabIndex        =   40
      Top             =   9480
      Width           =   2565
   End
   Begin MSComDlg.CommonDialog Cmdlog2 
      Left            =   3390
      Top             =   6990
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton xiugai 
      Caption         =   "修改"
      Enabled         =   0   'False
      Height          =   420
      Left            =   555
      TabIndex        =   38
      Top             =   4695
      Width           =   855
   End
   Begin VB.CommandButton dakai 
      Caption         =   "打开"
      Height          =   420
      Left            =   555
      TabIndex        =   37
      Top             =   5385
      Width           =   855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "返回三体模式"
      Height          =   360
      Left            =   285
      TabIndex        =   36
      Top             =   8955
      Width           =   2550
   End
   Begin VB.Frame Frame2 
      Caption         =   "说明"
      Height          =   2295
      Left            =   240
      TabIndex        =   31
      Top             =   6390
      Width           =   2700
      Begin VB.Label Label10 
         Caption         =   "5.另存数据为上次确定时的数据"
         Height          =   405
         Left            =   150
         TabIndex        =   39
         Top             =   1815
         Width           =   2355
      End
      Begin VB.Label Label9 
         Caption         =   "4.右下为2相对于1的运动"
         Height          =   300
         Left            =   150
         TabIndex        =   35
         Top             =   1500
         Width           =   2085
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "1.天体1为红色，天体2为蓝色"
         Height          =   180
         Left            =   150
         TabIndex        =   34
         Top             =   345
         Width           =   2340
      End
      Begin VB.Label Label7 
         Caption         =   "3.右上为1相对于2的运动"
         Height          =   270
         Left            =   150
         TabIndex        =   33
         Top             =   1125
         Width           =   2205
      End
      Begin VB.Label Label6 
         Caption         =   "2.中间大框为两天体运动"
         Height          =   270
         Left            =   150
         TabIndex        =   32
         Top             =   735
         Width           =   2205
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "控制"
      Height          =   1305
      Left            =   120
      TabIndex        =   26
      Top             =   195
      Width           =   2910
      Begin VB.CommandButton Command4 
         Caption         =   "清除轨迹"
         Height          =   420
         Left            =   1560
         TabIndex        =   30
         Top             =   765
         Width           =   1200
      End
      Begin VB.CommandButton Command3 
         Caption         =   "显示轨迹"
         Height          =   420
         Left            =   255
         TabIndex        =   29
         Top             =   765
         Width           =   1170
      End
      Begin VB.CommandButton Command2 
         Caption         =   "暂停"
         Height          =   420
         Left            =   1560
         TabIndex        =   28
         Top             =   285
         Width           =   1200
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000D&
         Caption         =   "开始"
         Height          =   420
         Left            =   255
         TabIndex        =   27
         Top             =   270
         Width           =   1170
      End
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      Height          =   5000
      Left            =   13275
      ScaleHeight     =   4935
      ScaleWidth      =   4950
      TabIndex        =   25
      Top             =   4980
      Width           =   5017
      Begin VB.Line Line8 
         X1              =   2595
         X2              =   2595
         Y1              =   15
         Y2              =   4920
      End
      Begin VB.Line Line7 
         X1              =   -30
         X2              =   4950
         Y1              =   2535
         Y2              =   2535
      End
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      Height          =   10000
      Left            =   3195
      ScaleHeight     =   9945
      ScaleWidth      =   9975
      TabIndex        =   24
      Top             =   0
      Width           =   10034
      Begin VB.Line Line4 
         X1              =   5010
         X2              =   4995
         Y1              =   0
         Y2              =   9945
      End
      Begin VB.Line Line3 
         X1              =   0
         X2              =   9975
         Y1              =   4960
         Y2              =   4960
      End
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   1395
      TabIndex        =   18
      Text            =   "1"
      Top             =   3930
      Width           =   1395
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1395
      TabIndex        =   15
      Text            =   "1"
      Top             =   3510
      Width           =   1395
   End
   Begin VB.CommandButton lingcun 
      Caption         =   "另存"
      Height          =   420
      Left            =   1710
      TabIndex        =   14
      Top             =   5385
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   270
      Left            =   1874
      TabIndex        =   11
      Text            =   "0.4"
      Top             =   3030
      Width           =   1104
   End
   Begin VB.TextBox Text10 
      Height          =   270
      Left            =   1874
      TabIndex        =   10
      Text            =   "0"
      Top             =   2745
      Width           =   1104
   End
   Begin VB.TextBox Text9 
      Height          =   270
      Left            =   1874
      TabIndex        =   9
      Text            =   "0"
      Top             =   2415
      Width           =   1104
   End
   Begin VB.TextBox Text8 
      Height          =   270
      Left            =   1874
      TabIndex        =   8
      Text            =   "-20"
      Top             =   2100
      Width           =   1104
   End
   Begin VB.TextBox Text7 
      Height          =   270
      Left            =   1875
      TabIndex        =   7
      Text            =   "1"
      Top             =   1755
      Width           =   1104
   End
   Begin VB.CommandButton queren1 
      Caption         =   "确认"
      Height          =   420
      Left            =   1695
      TabIndex        =   6
      Top             =   4695
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   270
      Left            =   614
      TabIndex        =   5
      Text            =   "-0.4"
      Top             =   3045
      Width           =   1104
   End
   Begin VB.TextBox Text5 
      Height          =   270
      Left            =   614
      TabIndex        =   4
      Text            =   "0"
      Top             =   2745
      Width           =   1104
   End
   Begin VB.TextBox Text4 
      Height          =   270
      Left            =   614
      TabIndex        =   3
      Text            =   "0"
      Top             =   2415
      Width           =   1104
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   614
      TabIndex        =   2
      Text            =   "10"
      Top             =   2085
      Width           =   1104
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   614
      TabIndex        =   1
      Text            =   "1"
      Top             =   1755
      Width           =   1104
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      Height          =   5000
      Left            =   13275
      ScaleHeight     =   4935
      ScaleWidth      =   4950
      TabIndex        =   0
      Top             =   -15
      Width           =   5017
      Begin VB.Line Line6 
         X1              =   -30
         X2              =   4950
         Y1              =   2430
         Y2              =   2430
      End
      Begin VB.Line Line5 
         X1              =   2565
         X2              =   2580
         Y1              =   -90
         Y2              =   4890
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2805
      Top             =   6180
   End
   Begin VB.Line Line2 
      X1              =   14.948
      X2              =   3198.97
      Y1              =   6135
      Y2              =   6135
   End
   Begin VB.Line Line1 
      X1              =   29.897
      X2              =   3169.073
      Y1              =   4425
      Y2              =   4425
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "速度Vy"
      Height          =   180
      Index           =   4
      Left            =   0
      TabIndex        =   23
      Top             =   3090
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "速度Vx"
      Height          =   180
      Index           =   3
      Left            =   0
      TabIndex        =   22
      Top             =   2775
      Width           =   540
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "坐标Y"
      Height          =   180
      Index           =   2
      Left            =   0
      TabIndex        =   21
      Top             =   2445
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "坐标X"
      Height          =   180
      Index           =   1
      Left            =   0
      TabIndex        =   20
      Top             =   2145
      Width           =   450
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "质量M："
      Height          =   180
      Index           =   0
      Left            =   0
      TabIndex        =   19
      Top             =   1785
      Width           =   630
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "迭代步长dt:"
      Height          =   180
      Left            =   375
      TabIndex        =   17
      Top             =   3960
      Width           =   990
   End
   Begin VB.Label Label3 
      Caption         =   "引力常量G:"
      Height          =   345
      Left            =   450
      TabIndex        =   16
      Top             =   3540
      Width           =   945
   End
   Begin VB.Label Label2 
      Caption         =   "天体2"
      Height          =   255
      Left            =   2175
      TabIndex        =   13
      Top             =   1530
      Width           =   1050
   End
   Begin VB.Label Label1 
      Caption         =   "天体1"
      Height          =   210
      Left            =   930
      TabIndex        =   12
      Top             =   1500
      Width           =   645
   End
End
Attribute VB_Name = "TwoBodyMode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim b1 As body, b2 As body
Dim oldb1 As body, oldb2 As body
Dim oldG As Single
Dim olddt As Single
Dim guiji As Boolean
Private Sub Command1_Click()
Timer1.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
ChangeEnabled (False)
End Sub
Private Sub Command2_Click()
Timer1.Enabled = False
Command1.Enabled = True
Command2.Enabled = False
ChangeEnabled (True)
End Sub
Private Sub Command3_Click()
guiji = Not guiji
If Command3.Caption = "显示轨迹" Then
    Command3.Caption = "隐藏轨迹"
Else
    Command3.Caption = "显示轨迹"
End If
End Sub

Private Sub Command4_Click()
Picture1.Cls
Picture2.Cls
Picture3.Cls
End Sub

Private Sub Command5_Click()
Unload Me
End Sub

Private Sub Command6_Click()
Dim i As Integer
    i = MsgBox("是否保存？", vbYesNoCancel)
    If i = 6 Then
        Call lingcun_Click
        End
    Else: If i = 7 Then End
          If i = 2 Then
            Exit Sub
          End If
    End If
End Sub

Private Sub dakai_Click()
Dim duqu As String
Command1.Enabled = True
Command2.Enabled = False
Timer1.Enabled = False
'CancelError 为 True。
On Error GoTo ErrHandler
Cmdlog2.InitDir = App.path & "\例子"
'设置过滤器。
Cmdlog2.Filter = "Twobody_Files (*.twb)|*.twb"
'指定缺省过滤器。
Cmdlog2.FilterIndex = 2
'显示“打开”对话框。
Cmdlog2.ShowOpen
'调用打开文件的过程。
Open Cmdlog2.FileName For Input As #1
    Line Input #1, duqu
    Text1.Text = duqu
    Line Input #1, duqu
    Text12.Text = duqu
    Line Input #1, duqu
    Text2.Text = duqu
    Line Input #1, duqu
    Text3.Text = duqu
    Line Input #1, duqu
    Text4.Text = duqu
    Line Input #1, duqu
    Text5.Text = duqu
    Line Input #1, duqu
    Text6.Text = duqu
    Line Input #1, duqu
    Text7.Text = duqu
    Line Input #1, duqu
    Text8.Text = duqu
    Line Input #1, duqu
    Text9.Text = duqu
    Line Input #1, duqu
    Text10.Text = duqu
    Line Input #1, duqu
    Text11.Text = duqu
'加载
Loading
Close 1
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Form_Load()
With b1
    .Posi.x = 10
    .Posi.y = 0
    .Posi.z = 0
    .M = 2
    .V.x = 0
    .V.y = -Sqr(10) / 30
    .V.z = 0
End With
With b2
    .Posi.x = -20
    .Posi.y = 0
    .Posi.z = 0
    .M = 1
    .V.x = 0
    .V.y = 2 * Sqr(10) / 30
    .V.z = 0
End With
g = 1
dt = 1
Picture1.Scale (-100, 100)-(100, -100)
Picture2.Scale (-100, 100)-(100, -100)
Picture3.Scale (-100, 100)-(100, -100)
Line3.X1 = -100: Line3.Y1 = 0: Line3.X2 = 100: Line3.Y2 = 0
Line4.X1 = 0: Line4.Y1 = -100: Line4.X2 = 0: Line4.Y2 = 100
Line5.X1 = -100: Line5.Y1 = 0: Line5.X2 = 100: Line5.Y2 = 0
Line6.X1 = 0: Line6.Y1 = -100: Line6.X2 = 0: Line6.Y2 = 100
Line7.X1 = -100: Line7.Y1 = 0: Line7.X2 = 100: Line7.Y2 = 0
Line8.X1 = 0: Line8.Y1 = -100: Line8.X2 = 0: Line8.Y2 = 100
Command2.Enabled = False
'Picture1.Line (-100, 0)-(100, 0), vbRed
'Picture1.Line (0, 100)-(0, -100), vbRed
'Picture2.Line (-100, 0)-(100, 0), vbRed
'Picture2.Line (0, 100)-(0, -100), vbRed
'Picture3.Line (-100, 0)-(100, 0), vbRed
'Picture3.Line (0, 100)-(0, -100), vbRed

Call suggest
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
    i = MsgBox("是否保存？", vbYesNoCancel)
    If i = 6 Then
        Call lingcun_Click
        Frmmain.Show
    Else: If i = 7 Then Frmmain.Show
          If i = 2 Then
            Cancel = 1
            Exit Sub
          End If
    End If
End Sub

Private Sub lingcun_Click()
On Error GoTo ErrHandler
Cmdlog2.InitDir = App.path & "\例子"
Cmdlog2.Filter = "Twobody_Files (*.twb)|*.twb"
Cmdlog2.ShowSave
Open Cmdlog2.FileName For Output As #2
Print #2, Text1.Text
Print #2, Text12.Text
Print #2, Text2.Text
Print #2, Text3.Text
Print #2, Text4.Text
Print #2, Text5.Text
Print #2, Text6.Text
Print #2, Text7.Text
Print #2, Text8.Text
Print #2, Text9.Text
Print #2, Text10.Text
Print #2, Text11.Text
Close 2
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub Picture1_DblClick()
Picture1.Circle (CurrentX, CurrentY), 10, vbRed
End Sub
Private Sub queren1_Click()
Dim judgeM1 As Single
Dim judgeM2 As Single
judgeM1 = Val(Text2.Text)
judgeM2 = Val(Text7.Text)
If judgeM1 <> 0 And judgeM2 <> 0 Then
    If jiance(b1, b2) = True Then
        Loading
        Command1.Enabled = True
        Command2.Enabled = False
        queren1.Enabled = False
        xiugai.Enabled = True
        Give
        ChangeEnabled (False)
    Else
        MsgBox "两天体位置重叠!", , "警告!"
    End If
Else
    MsgBox "质量不能为0或为负值！", , "警告!"
End If
End Sub

Private Sub queren2_Click()
g = Val(Text1.Text)
dt = Val(Text12.Text)
End Sub

Private Sub Timer1_Timer()
Call Count_2Body(b1, b2, g, dt)
Call suggest
Call pengzhaung2
If guiji = True Then
    Picture2.PSet (b1.Posi.x, b1.Posi.y), vbRed
    Picture2.PSet (b2.Posi.x, b2.Posi.y), vbBlue
    Picture1.PSet (b1.Posi.x - b2.Posi.x, b1.Posi.y - b2.Posi.y), vbRed
    Picture3.PSet (b2.Posi.x - b1.Posi.x, b2.Posi.y - b1.Posi.y), vbBlue
Else
    Picture1.Cls
    Picture2.Cls
    Picture3.Cls
    Picture2.Circle (b1.Posi.x, b1.Posi.y), 1, vbRed
    Picture2.Circle (b2.Posi.x, b2.Posi.y), 1, vbBlue
    Picture1.Circle (b1.Posi.x - b2.Posi.x, b1.Posi.y - b2.Posi.y), 1, vbRed
    Picture3.Circle (b2.Posi.x - b1.Posi.x, b2.Posi.y - b1.Posi.y), 1, vbBlue
End If
End Sub
Private Sub suggest()
Text2.Text = Str(b1.M)
Text3.Text = Str(b1.Posi.x)
Text4.Text = Str(b1.Posi.y)
Text5.Text = Str(b1.V.x)
Text6.Text = Str(b1.V.y)

Text7.Text = Str(b2.M)
Text8.Text = Str(b2.Posi.x)
Text9.Text = Str(b2.Posi.y)
Text10.Text = Str(b2.V.x)
Text11.Text = Str(b2.V.y)

Text1.Text = Str(g)
Text12.Text = Str(dt)
End Sub
Private Sub ChangeEnabled(bo As Boolean)
    Text2.Enabled = bo
    Text3.Enabled = bo
    Text4.Enabled = bo
    Text5.Enabled = bo
    Text6.Enabled = bo
    Text7.Enabled = bo
    Text8.Enabled = bo
    Text9.Enabled = bo
    Text10.Enabled = bo
    Text11.Enabled = bo
    xiugai.Enabled = Not bo
    queren1.Enabled = bo
End Sub
Private Sub Loading()
    g = Val(Text1.Text)
    dt = Val(Text12.Text)
    b1.M = Val(Text2.Text)
    b1.Posi.x = Val(Text3.Text)
    b1.Posi.y = Val(Text4.Text)
    b1.V.x = Val(Text5.Text)
    b1.V.y = Val(Text6.Text)
    b2.M = Val(Text7.Text)
    b2.Posi.x = Val(Text8.Text)
    b2.Posi.y = Val(Text9.Text)
    b2.V.x = Val(Text10.Text)
    b2.V.y = Val(Text11.Text)
End Sub
Private Sub Give()
oldb1 = b1
oldb2 = b2
oldG = g
olddt = dt
End Sub
Private Sub xiugai_Click()
Command1.Enabled = False
Command2.Enabled = False
Timer1.Enabled = False
queren1.Enabled = True
xiugai.Enabled = False
Text1.Enabled = True
Text2.Enabled = True
Text3.Enabled = True
Text4.Enabled = True
Text5.Enabled = True
Text6.Enabled = True
Text7.Enabled = True
Text8.Enabled = True
Text9.Enabled = True
Text10.Enabled = True
Text11.Enabled = True
Text12.Enabled = True
End Sub
Private Sub pengzhaung2()
Dim i As Single
i = (b1.Posi.x - b2.Posi.x) ^ 2 + (b1.Posi.y - b2.Posi.y) ^ 2
i = Sqr(i)
If i <= 2 Then
    MsgBox "两天体相撞！", 0, "提示"
    Timer1.Enabled = False
End If
End Sub
Private Function jiance(b1 As body, b2 As body) As Boolean
Dim dd12 As Single
dd12 = (b1.Posi.x - b2.Posi.x) ^ 2 + (b1.Posi.y + b2.Posi.y) ^ 2
If dd12 = 0 Then
jiance = False
Else
jiance = True
End If
End Function
