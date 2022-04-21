VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Frmmain 
   AutoRedraw      =   -1  'True
   Caption         =   "三体运动模拟系统"
   ClientHeight    =   9795
   ClientLeft      =   4395
   ClientTop       =   2235
   ClientWidth     =   15840
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   42
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "Frmmain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9795
   ScaleWidth      =   15840
   Begin 三体运动模拟系统.jcbutton jcbutton3 
      Height          =   495
      Left            =   6850
      TabIndex        =   16
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "轨迹清除"
      Picture         =   "Frmmain.frx":1BEA
      PictureHover    =   "Frmmain.frx":218A
   End
   Begin VB.PictureBox Picture1 
      Height          =   9435
      Left            =   -810
      ScaleHeight     =   9375
      ScaleWidth      =   17865
      TabIndex        =   15
      Top             =   480
      Width           =   17925
   End
   Begin MSComDlg.CommonDialog cmdlog 
      Left            =   14880
      Top             =   1230
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   8145
      Top             =   900
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   12750
      Top             =   645
   End
   Begin 三体运动模拟系统.jcbutton xinjian 
      Height          =   495
      Index           =   0
      Left            =   0
      TabIndex        =   0
      ToolTipText     =   "新建"
      Top             =   0
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "新建"
      Picture         =   "Frmmain.frx":272A
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton dakai 
      Height          =   495
      Index           =   0
      Left            =   840
      TabIndex        =   1
      ToolTipText     =   "打开"
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "打开"
      Picture         =   "Frmmain.frx":2B88
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton baocun 
      Height          =   495
      Index           =   0
      Left            =   1710
      TabIndex        =   2
      ToolTipText     =   "保存"
      Top             =   0
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "保存"
      Picture         =   "Frmmain.frx":3032
   End
   Begin 三体运动模拟系统.jcbutton jcbutton4 
      Height          =   495
      Index           =   0
      Left            =   2545
      TabIndex        =   3
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ""
   End
   Begin 三体运动模拟系统.jcbutton bofang 
      Height          =   495
      Index           =   0
      Left            =   2820
      TabIndex        =   4
      ToolTipText     =   "播放"
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   " 运动"
      Picture         =   "Frmmain.frx":3414
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton zanting 
      Height          =   495
      Index           =   1
      Left            =   4085
      TabIndex        =   5
      ToolTipText     =   "暂停"
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   " 暂停"
      Picture         =   "Frmmain.frx":394F
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton jcbutton4 
      Height          =   495
      Index           =   1
      Left            =   5320
      TabIndex        =   6
      Top             =   0
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ""
   End
   Begin 三体运动模拟系统.jcbutton jcbutton2 
      Height          =   495
      Left            =   5575
      TabIndex        =   7
      Top             =   0
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "轨迹隐藏"
      Value           =   -1  'True
      Picture         =   "Frmmain.frx":3EC3
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton jcbutton9 
      Height          =   495
      Left            =   8115
      TabIndex        =   8
      Top             =   0
      Width           =   1260
      _ExtentX        =   2223
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "恢复"
      Picture         =   "Frmmain.frx":4355
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton jcbutton10 
      Height          =   495
      Left            =   9405
      TabIndex        =   9
      Top             =   0
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ""
   End
   Begin 三体运动模拟系统.jcbutton jcbutton1 
      Height          =   495
      Left            =   9690
      TabIndex        =   10
      Top             =   0
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "提示"
      Picture         =   "Frmmain.frx":48F5
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton guanyu 
      Height          =   495
      Index           =   0
      Left            =   10770
      TabIndex        =   11
      ToolTipText     =   "关于"
      Top             =   0
      Width           =   1005
      _ExtentX        =   1773
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "关于"
      Picture         =   "Frmmain.frx":4DA5
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton bangzhu 
      Height          =   495
      Index           =   0
      Left            =   11820
      TabIndex        =   12
      ToolTipText     =   "帮助"
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "帮助"
      Picture         =   "Frmmain.frx":52C5
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton jcbutton11 
      Height          =   495
      Left            =   12840
      TabIndex        =   13
      Top             =   0
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   873
      ButtonStyle     =   4
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   "退出"
      Picture         =   "Frmmain.frx":581B
      UseMaskCOlor    =   -1  'True
      MaskColor       =   16777215
   End
   Begin 三体运动模拟系统.jcbutton jcbutton4 
      Height          =   495
      Index           =   2
      Left            =   -75
      TabIndex        =   14
      Top             =   0
      Width           =   25005
      _ExtentX        =   44106
      _ExtentY        =   873
      ButtonStyle     =   4
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   14935011
      Caption         =   ""
   End
   Begin VB.Menu wenjian 
      Caption         =   "文件(&F)"
      NegotiatePosition=   1  'Left
      Begin VB.Menu new 
         Caption         =   "新建(&N)"
         Shortcut        =   ^N
      End
      Begin VB.Menu open 
         Caption         =   "打开文件(&O)"
         Shortcut        =   ^O
      End
      Begin VB.Menu save 
         Caption         =   "保存(&S)"
         Shortcut        =   ^S
      End
      Begin VB.Menu m30 
         Caption         =   "-"
      End
      Begin VB.Menu end 
         Caption         =   "退出(&E)"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu twobody 
      Caption         =   "二体模式"
   End
   Begin VB.Menu watching 
      Caption         =   "视图(&W)"
      Begin VB.Menu watchmove 
         Caption         =   "运动参数查看"
         Shortcut        =   {F1}
      End
   End
   Begin VB.Menu run 
      Caption         =   "运行(&R)"
      Begin VB.Menu movebegin 
         Caption         =   "运动(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu stop 
         Caption         =   "暂停(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu renew 
         Caption         =   "恢复(&R)"
         Shortcut        =   ^R
      End
      Begin VB.Menu path 
         Caption         =   "轨迹查看(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu clear 
         Caption         =   "轨迹清除（&C）"
         Shortcut        =   ^C
      End
   End
   Begin VB.Menu help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu helpp 
         Caption         =   "帮助(&H)"
         Shortcut        =   ^H
      End
      Begin VB.Menu about 
         Caption         =   "关于(&A)"
         Shortcut        =   ^A
      End
   End
End
Attribute VB_Name = "Frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mose As Boolean
Dim changepoz As Boolean
Dim SugRoutine As Boolean

Private Sub about_Click()
Call guanyu_Click(0)
End Sub

Private Sub bangzhu_Click(Index As Integer)
Shell "hh.exe " & App.path & "\help.chm", vbMaximizedFocus
Timer1.Enabled = False
End Sub

Private Sub baocun_Click(Index As Integer)
Dim i As Integer
On Error GoTo ErrHandler
cmdlog.InitDir = App.path & "\例子"
cmdlog.Filter = "Threebody_Files (*.thrb)|*.thrb"
cmdlog.ShowSave
Open cmdlog.FileName For Output As #2
    For i = 1 To 3
        Print #2, Str(oldbody(i).M)
        Print #2, Str(oldbody(i).V.x)
        Print #2, Str(oldbody(i).V.y)
        Print #2, Str(oldbody(i).V.z)
        Print #2, Str(oldbody(i).Posi.x)
        Print #2, Str(oldbody(i).Posi.y)
        Print #2, Str(oldbody(i).Posi.z)
    Next i
Print #2, frmchangewatching.Text1.Text
Print #2, frmchangewatching.Text2.Text
Close #2
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub
Private Sub bofang_Click(Index As Integer)
Timer1.Enabled = True: jcbutton9.Enabled = True

End Sub

Private Sub clear_Click()
Call reSetRoutine
End Sub

Private Sub dakai_Click(Index As Integer)
Dim i As Integer
Dim duqu As String
'CancelError 为 True。
On Error GoTo ErrHandler
'设置过滤器。
cmdlog.InitDir = App.path & "\例子"
cmdlog.Filter = "Threebody_Files (*.thrb)|*.thrb"
'指定缺省过滤器。
cmdlog.FilterIndex = 2
'显示“打开”对话框。
cmdlog.ShowOpen
'调用打开文件的过程。
Open cmdlog.FileName For Input As #1
Do Until EOF(1)
        For i = 1 To 3
            Line Input #1, duqu
            body(i).M = Val(duqu)
            Line Input #1, duqu
            body(i).V.x = Val(duqu)
            Line Input #1, duqu
            body(i).V.y = Val(duqu)
            Line Input #1, duqu
            body(i).V.z = Val(duqu)
            Line Input #1, duqu
            body(i).Posi.x = Val(duqu)
            Line Input #1, duqu
            body(i).Posi.y = Val(duqu)
            Line Input #1, duqu
            body(i).Posi.z = Val(duqu)
        Next i
    Line Input #1, duqu
    g = Val(duqu)
    oldG = g
    frmchangewatching.Text1.Text = Str(g)
    Line Input #1, duqu
    dt = Val(duqu)
    olddt = dt
    frmchangewatching.Text2.Text = Str(dt)
Loop
Close 1
Timer1.Enabled = False
For i = 1 To 3
    oldbody(i) = body(i)
Next i
Call sugInform
Call reSetRoutine
baocun(0).Enabled = False
Exit Sub
ErrHandler:
'用户按“取消”按钮。
Exit Sub
End Sub

Private Sub end_Click()
End
End Sub
Private Sub Form_Load()
Dim i As Integer
NotSaved = False
SugRoutine = True
dt = 0.1
g = 62.7
oldG = g
olddt = dt
frmchangewatching.Text1.Text = Str(g)
frmchangewatching.Text2.Text = Str(dt)

mose = False

'加载初值
With body(1)
    .M = 50
    .Posi.x = 200
    .Posi.y = 150
    .Posi.z = 0
    .V.x = 0
    .V.y = 0
    .V.z = 5
End With
With body(2)
    .M = 50
    .Posi.x = 0
    .Posi.y = 0
    .Posi.z = 0
    .V.x = 0
    .V.y = 0
    .V.z = 0
End With
With body(3)
    .M = 50
    .Posi.x = -200
    .Posi.y = -150
    .Posi.z = 0
    .V.x = 0
    .V.y = 0
    .V.z = -5
End With
Call init_Tv3D
Call reSetRoutine
Me.Show
For i = 1 To 3 Step 1
    oldbody(i) = body(i)
Next i
Call sugInform
'设置查看窗体始终置顶
SetWindowPos frmchangewatching.hwnd, HWND_TOPMOST, 0, 0, 0, 0, Flag
frmchangewatching.Show
Tv.DisplayFPS True
Do '主循环
'
Inp.GetMouseState Mx, My, b1, b2, , , Roll           '接收鼠标信息
If mose = True Then
  CameraAngX = CameraAngX - 0.1 * Mx
  CameraAngY = CameraAngY - 0.1 * My

'    限制范围
     If CameraAngY > 90 Then CameraAngY = 90
     If CameraAngY < -90 Then CameraAngY = -90

    If Inp.IsKeyPressed(TV_KEY_Q) Then CameraPozY = CameraPozY + 0.1
    If Inp.IsKeyPressed(TV_KEY_E) Then CameraPozY = CameraPozY - 0.1
    If Inp.IsKeyPressed(TV_KEY_W) Then CameraPozZ = CameraPozZ - 0.1
    If Inp.IsKeyPressed(TV_KEY_S) Then CameraPozZ = CameraPozZ + 0.1
    If Inp.IsKeyPressed(TV_KEY_A) Then CameraPozX = CameraPozX + 0.1
    If Inp.IsKeyPressed(TV_KEY_D) Then CameraPozX = CameraPozX - 0.1
'设定摄像机
Camera.SetRotation CameraAngY, CameraAngX, 0
End If
'滚轮改变摄像机位置
If changepoz = True Then
If CameraPozZ - Roll / 100 > 1 Then CameraPozZ = CameraPozZ - Roll / 100 '是否越界
If CameraPozY - Roll / 100 > 1 Then CameraPozY = CameraPozY - Roll / 100
If CameraPozX - Roll / 100 > 1 Then CameraPozX = CameraPozX - Roll / 100
Camera.SetPosition CameraPozX, CameraPozY, CameraPozZ
strx.SetRotation CameraAngY, CameraAngX, 0
stry.SetRotation CameraAngY, CameraAngX, 0
strz.SetRotation CameraAngY, CameraAngX, 0
End If

For i = 1 To 3
    mesh(i).SetPosition body(i).Posi.x / 50, body(i).Posi.y / 50, body(i).Posi.z / 50
    mesh(i).RotateY 0.5
Next i

Tv.clear '清屏
Atmos.Fog_Enable False
  Atmos.Atmosphere_Render '渲染大气
Atmos.Fog_Enable True
For i = 1 To 3
    mesh(i).Render
Next i
If SugRoutine = True Then
    For i = 1 To 3
        Routine(i).Render
    Next i
Else
End If
strx.Render
stry.Render
strz.Render
Floor.Render
 Tv.RenderToScreen '把所得最终结果渲染到屏幕
DoEvents '这句是把线程空出来，使其他的的程序能运行，必加
Loop
End Sub
Private Sub Introduce_Click()
frmintro.Show
End Sub
Private Sub Form_LostFocus()
mose = False
End Sub

Private Sub Form_Resize()
Picture1.Left = 0
Picture1.Width = Me.ScaleWidth
Picture1.Top = jcbutton4(0).Height
Picture1.Height = Me.ScaleHeight - jcbutton4(0).Height
Tv.ResizeDevice
End Sub

Private Sub guanyu_Click(Index As Integer)
frmAbout.Show
End Sub

Private Sub helpp_Click()
Call bangzhu_Click(0)
End Sub

Private Sub jcbutton1_Click()
frmintro.Show
End Sub
Private Sub jcbutton11_Click()
Unload Me
End Sub

Private Sub jcbutton2_Click()
SugRoutine = Not SugRoutine
If jcbutton2.Caption = "轨迹显示" Then
    jcbutton2.Caption = "轨迹隐藏"
Else
    jcbutton2.Caption = "轨迹显示"
End If
End Sub

Private Sub jcbutton3_Click()
Call reSetRoutine
End Sub

Private Sub jcbutton9_Click()
Call renew_Click
End Sub

Private Sub movebegin_Click()
Call bofang_Click(0)
End Sub

Private Sub new_Click()
Call xinjian_Click(0)
End Sub

Private Sub open_Click()
Call dakai_Click(0)
End Sub

Private Sub path_Click()
Call jcbutton2_Click
End Sub

Private Sub picture1_GotFocus()
changepoz = True
End Sub

Private Sub picture1_LostFocus()
changepoz = False
End Sub

Private Sub picture1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
mose = True
Picture1.MousePointer = 15
End Sub

Private Sub picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
mose = False
Picture1.MousePointer = 0
End Sub
'恢复初状态
Private Sub renew_Click()
Dim i As Integer
Timer1.Enabled = False
If MsgBox("确定恢复？", vbYesNo) = vbYes Then
    For i = 1 To 3 Step 1
        body(i) = oldbody(i)
        Call reSetRoutine
    Next i
    jcbutton9.Enabled = False
Else
    Timer1.Enabled = True
End If
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then  '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub save_Click()
Call baocun_Click(0)
End Sub

Private Sub stop_Click()
Call zanting_Click(1)
End Sub

 '三体运动
Private Sub Timer1_Timer()
Call pengzhuang
Call Count_3Body(body(1), body(2), body(3), g, dt)
Call guiji(body(1), body(2), body(3))
Call countE_3Body(body(1), body(2), body(3), g)
Call sugInform
Call pengzhuang
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim i As Integer
If NotSaved = False Then
    If MsgBox("确定退出？", vbYesNo) = vbYes Then
    Tv.ReleaseAll
    End
    Else
    Cancel = 1
    End If
Else
    i = MsgBox("是否保存？", vbYesNoCancel)
    If i = 6 Then
        Call baocun_Click(0)
        Tv.ReleaseAll
        End
    Else: If i = 7 Then Tv.ReleaseAll: End
          If i = 2 Then
            Cancel = 1
            Exit Sub
          End If
    End If
End If
End Sub

Private Sub Timer2_Timer()
save.Enabled = baocun(0).Enabled
If Timer1.Enabled = True Then
    zanting(1).Enabled = True
    bofang(0).Enabled = False
Else
    Timer1.Enabled = False
    zanting(1).Enabled = False
    bofang(0).Enabled = True
End If
End Sub

Private Sub twobody_Click()
TwoBodyMode.Show
Me.Hide
Timer1.Enabled = False
frmchangewatching.Hide
End Sub

Private Sub watchmove_Click()
frmchangewatching.Show
End Sub

Private Sub xinjian_Click(Index As Integer)
Timer1.Enabled = False
frmAdd3Body.Show 1
End Sub

Private Sub zanting_Click(Index As Integer)
Timer1.Enabled = False
End Sub
Private Sub pengzhuang()
Dim D12 As Single, D23 As Single, D13 As Single
Dim i As Integer
D12 = (body(1).Posi.x - body(2).Posi.x) ^ 2 + (body(1).Posi.y - body(2).Posi.y) ^ 2 + (body(1).Posi.z - body(2).Posi.z) ^ 2
D12 = Sqr(D12)
D23 = (body(2).Posi.x - body(3).Posi.x) ^ 2 + (body(2).Posi.y - body(3).Posi.y) ^ 2 + (body(2).Posi.z - body(3).Posi.z) ^ 2
D23 = Sqr(D23)
D13 = (body(1).Posi.x - body(3).Posi.x) ^ 2 + (body(1).Posi.y - body(3).Posi.y) ^ 2 + (body(1).Posi.z - body(3).Posi.z) ^ 2
D13 = Sqr(D13)
If D12 > 5 Then
    If D23 > 5 Then
        If D13 > 5 Then
        Else
            Timer1.Enabled = False
            If MsgBox("天体1与天体3相撞！是否重新开始？", vbYesNo, "消息") = vbYes Then
                For i = 1 To 3
                    body(i) = oldbody(i)
                    Routine(i).AddVertex body(i).Posi.x, body(i).Posi.y, body(i).Posi.z, 0, 0, 0, 0, 0, 0, 0, vbBlack
                Next i
            Else
                bofang(0).Enabled = False
            End If
        End If
    Else
        Timer1.Enabled = False
        If MsgBox("天体2与天体3相撞！是否重新开始？", vbYesNo, "消息") = vbYes Then
            For i = 1 To 3
                body(i) = oldbody(i)
                Routine(i).AddVertex body(i).Posi.x, body(i).Posi.y, body(i).Posi.z, 0, 0, 0, 0, 0, 0, 0, vbBlack
            Next i
        Else
            bofang(0).Enabled = False
        End If
    End If
Else
    Timer1.Enabled = False
    If MsgBox("天体1与天体2相撞！是否重新开始？", vbYesNo, "消息") = vbYes Then
        For i = 1 To 3
            body(i) = oldbody(i)
            Routine(i).AddVertex body(i).Posi.x, body(i).Posi.y, body(i).Posi.z, 0, 0, 0, 0, 0, 0, 0, vbBlack
        Next i
    Else
        bofang(0).Enabled = False
    End If
End If
End Sub
