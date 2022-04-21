VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmchangewatching 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "运动状况查看器"
   ClientHeight    =   9315
   ClientLeft      =   12885
   ClientTop       =   2175
   ClientWidth     =   2790
   Icon            =   "frmchangewatching.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9315
   ScaleWidth      =   2790
   Begin VB.Frame Frame1 
      Caption         =   "参数设置"
      Height          =   3570
      Left            =   120
      TabIndex        =   16
      Top             =   4725
      Width           =   2550
      Begin VB.TextBox Text1 
         Enabled         =   0   'False
         Height          =   255
         Left            =   570
         TabIndex        =   21
         Text            =   "0"
         Top             =   540
         Width           =   975
      End
      Begin VB.TextBox Text2 
         Enabled         =   0   'False
         Height          =   270
         Left            =   570
         TabIndex        =   19
         Text            =   "0"
         Top             =   1920
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "修改"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   150
         TabIndex        =   18
         Top             =   3015
         Width           =   1035
      End
      Begin VB.CommandButton Command2 
         Caption         =   "确定"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   1395
         TabIndex        =   17
         Top             =   3000
         Width           =   1005
      End
      Begin MSComctlLib.Slider Slider1 
         Height          =   240
         Left            =   675
         TabIndex        =   20
         Top             =   900
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   423
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin MSComctlLib.Slider Slider2 
         Height          =   240
         Left            =   675
         TabIndex        =   22
         Top             =   2400
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   423
         _Version        =   393216
         LargeChange     =   1
         Min             =   1
         SelStart        =   1
         Value           =   1
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "引力常量G:"
         Height          =   180
         Index           =   5
         Left            =   810
         TabIndex        =   28
         Top             =   315
         Width           =   900
      End
      Begin VB.Line Line4 
         X1              =   -30
         X2              =   2550
         Y1              =   1410
         Y2              =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "迭代步长dt:"
         Height          =   180
         Index           =   5
         Left            =   735
         TabIndex        =   27
         Top             =   1560
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "倍数:"
         Height          =   180
         Index           =   5
         Left            =   210
         TabIndex        =   26
         Top             =   960
         Width           =   450
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "倍数:"
         Height          =   180
         Left            =   210
         TabIndex        =   25
         Top             =   2445
         Width           =   450
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "*1"
         Height          =   180
         Left            =   1710
         TabIndex        =   24
         Top             =   585
         Width           =   180
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "*1"
         Height          =   180
         Left            =   1740
         TabIndex        =   23
         Top             =   1965
         Width           =   180
      End
      Begin VB.Line Line3 
         X1              =   15
         X2              =   2535
         Y1              =   2880
         Y2              =   2880
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4410
      Left            =   45
      TabIndex        =   0
      Top             =   165
      Width           =   2700
      _ExtentX        =   4763
      _ExtentY        =   7779
      _Version        =   393216
      Tab             =   2
      TabHeight       =   520
      TabCaption(0)   =   "Body1"
      TabPicture(0)   =   "frmchangewatching.frx":0742
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(2)=   "Label1(4)"
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(4)=   "Label1(2)"
      Tab(0).Control(5)=   "Label1(1)"
      Tab(0).Control(6)=   "Label1(0)"
      Tab(0).ControlCount=   7
      TabCaption(1)   =   "Body2"
      TabPicture(1)   =   "frmchangewatching.frx":075E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Picture2"
      Tab(1).Control(1)=   "Label9"
      Tab(1).Control(2)=   "Label2(4)"
      Tab(1).Control(3)=   "Label2(3)"
      Tab(1).Control(4)=   "Label2(2)"
      Tab(1).Control(5)=   "Label2(1)"
      Tab(1).Control(6)=   "Label2(0)"
      Tab(1).Control(7)=   "Line26"
      Tab(1).ControlCount=   8
      TabCaption(2)   =   "Body3"
      TabPicture(2)   =   "frmchangewatching.frx":077A
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "Label3(0)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Label3(1)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Label3(2)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "Label3(3)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label3(4)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "Label10"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "Picture3"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).ControlCount=   7
      Begin VB.PictureBox Picture3 
         Height          =   795
         Left            =   1305
         Picture         =   "frmchangewatching.frx":0796
         ScaleHeight     =   735
         ScaleWidth      =   720
         TabIndex        =   35
         Top             =   3315
         Width           =   780
      End
      Begin VB.PictureBox Picture2 
         Height          =   795
         Left            =   -73695
         Picture         =   "frmchangewatching.frx":9088
         ScaleHeight     =   735
         ScaleWidth      =   720
         TabIndex        =   32
         Top             =   3315
         Width           =   780
      End
      Begin VB.PictureBox Picture1 
         Height          =   795
         Left            =   -73695
         Picture         =   "frmchangewatching.frx":E27E
         ScaleHeight     =   735
         ScaleWidth      =   720
         TabIndex        =   30
         Top             =   3315
         Width           =   780
      End
      Begin VB.Label Label10 
         Caption         =   "材质："
         Height          =   360
         Left            =   630
         TabIndex        =   34
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "材质："
         Height          =   360
         Left            =   -74370
         TabIndex        =   33
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label7 
         Caption         =   "材质："
         Height          =   360
         Left            =   -74370
         TabIndex        =   31
         Top             =   3660
         Width           =   975
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   4
         Left            =   570
         TabIndex        =   15
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   3
         Left            =   570
         TabIndex        =   14
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   13
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   12
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Label3"
         Height          =   180
         Index           =   0
         Left            =   570
         TabIndex        =   11
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   4
         Left            =   -74430
         TabIndex        =   10
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   3
         Left            =   -74430
         TabIndex        =   9
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   2
         Left            =   -74430
         TabIndex        =   8
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   1
         Left            =   -74430
         TabIndex        =   7
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Label2"
         Height          =   180
         Index           =   0
         Left            =   -74430
         TabIndex        =   6
         Top             =   600
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Index           =   4
         Left            =   -74430
         TabIndex        =   5
         Top             =   3000
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Index           =   3
         Left            =   -74430
         TabIndex        =   4
         Top             =   2400
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Index           =   2
         Left            =   -74430
         TabIndex        =   3
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Index           =   1
         Left            =   -74430
         TabIndex        =   2
         Top             =   1200
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Label1"
         Height          =   180
         Index           =   0
         Left            =   -74430
         TabIndex        =   1
         Top             =   600
         Width           =   540
      End
      Begin VB.Line Line26 
         BorderColor     =   &H00FF0000&
         X1              =   -75000
         X2              =   -75000
         Y1              =   555
         Y2              =   0
      End
   End
   Begin VB.Label Label8 
      Caption         =   "注意：迭代步长取值过大将导致误差变大"
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   29
      Top             =   8535
      Width           =   2310
   End
End
Attribute VB_Name = "frmchangewatching"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Call sugInform
Text1.Text = Str(oldG)
Text2.Text = Str(olddt)
End Sub

Private Sub Form_Unload(Cancel As Integer)
oldG = Val(Text1.Text)
olddt = Val(Text2.Text)
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If x < 940 And y < 300 Then
    SSTab1.Tab = 0
End If
If x > 940 And x < 1820 And y < 300 Then
    SSTab1.Tab = 1
End If
If x > 1820 And x < 2700 And y < 300 Then
    SSTab1.Tab = 2
End If
End Sub
'改变G倍数
Private Sub Slider1_Change()
g = Val(frmchangewatching.Text1.Text) * Slider1.Value
Label5.Caption = "*" & Str(Slider1.Value)
End Sub
'改变dt倍数
Private Sub Slider2_Change()
dt = Val(frmchangewatching.Text2.Text) * Slider2.Value
Label6.Caption = "*" & Str(Slider2.Value)
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then  '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub
Private Sub Command1_Click()
Text1.Enabled = True
Text2.Enabled = True
Command1.Enabled = False
Command2.Enabled = True
Frmmain.Timer1.Enabled = False
End Sub
Private Sub Command2_Click()
g = Val(frmchangewatching.Text1.Text)
dt = Val(frmchangewatching.Text2.Text)
Text1.Enabled = False
Text2.Enabled = False
Command2.Enabled = False
Command1.Enabled = True
Frmmain.Timer1.Enabled = True
Frmmain.jcbutton9.Enabled = True
Slider1.Value = 1
Slider2.Value = 1
End Sub

