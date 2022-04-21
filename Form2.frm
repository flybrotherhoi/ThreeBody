VERSION 5.00
Begin VB.Form frmintro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "提示"
   ClientHeight    =   8190
   ClientLeft      =   7665
   ClientTop       =   2655
   ClientWidth     =   8895
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":1BEA
   ScaleHeight     =   8190
   ScaleWidth      =   8895
   Begin VB.Label Label6 
      BackColor       =   &H8000000C&
      Caption         =   "2.按住鼠标左键，同时按W,A,S,D,Q,E可分别使镜头在X轴，Y轴，Z轴上移动"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   840
      TabIndex        =   5
      Top             =   2040
      Width           =   7440
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "3.单击恢复可回到上次新建或打开的状态"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   4
      Top             =   3240
      Width           =   5400
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "4.每次点击新建时输入框的内容都是当前值"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   3
      Top             =   4320
      Width           =   5700
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000C&
      Caption         =   "5.每次开始时总是从上次天体所在位置画轨迹，此步骤是为了快速找到天体位置，若想避免轨迹混淆，可在开始运动后马上点击“轨迹清除”"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   900
      Left            =   840
      TabIndex        =   2
      Top             =   5520
      Width           =   7740
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000C&
      Caption         =   "1.单击界面后可以使用滚轮调整摄像机位置"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   840
      TabIndex        =   1
      Top             =   1080
      Width           =   5700
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H8000000A&
      Caption         =   "操作提示"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3120
      TabIndex        =   0
      Top             =   360
      Width           =   1740
   End
End
Attribute VB_Name = "frmintro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

