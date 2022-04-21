VERSION 5.00
Begin VB.Form FormAdd2Body 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9435
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   9435
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   525
      Left            =   7080
      TabIndex        =   36
      Top             =   3150
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   525
      Left            =   7125
      TabIndex        =   35
      Top             =   2340
      Width           =   1455
   End
   Begin VB.Frame Frame4 
      Caption         =   "参数"
      Height          =   1365
      Left            =   6690
      TabIndex        =   30
      Top             =   540
      Width           =   2490
      Begin VB.TextBox Text22 
         Height          =   270
         Left            =   615
         TabIndex        =   32
         Text            =   "6.27"
         Top             =   390
         Width           =   1005
      End
      Begin VB.TextBox Text23 
         Height          =   270
         Left            =   1275
         TabIndex        =   31
         Text            =   "0.1"
         Top             =   855
         Width           =   1140
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "G:"
         Height          =   180
         Left            =   225
         TabIndex        =   34
         Top             =   435
         Width           =   180
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "迭代步长t:"
         Height          =   180
         Left            =   195
         TabIndex        =   33
         Top             =   885
         Width           =   900
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Body2"
      Height          =   1710
      Left            =   270
      TabIndex        =   15
      Top             =   2235
      Width           =   6060
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   4396
         TabIndex        =   22
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   2445
         TabIndex        =   21
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text10 
         Height          =   270
         Left            =   555
         TabIndex        =   20
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text11 
         Height          =   270
         Left            =   4396
         TabIndex        =   19
         Text            =   "0"
         Top             =   675
         Width           =   1245
      End
      Begin VB.TextBox Text12 
         Height          =   270
         Left            =   2445
         TabIndex        =   18
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text13 
         Height          =   270
         Left            =   570
         TabIndex        =   17
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text14 
         Height          =   270
         Left            =   660
         TabIndex        =   16
         Text            =   "1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   180
         Index           =   7
         Left            =   4080
         TabIndex        =   29
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   180
         Index           =   8
         Left            =   2115
         TabIndex        =   28
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   180
         Index           =   9
         Left            =   210
         TabIndex        =   27
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vz:"
         Height          =   180
         Index           =   10
         Left            =   4080
         TabIndex        =   26
         Top             =   735
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量:"
         Height          =   180
         Index           =   11
         Left            =   150
         TabIndex        =   25
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vx:"
         Height          =   180
         Index           =   12
         Left            =   195
         TabIndex        =   24
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vy:"
         Height          =   180
         Index           =   13
         Left            =   2100
         TabIndex        =   23
         Top             =   750
         Width           =   270
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Body1"
      Height          =   1710
      Left            =   255
      TabIndex        =   0
      Top             =   330
      Width           =   6060
      Begin VB.TextBox Text1 
         Height          =   270
         Left            =   660
         MouseIcon       =   "FormAdd2Body.frx":0000
         TabIndex        =   7
         Text            =   "1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Left            =   555
         TabIndex        =   6
         Text            =   "0"
         Top             =   720
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Left            =   2445
         TabIndex        =   5
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Left            =   4396
         TabIndex        =   4
         Text            =   "0"
         Top             =   675
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Left            =   555
         TabIndex        =   3
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Left            =   2445
         TabIndex        =   2
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Left            =   4396
         TabIndex        =   1
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vy:"
         Height          =   180
         Index           =   2
         Left            =   2100
         TabIndex        =   14
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vx:"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   13
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量:"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vz:"
         Height          =   180
         Index           =   3
         Left            =   4080
         TabIndex        =   11
         Top             =   735
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   10
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   180
         Index           =   5
         Left            =   2115
         TabIndex        =   9
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   180
         Index           =   6
         Left            =   4080
         TabIndex        =   8
         Top             =   1250
         Width           =   180
      End
   End
End
Attribute VB_Name = "FormAdd2Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub Command1_Click()
With body1
.M = Val(Text1.Text)
.V.X = Val(Text2.Text)
.V.Y = Val(Text3.Text)
.V.z = Val(Text4.Text)
.Posi.X = Val(Text5.Text)
.Posi.Y = Val(Text6.Text)
.Posi.z = Val(Text7.Text)
End With
With body2
.M = Val(Text14.Text)
.V.X = Val(Text13.Text)
.V.Y = Val(Text12.Text)
.V.z = Val(Text11.Text)
.Posi.X = Val(Text10.Text)
.Posi.Y = Val(Text9.Text)
.Posi.z = Val(Text8.Text)
End With
'G = Val(Text22.Text)
'dt = Val(Text23.Text)
panduan = 2
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
Private Sub Text1_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub
Private Sub Text10_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text14_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text22_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text23_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub
