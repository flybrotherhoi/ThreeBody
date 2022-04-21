VERSION 5.00
Begin VB.Form frmAdd3Body 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "输入"
   ClientHeight    =   7080
   ClientLeft      =   7950
   ClientTop       =   3525
   ClientWidth     =   6780
   Icon            =   "FormAdd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   6780
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command3 
      Caption         =   "随机生成"
      Height          =   525
      Left            =   3345
      TabIndex        =   52
      Top             =   6165
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Caption         =   "Frame2"
      Height          =   945
      Left            =   270
      TabIndex        =   47
      Top             =   5880
      Width           =   2850
      Begin VB.TextBox Text9 
         Height          =   270
         Left            =   1440
         TabIndex        =   51
         Text            =   "Text9"
         Top             =   540
         Width           =   1080
      End
      Begin VB.TextBox Text8 
         Height          =   270
         Left            =   1440
         TabIndex        =   49
         Text            =   "Text8"
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label Label3 
         Caption         =   "迭代步长dt:"
         Height          =   225
         Left            =   225
         TabIndex        =   50
         Top             =   540
         Width           =   1080
      End
      Begin VB.Label Label2 
         Caption         =   "引力常量G:"
         Height          =   255
         Left            =   300
         TabIndex        =   48
         Top             =   285
         Width           =   945
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Body1"
      Height          =   1710
      Index           =   3
      Left            =   345
      TabIndex        =   32
      Top             =   3930
      Width           =   6060
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   2
         Left            =   660
         TabIndex        =   39
         Text            =   "1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   2
         Left            =   570
         TabIndex        =   38
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   2
         Left            =   2445
         TabIndex        =   37
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   2
         Left            =   4396
         TabIndex        =   36
         Text            =   "0"
         Top             =   675
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   2
         Left            =   555
         TabIndex        =   35
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   2
         Left            =   2445
         TabIndex        =   34
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   2
         Left            =   4396
         TabIndex        =   33
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vy:"
         Height          =   180
         Index           =   20
         Left            =   2100
         TabIndex        =   46
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vx:"
         Height          =   180
         Index           =   19
         Left            =   195
         TabIndex        =   45
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量:"
         Height          =   180
         Index           =   18
         Left            =   150
         TabIndex        =   44
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vz:"
         Height          =   180
         Index           =   17
         Left            =   4080
         TabIndex        =   43
         Top             =   735
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   180
         Index           =   16
         Left            =   210
         TabIndex        =   42
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   180
         Index           =   15
         Left            =   2115
         TabIndex        =   41
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   180
         Index           =   14
         Left            =   4080
         TabIndex        =   40
         Top             =   1250
         Width           =   180
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Body1"
      Height          =   1710
      Index           =   2
      Left            =   360
      TabIndex        =   17
      Top             =   2115
      Width           =   6060
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   1
         Left            =   660
         TabIndex        =   24
         Text            =   "1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   1
         Left            =   570
         TabIndex        =   23
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   1
         Left            =   2445
         TabIndex        =   22
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   1
         Left            =   4396
         TabIndex        =   21
         Text            =   "0"
         Top             =   675
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   1
         Left            =   555
         TabIndex        =   20
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   1
         Left            =   2445
         TabIndex        =   19
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   1
         Left            =   4396
         TabIndex        =   18
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vy:"
         Height          =   180
         Index           =   13
         Left            =   2100
         TabIndex        =   31
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vx:"
         Height          =   180
         Index           =   12
         Left            =   195
         TabIndex        =   30
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量:"
         Height          =   180
         Index           =   11
         Left            =   150
         TabIndex        =   29
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vz:"
         Height          =   180
         Index           =   10
         Left            =   4080
         TabIndex        =   28
         Top             =   735
         Width           =   270
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
         Caption         =   "Y:"
         Height          =   180
         Index           =   8
         Left            =   2115
         TabIndex        =   26
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   180
         Index           =   7
         Left            =   4080
         TabIndex        =   25
         Top             =   1250
         Width           =   180
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   525
      Left            =   5460
      TabIndex        =   16
      Top             =   6150
      Width           =   915
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   525
      Left            =   4380
      TabIndex        =   15
      Top             =   6165
      Width           =   915
   End
   Begin VB.Frame Frame1 
      Caption         =   "Body1"
      Height          =   1710
      Index           =   1
      Left            =   360
      TabIndex        =   0
      Top             =   330
      Width           =   6060
      Begin VB.TextBox Text7 
         Height          =   270
         Index           =   0
         Left            =   4396
         TabIndex        =   14
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text6 
         Height          =   270
         Index           =   0
         Left            =   2445
         TabIndex        =   13
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text5 
         Height          =   270
         Index           =   0
         Left            =   555
         TabIndex        =   12
         Text            =   "0"
         Top             =   1200
         Width           =   1245
      End
      Begin VB.TextBox Text4 
         Height          =   270
         Index           =   0
         Left            =   4396
         TabIndex        =   11
         Text            =   "0"
         Top             =   675
         Width           =   1245
      End
      Begin VB.TextBox Text3 
         Height          =   270
         Index           =   0
         Left            =   2445
         TabIndex        =   10
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text2 
         Height          =   270
         Index           =   0
         Left            =   570
         TabIndex        =   9
         Text            =   "0"
         Top             =   700
         Width           =   1245
      End
      Begin VB.TextBox Text1 
         Height          =   270
         Index           =   0
         Left            =   660
         TabIndex        =   8
         Text            =   "1"
         Top             =   285
         Width           =   1425
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Z:"
         Height          =   180
         Index           =   6
         Left            =   4080
         TabIndex        =   7
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Y:"
         Height          =   180
         Index           =   5
         Left            =   2115
         TabIndex        =   6
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "X:"
         Height          =   180
         Index           =   4
         Left            =   210
         TabIndex        =   5
         Top             =   1250
         Width           =   180
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vz:"
         Height          =   180
         Index           =   3
         Left            =   4080
         TabIndex        =   4
         Top             =   735
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "质量:"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   3
         Top             =   330
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vx:"
         Height          =   180
         Index           =   1
         Left            =   195
         TabIndex        =   2
         Top             =   750
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vy:"
         Height          =   180
         Index           =   2
         Left            =   2100
         TabIndex        =   1
         Top             =   750
         Width           =   270
      End
   End
End
Attribute VB_Name = "frmAdd3Body"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim jilu1(3) As String
Dim jilu2(3) As String
Dim jilu3(3) As String
Dim jilu4(3) As String
Dim jilu5(3) As String
Dim jilu6(3) As String
Dim jilu7(3) As String
Dim jilu8 As String
Dim jilu9 As String
Dim zzbody(1 To 3) As body
Dim i As Integer
Dim j(1 To 6, 1 To 3) As Single
Dim n As Integer
Dim judge As Integer
Private Sub Command1_Click()
Dim judgeM(0 To 2) As Single
For i = 0 To 2
judgeM(i) = Val(Text1(i).Text)
Next i
For i = 1 To 3 Step 1
    With zzbody(i)
    .M = Val(Text1(i - 1).Text)
    .V.x = Val(Text2(i - 1).Text)
    .V.y = Val(Text3(i - 1).Text)
    .V.z = Val(Text4(i - 1).Text)
    .Posi.x = Val(Text5(i - 1).Text)
    .Posi.y = Val(Text6(i - 1).Text)
    .Posi.z = Val(Text7(i - 1).Text)
    End With
Next i
If judgeM(0) > 0 And judgeM(1) > 0 And judgeM(2) > 0 Then
    If jiance(zzbody(1), zzbody(2)) = True And jiance(zzbody(2), zzbody(3)) = True And jiance(zzbody(1), zzbody(3)) = True Then
        For i = 1 To 3 Step 1
            With body(i)
            .M = Val(Text1(i - 1).Text)
            .V.x = Val(Text2(i - 1).Text)
            .V.y = Val(Text3(i - 1).Text)
            .V.z = Val(Text4(i - 1).Text)
            .Posi.x = Val(Text5(i - 1).Text)
            .Posi.y = Val(Text6(i - 1).Text)
            .Posi.z = Val(Text7(i - 1).Text)
            End With
        Next i
        g = Text8.Text
        dt = Text9.Text
        frmchangewatching.Text1.Text = Str(g)
        frmchangewatching.Text2.Text = Str(dt)
        Call sugInform
        '判断数据是否改变
        For i = 0 To 2
            If jilu1(i) <> Text1(i).Text Then judge = judge + 1
            If jilu2(i) <> Text2(i).Text Then judge = judge + 1
            If jilu3(i) <> Text3(i).Text Then judge = judge + 1
            If jilu4(i) <> Text4(i).Text Then judge = judge + 1
            If jilu5(i) <> Text5(i).Text Then judge = judge + 1
            If jilu6(i) <> Text6(i).Text Then judge = judge + 1
            If jilu7(i) <> Text7(i).Text Then judge = judge + 1
        Next i
        If jilu8 <> Text8.Text Then judge = judge + 1
        If jilu9 <> Text9.Text Then judge = judge + 1
        If judge <> 0 Then
        NotSaved = True
        Frmmain.baocun(0).Enabled = True
        End If
        For i = 1 To 3 Step 1
            oldbody(i) = body(i)
        Next i
        Call reSetRoutine
        Unload Me
    Else
        MsgBox "位置重叠!请确认数据无误.", , "警告!"
    End If
Else
    MsgBox "质量不能为0或负值!请确认数据无误.", , "警告!"
End If
End Sub
Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Randomize
For i = 1 To 6
    For n = 1 To 3
        j(i, n) = 2 * Rnd - 1
    Next n
Next i
For i = 1 To 3 Step 1
Text2(i - 1).Text = Str(j(1, i))
Text3(i - 1).Text = Str(j(2, i))
Text4(i - 1).Text = Str(j(3, i))
Text5(i - 1).Text = Str(j(4, i)) * 500
Text6(i - 1).Text = Str(j(5, i)) * 500
Text7(i - 1).Text = Str(j(6, i)) * 500
Next i

End Sub

Private Sub Form_Load()
judge = 0
For i = 1 To 3 Step 1
Text1(i - 1).Text = Str(body(i).M)
Text2(i - 1).Text = Str(body(i).V.x)
Text3(i - 1).Text = Str(body(i).V.y)
Text4(i - 1).Text = Str(body(i).V.z)
Text5(i - 1).Text = Str(body(i).Posi.x)
Text6(i - 1).Text = Str(body(i).Posi.y)
Text7(i - 1).Text = Str(body(i).Posi.z)
Next i
Text8.Text = Str(g)
Text9.Text = Str(dt)
For i = 0 To 2
    jilu1(i) = Text1(i).Text
    jilu2(i) = Text2(i).Text
    jilu3(i) = Text3(i).Text
    jilu4(i) = Text4(i).Text
    jilu5(i) = Text5(i).Text
    jilu6(i) = Text6(i).Text
    jilu7(i) = Text7(i).Text
Next i
jilu8 = Text8.Text
jilu9 = Text9.Text
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

Private Sub Text1_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub


Private Sub Text2_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text3_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text4_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text5_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text6_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If KeyAscii >= 45 And KeyAscii <= 57 Or KeyAscii = 8 Then '数字0到9的ascii值为48到57
Else
KeyAscii = 0 '表示输入如果不数字则为NULL 即空
End If
End Sub

