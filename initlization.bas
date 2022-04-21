Attribute VB_Name = "Initialization"
Public body(1 To 3)  As body    '定义三个物体
Public oldbody(1 To 3) As body   '记录初状态
Public panduan As Integer
Public NotSaved As Boolean      '判断是否已经保存
Public g As Single          '引力常量
Public dt As Single         '迭代步长
Public oldG As Single
Public olddt As Single
Public Tv As New TVEngine '调用tv3d所必需的
Public Scene As New TVScene '调用tv3d所必需的
Public TF As New TVTextureFactory '添加一个贴图库
Public MF As New TVMaterialFactory ''添加一个材质库
Public LE As New TVLightEngine '添加一个灯光库
Public Atmos  As New TVAtmosphere '添加大气系统
Public Inp As New TVInputEngine
Public Mx As Long, My As Long, b1 As Boolean, b2 As Boolean, Roll As Long   '接收鼠标信息
Public Camera As New TVCamera '定义一个摄像机，相当于人的眼睛
Public CameraPozX As Single, CameraPozY As Single, CameraPozZ As Single '摄像机位置坐标
Public CameraAngX As Single, CameraAngY As Single '摄像机角度
Public Floor As TVMesh  '添加一个网格物体
Public Routine(1 To 3) As TVMesh
Public mesh(1 To 3) As TVMesh   '添加一个物体
Public strx As TVMesh   '坐标轴
Public stry As TVMesh
Public strz As TVMesh

  '''''初始化标签位置
Public Sub init_label()
Dim i As Integer
With frmchangewatching
.Label1(0).Left = 300
.Label2(0).Left = 300
.Label3(0).Left = 300
End With
With frmchangewatching
    For i = 1 To 4
    .Label1(i).Left = .Label1(0).Left
    .Label1(i).Top = .Label1(i - 1).Top + .Label1(i - 1).Height + 100
    Next i
    
    For i = 1 To 4
    .Label2(i).Left = .Label1(0).Left
    .Label2(i).Top = .Label2(i - 1).Top + .Label2(i - 1).Height + 100
    Next i
    
    For i = 1 To 4
    .Label3(i).Left = .Label1(0).Left
    .Label3(i).Top = .Label3(i - 1).Top + .Label3(i - 1).Height + 100
    Next i
End With

With frmchangewatching
    .Label1(0).Caption = "质量:0"
    .Label1(1).Caption = "速度:0"
    .Label1(2).Caption = "X:0"
    .Label1(3).Caption = "Y:0"
    .Label1(4).Caption = "Z:0"

    .Label2(0).Caption = "质量:0"
    .Label2(1).Caption = "速度:0"
    .Label2(2).Caption = "X:0"
    .Label2(3).Caption = "Y:0"
    .Label2(4).Caption = "Z:0"

    .Label3(0).Caption = "质量:0"
    .Label3(1).Caption = "速度:0"
    .Label3(2).Caption = "X:0"
    .Label3(3).Caption = "Y:0"
    .Label3(4).Caption = "Z:0"
End With

End Sub
Public Sub init_Tv3D()
'------------------------------------------------------------------Tv3d初始化
Tv.SetSearchDirectory App.path & "\Data\Picture" '设定贴图读取目录为当前目录
Tv.SetVSync True '垂直同步开关
Tv.Init3DWindowed Frmmain.Picture1.hwnd    '用窗口模式启动tv3d
Inp.Initialize
Tv.SetAngleSystem TV_ANGLE_DEGREE
 
TF.LoadTexture "body1.jpg", "1" '读取名为pic.jpg的贴图，并命名为pic
TF.LoadTexture "body2.jpg", "2" '读取名为pic.jpg的贴图，并命名为pic
TF.LoadTexture "body3.jpg", "3"
TF.LoadTexture "xing.jpg", "xing"

Atmos.SkyBox_Enable True '开启天空盒
  Atmos.SkyBox_SetTexture GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing") '设定贴图
Atmos.Fog_SetParameters 1, 200000, 0              '最近距离，最远距离，浓度

Scene.SetBackgroundColor 0.5, 0.5, 0.9  '背景颜色


'MF.CreateMaterialQuick 0, 1, 0, 0, "green"
'MF.CreateMaterialQuick 1, 0, 0, 0, "red"
'MF.CreateMaterialQuick 0, 0, 1, 0, "blue"
MF.CreateMaterial "2" '建立名为solid的材质
MF.SetAmbient GetMat("2"), 0, 0, 0, 1        '环境光
MF.SetDiffuse GetMat("2"), 1, 0, 0, 1 '扩散光，即物体的固有颜色
MF.SetEmissive GetMat("2"), 1, 0, 0, 1   '自发光
MF.SetOpacity GetMat("2"), 1 '不透明度
MF.SetSpecular GetMat("2"), 1, 1, 1, 1  '高光色
MF.SetPower GetMat("2"), 60 '散射强度

 MF.CreateMaterial "1" '建立蓝色的材质
MF.SetAmbient GetMat("1"), 0, 0, 1, 1       '环境光
MF.SetDiffuse GetMat("1"), 0, 1, 1, 1 '扩散光，即物体的固有颜色
MF.SetEmissive GetMat("1"), 0, 1, 1, 1   '自发光
MF.SetOpacity GetMat("1"), 1 '不透明度
MF.SetSpecular GetMat("1"), 1, 1, 1, 1  '高光色
MF.SetPower GetMat("1"), 60 '散射强度

  MF.CreateMaterial "3" '建立蓝色的材质
MF.SetAmbient GetMat("3"), 1, 1, 0, 1       '环境光
MF.SetDiffuse GetMat("3"), 1, 1, 0, 1 '扩散光，即物体的固有颜色
MF.SetEmissive GetMat("3"), 1, 1, 0, 1  '自发光
MF.SetOpacity GetMat("3"), 1 '不透明度
MF.SetSpecular GetMat("3"), 1, 1, 1, 1  '高光色
MF.SetPower GetMat("3"), 60 '散射强度

 LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, , 1  '添加一个平行光
 LE.SetSpecularLighting True  '高光开关
'三条轨迹

Set Floor = Scene.CreateMeshBuilder '网格物体初始化，必加
Floor.SetMeshFormat CONST_TV_MESHFORMAT.TV_MESHFORMAT_DIFFUSE + CONST_TV_MESHFORMAT.TV_MESHFORMAT_NOLIGHTING
Floor.SetLightingMode (CONST_TV_LIGHTINGMODE.TV_LIGHTING_NONE)
Floor.SetPrimitiveType (CONST_TV_PRIMITIVETYPE.TV_LINELIST)
'创建网格

Dim x
Dim y
For x = -100 To 100 Step 10
If x = 0 Then
Else
        Floor.AddVertex x, 0, -100, 0, 1, 0, 0, 0, 0, 0, -1
        Floor.AddVertex x, 0, 100, 0, 1, 0, 0, 1, 0, 0, -1
End If
Next x
For y = -100 To 100 Step 10
    If y = 0 Then
    Else
        Floor.AddVertex -100, 0, y, 0, 1, 0, 1, 0, 0, 0, -1
        Floor.AddVertex 100, 0, y, 0, 1, 0, 1, 1, 0, 0, -1
    End If
Next y
Floor.AddVertex 0, -100, 0, 1, 0, 0, 0, 0, 0, 0, 1000000
Floor.AddVertex 0, 100, 0, 1, 0, 0, 0, 0, 0, 0, 1000000
Floor.AddVertex -100, 0, 0, 1, 0, 0, 0, 0, 0, 0, 1003
Floor.AddVertex 100, 0, 0, 1, 0, 0, 0, 0, 0, 0, 1003
Floor.AddVertex 0, 0, -100, 1, 0, 0, 0, 0, 0, 0, -1030000
Floor.AddVertex 0, 0, 100, 1, 0, 0, 0, 0, 0, 0, -1900000


'Floor.AddVertex 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, -30000
'Floor.AddVertex -100, 0, 0, 1, 0, 0, 0, 0, 0, 0, -30000
'Floor.AddVertex 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, -30000
'Floor.AddVertex 0, 0, -100, 1, 0, 0, 0, 0, 0, 0, -30000

Set mesh(1) = Scene.CreateMeshBuilder '网格物体初始化，必加
mesh(1).CreateSphere 0.25   '建立一个半径为1的球
mesh(1).SetTexture GetTex("1") '赋予物体pic贴图
mesh(1).SetMaterial GetMat("1") '赋予物体solid材质
mesh(1).SetLightingMode TV_LIGHTING_NORMAL      '这个最常用的灯光模式
Set mesh(2) = Scene.CreateMeshBuilder '网格物体初始化，必加
mesh(2).CreateSphere 0.25 '建立一个半径为1的球
mesh(2).SetTexture GetTex("2") '赋予物体pic贴图
mesh(2).SetMaterial GetMat("2") '赋予物体solid材质
mesh(2).SetLightingMode TV_LIGHTING_NORMAL    '这个最常用的灯光模式
Set mesh(3) = Scene.CreateMeshBuilder '网格物体初始化，必加
mesh(3).CreateSphere 0.25  '建立一个半径为1的球
mesh(3).SetTexture GetTex("3") '赋予物体pic贴图
mesh(3).SetMaterial GetMat("3") '赋予物体solid材质
mesh(3).SetLightingMode TV_LIGHTING_NORMAL    '这个最常用的灯光模式
'X,Y,Z的位置
Set strx = Scene.CreateMeshBuilder
strx.Create3DText "X", 1, 10, 0
strx.SetPosition 11, 0, 0
Set stry = Scene.CreateMeshBuilder
stry.Create3DText "Y", 1, 10, 0
stry.SetPosition 0, 11, 0
Set strz = Scene.CreateMeshBuilder
strz.Create3DText "Z", 1, 10, 0
strz.SetPosition 0, 0, 11
CameraAngX = -140
CameraAngY = 22
CameraPozX = 2
CameraPozY = 2
CameraPozZ = 2
Camera.SetRotation CameraAngY, CameraAngX, 0
Camera.SetPosition CameraPozX, CameraPozY, CameraPozZ
End Sub
Public Sub reSetRoutine()
Dim i As Integer
For i = 1 To 3
    Set Routine(i) = Nothing
Next i
For i = 1 To 3
    Set Routine(i) = Scene.CreateMeshBuilder '网格物体初始化，必加
    Routine(i).SetMeshFormat CONST_TV_MESHFORMAT.TV_MESHFORMAT_DIFFUSE + CONST_TV_MESHFORMAT.TV_MESHFORMAT_NOLIGHTING
    Routine(i).SetLightingMode (CONST_TV_LIGHTINGMODE.TV_LIGHTING_NONE)
    Routine(i).SetPrimitiveType (CONST_TV_PRIMITIVETYPE.TV_LINELIST)
Next i
End Sub
