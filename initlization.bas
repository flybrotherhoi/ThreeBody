Attribute VB_Name = "Initialization"
Public body(1 To 3)  As body    '������������
Public oldbody(1 To 3) As body   '��¼��״̬
Public panduan As Integer
Public NotSaved As Boolean      '�ж��Ƿ��Ѿ�����
Public g As Single          '��������
Public dt As Single         '��������
Public oldG As Single
Public olddt As Single
Public Tv As New TVEngine '����tv3d�������
Public Scene As New TVScene '����tv3d�������
Public TF As New TVTextureFactory '���һ����ͼ��
Public MF As New TVMaterialFactory ''���һ�����ʿ�
Public LE As New TVLightEngine '���һ���ƹ��
Public Atmos  As New TVAtmosphere '��Ӵ���ϵͳ
Public Inp As New TVInputEngine
Public Mx As Long, My As Long, b1 As Boolean, b2 As Boolean, Roll As Long   '���������Ϣ
Public Camera As New TVCamera '����һ����������൱���˵��۾�
Public CameraPozX As Single, CameraPozY As Single, CameraPozZ As Single '�����λ������
Public CameraAngX As Single, CameraAngY As Single '������Ƕ�
Public Floor As TVMesh  '���һ����������
Public Routine(1 To 3) As TVMesh
Public mesh(1 To 3) As TVMesh   '���һ������
Public strx As TVMesh   '������
Public stry As TVMesh
Public strz As TVMesh

  '''''��ʼ����ǩλ��
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
    .Label1(0).Caption = "����:0"
    .Label1(1).Caption = "�ٶ�:0"
    .Label1(2).Caption = "X:0"
    .Label1(3).Caption = "Y:0"
    .Label1(4).Caption = "Z:0"

    .Label2(0).Caption = "����:0"
    .Label2(1).Caption = "�ٶ�:0"
    .Label2(2).Caption = "X:0"
    .Label2(3).Caption = "Y:0"
    .Label2(4).Caption = "Z:0"

    .Label3(0).Caption = "����:0"
    .Label3(1).Caption = "�ٶ�:0"
    .Label3(2).Caption = "X:0"
    .Label3(3).Caption = "Y:0"
    .Label3(4).Caption = "Z:0"
End With

End Sub
Public Sub init_Tv3D()
'------------------------------------------------------------------Tv3d��ʼ��
Tv.SetSearchDirectory App.path & "\Data\Picture" '�趨��ͼ��ȡĿ¼Ϊ��ǰĿ¼
Tv.SetVSync True '��ֱͬ������
Tv.Init3DWindowed Frmmain.Picture1.hwnd    '�ô���ģʽ����tv3d
Inp.Initialize
Tv.SetAngleSystem TV_ANGLE_DEGREE
 
TF.LoadTexture "body1.jpg", "1" '��ȡ��Ϊpic.jpg����ͼ��������Ϊpic
TF.LoadTexture "body2.jpg", "2" '��ȡ��Ϊpic.jpg����ͼ��������Ϊpic
TF.LoadTexture "body3.jpg", "3"
TF.LoadTexture "xing.jpg", "xing"

Atmos.SkyBox_Enable True '������պ�
  Atmos.SkyBox_SetTexture GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing"), GetTex("xing") '�趨��ͼ
Atmos.Fog_SetParameters 1, 200000, 0              '������룬��Զ���룬Ũ��

Scene.SetBackgroundColor 0.5, 0.5, 0.9  '������ɫ


'MF.CreateMaterialQuick 0, 1, 0, 0, "green"
'MF.CreateMaterialQuick 1, 0, 0, 0, "red"
'MF.CreateMaterialQuick 0, 0, 1, 0, "blue"
MF.CreateMaterial "2" '������Ϊsolid�Ĳ���
MF.SetAmbient GetMat("2"), 0, 0, 0, 1        '������
MF.SetDiffuse GetMat("2"), 1, 0, 0, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("2"), 1, 0, 0, 1   '�Է���
MF.SetOpacity GetMat("2"), 1 '��͸����
MF.SetSpecular GetMat("2"), 1, 1, 1, 1  '�߹�ɫ
MF.SetPower GetMat("2"), 60 'ɢ��ǿ��

 MF.CreateMaterial "1" '������ɫ�Ĳ���
MF.SetAmbient GetMat("1"), 0, 0, 1, 1       '������
MF.SetDiffuse GetMat("1"), 0, 1, 1, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("1"), 0, 1, 1, 1   '�Է���
MF.SetOpacity GetMat("1"), 1 '��͸����
MF.SetSpecular GetMat("1"), 1, 1, 1, 1  '�߹�ɫ
MF.SetPower GetMat("1"), 60 'ɢ��ǿ��

  MF.CreateMaterial "3" '������ɫ�Ĳ���
MF.SetAmbient GetMat("3"), 1, 1, 0, 1       '������
MF.SetDiffuse GetMat("3"), 1, 1, 0, 1 '��ɢ�⣬������Ĺ�����ɫ
MF.SetEmissive GetMat("3"), 1, 1, 0, 1  '�Է���
MF.SetOpacity GetMat("3"), 1 '��͸����
MF.SetSpecular GetMat("3"), 1, 1, 1, 1  '�߹�ɫ
MF.SetPower GetMat("3"), 60 'ɢ��ǿ��

 LE.CreateDirectionalLight Vector(1, -1, 1), 1, 1, 1, , 1  '���һ��ƽ�й�
 LE.SetSpecularLighting True  '�߹⿪��
'�����켣

Set Floor = Scene.CreateMeshBuilder '���������ʼ�����ؼ�
Floor.SetMeshFormat CONST_TV_MESHFORMAT.TV_MESHFORMAT_DIFFUSE + CONST_TV_MESHFORMAT.TV_MESHFORMAT_NOLIGHTING
Floor.SetLightingMode (CONST_TV_LIGHTINGMODE.TV_LIGHTING_NONE)
Floor.SetPrimitiveType (CONST_TV_PRIMITIVETYPE.TV_LINELIST)
'��������

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

Set mesh(1) = Scene.CreateMeshBuilder '���������ʼ�����ؼ�
mesh(1).CreateSphere 0.25   '����һ���뾶Ϊ1����
mesh(1).SetTexture GetTex("1") '��������pic��ͼ
mesh(1).SetMaterial GetMat("1") '��������solid����
mesh(1).SetLightingMode TV_LIGHTING_NORMAL      '�����õĵƹ�ģʽ
Set mesh(2) = Scene.CreateMeshBuilder '���������ʼ�����ؼ�
mesh(2).CreateSphere 0.25 '����һ���뾶Ϊ1����
mesh(2).SetTexture GetTex("2") '��������pic��ͼ
mesh(2).SetMaterial GetMat("2") '��������solid����
mesh(2).SetLightingMode TV_LIGHTING_NORMAL    '�����õĵƹ�ģʽ
Set mesh(3) = Scene.CreateMeshBuilder '���������ʼ�����ؼ�
mesh(3).CreateSphere 0.25  '����һ���뾶Ϊ1����
mesh(3).SetTexture GetTex("3") '��������pic��ͼ
mesh(3).SetMaterial GetMat("3") '��������solid����
mesh(3).SetLightingMode TV_LIGHTING_NORMAL    '�����õĵƹ�ģʽ
'X,Y,Z��λ��
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
    Set Routine(i) = Scene.CreateMeshBuilder '���������ʼ�����ؼ�
    Routine(i).SetMeshFormat CONST_TV_MESHFORMAT.TV_MESHFORMAT_DIFFUSE + CONST_TV_MESHFORMAT.TV_MESHFORMAT_NOLIGHTING
    Routine(i).SetLightingMode (CONST_TV_LIGHTINGMODE.TV_LIGHTING_NONE)
    Routine(i).SetPrimitiveType (CONST_TV_PRIMITIVETYPE.TV_LINELIST)
Next i
End Sub
