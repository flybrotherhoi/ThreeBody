Attribute VB_Name = "Module1"
Option Explicit
'�ö�����
Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const SWP_NOMOVE = &H2 '���ƶ�����
Public Const SWP_NOSIZE = &H1 '���ı䴰��ߴ�
Public Const Flag = SWP_NOMOVE Or SWP_NOSIZE
Public Const HWND_TOPMOST = -1 '����������ǰ��
Public Sub sugInform()
frmchangewatching.Label1(0).Caption = "����:" & body(1).M
frmchangewatching.Label1(1).Caption = "����:" & Sqr(body(1).V.x ^ 2 + body(1).V.y ^ 2 + body(1).V.z ^ 2)
frmchangewatching.Label1(2).Caption = "X:" & body(1).Posi.x
frmchangewatching.Label1(3).Caption = "Y:" & body(1).Posi.y
frmchangewatching.Label1(4).Caption = "Z:" & body(1).Posi.z
frmchangewatching.Label2(0).Caption = "����:" & body(2).M
frmchangewatching.Label2(1).Caption = "����:" & Sqr(body(2).V.x ^ 2 + body(2).V.y ^ 2 + body(2).V.z ^ 2)
frmchangewatching.Label2(2).Caption = "X:" & body(2).Posi.x
frmchangewatching.Label2(3).Caption = "Y:" & body(2).Posi.y
frmchangewatching.Label2(4).Caption = "Z:" & body(2).Posi.z
frmchangewatching.Label3(0).Caption = "����:" & body(3).M
frmchangewatching.Label3(1).Caption = "����:" & Sqr(body(3).V.x ^ 2 + body(3).V.y ^ 2 + body(3).V.z ^ 2)
frmchangewatching.Label3(2).Caption = "X:" & body(3).Posi.x
frmchangewatching.Label3(3).Caption = "Y:" & body(3).Posi.y
frmchangewatching.Label3(4).Caption = "Z:" & body(3).Posi.z
End Sub

