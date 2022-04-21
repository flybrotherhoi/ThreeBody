Attribute VB_Name = "calculate"
'定义新类型，body
Public Type body
Posi As TV_3DVECTOR    '位置
oldpos As TV_3DVECTOR
V As TV_3DVECTOR    '速度
A As TV_3DVECTOR    '加速度
M As Single
Ek As Single
Ep As Single
E As Single
End Type

'计算轨迹
Public Sub Count_2Body(b1 As body, b2 As body, g As Single, dt As Single) ''''''''''''''''''''''
Dim DD As Single, F As Single, acc As Single     'acc为加速度
Dim xx As Single, yy As Single, zz As Single
Dim outX As Single, outY As Single, outZ As Single
xx = b1.Posi.x - b2.Posi.x
yy = b1.Posi.y - b2.Posi.y
zz = b1.Posi.z - b2.Posi.z
DD = xx ^ 2 + yy ^ 2 + zz ^ 2 '计算距离

F = g * b1.M * b2.M / DD
DD = Sqr(DD)
acc = F / b1.M                                   '计算加速度
b1.A.x = -acc * xx / DD                          '加速度分解
b1.A.y = -acc * yy / DD
b1.A.z = -acc * zz / DD
outX = b1.Posi.x + b1.V.x * dt + 1 / 2 * b1.A.x * dt ^ 2    '变量储存三个过程量
outY = b1.Posi.y + b1.V.y * dt + 1 / 2 * b1.A.y * dt ^ 2
outZ = b1.Posi.z + b1.V.z * dt + 1 / 2 * b1.A.z * dt ^ 2
b1.V.x = b1.V.x + b1.A.x * dt                               '更新速度
b1.V.y = b1.V.y + b1.A.y * dt
b1.V.z = b1.V.z + b1.A.z * dt
b1.Posi.x = outX                                            '更新位置
b1.Posi.y = outY
b1.Posi.z = outZ

acc = F / b2.M
b2.A.x = acc * xx / DD                          '加速度分解
b2.A.y = acc * yy / DD
b2.A.z = acc * zz / DD
outX = b2.Posi.x + b2.V.x * dt + 1 / 2 * b2.A.x * dt ^ 2    '变量储存三个过程量
outY = b2.Posi.y + b2.V.y * dt + 1 / 2 * b2.A.y * dt ^ 2
outZ = b2.Posi.z + b2.V.z * dt + 1 / 2 * b2.A.z * dt ^ 2
b2.V.x = b2.V.x + b2.A.x * dt                               '更新速度
b2.V.y = b2.V.y + b2.A.y * dt
b2.V.z = b2.V.z + b2.A.z * dt
b2.Posi.x = outX                                            '更新位置
b2.Posi.y = outY
b2.Posi.z = outZ
End Sub
Public Sub Count_3Body(b1 As body, b2 As body, b3 As body, g As Single, dt As Single) '''''''''''''''''''''''
Dim DD12 As Single, F12 As Single
Dim DD23 As Single, F23 As Single
Dim DD13 As Single, F13 As Single
Dim acc1 As Single, acc2 As Single  '加速度
Dim xx12 As Single, yy12 As Single, zz12 As Single
Dim xx23 As Single, yy23 As Single, zz23 As Single
Dim xx13 As Single, yy13 As Single, zz13 As Single
Dim outX As Single, outY As Single, outZ As Single
xx12 = b1.Posi.x - b2.Posi.x
yy12 = b1.Posi.y - b2.Posi.y
zz12 = b1.Posi.z - b2.Posi.z
xx23 = b2.Posi.x - b3.Posi.x
yy23 = b2.Posi.y - b3.Posi.y
zz23 = b2.Posi.z - b3.Posi.z
xx13 = b1.Posi.x - b3.Posi.x
yy13 = b1.Posi.y - b3.Posi.y
zz13 = b1.Posi.z - b3.Posi.z

DD12 = xx12 ^ 2 + yy12 ^ 2 + zz12 ^ 2 '计算距离
DD23 = xx23 ^ 2 + yy23 ^ 2 + zz23 ^ 2
DD13 = xx13 ^ 2 + yy13 ^ 2 + zz13 ^ 2
F12 = g * b1.M * b2.M / DD12
F23 = g * b2.M * b3.M / DD23
F13 = g * b1.M * b3.M / DD13
DD12 = Sqr(DD12)
DD23 = Sqr(DD23)
DD13 = Sqr(DD13)
'b1
acc1 = F12 / b1.M                                   '计算加速度
acc2 = F13 / b1.M
b1.A.x = (-acc1 * xx12 / DD12 - acc2 * xx13 / DD13)                    '加速度分解
b1.A.y = (-acc1 * yy12 / DD12 - acc2 * yy13 / DD13)
b1.A.z = (-acc1 * zz12 / DD12 - acc2 * zz13 / DD13)
outX = b1.Posi.x + b1.V.x * dt + 1 / 2 * b1.A.x * dt ^ 2   '变量储存三个过程量
outY = b1.Posi.y + b1.V.y * dt + 1 / 2 * b1.A.y * dt ^ 2
outZ = b1.Posi.z + b1.V.z * dt + 1 / 2 * b1.A.z * dt ^ 2
b1.V.x = b1.V.x + b1.A.x * dt                               '更新速度
b1.V.y = b1.V.y + b1.A.y * dt
b1.V.z = b1.V.z + b1.A.z * dt
b1.Posi.x = outX                                            '更新位置
b1.Posi.y = outY
b1.Posi.z = outZ
'b2
acc1 = F12 / b2.M                                   '计算加速度
acc2 = F23 / b2.M
b2.A.x = (acc1 * xx12 / DD12 - acc2 * xx23 / DD23)                   '加速度分解
b2.A.y = (acc1 * yy12 / DD12 - acc2 * yy23 / DD23)
b2.A.z = (acc1 * zz12 / DD12 - acc2 * zz23 / DD23)
outX = b2.Posi.x + b2.V.x * dt + 1 / 2 * b2.A.x * dt ^ 2    '变量储存三个过程量
outY = b2.Posi.y + b2.V.y * dt + 1 / 2 * b2.A.y * dt ^ 2
outZ = b2.Posi.z + b2.V.z * dt + 1 / 2 * b2.A.z * dt ^ 2
b2.V.x = b2.V.x + b2.A.x * dt                               '更新速度
b2.V.y = b2.V.y + b2.A.y * dt
b2.V.z = b2.V.z + b2.A.z * dt
b2.Posi.x = outX                                            '更新位置
b2.Posi.y = outY
b2.Posi.z = outZ
'b3
acc1 = F23 / b3.M                                   '计算加速度
acc2 = F13 / b3.M
b3.A.x = (acc1 * xx23 / DD23 + acc2 * xx13 / DD13)                   '加速度分解
b3.A.y = (acc1 * yy23 / DD23 + acc2 * yy13 / DD13)
b3.A.z = (acc1 * zz23 / DD23 + acc2 * zz13 / DD13)
outX = b3.Posi.x + b3.V.x * dt + 1 / 2 * b3.A.x * dt ^ 2    '变量储存三个过程量
outY = b3.Posi.y + b3.V.y * dt + 1 / 2 * b3.A.y * dt ^ 2
outZ = b3.Posi.z + b3.V.z * dt + 1 / 2 * b3.A.z * dt ^ 2
b3.V.x = b3.V.x + b3.A.x * dt                               '更新速度
b3.V.y = b3.V.y + b3.A.y * dt
b3.V.z = b3.V.z + b3.A.z * dt
b3.Posi.x = outX                                            '更新位置
b3.Posi.y = outY
b3.Posi.z = outZ
End Sub
Public Sub countE_3Body(b1 As body, b2 As body, b3 As body, g As Single)
Dim vv As Single
Dim DD12 As Single
Dim DD23 As Single
Dim DD13 As Single
Dim xx12 As Single, yy12 As Single, zz12 As Single
Dim xx23 As Single, yy23 As Single, zz23 As Single
Dim xx13 As Single, yy13 As Single, zz13 As Single
'计算Ek
vv = ((b2.V.x + b3.V.x) / 2 - b1.V.x) ^ 2 + ((b2.V.y + b3.V.y) / 2 - b1.V.y) ^ 2 + ((b2.V.z + b3.V.z) / 2 - b1.V.z) ^ 2
b1.Ek = 1 / 2 * b1.M * vv
vv = ((b1.V.x + b3.V.x) / 2 - b2.V.x) ^ 2 + ((b1.V.y + b3.V.y) / 2 - b2.V.y) ^ 2 + ((b1.V.z + b3.V.z) / 2 - b2.V.z) ^ 2
b2.Ek = 1 / 2 * b2.M * vv
vv = ((b1.V.x + b2.V.x) / 2 - b3.V.x) ^ 2 + ((b1.V.y + b2.V.y) / 2 - b3.V.y) ^ 2 + ((b1.V.z + b2.V.z) / 2 - b3.V.z) ^ 2
b3.Ek = 1 / 2 * b3.M * vv

xx12 = b1.Posi.x - b2.Posi.x
yy12 = b1.Posi.y - b2.Posi.y
zz12 = b1.Posi.z - b2.Posi.z
xx23 = b2.Posi.x - b3.Posi.x
yy23 = b2.Posi.y - b3.Posi.y
zz23 = b2.Posi.z - b3.Posi.z
xx13 = b1.Posi.x - b3.Posi.x
yy13 = b1.Posi.y - b3.Posi.y
zz13 = b1.Posi.z - b3.Posi.z

DD12 = Sqr(xx12 ^ 2 + yy12 ^ 2 + zz12 ^ 2) '计算距离
DD23 = Sqr(xx23 ^ 2 + yy23 ^ 2 + zz23 ^ 2)
DD13 = Sqr(xx13 ^ 2 + yy13 ^ 2 + zz13 ^ 2)
b1.Ep = g * b1.M * b2.M / DD12 + g * b1.M * b3.M / DD13
b2.Ep = g * b1.M * b2.M / DD12 + g * b2.M * b3.M / DD23
b3.Ep = g * b2.M * b3.M / DD23 + g * b1.M * b3.M / DD13
b1.E = b1.Ek - b1.Ep
b2.E = b2.Ek - b2.Ep
b3.E = b3.Ek - b3.Ep
End Sub
Public Sub guiji(b1 As body, b2 As body, b3 As body)
Routine(1).AddVertex b1.oldpos.x / 50, b1.oldpos.y / 50, b1.oldpos.z / 50, 0, 1, 0, 1, 1, 0, 0, 55000
Routine(1).AddVertex b1.Posi.x / 50, b1.Posi.y / 50, b1.Posi.z / 50, 0, 1, 0, 1, 1, 0, 0, 55000
Routine(2).AddVertex b2.oldpos.x / 50, b2.oldpos.y / 50, b2.oldpos.z / 50, 0, 1, 0, 1, 1, 0, 0, vbBlue
Routine(2).AddVertex b2.Posi.x / 50, b2.Posi.y / 50, b2.Posi.z / 50, 0, 1, 0, 1, 1, 0, 0, vbBlue
Routine(3).AddVertex b3.oldpos.x / 50, b3.oldpos.y / 50, b3.oldpos.z / 50, 0, 1, 0, 1, 1, 0, 0, -5500
Routine(3).AddVertex b3.Posi.x / 50, b3.Posi.y / 50, b3.Posi.z / 50, 0, 1, 0, 1, 1, 0, 0, -5500

b1.oldpos = b1.Posi
b2.oldpos = b2.Posi
b3.oldpos = b3.Posi
End Sub
