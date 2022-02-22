Attribute VB_Name = "modMath"
Option Explicit

Type vec4_t
    X As Single
    Y As Single
    Z As Single
    W As Single
End Type

Type mat4x4_t
    X As vec4_t
    Y As vec4_t
    Z As vec4_t
    W As vec4_t
End Type

Global Const Single_Epsilon As Single = 0.000001!
Global Const PI As Double = 3.14159265358979
Global Const Sqr2 As Double = 2 ^ 0.5
Global Const Sqr2_Half As Double = Sqr2 * 0.5
Global Const FLT_MAX As Single = 3.402823466E+38
Global Const DBL_MAX As Double = 1.79769313486231E+308

Global g_ZeroVec4 As vec4_t
Global g_ZeroMat4x4 As mat4x4_t

#Const MinMax_UseIf = 1
#Const Clamp_UseIf = 1

Function clamp(ByVal V As Single, Optional ByVal min_ As Single = 0, Optional ByVal max_ As Single = 1) As Single
#If Clamp_UseIf Then
    clamp = V
    If clamp < min_ Then
        clamp = min_
    ElseIf clamp > max_ Then
        clamp = max_
    End If
#Else
    clamp = max(min(V, max_), min_)
#End If
End Function

Function clamp_abs(ByVal V As Single, Optional ByVal min_ As Single = 0, Optional ByVal max_ As Single = 1) As Single
If V < 0 Then clamp_abs = clamp(V, -max_, -min_) Else clamp_abs = clamp(V, min_, max_)
End Function

Function lerp(ByVal A As Single, ByVal B As Single, ByVal s As Single) As Single
lerp = A + (B - A) * s
End Function

Function max(ByVal A As Single, ByVal B As Single) As Single
'max = IIf(A > B, A, B)
#If MinMax_UseIf Then
    If A > B Then max = A Else max = B
#Else
    max = (A + B + Abs(A - B)) * 0.5
#End If
End Function

Function min(ByVal A As Single, ByVal B As Single) As Single
'min = IIf(A < B, A, B)
#If MinMax_UseIf Then
    If A < B Then min = A Else min = B
#Else
    min = (A + B - Abs(A - B)) * 0.5
#End If
End Function

Function s_hermite(ByVal s As Single) As Single
s_hermite = s * s * (3 - 2 * s)
End Function

Function lerp_hermite(ByVal A As Single, ByVal B As Single, ByVal s As Single) As Single
s = s * s * (3 - 2 * s)
lerp_hermite = A + (B - A) * s
End Function

'根据输入的数字来构造一个结构体
Function vec4(Optional ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single, Optional ByVal W As Single = 1) As vec4_t
vec4.X = X
vec4.Y = Y
vec4.Z = Z
vec4.W = W
End Function

'根据输入的字符串来构造一个结构体
Function vec4_parse(Optional Arg As String = "0.0, 0.0, 0.0, 1.0") As vec4_t
Dim SArr() As String
SArr = Split(Arg, ",")
If UBound(SArr) >= 0 Then vec4_parse.X = Val(SArr(0))
If UBound(SArr) >= 1 Then vec4_parse.Y = Val(SArr(1))
If UBound(SArr) >= 2 Then vec4_parse.Z = Val(SArr(2))
If UBound(SArr) >= 3 Then vec4_parse.W = Val(SArr(3))
End Function

Function vec4_fromarray(arr() As Single) As vec4_t
With vec4_fromarray
    .X = arr(0)
    .Y = arr(1)
    .Z = arr(2)
    .W = arr(3)
End With
End Function

Function vec4_toarray(V4 As vec4_t) As Single()
Dim ret_arr(3) As Single

ret_arr(0) = V4.X
ret_arr(1) = V4.Y
ret_arr(2) = V4.Z
ret_arr(3) = V4.W

vec4_toarray = ret_arr
End Function

Function vec4_max_comp(V As vec4_t, Optional CompIdOut As Long) As Single
vec4_max_comp = V.X
CompIdOut = 0
If V.Y > vec4_max_comp Then
    vec4_max_comp = V.Y
    CompIdOut = 1
End If
If V.Z > vec4_max_comp Then
    vec4_max_comp = V.Z
    CompIdOut = 2
End If
If V.W > vec4_max_comp Then
    vec4_max_comp = V.W
    CompIdOut = 3
End If
End Function

Function vec4_max_xyz(V As vec4_t, Optional CompIdOut As Long) As Single
vec4_max_xyz = V.X
CompIdOut = 0
If V.Y > vec4_max_xyz Then
    vec4_max_xyz = V.Y
    CompIdOut = 1
End If
If V.Z > vec4_max_xyz Then
    vec4_max_xyz = V.Z
    CompIdOut = 2
End If
End Function

Function vec4_min_comp(V As vec4_t, Optional CompIdOut As Long) As Single
vec4_min_comp = V.X
CompIdOut = 0
If V.Y < vec4_min_comp Then
    vec4_min_comp = V.Y
    CompIdOut = 1
End If
If V.Z < vec4_min_comp Then
    vec4_min_comp = V.Z
    CompIdOut = 2
End If
If V.W < vec4_min_comp Then
    vec4_min_comp = V.W
    CompIdOut = 3
End If
End Function

Function vec4_min_xyz(V As vec4_t, Optional CompIdOut As Long) As Single
vec4_min_xyz = V.X
CompIdOut = 0
If V.Y < vec4_min_xyz Then
    vec4_min_xyz = V.Y
    CompIdOut = 1
End If
If V.Z < vec4_min_xyz Then
    vec4_min_xyz = V.Z
    CompIdOut = 2
End If
End Function

Function vec4_sel_comp(ByVal CompId As Long) As vec4_t
Select Case CompId
Case 0
    vec4_sel_comp.X = 1
Case 1
    vec4_sel_comp.Y = 1
Case 2
    vec4_sel_comp.Z = 1
Case 3
    vec4_sel_comp.W = 1
End Select
End Function

Function vec4_abs(V As vec4_t) As vec4_t
vec4_abs = vec4(Abs(V.X), Abs(V.Y), Abs(V.Z), Abs(V.W))
End Function

Function vec4_sgn(V As vec4_t) As vec4_t
vec4_sgn = vec4(Sgn(V.X), Sgn(V.Y), Sgn(V.Z), Sgn(V.W))
End Function

Function vec4_pow(V As vec4_t, ByVal n As Single) As vec4_t
vec4_pow = vec4(V.X ^ n, V.Y ^ n, V.Z ^ n, V.W ^ n)
End Function

Function vec4_max(A As vec4_t, B As vec4_t) As vec4_t
vec4_max.X = max(A.X, B.X)
vec4_max.Y = max(A.Y, B.Y)
vec4_max.Z = max(A.Z, B.Z)
vec4_max.W = max(A.W, B.W)
End Function

Function vec4_lerp(A As vec4_t, B As vec4_t, ByVal s As Single) As vec4_t
vec4_lerp.X = A.X + (B.X - A.X) * s
vec4_lerp.Y = A.Y + (B.Y - A.Y) * s
vec4_lerp.Z = A.Z + (B.Z - A.Z) * s
vec4_lerp.W = A.W + (B.W - A.W) * s
End Function

Sub vec4_print(V4 As vec4_t)
Debug.Print Format$(V4.X, "0.00000"), Format$(V4.Y, "0.00000"), Format$(V4.Z, "0.00000"), Format$(V4.W, "0.00000")
End Sub

'向量相加
Function vec4_add(L As vec4_t, R As vec4_t) As vec4_t
vec4_add.X = L.X + R.X
vec4_add.Y = L.Y + R.Y
vec4_add.Z = L.Z + R.Z
vec4_add.W = L.W + R.W
End Function

'向量相减
Function vec4_sub(L As vec4_t, R As vec4_t) As vec4_t
vec4_sub.X = L.X - R.X
vec4_sub.Y = L.Y - R.Y
vec4_sub.Z = L.Z - R.Z
vec4_sub.W = L.W - R.W
End Function

'向量分别乘
Function vec4_mul(L As vec4_t, R As vec4_t) As vec4_t
vec4_mul.X = L.X * R.X
vec4_mul.Y = L.Y * R.Y
vec4_mul.Z = L.Z * R.Z
vec4_mul.W = L.W * R.W
End Function

'向量分别除
Function vec4_div(L As vec4_t, R As vec4_t) As vec4_t
'On Local Error Resume Next
If Abs(R.X) >= Single_Epsilon Then vec4_div.X = L.X / R.X Else vec4_div.X = FLT_MAX * Sgn(L.X) * Sgn(R.X)
If Abs(R.Y) >= Single_Epsilon Then vec4_div.Y = L.Y / R.Y Else vec4_div.Y = FLT_MAX * Sgn(L.Y) * Sgn(R.Y)
If Abs(R.Z) >= Single_Epsilon Then vec4_div.Z = L.Z / R.Z Else vec4_div.Z = FLT_MAX * Sgn(L.Z) * Sgn(R.Z)
If Abs(R.W) >= Single_Epsilon Then vec4_div.W = L.W / R.W Else vec4_div.W = FLT_MAX * Sgn(L.W) * Sgn(R.W)
'Err.Clear
End Function

'向量缩放
Function vec4_scale(V As vec4_t, ByVal s As Single) As vec4_t
vec4_scale.X = V.X * s
vec4_scale.Y = V.Y * s
vec4_scale.Z = V.Z * s
vec4_scale.W = V.W * s
End Function

'向量点乘
Function vec4_dot(L As vec4_t, R As vec4_t) As Single
vec4_dot = L.X * R.X + L.Y * R.Y + L.Z * R.Z + L.W * R.W
End Function

'向量求模
Function vec4_length(V As vec4_t) As Single
vec4_length = Sqr(vec4_dot(V, V))
End Function

'向量单位化
Function vec4_normalize(V As vec4_t) As vec4_t
Dim lov As Single
lov = vec4_length(V)
If lov >= Single_Epsilon Then vec4_normalize = vec4_scale(V, 1 / lov)
End Function

'把向量当成三维向量，然后算叉乘
Function vec4_cross3(L As vec4_t, R As vec4_t) As vec4_t
vec4_cross3.X = L.Y * R.Z - L.Z * R.Y
vec4_cross3.Y = L.Z * R.X - L.X * R.Z
vec4_cross3.Z = L.X * R.Y - L.Y * R.X
End Function

'向量镜像
Function vec4_reflect(I As vec4_t, n As vec4_t) As vec4_t
vec4_reflect = vec4_sub(I, vec4_scale(n, vec4_dot(I, n) * 2))
End Function

'向量乘矩阵，得到经过变换后的向量——矩阵代表了坐标系
Function vec4_mult_matrix(V As vec4_t, Matrix As mat4x4_t) As vec4_t
vec4_mult_matrix.X = V.X * Matrix.X.X + V.Y * Matrix.Y.X + V.Z * Matrix.Z.X + V.W * Matrix.W.X
vec4_mult_matrix.Y = V.X * Matrix.X.Y + V.Y * Matrix.Y.Y + V.Z * Matrix.Z.Y + V.W * Matrix.W.Y
vec4_mult_matrix.Z = V.X * Matrix.X.Z + V.Y * Matrix.Y.Z + V.Z * Matrix.Z.Z + V.W * Matrix.W.Z
vec4_mult_matrix.W = V.X * Matrix.X.W + V.Y * Matrix.Y.W + V.Z * Matrix.Z.W + V.W * Matrix.W.W
End Function

'将矩阵转置了再乘。如果只是一个旋转矩阵的话，相当于以相反的角度进行旋转
Function vec4_mult_matrix_transpose(V As vec4_t, Matrix As mat4x4_t) As vec4_t
vec4_mult_matrix_transpose.X = vec4_dot(V, Matrix.X)
vec4_mult_matrix_transpose.Y = vec4_dot(V, Matrix.Y)
vec4_mult_matrix_transpose.Z = vec4_dot(V, Matrix.Z)
vec4_mult_matrix_transpose.W = vec4_dot(V, Matrix.W)
End Function

Function vec4_to_rgb(V As vec4_t) As Long
Dim R As Single, G As Single, B As Single
R = V.X * 255
G = V.Y * 255
B = V.Z * 255
If R < 0 Then R = 0
'If R > 255 Then R = 255
If G < 0 Then G = 0
'If G > 255 Then G = 255
If B < 0 Then B = 0
'If B > 255 Then B = 255
vec4_to_rgb = RGB(R, G, B)
End Function

Function vec4_from_rgb(RGB As Long) As vec4_t
vec4_from_rgb = vec4((RGB And &HFF&) / 255, ((RGB And &HFF00&) \ &H100&) / 255, ((RGB And &HFF0000) \ &H10000) / 255, 0)
End Function

Function vec4_from_rgba(RGBA As Long) As vec4_t
vec4_from_rgba = vec4((RGBA And &HFF&) / 255, ((RGBA And &HFF00&) \ &H100&) / 255, ((RGBA And &HFF0000) \ &H10000) / 255, (((RGBA And &HFF000000) \ &H1000000) And &HFF&) / 255)
End Function

Function quat_from_axis_angle(Axis As vec4_t, ByVal Angle As Single) As vec4_t
Dim Half_Ang As Single
Half_Ang = Angle * 0.5

Dim Sin_HA As Single
Sin_HA = Sin(Half_Ang)
quat_from_axis_angle.X = Axis.X * Sin_HA
quat_from_axis_angle.Y = Axis.Y * Sin_HA
quat_from_axis_angle.Z = Axis.Z * Sin_HA
quat_from_axis_angle.W = Cos(Half_Ang)
End Function

Function quat_mult(q1 As vec4_t, q2 As vec4_t) As vec4_t
quat_mult.X = (q1.W * q2.X) + (q1.X * q2.W) + (q1.Y * q2.Z) - (q1.Z * q2.Y)
quat_mult.Y = (q1.W * q2.Y) - (q1.X * q2.Z) + (q1.Y * q2.W) + (q1.Z * q2.X)
quat_mult.Z = (q1.W * q2.Z) + (q1.X * q2.Y) - (q1.Y * q2.X) + (q1.Z * q2.W)
quat_mult.W = (q1.W * q2.W) - (q1.X * q2.X) - (q1.Y * q2.Y) - (q1.Z * q2.Z)
End Function

'假的。
'Function quat_euler(ByVal yaw As Single, Optional ByVal pitch As Single, Optional ByVal roll As Single) As vec4_t
'Dim hy As Single, hp As Single, hr As Single
'hy = yaw * 0.5
'hp = pitch * 0.5
'hr = roll * 0.5
'
'Dim sy As Single, cy As Single, sp As Single, cp As Single, sr As Single, cr As Single
'sy = Sin(hy)
'cy = Cos(hy)
'sp = Sin(hp)
'cp = Cos(hp)
'sr = Sin(hr)
'cr = Cos(hr)
'quat_euler.X = sr * cp * cy - cr * sp * sy
'quat_euler.Y = cr * sp * cy + sr * cp * sy
'quat_euler.Z = cr * cp * sy - sr * sp * cy
'quat_euler.W = cr * cp * cy + sr * sp * sy
'End Function

Function vec4_rot_quat(V As vec4_t, q As vec4_t) As vec4_t
vec4_rot_quat = quat_mult(quat_mult(q, vec4(V.X, V.Y, V.Z, 0)), vec4(-q.X, -q.Y, -q.Z, q.W))
End Function

Function quat_add_vec(q As vec4_t, V As vec4_t, ByVal scaling As Single) As vec4_t
quat_add_vec = vec4_add(vec4_scale(quat_mult(vec4(V.X * scaling, V.Y * scaling, V.Z * scaling, 0), q), 0.5), q)
End Function

Function mat_from_quat(q As vec4_t) As mat4x4_t
mat_from_quat.X = vec4(1 - 2 * q.Y * q.Y - 2 * q.Z * q.Z, 0 + 2 * q.X * q.Y + 2 * q.Z * q.W, 0 + 2 * q.X * q.Z - 2 * q.Y * q.W, 0)
mat_from_quat.Y = vec4(0 + 2 * q.X * q.Y - 2 * q.Z * q.W, 1 - 2 * q.X * q.X - 2 * q.Z * q.Z, 0 + 2 * q.Y * q.Z + 2 * q.X * q.W, 0)
mat_from_quat.Z = vec4(0 + 2 * q.X * q.Z + 2 * q.Y * q.W, 0 + 2 * q.Y * q.Z - 2 * q.X * q.W, 1 - 2 * q.X * q.X - 2 * q.Y * q.Y, 0)
mat_from_quat.W = vec4(0, 0, 0, 1)
End Function

Function mat_from_quat_transpose(q As vec4_t) As mat4x4_t
mat_from_quat_transpose.X = vec4(1 - 2 * q.Y * q.Y - 2 * q.Z * q.Z, 0 + 2 * q.X * q.Y - 2 * q.Z * q.W, 0 + 2 * q.X * q.Z + 2 * q.Y * q.W, 0)
mat_from_quat_transpose.Y = vec4(0 + 2 * q.X * q.Y + 2 * q.Z * q.W, 1 - 2 * q.X * q.X - 2 * q.Z * q.Z, 0 + 2 * q.Y * q.Z - 2 * q.X * q.W, 0)
mat_from_quat_transpose.Z = vec4(0 + 2 * q.X * q.Z - 2 * q.Y * q.W, 0 + 2 * q.Y * q.Z + 2 * q.X * q.W, 1 - 2 * q.X * q.X - 2 * q.Y * q.Y, 0)
mat_from_quat_transpose.W = vec4(0, 0, 0, 1)
End Function

Sub mat_print(m As mat4x4_t, Optional Title As String)
If Len(Title) Then Debug.Print Title
vec4_print m.X
vec4_print m.Y
vec4_print m.Z
vec4_print m.W
End Sub

Function mat_fromarray(arr() As Single) As mat4x4_t
With mat_fromarray
    .X.X = arr(0)
    .X.Y = arr(1)
    .X.Z = arr(2)
    .X.W = arr(3)
    
    .Y.X = arr(4)
    .Y.Y = arr(5)
    .Y.Z = arr(6)
    .Y.W = arr(7)
    
    .Z.X = arr(8)
    .Z.Y = arr(9)
    .Z.Z = arr(10)
    .Z.W = arr(11)
    
    .W.X = arr(12)
    .W.Y = arr(13)
    .W.Z = arr(14)
    .W.W = arr(15)
End With
End Function

Function mat_fromarray_transpose(arr() As Single) As mat4x4_t
With mat_fromarray_transpose
    .X.X = arr(0)
    .Y.X = arr(1)
    .Z.X = arr(2)
    .W.X = arr(3)
      
    .X.Y = arr(4)
    .Y.Y = arr(5)
    .Z.Y = arr(6)
    .W.Y = arr(7)
      
    .X.Z = arr(8)
    .Y.Z = arr(9)
    .Z.Z = arr(10)
    .W.Z = arr(11)
      
    .X.W = arr(12)
    .Y.W = arr(13)
    .Z.W = arr(14)
    .W.W = arr(15)
End With
End Function

Function mat_toarray(m As mat4x4_t) As Single()
Dim ret_arr(15) As Single
With m
    ret_arr(0) = .X.X
    ret_arr(1) = .X.Y
    ret_arr(2) = .X.Z
    ret_arr(3) = .X.W
                     
    ret_arr(4) = .Y.X
    ret_arr(5) = .Y.Y
    ret_arr(6) = .Y.Z
    ret_arr(7) = .Y.W
                     
    ret_arr(8) = .Z.X
    ret_arr(9) = .Z.Y
    ret_arr(10) = .Z.Z
    ret_arr(11) = .Z.W
                     
    ret_arr(12) = .W.X
    ret_arr(13) = .W.Y
    ret_arr(14) = .W.Z
    ret_arr(15) = .W.W
End With
mat_toarray = ret_arr
End Function

Function mat_toarray_transpose(m As mat4x4_t) As Single()
Dim ret_arr(15) As Single
With m
    ret_arr(0) = .X.X
    ret_arr(1) = .Y.X
    ret_arr(2) = .Z.X
    ret_arr(3) = .W.X
      
    ret_arr(4) = .X.Y
    ret_arr(5) = .Y.Y
    ret_arr(6) = .Z.Y
    ret_arr(7) = .W.Y
      
    ret_arr(8) = .X.Z
    ret_arr(9) = .Y.Z
    ret_arr(10) = .Z.Z
    ret_arr(11) = .W.Z
      
    ret_arr(12) = .X.W
    ret_arr(13) = .Y.W
    ret_arr(14) = .Z.W
    ret_arr(15) = .W.W
End With
mat_toarray_transpose = ret_arr
End Function

'取得单位矩阵
Function mat_identity() As mat4x4_t
mat_identity.X.X = 1
mat_identity.Y.Y = 1
mat_identity.Z.Z = 1
mat_identity.W.W = 1
End Function

Function mat_scaling(Optional ByVal X As Single = 1, Optional ByVal Y As Single = 1, Optional ByVal Z As Single = 1) As mat4x4_t
mat_scaling.X.X = X
mat_scaling.Y.Y = Y
mat_scaling.Z.Z = Z
mat_scaling.W.W = 1
End Function

Function mat_translate(Optional ByVal X As Single, Optional ByVal Y As Single, Optional ByVal Z As Single) As mat4x4_t
mat_translate.X.X = 1
mat_translate.Y.Y = 1
mat_translate.Z.Z = 1
mat_translate.W.X = X
mat_translate.W.Y = Y
mat_translate.W.Z = Z
mat_translate.W.W = 1
End Function

'旋转矩阵
'mat_rot_x 绕X轴旋转
'mat_rot_y 绕Y轴旋转
'mat_rot_z 绕Z轴旋转
'对于左手坐标系：X轴朝右，Y轴朝上，Z轴朝前
'对于右手坐标系：X轴朝右，Y轴朝上，Z轴朝后
'当你视线顺着旋转轴看过去的时候，这三个函数生成的矩阵，旋转方向，根据坐标系有以下规则：
'左手坐标系时：顺时针旋转
'右手坐标系时：逆时针旋转
Function mat_rot_x(ByVal Angle As Single) As mat4x4_t
Dim SA As Single
Dim CA As Single
SA = Sin(Angle)
CA = Cos(Angle)

mat_rot_x.X.X = 1

mat_rot_x.Y.Y = CA
mat_rot_x.Y.Z = SA

mat_rot_x.Z.Y = -SA
mat_rot_x.Z.Z = CA

mat_rot_x.W.W = 1
End Function

Function mat_rot_y(ByVal Angle As Single) As mat4x4_t
Dim SA As Single
Dim CA As Single
SA = Sin(Angle)
CA = Cos(Angle)

mat_rot_y.X.X = CA
mat_rot_y.X.Z = -SA

mat_rot_y.Y.Y = 1

mat_rot_y.Z.X = SA
mat_rot_y.Z.Z = CA

mat_rot_y.W.W = 1
End Function

Function mat_rot_z(ByVal Angle As Single) As mat4x4_t
Dim SA As Single
Dim CA As Single
SA = Sin(Angle)
CA = Cos(Angle)

mat_rot_z.X.X = CA
mat_rot_z.X.Y = SA

mat_rot_z.Y.X = -SA
mat_rot_z.Y.Y = CA

mat_rot_z.Z.Z = 1

mat_rot_z.W.W = 1
End Function

'按照指定轴旋转
Function mat_rot_axis(Axis As vec4_t, ByVal Angle As Single) As mat4x4_t
Dim SA As Single
Dim CA As Single
SA = Sin(Angle)
CA = Cos(Angle)

Dim V As vec4_t
V = vec4_normalize(Axis)

mat_rot_axis.X.X = (1 - CA) * V.X * V.X + CA
mat_rot_axis.X.Y = (1 - CA) * V.Y * V.X + SA * V.Z
mat_rot_axis.X.Z = (1 - CA) * V.Z * V.X - SA * V.Y

mat_rot_axis.Y.X = (1 - CA) * V.X * V.Y - SA * V.Z
mat_rot_axis.Y.Y = (1 - CA) * V.Y * V.Y + CA
mat_rot_axis.Y.Z = (1 - CA) * V.Z * V.Y + SA * V.X

mat_rot_axis.Z.X = (1 - CA) * V.X * V.Z + SA * V.Y
mat_rot_axis.Z.Y = (1 - CA) * V.Y * V.Z - SA * V.X
mat_rot_axis.Z.Z = (1 - CA) * V.Z * V.Z + CA

mat_rot_axis.W.W = 1
End Function

'欧拉角旋转
Function mat_rot_euler(ByVal yaw As Single, Optional ByVal pitch As Single, Optional ByVal roll As Single) As mat4x4_t
's表示Sin，c表示Cos，r p y是roll pitch yaw的缩写
Dim sr As Single, cr As Single
Dim sp As Single, cp As Single
Dim sy As Single, cy As Single
sy = Sin(yaw)
cy = Cos(yaw)
sp = Sin(pitch)
cp = Cos(pitch)
sr = Sin(roll)
cr = Cos(roll)

'某些组合的乘积
Dim srcp As Single, srsp As Single
Dim crcp As Single, crsp As Single
srcp = sr * cp
srsp = sr * sp
crcp = cr * cp
crsp = cr * sp

'roll:
' c, s, 0
'-s, c, 0
' 0, 0, 1
'
'pitch:
' 1, 0, 0
' 0, c, s
' 0,-s, c
'
'yaw:
' c, 0,-s
' 0, 1, 0
' s, 0, c
'
'rp
' cr,  sr cp,  sr sp
'-sr,  cr cp,  cr sp
' 0,     -sp,     cp
'
'rpy
' cr cy + sr sp sy,  sr cp, -sy cr + sr sp cy
'-sr cy + cr sp sy,  cr cp,  sy sr + cr sp cy
'            sy cp,    -sp,             cp cy

'欧拉角矩阵
mat_rot_euler.X.X = (cr * cy + srsp * sy)
mat_rot_euler.X.Y = (srcp)
mat_rot_euler.X.Z = (-sy * cr + srsp * cy)

mat_rot_euler.Y.X = (-sr * cy + crsp * sy)
mat_rot_euler.Y.Y = (crcp)
mat_rot_euler.Y.Z = (sy * sr + crsp * cy)

mat_rot_euler.Z.X = (sy * cp)
mat_rot_euler.Z.Y = (-sp)
mat_rot_euler.Z.Z = (cp * cy)

mat_rot_euler.W.W = 1

'mat_rot_euler = mat_mult(mat_mult(mat_rot_z(roll), mat_rot_x(pitch)), mat_rot_y(yaw))
End Function

'观察矩阵
Function mat_view(eye_pos As vec4_t, Rot_Euler As mat4x4_t) As mat4x4_t
mat_view = mat_transpose(Rot_Euler)
mat_view.W = vec4_scale(vec4_mult_matrix(eye_pos, mat_view), -1)
mat_view.W.W = 1
End Function

'正交投影
Function mat_ortho(Optional ByVal Left As Single = -1, Optional ByVal Right As Single = 1, Optional ByVal Bottom As Single = -1, Optional ByVal Top As Single = 1, Optional ByVal Near As Single = 1, Optional ByVal Far As Single = 1000) As mat4x4_t
Dim Width As Single, Height As Single, Depth As Single

Width = Right - Left
Height = Top - Bottom
Depth = Far - Near

mat_ortho.X.X = 2 / Width
mat_ortho.Y.Y = 2 / Height
mat_ortho.Z.Z = -2 / Depth

mat_ortho.W.X = -(Right + Left) / Width
mat_ortho.W.Y = -(Top + Bottom) / Height
mat_ortho.W.Z = -(Far + Near) / Depth
mat_ortho.W.W = 1
End Function

'透视投影
Function mat_persp(Optional ByVal Left As Single = -1, Optional ByVal Right As Single = 1, Optional ByVal Bottom As Single = -1, Optional ByVal Top As Single = 1, Optional ByVal Near As Single = 1, Optional ByVal Far As Single = 1000) As mat4x4_t
Dim Width As Single, Height As Single, Depth As Single

Width = Right - Left
Height = Top - Bottom
Depth = Far - Near

mat_persp.X.X = 2 * Near / Width
mat_persp.Y.Y = 2 * Near / Height
mat_persp.Z.Z = -Far / Depth
mat_persp.Z.W = -1
mat_persp.W.Z = -2 * Near * Far / Depth
End Function

'矩阵转置
Function mat_transpose(m As mat4x4_t) As mat4x4_t
mat_transpose.X.X = m.X.X
mat_transpose.Y.X = m.X.Y
mat_transpose.Z.X = m.X.Z
mat_transpose.W.X = m.X.W

mat_transpose.X.Y = m.Y.X
mat_transpose.Y.Y = m.Y.Y
mat_transpose.Z.Y = m.Y.Z
mat_transpose.W.Y = m.Y.W

mat_transpose.X.Z = m.Z.X
mat_transpose.Y.Z = m.Z.Y
mat_transpose.Z.Z = m.Z.Z
mat_transpose.W.Z = m.Z.W

mat_transpose.X.W = m.W.X
mat_transpose.Y.W = m.W.Y
mat_transpose.Z.W = m.W.Z
mat_transpose.W.W = m.W.W
End Function

'矩阵加标量
Function mat_add_scalar(m As mat4x4_t, ByVal s As Single) As mat4x4_t
Dim vs As vec4_t
vs = vec4(s, s, s, s)
mat_add_scalar.X = vec4_add(m.X, vs)
mat_add_scalar.Y = vec4_add(m.Y, vs)
mat_add_scalar.Z = vec4_add(m.Z, vs)
mat_add_scalar.W = vec4_add(m.W, vs)
End Function

'矩阵减标量
Function mat_sub_scalar(m As mat4x4_t, ByVal s As Single) As mat4x4_t
Dim vs As vec4_t
vs = vec4(s, s, s, s)
mat_sub_scalar.X = vec4_sub(m.X, vs)
mat_sub_scalar.Y = vec4_sub(m.Y, vs)
mat_sub_scalar.Z = vec4_sub(m.Z, vs)
mat_sub_scalar.W = vec4_sub(m.W, vs)
End Function

'矩阵乘标量
Function mat_mult_scalar(m As mat4x4_t, ByVal s As Single) As mat4x4_t
mat_mult_scalar.X = vec4_scale(m.X, s)
mat_mult_scalar.Y = vec4_scale(m.Y, s)
mat_mult_scalar.Z = vec4_scale(m.Z, s)
mat_mult_scalar.W = vec4_scale(m.W, s)
End Function

'矩阵加矩阵
Function mat_add(L As mat4x4_t, R As mat4x4_t) As mat4x4_t
mat_add.X = vec4_add(L.X, R.X)
mat_add.Y = vec4_add(L.Y, R.Y)
mat_add.Z = vec4_add(L.Z, R.Z)
mat_add.W = vec4_add(L.W, R.W)
End Function

'矩阵减矩阵
Function mat_sub(L As mat4x4_t, R As mat4x4_t) As mat4x4_t
mat_sub.X = vec4_sub(L.X, R.X)
mat_sub.Y = vec4_sub(L.Y, R.Y)
mat_sub.Z = vec4_sub(L.Z, R.Z)
mat_sub.W = vec4_sub(L.W, R.W)
End Function

'矩阵乘矩阵
Function mat_mult(L As mat4x4_t, R As mat4x4_t) As mat4x4_t
mat_mult.X = vec4_mult_matrix(L.X, R)
mat_mult.Y = vec4_mult_matrix(L.Y, R)
mat_mult.Z = vec4_mult_matrix(L.Z, R)
mat_mult.W = vec4_mult_matrix(L.W, R)
End Function

'矩阵乘转置矩阵
Function mat_mult_transpose(L As mat4x4_t, R As mat4x4_t) As mat4x4_t
mat_mult_transpose.X = vec4_mult_matrix_transpose(L.X, R)
mat_mult_transpose.Y = vec4_mult_matrix_transpose(L.Y, R)
mat_mult_transpose.Z = vec4_mult_matrix_transpose(L.Z, R)
mat_mult_transpose.W = vec4_mult_matrix_transpose(L.W, R)
End Function

Function mat_copy3x3(m4x4 As mat4x4_t) As mat4x4_t
mat_copy3x3.X.X = m4x4.X.X
mat_copy3x3.X.Y = m4x4.X.Y
mat_copy3x3.X.Z = m4x4.X.Z

mat_copy3x3.Y.X = m4x4.Y.X
mat_copy3x3.Y.Y = m4x4.Y.Y
mat_copy3x3.Y.Z = m4x4.Y.Z

mat_copy3x3.Z.X = m4x4.Z.X
mat_copy3x3.Z.Y = m4x4.Z.Y
mat_copy3x3.Z.Z = m4x4.Z.Z

mat_copy3x3.W.W = 1
End Function

'3x3矩阵求逆
Function mat_inverse_3x3(Out_mat As mat4x4_t, Out_Determinant As Single, In_Matrix As mat4x4_t) As Boolean
Dim m As mat4x4_t
m = In_Matrix 'mat_transpose(In_Matrix)

Dim t4 As Single
Dim t6 As Single
Dim t8 As Single
Dim t10 As Single
Dim t12 As Single
Dim t14 As Single
Dim t16 As Single
Dim t17 As Single

t4 = m.X.X * m.Y.Y
t6 = m.X.X * m.Z.Y
t8 = m.Y.X * m.X.Y
t10 = m.Z.X * m.X.Y
t12 = m.Y.X * m.X.Z
t14 = m.Z.X * m.X.Z

' Calculate the determinant
t16 = (t4 * m.Z.Z - t6 * m.Y.Z - t8 * m.Z.Z + t10 * m.Y.Z + t12 * m.Z.Y - t14 * m.Y.Y)

' Make sure the determinant is non-zero.
If t16 < Single_Epsilon Then Exit Function
t17 = 1 / t16

Out_mat.X.X = (m.Y.Y * m.Z.Z - m.Z.Y * m.Y.Z) * t17
Out_mat.Y.X = -(m.Y.X * m.Z.Z - m.Z.X * m.Y.Z) * t17
Out_mat.Z.X = (m.Y.X * m.Z.Y - m.Z.X * m.Y.Y) * t17
Out_mat.X.Y = -(m.X.Y * m.Z.Z - m.Z.Y * m.X.Z) * t17
Out_mat.Y.Y = (m.X.X * m.Z.Z - t14) * t17
Out_mat.Z.Y = -(t6 - t10) * t17
Out_mat.X.Z = (m.X.Y * m.Y.Z - m.Y.Y * m.X.Z) * t17
Out_mat.Y.Z = -(m.X.X * m.Y.Z - t12) * t17
Out_mat.Z.Z = (t4 - t8) * t17
Out_mat.W.W = 1
Out_Determinant = t16
mat_inverse_3x3 = True
End Function

'3x3矩阵——用于向量叉乘
Function mat_skew_symmetric_3x3(V As vec4_t) As mat4x4_t
mat_skew_symmetric_3x3.X = vec4(0, -V.Z, V.Y, 0)
mat_skew_symmetric_3x3.Y = vec4(V.Z, 0, -V.X, 0)
mat_skew_symmetric_3x3.Z = vec4(-V.Y, V.X, 0, 0)
mat_skew_symmetric_3x3.W.W = 1
End Function

'矩阵求逆
Function mat_inverse(Out_mat As mat4x4_t, Out_Determinant As Single, In_Matrix As mat4x4_t) As Boolean
Dim m As mat4x4_t
m = In_Matrix 'mat_transpose(In_Matrix)

Dim m1234 As Single
Dim m4321 As Single
Dim m2143 As Single
Dim m3412 As Single
Dim m1342 As Single
Dim m1423 As Single
Dim m4132 As Single
Dim m4213 As Single
Dim m3241 As Single
Dim m2431 As Single
Dim m2314 As Single
Dim m3124 As Single

Dim m2134 As Single
Dim m3421 As Single
Dim m1243 As Single
Dim m4312 As Single
Dim m3214 As Single
Dim m2341 As Single
Dim m1432 As Single
Dim m4123 As Single
Dim m2413 As Single
Dim m3142 As Single
Dim m1324 As Single
Dim m4231 As Single

m1234 = m.X.X * m.Y.Y * m.Z.Z * m.W.W
m4321 = m.X.W * m.Y.Z * m.Z.Y * m.W.X
m2143 = m.X.Y * m.Y.X * m.Z.W * m.W.Z
m3412 = m.X.Z * m.Y.W * m.Z.X * m.W.Y
m1342 = m.X.X * m.Y.Z * m.Z.W * m.W.Y
m1423 = m.X.X * m.Y.W * m.Z.Y * m.W.Z
m4132 = m.X.W * m.Y.X * m.Z.Z * m.W.Y
m4213 = m.X.W * m.Y.Y * m.Z.X * m.W.Z
m3241 = m.X.Z * m.Y.Y * m.Z.W * m.W.X
m2431 = m.X.Y * m.Y.W * m.Z.Z * m.W.X
m2314 = m.X.Y * m.Y.Z * m.Z.X * m.W.W
m3124 = m.X.Z * m.Y.X * m.Z.Y * m.W.W

m2134 = m.X.Y * m.Y.X * m.Z.Z * m.W.W
m3421 = m.X.Z * m.Y.W * m.Z.Y * m.W.X
m1243 = m.X.X * m.Y.Y * m.Z.W * m.W.Z
m4312 = m.X.W * m.Y.Z * m.Z.X * m.W.Y
m3214 = m.X.Z * m.Y.Y * m.Z.X * m.W.W
m2341 = m.X.Y * m.Y.Z * m.Z.W * m.W.X
m1432 = m.X.X * m.Y.W * m.Z.Z * m.W.Y
m4123 = m.X.W * m.Y.X * m.Z.Y * m.W.Z
m2413 = m.X.Y * m.Y.W * m.Z.X * m.W.Z
m3142 = m.X.Z * m.Y.X * m.Z.W * m.W.Y
m1324 = m.X.X * m.Y.Z * m.Z.Y * m.W.W
m4231 = m.X.W * m.Y.Y * m.Z.Z * m.W.X

Out_Determinant = 0 _
    + m1234 + m4321 + m2143 + m3412 _
    + m1342 + m1423 + m4132 + m4213 _
    + m3241 + m2431 + m2314 + m3124 _
    - m2134 - m3421 - m1243 - m4312 _
    - m3214 - m2341 - m1432 - m4123 _
    - m2413 - m3142 - m1324 - m4231

If Abs(Out_Determinant) >= Single_Epsilon Then
    Out_mat.X.X = (m1234 + m1342 + m1423 - m1243 - m1432 - m1324) / m.X.X
    Out_mat.X.Y = (m2143 + m3124 + m4132 - m2134 - m4123 - m3142) / m.Y.X
    Out_mat.X.Z = (m3412 + m2314 + m4213 - m4312 - m3214 - m2413) / m.Z.X
    Out_mat.X.W = (m4321 + m2431 + m3241 - m3421 - m2341 - m4231) / m.W.X
    
    Out_mat.Y.X = (m2143 + m2314 + m2431 - m2134 - m2341 - m2413) / m.X.Y
    Out_mat.Y.Y = (m1234 + m3241 + m4213 - m1243 - m3214 - m4231) / m.Y.Y
    Out_mat.Y.Z = (m4321 + m3124 + m1423 - m3421 - m4123 - m1324) / m.Z.Y
    Out_mat.Y.W = (m3412 + m1342 + m4132 - m4312 - m1432 - m3142) / m.W.Y
    
    Out_mat.Z.X = (m3412 + m3124 + m3241 - m3421 - m3214 - m3142) / m.X.Z
    Out_mat.Z.Y = (m4321 + m1342 + m2314 - m4312 - m2341 - m1324) / m.Y.Z
    Out_mat.Z.Z = (m1234 + m2431 + m4132 - m2134 - m1432 - m4231) / m.Z.Z
    Out_mat.Z.W = (m2143 + m1423 + m4213 - m1243 - m4123 - m2413) / m.W.Z
    
    Out_mat.W.X = (m4321 + m4132 + m4213 - m4312 - m4123 - m4231) / m.X.W
    Out_mat.W.Y = (m3412 + m1423 + m2431 - m3421 - m1432 - m2413) / m.Y.W
    Out_mat.W.Z = (m2143 + m1342 + m3241 - m1243 - m2341 - m3142) / m.Z.W
    Out_mat.W.W = (m1234 + m2314 + m3124 - m2134 - m3214 - m1324) / m.W.W
    
    mat_inverse = True
End If
End Function



