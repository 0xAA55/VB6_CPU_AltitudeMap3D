Attribute VB_Name = "modMap"
Option Explicit

Global g_Map_Size_X As Long
Global g_Map_Size_Y As Long
Global g_Map_MaxX As Long
Global g_Map_MaxY As Long
Global g_Map_HalfX As Long
Global g_Map_HalfY As Long
Global g_Map_Dir As String

Global Const g_MetaDataFilePath As String = "\meta.ini"
Private Const m_MinDistThreshold As Single = 0.1

#Const DontUseKMap = 0 '是否不使用K值图
#Const UseInterpolatedRaycast = 1 '是否对地形和K值图使用插值
#Const UseDichotomyWhenNear = 0 '是否在射线靠近的时候使用二分法

'K值图还是有必要用的
'不使用K值图，但启用二分法也可以拯救一下运行效果
'地形越平坦，K值图的K值就越小，射线相交算法就越快

#If UseDichotomyWhenNear Then
Private Const m_DichotomyThreshold = 2 '二分法开始阈值
#End If

Sub Map_Init(Map_Dir As String)
g_Map_Dir = Map_Dir
MetaData_Load Map_Dir & g_MetaDataFilePath

g_Map_Size_X = Val(MetaData_Query("[map]", "size_x"))
g_Map_Size_Y = Val(MetaData_Query("[map]", "size_y"))
Debug.Assert Is2Pow(g_Map_Size_X)
Debug.Assert Is2Pow(g_Map_Size_Y)
g_Map_MaxX = g_Map_Size_X - 1
g_Map_MaxY = g_Map_Size_Y - 1
g_Map_HalfX = g_Map_Size_X \ 2
g_Map_HalfY = g_Map_Size_Y \ 2

'AltMap_FromPictureBox SrcPic
If AltMap_TryLoadExisting = False Then AltMap_Generate
If KMap_TryLoadExisting = False Then KMap_Generate
If WalkMap_TryLoadExisting = False Then WalkMap_Generate
If NormalMap_TryLoadExisting = False Then NormalMap_Generate
If WalkNormalMap_TryLoadExisting = False Then WalkNormalMap_Generate
If TextureMap_TryLoadExisting = False Then TextureMap_Generate
End Sub

Function Map_Raycast(Orig As vec4_t, RayDir As vec4_t, Out_CastPos As vec4_t, Optional ByVal MaxIterCount As Long = 64, Optional ByVal MaxDist As Single = 1024, Optional Out_CastDist As Single) As Boolean
On Error Resume Next

Dim I As Long
Dim CurOrig As vec4_t

Dim CurAlt As Single, CurK As Single, ForwardLength As Single, LengthTotal As Single
Dim RayK As Single, RayHor As Single, RayLengthRatio As Single
Dim KVal As Single

CurOrig = Orig

RayHor = (RayDir.X * RayDir.X + RayDir.Z * RayDir.Z)
If RayHor <= Single_Epsilon Then
    Map_Raycast = True
#If UseInterpolatedRaycast Then
    CurOrig.Y = AltMap_GetVal_Interpolated(CurOrig.X, CurOrig.Z)
#Else
    CurOrig.Y = AltMap_GetVal(CurOrig.X, CurOrig.Z)
#End If
    Out_CastPos = CurOrig
    Out_CastDist = Orig.Y - CurOrig.Y
    Exit Function
End If
RayHor = Sqr(RayHor)
RayK = -RayDir.Y / RayHor
RayLengthRatio = 1 / RayHor
Out_CastDist = 0

#If UseDichotomyWhenNear Then
ForwardLength = m_DichotomyThreshold
#End If

#If DontUseKMap Then
#If UseDichotomyWhenNear = 0 Then
ForwardLength = 1
#End If 'UseDichotomyWhenNear
For I = 1 To MaxIterCount * 8
#Else 'DontUseKMap
For I = 1 To MaxIterCount
#End If 'DontUseKMap
    If CurOrig.Y >= g_AltScale And RayDir.Y > 0 Then Exit Function
    
#If UseInterpolatedRaycast Then
    CurAlt = AltMap_GetVal_Interpolated(CurOrig.X, CurOrig.Z)
#Else 'UseInterpolatedRaycast
    CurAlt = AltMap_GetVal(CurOrig.X, CurOrig.Z)
#End If 'UseInterpolatedRaycast

#If DontUseKMap Then
#If UseDichotomyWhenNear Then
    If CurOrig.Y > CurAlt Then
        If CurOrig.Y > CurAlt + m_MinDistThreshold Then
            ForwardLength = Abs(ForwardLength) + m_MinDistThreshold
        Else
            ForwardLength = Abs(ForwardLength)
        End If
    Else
        ForwardLength = -Abs(ForwardLength) / 2
        Map_Raycast = True
    End If
#Else 'UseDichotomyWhenNear
    If CurOrig.Y <= CurAlt Then
        Map_Raycast = True
        Exit For
    End If
#End If 'UseDichotomyWhenNear
#Else 'DontUseKMap
#If UseDichotomyWhenNear Then
    If Abs(ForwardLength) < m_DichotomyThreshold Then
        If CurOrig.Y > CurAlt Then
            If CurOrig.Y > CurAlt + m_MinDistThreshold Then
                ForwardLength = Abs(ForwardLength) + m_MinDistThreshold
            Else
                ForwardLength = Abs(ForwardLength)
            End If
        Else
            ForwardLength = -Abs(ForwardLength) / 2
        End If
    Else
'#Else
'    If CurOrig.Y <= CurAlt Then
'        Map_Raycast = True
'        Exit For
'    End If
#End If 'UseDichotomyWhenNear
#If UseInterpolatedRaycast Then
    CurK = KMap_GetVal_Interpolated(CurOrig.X, CurOrig.Z)
#Else 'UseInterpolatedRaycast
    CurK = KMap_GetVal(CurOrig.X, CurOrig.Z)
#End If 'UseInterpolatedRaycast
    KVal = CurK + RayK
    If RayK < 0 And -RayK > CurK Then Exit Function
    
    If KVal > Single_Epsilon Then
        ForwardLength = (CurOrig.Y - CurAlt) * RayLengthRatio / KVal
    Else
        Exit Function
    End If
#If UseDichotomyWhenNear Then
    End If
#End If 'UseDichotomyWhenNear
#End If 'DontUseKMap
    
    CurOrig.X = CurOrig.X + RayDir.X * ForwardLength
    CurOrig.Y = CurOrig.Y + RayDir.Y * ForwardLength
    CurOrig.Z = CurOrig.Z + RayDir.Z * ForwardLength
    Out_CastDist = Out_CastDist + ForwardLength
#If UseDichotomyWhenNear Then
    If ForwardLength < -m_DichotomyThreshold Then Exit For
    If Abs(ForwardLength) <= m_MinDistThreshold Then Exit For
#ElseIf DontUseKMap = 0 Then 'UseDichotomyWhenNear
    If ForwardLength <= m_MinDistThreshold And CurOrig.Y <= CurAlt + m_MinDistThreshold Then Exit For
#End If 'DontUseKMap
    If Out_CastDist > MaxDist Then Exit Function
Next
Out_CastPos = CurOrig
#If DontUseKMap = 0 Then
Map_Raycast = True
#End If 'DontUseKMap
End Function

Sub Map_Cleanup()
AltMap_Cleanup
KMap_Cleanup
WalkMap_Cleanup
NormalMap_Cleanup
WalkNormalMap_Cleanup
End Sub
