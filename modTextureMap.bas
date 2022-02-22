Attribute VB_Name = "modTextureMap"
Option Explicit

#Const DoNotInterpolate = 0
#Const UseHermiteInterpolation = 1
#Const DoUnpack = 1

Private m_TexPackedMap() As Long
Private m_TextureSampScale As Single
Private m_TexMap_SizeX As Long
Private m_TexMap_SizeY As Long
Private m_TexMap_MaxX As Long
Private m_TexMap_MaxY As Long
Global Const g_TextureMapFilePath As String = "\texturemap.bmp"

Private Type TextureSrc_t
    XRes As Long
    YRes As Long
    Pixels() As Long
End Type

Private m_TextureSrc() As TextureSrc_t
Private m_TexturePositions() As Single

#If DoUnpack Then
Private m_TextureMap() As vec4_t
#End If

Private Sub TextureMap_GetMeta()
m_TextureSampScale = Val(MetaData_Query("[texture]", "scale", "1"))
m_TexMap_SizeX = g_Map_Size_X * m_TextureSampScale
m_TexMap_SizeY = g_Map_Size_Y * m_TextureSampScale
m_TexMap_MaxX = m_TexMap_SizeX - 1
m_TexMap_MaxY = m_TexMap_SizeY - 1
End Sub

Private Sub TextureMap_MultNormal()
Dim X As Long, Y As Long
For Y = 0 To m_TexMap_MaxY
    For X = 0 To m_TexMap_MaxX
        m_TexPackedMap(X, Y) = vec4_to_rgb(vec4_scale(vec4_from_rgb(m_TexPackedMap(X, Y)), NormalMap_GetVal(X, Y).Y))
    Next
Next
End Sub

Function TextureMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

Dim DoMultNormal As Boolean

TextureMap_GetMeta

If Len(Dir$(g_Map_Dir & g_TextureMapFilePath)) = 0 Then GoTo ErrReturn
If LCase$("\" & MetaData_Query("[texture]", "source")) <> LCase$(g_TextureMapFilePath) Then
    DoMultNormal = False
    If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapFilePath), FileDateTime(g_Map_Dir & g_TextureMapFilePath)) <= 0 Then GoTo ErrReturn
    If DateDiff("s", FileDateTime(g_Map_Dir & g_NormalMapFilePath), FileDateTime(g_Map_Dir & g_TextureMapFilePath)) <= 0 Then GoTo ErrReturn
End If

TextureMap_TryLoadExisting = Bmp_ReadColors(g_Map_Dir & g_TextureMapFilePath, m_TexMap_SizeX, m_TexMap_SizeY, m_TexPackedMap)
If DoMultNormal Then TextureMap_MultNormal
#If DoUnpack Then
    If TextureMap_TryLoadExisting Then TextureMap_Unpack
#End If
Exit Function
ErrReturn:
End Function

Private Sub TextureMap_SaveToFile()
If Bmp_WriteColors(g_Map_Dir & g_TextureMapFilePath, m_TexMap_SizeX, m_TexMap_SizeY, m_TexPackedMap) = False Then Beep
End Sub

Private Function TexSrc_Sampling(ByVal Index As Long, ByVal pX As Long, ByVal pY As Long) As Long
TexSrc_Sampling = m_TextureSrc(Index).Pixels((pX) Mod m_TextureSrc(Index).XRes, (pY) Mod m_TextureSrc(Index).YRes)
End Function

Private Function TextureMap_Calculate(ByVal pX As Single, ByVal pY As Single) As Long
Dim Alt As Double
Alt = AltMap_GetVal_Interpolated(pX / m_TextureSampScale, pY / m_TextureSampScale)

Dim I As Long, TexInd As Long
Dim Color As vec4_t

For TexInd = 0 To UBound(m_TextureSrc)
    I = TexInd * 2
    If Alt >= m_TexturePositions(I) And Alt < m_TexturePositions(I + 1) Then
        Color = vec4_from_rgb(m_TextureSrc(TexInd).Pixels(pX Mod m_TextureSrc(TexInd).XRes, pY Mod m_TextureSrc(TexInd).YRes))
        GoTo Sampling_Finished
    End If
    If Alt >= m_TexturePositions(I + 1) And Alt < m_TexturePositions(I + 2) Then
        Color = vec4_lerp _
        ( _
            vec4_from_rgb(TexSrc_Sampling(TexInd, pX, pY)), _
            vec4_from_rgb(TexSrc_Sampling(TexInd + 1, pX, pY)), _
            s_hermite((Alt - m_TexturePositions(I + 1)) / (m_TexturePositions(I + 2) - m_TexturePositions(I + 1))) _
        )
        GoTo Sampling_Finished
    End If
Next

TexInd = UBound(m_TextureSrc)
Color = vec4_from_rgb(m_TextureSrc(TexInd).Pixels(pX Mod m_TextureSrc(TexInd).XRes, pY Mod m_TextureSrc(TexInd).YRes))

Sampling_Finished:
Color = vec4_scale(Color, NormalMap_GetVal_Interpolated(pX / m_TextureSampScale, pY / m_TextureSampScale).Y)

TextureMap_Calculate = vec4_to_rgb(Color)
End Function

Sub TextureMap_Generate()
TextureMap_GetMeta

ReDim m_TexPackedMap(m_TexMap_MaxX, m_TexMap_MaxY)

If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating Texture Map"

Dim SrcTexFiles() As String, I As Long
SrcTexFiles = Split(MetaData_Query("[texture]", "source"), "|")
ReDim m_TextureSrc(UBound(SrcTexFiles))
For I = 0 To UBound(SrcTexFiles)
    If Bmp_ReadColors(g_Map_Dir & "\" & SrcTexFiles(I), m_TextureSrc(I).XRes, m_TextureSrc(I).YRes, m_TextureSrc(I).Pixels) = False Then
        MsgBox "Could not load textures", vbCritical
        End
    End If
Next
Erase SrcTexFiles

Dim SrcTexPos() As String
SrcTexPos = Split(MetaData_Query("[texture]", "positions") & "," & FLT_MAX, ",")
ReDim m_TexturePositions(UBound(SrcTexPos))
For I = 0 To UBound(SrcTexPos)
    m_TexturePositions(I) = Val(SrcTexPos(I))
Next
Erase SrcTexPos

Dim X As Long, Y As Long
If UBound(m_TextureSrc) Then
    For Y = 0 To m_TexMap_MaxY
        For X = 0 To m_TexMap_MaxX
            m_TexPackedMap(X, Y) = TextureMap_Calculate(X, Y)
        Next
        If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, m_TexMap_MaxY
    Next
Else
    For Y = 0 To m_TexMap_MaxY
        For X = 0 To m_TexMap_MaxX
            m_TexPackedMap(X, Y) = vec4_to_rgb(vec4_scale(vec4_from_rgb(TexSrc_Sampling(0, X, Y)), NormalMap_GetVal_Interpolated(X / m_TextureSampScale, Y / m_TextureSampScale).Y))
        Next
        If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, m_TexMap_MaxY
    Next
End If

Erase m_TextureSrc
Erase m_TexturePositions

TextureMap_SaveToFile
#If DoUnpack Then
    TextureMap_Unpack
#End If
End Sub

#If DoUnpack Then
Private Sub TextureMap_Unpack()
If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Unpacking Texture Map"

ReDim m_TextureMap(m_TexMap_MaxX, m_TexMap_MaxY)

Dim X As Long, Y As Long
For Y = 0 To m_TexMap_MaxY
    For X = 0 To m_TexMap_MaxX
        m_TextureMap(X, Y) = vec4_from_rgb(m_TexPackedMap(X, Y))
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, m_TexMap_MaxY
Next
Erase m_TexPackedMap
End Sub

Function TextureMap_GetVal(ByVal X As Single, ByVal Y As Single) As vec4_t
TextureMap_GetVal = m_TextureMap(Val(X * m_TextureSampScale) And m_TexMap_MaxX, Val(Y * m_TextureSampScale) And m_TexMap_MaxY)
End Function

Function TextureMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As vec4_t
#If DoNotInterpolate Then
TextureMap_GetVal_Interpolated = TextureMap_GetVal(X, Y)
#Else 'DoNotInterpolate
Dim V1 As vec4_t, V2 As vec4_t, V3 As vec4_t, V4 As vec4_t, V5 As vec4_t, V6 As vec4_t
Dim X1 As Long, Y1 As Long, SX As Single, SY As Single

X = X * m_TextureSampScale
Y = Y * m_TextureSampScale

X1 = Int(X)
Y1 = Int(Y)
SX = X - X1
SY = Y - Y1

#If UseHermiteInterpolation Then
SX = SX * SX * (3 - 2 * SX)
SY = SY * SY * (3 - 2 * SY)
#End If

V1 = m_TextureMap((X1) And m_TexMap_MaxX, (Y1) And m_TexMap_MaxY)
V2 = m_TextureMap((X1 + 1) And m_TexMap_MaxX, (Y1) And m_TexMap_MaxY)
V3 = m_TextureMap((X1) And m_TexMap_MaxX, (Y1 + 1) And m_TexMap_MaxY)
V4 = m_TextureMap((X1 + 1) And m_TexMap_MaxX, (Y1 + 1) And m_TexMap_MaxY)

V5 = vec4_add(V1, vec4_scale(vec4_sub(V2, V1), SX))
V6 = vec4_add(V3, vec4_scale(vec4_sub(V4, V3), SX))
TextureMap_GetVal_Interpolated = vec4_add(V5, vec4_scale(vec4_sub(V6, V5), SY))
#End If 'DoNotInterpolate
End Function

#Else 'DoUnpack
Function TextureMap_GetVal(ByVal X As Single, ByVal Y As Single) As vec4_t
TextureMap_GetVal = vec4_from_rgb(m_TexPackedMap(Val(X * m_TextureSampScale) And m_TexMap_MaxX, Val(Y * m_TextureSampScale) And m_TexMap_MaxY))
End Function

Function TextureMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As vec4_t
#If DoNotInterpolate Then
TextureMap_GetVal_Interpolated = TextureMap_GetVal(X, Y)
#Else 'DoNotInterpolate
Dim V1 As vec4_t, V2 As vec4_t, V3 As vec4_t, V4 As vec4_t, V5 As vec4_t, V6 As vec4_t
Dim X1 As Long, Y1 As Long, SX As Single, SY As Single

X = X * m_TextureSampScale
Y = Y * m_TextureSampScale

X1 = Int(X)
Y1 = Int(Y)
SX = X - X1
SY = Y - Y1

#If UseHermiteInterpolation Then
SX = SX * SX * (3 - 2 * SX)
SY = SY * SY * (3 - 2 * SY)
#End If

V1 = vec4_from_rgb(m_TexPackedMap((X1) And m_TexMap_MaxX, (Y1) And m_TexMap_MaxY))
V2 = vec4_from_rgb(m_TexPackedMap((X1 + 1) And m_TexMap_MaxX, (Y1) And m_TexMap_MaxY))
V3 = vec4_from_rgb(m_TexPackedMap((X1) And m_TexMap_MaxX, (Y1 + 1) And m_TexMap_MaxY))
V4 = vec4_from_rgb(m_TexPackedMap((X1 + 1) And m_TexMap_MaxX, (Y1 + 1) And m_TexMap_MaxY))

V5 = vec4_add(V1, vec4_scale(vec4_sub(V2, V1), SX))
V6 = vec4_add(V3, vec4_scale(vec4_sub(V4, V3), SX))
TextureMap_GetVal_Interpolated = vec4_add(V5, vec4_scale(vec4_sub(V6, V5), SY))
#End If 'DoNotInterpolate
End Function
#End If 'DoUnpack

Sub TextureMap_Cleanup()
Erase m_TexPackedMap
#If DoUnpack Then
    Erase m_TextureMap
#End If
End Sub

