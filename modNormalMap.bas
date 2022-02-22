Attribute VB_Name = "modNormalMap"
Option Explicit

Private m_NormalMap() As vec4_t
Global Const g_NormalMapFilePath As String = "\normalmap.bin"

#Const UseHermiteInterpolation = 1

Function NormalMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

If Len(Dir$(g_Map_Dir & g_NormalMapFilePath)) = 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapFilePath), FileDateTime(g_Map_Dir & g_NormalMapFilePath)) <= 0 Then GoTo ErrReturn

ReDim m_NormalMap(g_Map_MaxX, g_Map_MaxY)
Open g_Map_Dir & g_NormalMapFilePath For Binary Access Read As #1
Get #1, , m_NormalMap
Close #1

NormalMap_TryLoadExisting = True
Exit Function
ErrReturn:
End Function

Private Sub NormalMap_SaveToFile()
Open g_Map_Dir & g_NormalMapFilePath For Binary Access Write As #1
Put #1, , m_NormalMap
Close #1
End Sub

Private Function NormalMap_Calculate(ByVal pX As Long, ByVal pY As Long) As vec4_t
NormalMap_Calculate.X = AltMap_GetVal(pX - 1, pY) - AltMap_GetVal(pX + 1, pY)
NormalMap_Calculate.Z = AltMap_GetVal(pX, pY - 1) - AltMap_GetVal(pX, pY + 1)
NormalMap_Calculate.Y = 2
NormalMap_Calculate = vec4_normalize(NormalMap_Calculate)
End Function

Sub NormalMap_Generate()
ReDim m_NormalMap(g_Map_MaxX, g_Map_MaxY)

If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating Normal Map"

Dim X As Long, Y As Long
For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        m_NormalMap(X, Y) = NormalMap_Calculate(X, Y)
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, g_Map_MaxY
Next

NormalMap_SaveToFile
End Sub

Function NormalMap_GetVal(ByVal X As Long, ByVal Y As Long) As vec4_t
NormalMap_GetVal = m_NormalMap(X And g_Map_MaxX, Y And g_Map_MaxY)
End Function

Function NormalMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As vec4_t
Dim V1 As vec4_t, V2 As vec4_t, V3 As vec4_t, V4 As vec4_t, V5 As vec4_t, V6 As vec4_t
Dim X1 As Long, Y1 As Long, SX As Single, SY As Single

X1 = Int(X)
Y1 = Int(Y)
SX = X - X1
SY = Y - Y1

#If UseHermiteInterpolation Then
SX = SX * SX * (3 - 2 * SX)
SY = SY * SY * (3 - 2 * SY)
#End If

V1 = m_NormalMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_NormalMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_NormalMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_NormalMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

V5 = vec4_add(V1, vec4_scale(vec4_sub(V2, V1), SX))
V6 = vec4_add(V3, vec4_scale(vec4_sub(V4, V3), SX))
NormalMap_GetVal_Interpolated = vec4_add(V5, vec4_scale(vec4_sub(V6, V5), SY))
End Function

Sub NormalMap_Cleanup()
Erase m_NormalMap
End Sub
