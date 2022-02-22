Attribute VB_Name = "modWalkNormalMap"
Option Explicit

Private m_WalkNormalMap() As vec4_t
Global Const g_WalkNormalMapFilePath As String = "\walknormalmap.bin"

#Const UseHermiteInterpolation = 1

Function WalkNormalMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

If Len(Dir$(g_Map_Dir & g_WalkNormalMapFilePath)) = 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapFilePath), FileDateTime(g_Map_Dir & g_WalkNormalMapFilePath)) <= 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_WalkMapFilePath), FileDateTime(g_Map_Dir & g_WalkNormalMapFilePath)) <= 0 Then GoTo ErrReturn

ReDim m_WalkNormalMap(g_Map_MaxX, g_Map_MaxY)
Open g_Map_Dir & g_WalkNormalMapFilePath For Binary Access Read As #1
Get #1, , m_WalkNormalMap
Close #1

WalkNormalMap_TryLoadExisting = True
Exit Function
ErrReturn:
End Function

Private Sub WalkNormalMap_SaveToFile()
Open g_Map_Dir & g_WalkNormalMapFilePath For Binary Access Write As #1
Put #1, , m_WalkNormalMap
Close #1
End Sub

Private Function WalkNormalMap_Calculate(ByVal pX As Long, ByVal pY As Long) As vec4_t
WalkNormalMap_Calculate.X = WalkMap_GetVal(pX - 1, pY) - WalkMap_GetVal(pX + 1, pY)
WalkNormalMap_Calculate.Z = WalkMap_GetVal(pX, pY - 1) - WalkMap_GetVal(pX, pY + 1)
WalkNormalMap_Calculate.Y = 2
WalkNormalMap_Calculate = vec4_normalize(WalkNormalMap_Calculate)
End Function

Sub WalkNormalMap_Generate()
ReDim m_WalkNormalMap(g_Map_MaxX, g_Map_MaxY)

If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating WalkNormal Map"

Dim X As Long, Y As Long
For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        m_WalkNormalMap(X, Y) = WalkNormalMap_Calculate(X, Y)
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, g_Map_MaxY
Next

WalkNormalMap_SaveToFile
End Sub

Function WalkNormalMap_GetVal(ByVal X As Long, ByVal Y As Long) As vec4_t
WalkNormalMap_GetVal = m_WalkNormalMap(X And g_Map_MaxX, Y And g_Map_MaxY)
End Function

Function WalkNormalMap_GetVal_Interpolated(ByVal X As Long, ByVal Y As Long) As vec4_t
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

V1 = m_WalkNormalMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_WalkNormalMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_WalkNormalMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_WalkNormalMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

V5 = vec4_add(V1, vec4_scale(vec4_sub(V2, V1), SX))
V6 = vec4_add(V3, vec4_scale(vec4_sub(V4, V3), SX))
WalkNormalMap_GetVal_Interpolated = vec4_add(V5, vec4_scale(vec4_sub(V6, V5), SY))
End Function

Sub WalkNormalMap_Cleanup()
Erase m_WalkNormalMap
End Sub
