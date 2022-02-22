Attribute VB_Name = "modWalkMap"
Option Explicit

Private m_WalkMap() As Single
Private m_MaxMap() As Single
Global Const g_WalkMapFootSize As Single = 8
Global Const g_WalkMapFilePath As String = "\walkmap.bin"

#Const UseHermiteInterpolation = 1

Function WalkMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

If Len(Dir$(g_Map_Dir & g_WalkMapFilePath)) = 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapFilePath), FileDateTime(g_Map_Dir & g_WalkMapFilePath)) <= 0 Then GoTo ErrReturn

ReDim m_WalkMap(g_Map_MaxX, g_Map_MaxY)
Open g_Map_Dir & g_WalkMapFilePath For Binary Access Read As #1
Get #1, , m_WalkMap
Close #1

WalkMap_TryLoadExisting = True
Exit Function
ErrReturn:
End Function

Private Sub WalkMap_SaveToFile()
Open g_Map_Dir & g_WalkMapFilePath For Binary Access Write As #1
Put #1, , m_WalkMap
Close #1
End Sub

Private Function WalkMap_Calculate1(ByVal pX As Long, ByVal pY As Long) As Single
Const WalkMapHalfSize As Long = g_WalkMapFootSize \ 2
'Const WalkMapArea As Single = (WalkMapHalfSize * 2 + 1) ^ 2

Dim X As Long, Y As Long
Dim CX As Long, CY As Long

Dim Alt As Single
For Y = -WalkMapHalfSize To WalkMapHalfSize
    CY = pY + Y
    For X = -WalkMapHalfSize To WalkMapHalfSize
        CX = pX + X
        
        Alt = AltMap_GetVal(CX, CY)
        If WalkMap_Calculate1 < Alt Then WalkMap_Calculate1 = Alt
    Next
Next
End Function

Private Function WalkMap_Calculate2(ByVal pX As Long, ByVal pY As Long) As Single
Const WalkMapHalfSize As Long = g_WalkMapFootSize \ 2
Const WalkMapArea As Single = (WalkMapHalfSize * 2 + 1) ^ 2

Dim X As Long, Y As Long
Dim CX As Long, CY As Long

Dim Alt As Single
For Y = -WalkMapHalfSize To WalkMapHalfSize
    CY = pY + Y
    For X = -WalkMapHalfSize To WalkMapHalfSize
        WalkMap_Calculate2 = WalkMap_Calculate2 + m_MaxMap((pX + X) And g_Map_MaxX, (pY + Y) And g_Map_MaxY)
    Next
Next
WalkMap_Calculate2 = WalkMap_Calculate2 / WalkMapArea
End Function

Sub WalkMap_Generate()
ReDim m_MaxMap(g_Map_MaxX, g_Map_MaxY)
ReDim m_WalkMap(g_Map_MaxX, g_Map_MaxY)

If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating Walk Map"

Dim X As Long, Y As Long
For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        m_MaxMap(X, Y) = WalkMap_Calculate1(X, Y)
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, g_Map_Size_Y * 2 - 1
Next
For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        m_WalkMap(X, Y) = WalkMap_Calculate2(X, Y)
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y + g_Map_Size_Y, g_Map_Size_Y * 2 - 1
Next
Erase m_MaxMap
WalkMap_SaveToFile
End Sub

Function WalkMap_GetVal(ByVal X As Long, ByVal Y As Long) As Single
WalkMap_GetVal = m_WalkMap(X And g_Map_MaxX, Y And g_Map_MaxY)
End Function

Function WalkMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
Dim V1 As Single, V2 As Single, V3 As Single, V4 As Single, V5 As Single, V6 As Single
Dim X1 As Long, Y1 As Long, SX As Single, SY As Single

X1 = Int(X)
Y1 = Int(Y)
SX = X - X1
SY = Y - Y1

#If UseHermiteInterpolation Then
SX = SX * SX * (3 - 2 * SX)
SY = SY * SY * (3 - 2 * SY)
#End If

V1 = m_WalkMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_WalkMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_WalkMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_WalkMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

V5 = V1 + (V2 - V1) * SX
V6 = V3 + (V4 - V3) * SX
WalkMap_GetVal_Interpolated = V5 + (V6 - V5) * SY
End Function

Sub WalkMap_Cleanup()
Erase m_WalkMap
End Sub
