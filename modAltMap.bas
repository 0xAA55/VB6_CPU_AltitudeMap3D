Attribute VB_Name = "modAltMap"
Option Explicit

Private m_AltMap() As Single

Global g_AltScale As Single
Global g_MaxAlt As Single
Global Const g_AltMapSourceFilePath As String = "\altitudemap.bmp"
Global Const g_AltMapFilePath As String = "\altmap.bin"

#Const UseHermiteInterpolation = 1
#Const UseFindMaxInterpolation = 0
#Const UseNearestInterpolation = 0

'Sub AltMap_FromPictureBox(SrcPic As PictureBox)
'Dim X As Long, Y As Long
'
'g_Map_Size_X = SrcPic.ScaleX(SrcPic.Picture.Width, vbHimetric, vbPixels)
'g_Map_Size_Y = SrcPic.ScaleY(SrcPic.Picture.Height, vbHimetric, vbPixels)
'
'Debug.Assert Is2Pow(g_Map_Size_X)
'Debug.Assert Is2Pow(g_Map_Size_Y)
'g_Map_MaxX = g_Map_Size_X - 1
'g_Map_MaxY = g_Map_Size_Y - 1
'g_Map_HalfX = g_Map_Size_X \ 2
'g_Map_HalfY = g_Map_Size_Y \ 2
'
'g_AltScale = Val(MetaData_Query("[altmap]", "alt_scale"))
'If g_AltScale <= Single_Epsilon Then g_AltScale = 256
'
'g_MaxAlt = 0
'ReDim m_AltMap(g_Map_MaxX, g_Map_MaxY)
'For Y = 0 To g_Map_MaxY
'    For X = 0 To g_Map_MaxX
'        m_AltMap(X, Y) = (SrcPic.Point(X, Y) And &HFF&) * g_AltScale / 255
'        If g_MaxAlt < m_AltMap(X, Y) Then g_MaxAlt = m_AltMap(X, Y)
'    Next
'Next
'End Sub

Function AltMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

If Len(Dir$(g_Map_Dir & g_AltMapFilePath)) = 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapSourceFilePath), FileDateTime(g_Map_Dir & g_AltMapFilePath)) <= 0 Then GoTo ErrReturn

g_AltScale = Val(MetaData_Query("[altmap]", "alt_scale", "256"))

ReDim m_AltMap(g_Map_MaxX, g_Map_MaxY)
Open g_Map_Dir & g_AltMapFilePath For Binary Access Read As #1
Get #1, , m_AltMap
Close #1

AltMap_TryLoadExisting = True
Exit Function
ErrReturn:
End Function

Private Sub AltMap_SaveToFile()
Open g_Map_Dir & g_AltMapFilePath For Binary Access Write As #1
Put #1, , m_AltMap
Close #1
End Sub

Sub AltMap_Generate()

Dim AltMap_Src() As Long, XRes As Long, YRes As Long
If Bmp_ReadColors(g_Map_Dir & g_AltMapSourceFilePath, XRes, YRes, AltMap_Src) = False Then
    MsgBox "Could not load BMP source file.", vbCritical
    End
End If
If XRes <> g_Map_Size_X Or YRes <> g_Map_Size_Y Then
    g_Map_Size_X = XRes
    g_Map_Size_Y = YRes
    Debug.Assert Is2Pow(XRes)
    Debug.Assert Is2Pow(YRes)
    g_Map_MaxX = XRes - 1
    g_Map_MaxY = YRes - 1
    g_Map_HalfX = XRes \ 2
    g_Map_HalfY = YRes \ 2
End If

Dim X As Long, Y As Long, C As Long
If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating Altitude Map"

g_AltScale = Val(MetaData_Query("[altmap]", "alt_scale", "256"))

Dim Blur As Long
Blur = CLng(Val(MetaData_Query("[altmap]", "blur"))) \ 2

Dim AltMap() As Single
ReDim AltMap(g_Map_MaxX, g_Map_MaxY)

For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        C = AltMap_Src(X, Y)
        AltMap(X, Y) = CSng((C And &HFF&) + ((C And &HFF00&) \ &H100&) + ((C And &HFF0000) \ &H10000)) * g_AltScale / 765
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y, g_Map_Size_Y * 2 - 1
Next

ReDim m_AltMap(g_Map_MaxX, g_Map_MaxY)
Dim SumVal As Single
Dim SumMax As Single
Dim XB As Long, YB As Long
SumMax = Blur * 2 + 1
SumMax = SumMax * SumMax

For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        SumVal = 0
        For YB = -Blur To Blur
            For XB = -Blur To Blur
                SumVal = SumVal + AltMap((X + XB) And g_Map_MaxX, (Y + YB) And g_Map_MaxY)
            Next
        Next
        m_AltMap(X, Y) = SumVal / SumMax
    Next
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y + g_Map_MaxY, g_Map_Size_Y * 2 - 1
Next

Erase AltMap
AltMap_SaveToFile
End Sub

Function AltMap_GetVal(ByVal X As Long, ByVal Y As Long) As Single
AltMap_GetVal = m_AltMap(X And g_Map_MaxX, Y And g_Map_MaxY)
End Function

#If UseNearestInterpolation Then
Function AltMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
AltMap_GetVal_Interpolated = m_AltMap(CLng(Int(X)) And g_Map_MaxX, CLng(Int(Y)) And g_Map_MaxY)
End Function

#ElseIf UseFindMaxInterpolation Then
Function AltMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
Dim V1 As Single, V2 As Single, V3 As Single, V4 As Single
Dim X1 As Long, Y1 As Long

X1 = Int(X)
Y1 = Int(Y)

V1 = m_AltMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_AltMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_AltMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_AltMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

AltMap_GetVal_Interpolated = max(max(max(V1, V2), V3), V4)
End Function

#Else

Function AltMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
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

V1 = m_AltMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_AltMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_AltMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_AltMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

V5 = V1 + (V2 - V1) * SX
V6 = V3 + (V4 - V3) * SX
AltMap_GetVal_Interpolated = V5 + (V6 - V5) * SY
End Function
#End If

Sub AltMap_Cleanup()
Erase m_AltMap
End Sub
