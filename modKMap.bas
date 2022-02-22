Attribute VB_Name = "modKMap"
Option Explicit

Private m_KMap() As Single
Global Const g_KMapFilePath As String = "\kmap.bin"
Global Const g_KMapPreviewFilePath As String = "\kmap.bmp"
Private m_DistMap() As Single

#Const UseGPUKMAP = 1
#Const UseHermiteInterpolation = 0
#Const UseFindMaxInterpolation = 1
#Const UseNearestInterpolation = 0

Private Const m_GPUKMap_WaitMS = 200
Private Const m_GPUKMap_WaitEnd = 1000
#If UseGPUKMAP Then
Private Const m_GPUKMap_ExceptedTimeCost_PerTexel = 20000 / (1024 ^ 2)
Private m_GPUKMap_ExceptedTimeCost As Single
#End If

Function KMap_TryLoadExisting() As Boolean
On Error GoTo ErrReturn

If Len(Dir$(g_Map_Dir & g_KMapFilePath)) = 0 Then GoTo ErrReturn
If DateDiff("s", FileDateTime(g_Map_Dir & g_AltMapFilePath), FileDateTime(g_Map_Dir & g_KMapFilePath)) <= 0 Then GoTo ErrReturn

ReDim m_KMap(g_Map_MaxX, g_Map_MaxY)
Open g_Map_Dir & g_KMapFilePath For Binary Access Read As #1
Get #1, , m_KMap
Close #1

KMap_TryLoadExisting = True
Exit Function
ErrReturn:
End Function

Private Sub KMap_SaveToFile()
Open g_Map_Dir & g_KMapFilePath For Binary Access Write As #1
Put #1, , m_KMap
Close #1
End Sub

Private Function KMap_Generate_By_gpukmap() As Boolean
#If UseGPUKMAP = 0 Then
    Exit Function
#End If
If Len(Dir$(App.Path & "\gpukmap.exe")) = 0 Then Exit Function

Dim CmdLine As String
Dim Created As Boolean
Dim SI As STARTUPINFO
Dim P_I As PROCESS_INFORMATION
SI.cb = Len(SI)
CmdLine = """" & App.Path & "\gpukmap.exe """ & g_Map_Size_X & " " & g_Map_Size_Y & " """ & g_Map_Dir & g_AltMapFilePath & """ """ & g_Map_Dir & g_KMapFilePath & """ 64 1 " & g_Map_Dir & g_KMapPreviewFilePath & vbNullChar
Created = CreateProcess(ByVal 0, ByVal StrPtr(CmdLine), ByVal 0, ByVal 0, 0, NORMAL_PRIORITY_CLASS, ByVal 0, ByVal 0, SI, P_I)
If Created = False Then
    Debug.Print GetLastError
    Exit Function
End If

Dim ExitCode As Long
Dim Progress_Val As Long

m_GPUKMap_ExceptedTimeCost = m_GPUKMap_ExceptedTimeCost_PerTexel * g_Map_Size_X * g_Map_Size_Y
Do
    
    If GetExitCodeProcess(P_I.hProcess, ExitCode) Then
        If ExitCode <> STILL_ACTIVE Then
            KMap_Generate_By_gpukmap = True
            Exit Do
        End If
    End If
    Select Case WaitForSingleObject(P_I.hProcess, m_GPUKMap_WaitMS)
    Case WAIT_OBJECT_0
        'Continue looping
    Case WAIT_ABANDONED
        'MT bug
    Case WAIT_FAILED
        Sleep m_GPUKMap_WaitMS / 2
    End Select
    If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Progress_Val, m_GPUKMap_ExceptedTimeCost
    Progress_Val = Progress_Val + m_GPUKMap_WaitMS
    If Progress_Val > m_GPUKMap_ExceptedTimeCost Then Progress_Val = m_GPUKMap_ExceptedTimeCost
Loop While KMap_Generate_By_gpukmap = False
CloseHandle P_I.hThread
CloseHandle P_I.hProcess

'AppActivate Shell(App.Path & "\gpukmap.exe " & g_Map_Size_X & " " & g_Map_Size_Y & " " & g_Map_Dir & g_AltMapFilePath & " " & g_Map_Dir & g_KMapFilePath & " 32 1", vbMinimizedNoFocus), True

Sleep m_GPUKMap_WaitEnd

If KMap_Generate_By_gpukmap Then KMap_Generate_By_gpukmap = KMap_TryLoadExisting
If KMap_Generate_By_gpukmap = False Then
    Debug.Print "gpukmap failed."
End If
End Function

Sub KMap_Generate()
ReDim m_KMap(g_Map_MaxX, g_Map_MaxY)

Dim X As Long, Y As Long
If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating K Map by gpukmap"

If KMap_Generate_By_gpukmap Then Exit Sub
If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnStartProgress "Generating K Map. It's very slow, so it will quit after done."

If DoEvents = 0 Then End
ReDim m_DistMap(g_Map_MaxX, g_Map_MaxY)
Dim dX As Single, dY As Single
For Y = 0 To g_Map_MaxY
    dY = Y - g_Map_HalfY
    For X = 0 To g_Map_MaxX
        dX = X - g_Map_HalfX
        m_DistMap(X, Y) = Sqr(dX * dX + dY * dY)
        If m_DistMap(X, Y) < 1 Then m_DistMap(X, Y) = 1
    Next
Next
If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress 10, g_Map_MaxY + 10

For Y = 0 To g_Map_MaxY
    For X = 0 To g_Map_MaxX
        m_KMap(X, Y) = KMap_Calculate(X, Y)
        If (X And &HFF) = 0 Then
            If Not (g_ProgressCallback Is Nothing) Then g_ProgressCallback.OnProgress Y + 10 + (X / g_Map_MaxX), g_Map_MaxY + 10
        End If
    Next
Next

Erase m_DistMap
KMap_SaveToFile
End
End Sub

Private Function KMap_Calculate(ByVal pX As Long, ByVal pY As Long) As Single
Dim X As Long, Y As Long

Dim KVal As Single
Dim MaxKVal As Single
Dim Altitude As Single
Dim CurAltitude As Single
Dim Dist As Single

CurAltitude = AltMap_GetVal(pX, pY)

For Y = -g_Map_HalfY To g_Map_HalfY - 1
    For X = -g_Map_HalfX To g_Map_HalfX - 1
        Altitude = AltMap_GetVal(X + pX, Y + pY)
        If Altitude <= CurAltitude Then GoTo Continue
        
        'Dist = Sqr(X * X + Y * Y)
        Dist = m_DistMap(X + g_Map_HalfX, Y + g_Map_HalfY)
        KVal = (Altitude - CurAltitude) / Dist
        
        If KVal > MaxKVal Then
            MaxKVal = KVal
        End If
Continue:
    Next
Next

KMap_Calculate = MaxKVal
End Function

Function KMap_GetVal(ByVal X As Long, ByVal Y As Long) As Single
KMap_GetVal = m_KMap(X And g_Map_MaxX, Y And g_Map_MaxY)
End Function

#If UseNearestInterpolation Then
Function KMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
KMap_GetVal_Interpolated = m_KMap(CLng(Int(X)) And g_Map_MaxX, CLng(Int(Y)) And g_Map_MaxY)
End Function

#ElseIf UseFindMaxInterpolation Then
Function KMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
Dim V1 As Single, V2 As Single, V3 As Single, V4 As Single
Dim X1 As Long, Y1 As Long

X1 = Int(X)
Y1 = Int(Y)

V1 = m_KMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_KMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_KMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_KMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

KMap_GetVal_Interpolated = max(V1, max(V2, max(V3, V4)))
End Function
#Else 'UseFindMaxInterpolation

Function KMap_GetVal_Interpolated(ByVal X As Single, ByVal Y As Single) As Single
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

V1 = m_KMap((X1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V2 = m_KMap((X1 + 1) And g_Map_MaxX, (Y1) And g_Map_MaxY)
V3 = m_KMap((X1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)
V4 = m_KMap((X1 + 1) And g_Map_MaxX, (Y1 + 1) And g_Map_MaxY)

V5 = V1 + (V2 - V1) * SX
V6 = V3 + (V4 - V3) * SX
KMap_GetVal_Interpolated = V5 + (V6 - V5) * SY
End Function
#End If 'UseFindMaxInterpolation

Sub KMap_Cleanup()
Erase m_KMap
End Sub
