Attribute VB_Name = "modMisc"
Option Explicit

'#Const FullOptimize = 0

Type Point_t
    X As Long
    Y As Long
End Type

Type Rect_t
    X1 As Long
    Y1 As Long
    X2 As Long
    Y2 As Long
End Type

Type BitmapFileHeader_t
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Type BitmapInfoHeader_t
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Declare Function StretchDIBits Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal dX As Long, ByVal dY As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, lpBits As Any, lpBitsInfo As Any, ByVal wUsage As Long, ByVal dwRop As Long) As Long
Global Const BI_bitfields = 3&

Global g_ProgressCallback

Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Declare Function GetCursorPos Lib "user32" (lpPoint As Point_t) As Long
Declare Function ClientToScreen Lib "user32" (ByVal hWnd As Long, lpPoint As Point_t) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As Rect_t) As Long
Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Declare Function GetForegroundWindow Lib "user32" () As Long
Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Declare Sub ZeroMemory Lib "KERNEL32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Declare Function GetLastError Lib "KERNEL32" () As Long
Declare Sub Sleep Lib "KERNEL32" (ByVal dwMilliseconds As Long)

Type STARTUPINFO
    cb As Long
    lpReserved As Long
    lpDesktop As Long
    lpTitle As Long
    dwX As Long
    dwY As Long
    dwXSize As Long
    dwYSize As Long
    dwXCountChars As Long
    dwYCountChars As Long
    dwFillAttribute As Long
    dwFlags As Long
    wShowWindow As Integer
    cbReserved2 As Integer
    lpReserved2 As Long
    hStdInput As Long
    hStdOutput As Long
    hStdError As Long
End Type
Type PROCESS_INFORMATION
    hProcess As Long
    hThread As Long
    dwProcessId As Long
    dwThreadId As Long
End Type
Declare Function CreateProcess Lib "KERNEL32" Alias "CreateProcessW" (lpApplicationName As Any, lpCommandLine As Any, lpProcessAttributes As Any, lpThreadAttributes As Any, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, lpCurrentDriectory As Any, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Declare Function WaitForSingleObject Lib "KERNEL32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Declare Function GetExitCodeProcess Lib "KERNEL32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Global Const NORMAL_PRIORITY_CLASS = &H20&
Global Const WAIT_TIMEOUT = &H102&
Global Const WAIT_OBJECT_0 = &H0&
Global Const WAIT_ABANDONED = &H80&
Global Const WAIT_FAILED = &HFFFFFFFF
Global Const STILL_ACTIVE = 259

Global g_Args() As String

Sub ParseArgs()
g_Args = SplitCmdArgs(Command)
End Sub

Function Args_Search(ByVal Keyword As String) As Long
Dim I As Long
Keyword = LCase$(Keyword)
For I = 0 To UBound(g_Args)
    If LCase$(g_Args(I)) = Keyword Then
        Args_Search = I + 1
        Exit Function
    End If
Next
End Function

Function Args_Query(Keyword As String, Optional Delimiter As String = "=") As String
Dim I As Long
I = Args_Search(Keyword)
If I Then Args_Query = Split(g_Args(I - 1), Delimiter)(0)
End Function

Function Is2Pow(ByVal Value As Long) As Boolean
Is2Pow = (Value = &H1 Or Value = &H2 Or Value = &H4 Or Value = &H8 Or _
          Value = &H10 Or Value = &H20 Or Value = &H40 Or Value = &H80 Or _
          Value = &H100 Or Value = &H200 Or Value = &H400 Or Value = &H800 Or _
          Value = &H1000 Or Value = &H2000 Or Value = &H4000 Or Value = &H8000& Or _
          Value = &H10000 Or Value = &H20000 Or Value = &H40000 Or Value = &H80000 Or _
          Value = &H100000 Or Value = &H200000 Or Value = &H400000 Or Value = &H800000 Or _
          Value = &H1000000 Or Value = &H2000000 Or Value = &H4000000 Or Value = &H8000000 Or _
          Value = &H10000000 Or Value = &H20000000 Or Value = &H40000000 Or Value = &H80000000)
End Function

Function RGBDim(ByVal Color As Long, ByVal Alpha As Long) As Long
#If FullOptimize Then
RGBDim = (((Color And &HFF&) * Alpha \ 255&) And &HFF&) Or (((Color And &HFF00&) * Alpha \ 255&) And &HFF00&) Or ((((Color And &HFF0000)) * Alpha \ 255&) And &HFF0000)
#Else
RGBDim = (((Color And &HFF&) * Alpha \ 255&) And &HFF&) Or (((Color And &HFF00&) * Alpha \ 255&) And &HFF00&) Or (((((Color And &HFF0000) \ &H10000) * Alpha \ 255&) * &H10000) And &HFF0000)
#End If
End Function

Function RGBLerp(ByVal C1 As Long, ByVal C2 As Long, ByVal S As Single) As Long
S = clamp(S * S * (3 - 2 * S))
Dim AVal As Long
AVal = S * 255
RGBLerp = RGBDim(C1, 255 - AVal) + RGBDim(C2, AVal)
End Function

Function RGBAdd(ByVal C1 As Long, ByVal C2 As Long) As Long
RGBAdd = ((C1 And &HFEFEFE) + (C2 And &HFEFEFE))
Dim Carrys As Long
Carrys = RGBAdd And &H1010100
RGBAdd = (RGBAdd Or (Carrys - Carrys \ &H100)) And &HFFFFFF
End Function

Function RGBA(ByVal R As Long, ByVal G As Long, ByVal B As Long, ByVal A As Long) As Long
If R > 255 Then R = 255
If G > 255 Then G = 255
If B > 255 Then B = 255
If A > 255 Then A = 255
#If FullOptimize = 0 Then
If A >= 128 Then A = &HFFFFFF00 Or A
#End If
RGBA = R Or (G * &H100&) Or (B * &H10000) Or (A * &H1000000)
End Function

Function Rainbow(ByVal XPos As Single, Optional ByVal YPos As Single = 0) As Long
Dim LPos As Long, R As Long, G As Long, B As Long
LPos = Int(XPos * 1536) Mod 1536

If LPos < 256 Then
    R = 255
    G = LPos
    B = 0
ElseIf LPos < 512 Then
    R = 511 - LPos
    G = 255
    B = 0
ElseIf LPos < 768 Then
    R = 0
    G = 255
    B = LPos - 512
ElseIf LPos < 1024 Then
    R = 0
    G = 1023 - LPos
    B = 255
ElseIf LPos < 1280 Then
    R = LPos - 1024
    G = 0
    B = 255
Else
    R = 255
    G = 0
    B = 1535 - LPos
End If

LPos = Int(YPos * 128)
If LPos > 127 Then LPos = 127

Rainbow = RGB(R * (127 - LPos) \ 127 + LPos, G * (127 - LPos) \ 127 + LPos, B * (127 - LPos) \ 127 + LPos)
End Function

Function ParseRGB(Optional R_G_B As String = "255,255,255") As Long
Dim RGBComp() As String
RGBComp = Split(R_G_B, ",")
ParseRGB = RGB(Val(RGBComp(0)), Val(RGBComp(1)), Val(RGBComp(2)))
End Function

Function SplitCmdArgs(CommandsVal As String) As String()
Dim I As Long, NB As Long, L As Long
Dim ReArrange As String
I = 1
L = Len(CommandsVal)
Do
    Select Case Mid$(CommandsVal, I, 1)
    Case """"
        NB = InStr(I + 1, CommandsVal, """")
        If NB Then
            ReArrange = ReArrange & vbNullChar & Mid$(CommandsVal, I + 1, NB - (I + 1))
            I = NB + 1
        Else
            Exit Do
        End If
    Case " "
        I = I + 1
    Case Else
        NB = InStr(I, CommandsVal, " ")
        If NB = 0 Then
            ReArrange = ReArrange & vbNullChar & Mid$(CommandsVal, I)
            Exit Do
        End If
        ReArrange = ReArrange & vbNullChar & Mid$(CommandsVal, I, NB - I)
        I = NB
    End Select
Loop While I <= L
SplitCmdArgs = Split(Mid$(ReArrange, 2), vbNullChar)
End Function

Function bar_indicator(ByVal charge As Single, Optional ByVal count As Long = 10, Optional char_a As String = "|", Optional char_b As String = "_", Optional char_c As String = "!", Optional ByVal max_count As Long = 20) As String
Dim count_a As Long, count_b As Long, count_c As Long
count_a = max(charge * count, 0)
If charge >= Single_Epsilon Then count_a = max(count_a, 1)
count_c = clamp(count_a - count, 0, max_count - count)
count_a = min(count_a, count)
count_b = count - count_a
bar_indicator = String(count_a, char_a) & String(count_b, char_b) & String(count_c, char_c)
End Function
