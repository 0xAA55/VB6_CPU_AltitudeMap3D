Attribute VB_Name = "modGL"
Option Explicit

Private g_hWnd As Long
Private g_hDC As Long
Private g_hRC As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Sub GL_Init(ByVal hWnd As Long, ByVal hDC As Long)
g_hWnd = hWnd
g_hDC = hDC

InitOpenGL hDC
End Sub

Sub GL_Term()
TermOpenGL
End Sub

Sub GL_Refresh()
SwapBuffers g_hDC
End Sub

Private Sub InitOpenGL(ByVal hDC As Long)
Dim FLog As Long

Dim PFD As PIXELFORMATDESCRIPTOR
With PFD
    .nSize = Len(PFD)
    .nVersion = 1
    .dwFlags = PFD_DRAW_TO_WINDOW Or PFD_SUPPORT_OPENGL Or PFD_DOUBLEBUFFER
    .iPixelType = PFD_TYPE_RGBA
    .cAlphaBits = 8
    .cBlueBits = 8
    .cGreenBits = 8
    .cRedBits = 8
    .cDepthBits = 32
    .iLayerType = PFD_MAIN_PLANE
End With

Dim nPixelFormat As Long
nPixelFormat = ChoosePixelFormat(hDC, PFD)
SetPixelFormat hDC, nPixelFormat, PFD

g_hRC = wglCreateContext(hDC)
If g_hRC = 0 Then
    Beep
    End
End If
wglMakeCurrent hDC, g_hRC
glewInit
'https://www.khronos.org/opengl/wiki/Swap_Interval

End Sub

Private Sub TermOpenGL()
wglMakeCurrent 0, 0
If g_hRC Then
    wglDeleteContext g_hRC
    g_hRC = 0
End If
g_hDC = 0
g_hWnd = 0
End Sub

Sub GL_GetClientSize(Optional Out_Width As Long, Optional Out_Height As Long)
Dim rc As RECT
GetClientRect g_hWnd, rc
Out_Width = rc.Right - rc.Left
Out_Height = rc.Bottom - rc.Top
End Sub
