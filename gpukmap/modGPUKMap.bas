Attribute VB_Name = "modGPUKMap"
Option Explicit

Global g_Map_Size_X As Long
Global g_Map_Size_Y As Long
Global g_Map() As Single
Global g_MaxAlt As Single
Global g_KMap() As Single
Global g_AltFile As String
Global g_KFile As String

Private m_Quad_VBO As Long

Private m_AltMap_Tex As Long

Private m_FBuf_Size_X As Long
Private m_FBuf_Size_Y As Long

Private m_FBuf As Long
Private m_FBuf_Tex(1) As Long

Private m_Shader_PixelK As Long

Private m_Batch_Size As Long
Global g_HideWindow As Boolean

Global g_WriteBmp As Boolean
Global g_BmpPath As String

#Const DoFitClamp = 0
#Const Use_PBO = 1
#Const ShadeOnce = 0

Private Function GPUK_InitAltMapTex() As Boolean

#If DoFitClamp Then
    Dim X As Long, Y As Long
    For Y = 0 To g_Map_Size_Y - 1
        For X = 0 To g_Map_Size_X - 1
            g_Map(X, Y) = g_Map(X, Y) / g_MaxAlt
        Next
    Next
#End If

#If Use_PBO Then
    If GL_ARB_pixel_buffer_object Then
        Dim PBO_Upload As Long
        frmMain.Output "Using PBO to upload data"
        glGenBuffers 1, PBO_Upload
        glBindBuffer GL_PIXEL_UNPACK_BUFFER, PBO_Upload
        glBufferData GL_PIXEL_UNPACK_BUFFER, g_Map_Size_X * g_Map_Size_Y * 4, ByVal 0, GL_STATIC_DRAW
        CopyMemory ByVal glMapBuffer(GL_PIXEL_UNPACK_BUFFER, GL_WRITE_ONLY), g_Map(0, 0), g_Map_Size_X * g_Map_Size_Y * 4
        glUnmapBuffer GL_PIXEL_UNPACK_BUFFER
    End If
#End If

glGenTextures 1, m_AltMap_Tex
glBindTexture GL_TEXTURE_2D, m_AltMap_Tex
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
#If Use_PBO Then
    If GL_ARB_pixel_buffer_object Then
        glTexImage2D GL_TEXTURE_2D, 0, GL_R32F, g_Map_Size_X, g_Map_Size_Y, 0, GL_RED, GL_FLOAT, ByVal 0
        glBindBuffer GL_PIXEL_UNPACK_BUFFER, 0
        glDeleteBuffers 1, PBO_Upload
    Else
        glTexImage2D GL_TEXTURE_2D, 0, GL_R32F, g_Map_Size_X, g_Map_Size_Y, 0, GL_RED, GL_FLOAT, g_Map(0, 0)
    End If
#Else
    glTexImage2D GL_TEXTURE_2D, 0, GL_R32F, g_Map_Size_X, g_Map_Size_Y, 0, GL_RED, GL_FLOAT, g_Map(0, 0)
#End If
glBindTexture GL_TEXTURE_2D, 0
GPUK_InitAltMapTex = True
End Function

Private Sub GPUK_InitFBufTextures()
glGenTextures 2, m_FBuf_Tex(0)
glBindTexture GL_TEXTURE_2D, m_FBuf_Tex(0)
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
glTexImage2D GL_TEXTURE_2D, 0, GL_R32F, m_FBuf_Size_X, m_FBuf_Size_Y, 0, GL_RED, GL_FLOAT, ByVal 0
glBindTexture GL_TEXTURE_2D, m_FBuf_Tex(1)
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MAG_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_MIN_FILTER, GL_NEAREST
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_S, GL_REPEAT
glTexParameteri GL_TEXTURE_2D, GL_TEXTURE_WRAP_T, GL_REPEAT
glTexImage2D GL_TEXTURE_2D, 0, GL_R32F, m_FBuf_Size_X, m_FBuf_Size_Y, 0, GL_RED, GL_FLOAT, ByVal 0
glBindTexture GL_TEXTURE_2D, 0
End Sub

Private Function GPUK_InitFrameBufs() As Boolean
If GL_ARB_framebuffer_object = 0 Then
    frmMain.Output "The Framebuffer Object is not supported."
    Exit Function
End If

m_FBuf_Size_X = g_Map_Size_X
m_FBuf_Size_Y = g_Map_Size_Y

glGenFramebuffers 1, m_FBuf
glBindFramebuffer GL_FRAMEBUFFER, m_FBuf

GPUK_InitFBufTextures
GPUK_SetFBufOutputTex m_FBuf_Tex(0)

If glCheckFramebufferStatus(GL_FRAMEBUFFER) <> GL_FRAMEBUFFER_COMPLETE Then
    '不可用，则清理残局
    frmMain.Output "glCheckFramebufferStatus() failed."
    Exit Function
End If

glBindFramebuffer GL_FRAMEBUFFER, 0

GPUK_InitFrameBufs = True
End Function

Private Sub GPUK_SetFBufOutputTex(ByVal TexName As Long)
Dim DrawBuffers(0) As Long
DrawBuffers(0) = GL_COLOR_ATTACHMENT0
glFramebufferTexture2D GL_FRAMEBUFFER, DrawBuffers(0), GL_TEXTURE_2D, TexName, 0
glDrawBuffers 1, DrawBuffers(0)
End Sub

Private Function GPUK_InitQuadModel() As Boolean
If GL_ARB_vertex_buffer_object = 0 Then
    frmMain.Output "The Vertex Buffer Object is not supported."
    Exit Function
End If

Dim QuadVertices(7) As Single
Const L As Single = -1
Const R As Single = 1
Const U As Single = -1
Const D As Single = 1

QuadVertices(0) = L
QuadVertices(1) = U

QuadVertices(2) = R
QuadVertices(3) = U

QuadVertices(4) = R
QuadVertices(5) = D

QuadVertices(6) = L
QuadVertices(7) = D

glGenBuffers 1, m_Quad_VBO
glBindBuffer GL_ARRAY_BUFFER, m_Quad_VBO
glBufferData GL_ARRAY_BUFFER, 8 * 4, ByVal 0, GL_STATIC_DRAW

CopyMemory ByVal glMapBuffer(GL_ARRAY_BUFFER, GL_WRITE_ONLY), QuadVertices(0), 8 * 4
glUnmapBuffer GL_ARRAY_BUFFER
glBindBuffer GL_ARRAY_BUFFER, 0

GPUK_InitQuadModel = True
End Function

Private Function GPUK_InitShaders() As Boolean
m_Shader_PixelK = CreateShaderProgramFromFile(App.Path & "\drawquad.vsh", App.Path & "\pixelk.fsh")
If m_Shader_PixelK = 0 Then Exit Function

SelectShader 0
GPUK_InitShaders = True
End Function

Private Function GPUK_SetVSHAttr(AttrName As String, ByVal Size As Long, ByVal Stride As Long, ByVal Offset As Long) As Boolean
Dim Location As Long

Location = glGetAttribLocation(g_Current_Shader, AttrName)
If Location = -1 Then
    frmMain.Output "[WARN] Location of attrib """ & AttrName & """ not found in the shader program."
    Exit Function
End If
glEnableVertexAttribArray Location
glVertexAttribPointer Location, Size, GL_FLOAT, GL_FALSE, Stride, ByVal Offset
GPUK_SetVSHAttr = True
End Function

Function GPUK_CalcKVal() As Boolean
glBindFramebuffer GL_FRAMEBUFFER, m_FBuf
glViewport 0, 0, g_Map_Size_X, g_Map_Size_Y
glClearColor 0, 0, 0, 0
glClear GL_COLOR_BUFFER_BIT

SelectShader m_Shader_PixelK
SetShaderUniformValue "Size_X", g_Map_Size_X
SetShaderUniformValue "Size_Y", g_Map_Size_Y
SetShaderUniformValue "Batch_W", m_Batch_Size
SetShaderUniformValue "Batch_H", m_Batch_Size
If GL_EXT_gpu_shader4 Then glBindFragDataLocation m_Shader_PixelK, 0, "PixelK"

glBindBuffer GL_ARRAY_BUFFER, m_Quad_VBO
GPUK_SetVSHAttr "iPosition", 2, 2 * 4, 0

frmMain.Output "Calculating K Val"

Dim Map_Half_W As Long, Map_Half_H As Long
Map_Half_W = g_Map_Size_X \ 2
Map_Half_H = g_Map_Size_Y \ 2

Dim CurTexInd As Long

SetShaderUniformInt "Initial", GL_TRUE

Dim X As Long, Y As Long
#If ShadeOnce = 0 Then
For Y = 0 To g_Map_Size_Y - 1 Step m_Batch_Size
    For X = 0 To g_Map_Size_X - 1 Step m_Batch_Size
#End If
        SelectShader m_Shader_PixelK
        SetShaderUniformTexture "AltMapTex", m_AltMap_Tex
        SetShaderUniformTexture "PrevKMapTex", m_FBuf_Tex(CurTexInd)
        CurTexInd = 1 - CurTexInd
        GPUK_SetFBufOutputTex m_FBuf_Tex(CurTexInd)
        glClear GL_COLOR_BUFFER_BIT
        SetShaderUniformValue "Start_X", X - Map_Half_W
        SetShaderUniformValue "Start_Y", Y - Map_Half_H
        glDrawArrays GL_QUADS, 0, 4
        glFinish
        SetShaderUniformInt "Initial", GL_FALSE
#If ShadeOnce = 0 Then
    Next
    frmMain.Output Format$(Y * 100 / (g_Map_Size_Y - 1), "0.0") & " %"
    frmMain.Output_WriteToLog
    If g_HideWindow = False And DoEvents = 0 Then Exit Function
Next
#End If

SelectShader 0
glBindBuffer GL_ARRAY_BUFFER, 0

frmMain.Output "100 %"
ReDim g_KMap(g_Map_Size_X - 1, g_Map_Size_Y - 1)
#If Use_PBO Then
    If GL_ARB_pixel_buffer_object Then
        Dim PBO_Download As Long
        frmMain.Output "Using PBO to download data"
        glGenBuffers 1, PBO_Download
        glBindBuffer GL_PIXEL_PACK_BUFFER, PBO_Download
        glBufferData GL_PIXEL_PACK_BUFFER, g_Map_Size_X * g_Map_Size_Y * 4, ByVal 0, GL_STATIC_READ
        glReadPixels 0, 0, g_Map_Size_X, g_Map_Size_Y, GL_RED, GL_FLOAT, ByVal 0
        CopyMemory g_KMap(0, 0), ByVal glMapBuffer(GL_PIXEL_PACK_BUFFER, GL_READ_ONLY), g_Map_Size_X * g_Map_Size_Y * 4
        glUnmapBuffer GL_PIXEL_PACK_BUFFER
        glBindBuffer GL_PIXEL_PACK_BUFFER, 0
        glDeleteBuffers 1, PBO_Download
    End If
#Else
    glReadPixels 0, 0, g_Map_Size_X, g_Map_Size_Y, GL_RED, GL_FLOAT, g_KMap(0, 0)
#End If
#If DoFitClamp Then
    For Y = 0 To g_Map_Size_Y - 1
        For X = 0 To g_Map_Size_X - 1
            g_KMap(X, Y) = g_KMap(X, Y) * g_MaxAlt
        Next
    Next
#End If

glBindFramebuffer GL_FRAMEBUFFER, 0
GPUK_CalcKVal = True

End Function

Function GPUK_ParseArgs() As Boolean
Dim CmdSplit() As String
Dim I As Long

CmdSplit = SplitCmdArgs(Command)
If UBound(CmdSplit) < 3 Then
    frmMain.Output "No command parameters."
    frmMain.Output "<size_x> <size_y> <altitude map file path> <output K file path> [batch size] [1=hide,0=show] [bmp preview]"
    GoTo ErrOut
End If
'frmMain.Output "Parsing commands"
'For I = 0 To UBound(CmdSplit)
'    frmMain.Output "arg" & I & " = " & CmdSplit(I)
'Next

g_Map_Size_X = Val(CmdSplit(0))
g_Map_Size_Y = Val(CmdSplit(1))
If g_Map_Size_X = 0 Or g_Map_Size_Y = 0 Then GoTo ErrOut

g_AltFile = CmdSplit(2)
g_KFile = CmdSplit(3)

m_Batch_Size = 64
If UBound(CmdSplit) >= 4 Then
    m_Batch_Size = Abs(Val(CmdSplit(4)))
    If m_Batch_Size = 8 Then
        frmMain.Output "[WARN] Batch size should not less than 8."
        m_Batch_Size = 8
    End If
End If

If UBound(CmdSplit) >= 5 Then
    g_HideWindow = CBool(CLng(Val(CmdSplit(5))))
    If g_HideWindow Then frmMain.Output "Will quit after works done."
End If

If UBound(CmdSplit) >= 6 Then
    g_WriteBmp = True
    g_BmpPath = CmdSplit(6)
End If

GPUK_ParseArgs = True
Exit Function
ErrOut:
End Function

Function GPUK_Init() As Boolean
On Error GoTo BugOut

frmMain.Output "Loading map from " & g_AltFile
ReDim g_Map(g_Map_Size_X - 1, g_Map_Size_Y - 1)
Open g_AltFile For Binary As #1
Get #1, , g_Map
Close #1
frmMain.Output "Map file loaded"

g_MaxAlt = 0
Dim X As Long, Y As Long
For Y = 0 To g_Map_Size_Y - 1
    For X = 0 To g_Map_Size_X - 1
        If g_Map(X, Y) > g_MaxAlt Then g_MaxAlt = g_Map(X, Y)
    Next
Next
frmMain.Output "Max altitude is " & Format$(g_MaxAlt, "0.00")
If g_MaxAlt <= 0.000001 Then
    frmMain.Output "Nothing to do."
    GoTo BugOut
End If

frmMain.Output "Creating Quad model"
If GPUK_InitQuadModel = False Then GoTo BugOut
frmMain.Output "Creating altitude map texture"
If GPUK_InitAltMapTex = False Then GoTo BugOut
frmMain.Output "Creating framebuffers"
If GPUK_InitFrameBufs = False Then GoTo BugOut
frmMain.Output "Compiling shaders"
If GPUK_InitShaders = False Then GoTo BugOut

GPUK_Init = True
Exit Function
BugOut:
Beep
If Err Then frmMain.Output "VB6 Runtime error: " & Err.Number & vbCrLf & Err.Description & vbCrLf & Err.Source
End Function
