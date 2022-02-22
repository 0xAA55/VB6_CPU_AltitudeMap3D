Attribute VB_Name = "modRenderer"
Option Explicit

#Const UseInterleavedPixelDrawing = 0 '是否强制使用像素交错方式提高帧数
#Const UsePrevAverageColor = 1 '是否对上一帧做平均色处理
#Const UseDynamicIterSteps = 0 '是否动态调整迭代次数
#Const DoFrameMoveCheck = 1 '是否动态做场景处理

Public m_RendererTarget As PictureBox
Public m_RendererTargetXRes As Long
Public m_RendererTargetYRes As Long
Public m_RendererTargetAspect As Single

Private Type BMIF_32bitRGB_t
    BMIF As BitmapInfoHeader_t
    BitFields(2) As Long
End Type
Private m_BMIF As BMIF_32bitRGB_t
Public m_VoxelCanvasXRes As Long
Public m_VoxelCanvasYRes As Long
Public m_VoxelCanvasBuffer() As Long
Public m_VoxelCanvasZBuffer() As Single

Public m_SkyColor1 As Long
Public m_SkyColor2 As Long
Public m_RenderDist As Single
Public m_IterCount As Long

Public m_FrameCount As Long

#If UseDynamicIterSteps Then
Public m_IterCountMin As Long
Public m_IterCountMax As Long
Private Const m_TargetFPS = 30
#End If 'UseDynamicIterSteps

#If DoFrameMoveCheck Then
Private Const m_FrameMoveCheckInterval As Long = 8
#End If 'DoFrameMoveCheck

Private m_RenderTimer As New clsTimer
Private m_RenderTimeDelta As Double
Private m_RenderTimeUpdate As Double
Private m_RenderFPS As Double

Private m_PixelAspect As Single

Private m_CameraPos As vec4_t
Private m_CameraOrient As mat4x4_t
Private m_CameraOrient_Inverse As mat4x4_t
Private m_CameraFov As Single
Private m_CameraProjX As Single
Private m_CameraProjY As Single
Private m_CameraProjZ As Single
Private m_ProjCenterX As Single
Private m_ProjCenterY As Single

Private m_ZEnabled As Boolean
Private m_ZWriteEnabled As Boolean

Private m_MaskBmpEnabled As Boolean
Private m_MaskBmpOffsetX As Long
Private m_MaskBmpOffsetY As Long
Private m_MaskBmpWidth As Long
Private m_MaskBmpHeight As Long
Private m_MaskBorder As Long
Private m_MaskBmp() As Long
Private m_MaskBmpRight As Long
Private m_MaskBmpBottom As Long
Private m_MaskBmpXor As Long

Enum Sampler_Methods
    SM_Nearest = 0
    SM_Linear = 1
    SM_Hermite = 2
End Enum
Private m_TextureInterpolate As Boolean
Private m_DoInterleavedPixelRendering As Boolean
Private m_LockTick As Boolean
Private m_TickLength As Double
Private m_TickLengthMax As Double
#If UsePrevAverageColor Then
Private m_PrevBlending As Boolean
Private m_VoxelCanvasBuffer_Prev() As Long
#End If

Sub Renderer_Init(RenderTarget As PictureBox, ByVal VoxelCanvasXRes As Long, ByVal VoxelCanvasYRes As Long, ByVal PixelAspect As Single)
Set m_RendererTarget = RenderTarget
m_RendererTargetXRes = m_RendererTarget.ScaleX(m_RendererTarget.ScaleWidth, m_RendererTarget.ScaleMode, vbPixels)
m_RendererTargetYRes = m_RendererTarget.ScaleY(m_RendererTarget.ScaleHeight, m_RendererTarget.ScaleMode, vbPixels)
m_RendererTargetAspect = m_RendererTargetXRes / m_RendererTargetYRes

m_PixelAspect = PixelAspect

m_VoxelCanvasXRes = VoxelCanvasXRes
m_VoxelCanvasYRes = VoxelCanvasYRes
With m_BMIF.BMIF
    .biSize = 40
    .biWidth = m_VoxelCanvasXRes
    .biHeight = -m_VoxelCanvasYRes
    .biPlanes = 1
    .biBitCount = 32
    .biCompression = BI_bitfields '让StretchDIBits以RGB，而非BGR的方式来识别像素
End With
m_BMIF.BitFields(0) = &HFF& '红色位域
m_BMIF.BitFields(1) = &HFF00& '绿色位域
m_BMIF.BitFields(2) = &HFF0000 '蓝色位域
ReDim m_VoxelCanvasBuffer(m_VoxelCanvasXRes - 1, m_VoxelCanvasYRes - 1)
ReDim m_VoxelCanvasZBuffer(m_VoxelCanvasXRes - 1, m_VoxelCanvasYRes - 1)

m_SkyColor1 = ParseRGB(MetaData_Query("[scene]", "skycolor1"))
m_SkyColor2 = ParseRGB(MetaData_Query("[scene]", "skycolor2"))
m_RenderDist = Val(MetaData_Query("[scene]", "render_dist", "512"))
m_IterCount = Val(MetaData_Query("[scene]", "iter_count", "64"))

#If UseDynamicIterSteps Then
m_IterCountMin = Val(MetaData_Query("[scene]", "iter_count_min"))
m_IterCountMax = Val(MetaData_Query("[scene]", "iter_count_max"))
If m_IterCountMin = 0 Then m_IterCountMin = 32
If m_IterCountMax = 0 Then m_IterCountMax = 128
#End If

m_TextureInterpolate = g_Cfg_Render_Interpolate
m_DoInterleavedPixelRendering = g_Cfg_Render_Interleaved
m_LockTick = g_Cfg_Game_Tick_Lock
m_TickLength = 1# / g_Cfg_Game_Tick_Freq
m_TickLengthMax = m_TickLength * g_Cfg_Game_Tick_Skip
#If UsePrevAverageColor Then
m_PrevBlending = g_Cfg_Render_Blending
If m_PrevBlending Then ReDim m_VoxelCanvasBuffer_Prev(m_VoxelCanvasXRes - 1, m_VoxelCanvasYRes - 1)
#End If

m_RenderTimer.Value = 0
m_RenderTimer.Start

m_FrameCount = 0
End Sub

Sub Renderer_ClearZ()
Dim X As Long, Y As Long
For Y = 0 To m_VoxelCanvasYRes - 1
    For X = 0 To m_VoxelCanvasXRes - 1
        m_VoxelCanvasZBuffer(X, Y) = FLT_MAX 'm_RenderDist
    Next
Next
End Sub

Sub Renderer_SetZ(ByVal DoTestEnabled As Boolean, ByVal WriteEnabled As Boolean)
m_ZEnabled = DoTestEnabled
m_ZWriteEnabled = WriteEnabled
End Sub

Sub Renderer_SetCamera(CameraPos As vec4_t, CameraOrient As mat4x4_t, ByVal Fov As Single)
m_CameraPos = CameraPos
m_CameraOrient = CameraOrient
m_CameraOrient_Inverse = mat_transpose(CameraOrient)
'mat_inverse m_CameraOrient_Inverse, 0, CameraOrient
m_CameraFov = Fov
m_CameraProjX = Sin(Fov * 0.5) * m_RendererTargetAspect
m_CameraProjY = Sin(Fov * 0.5)
m_CameraProjZ = Cos(Fov * 0.5)
End Sub

Sub Renderer_SetProjCenter(ByVal X As Single, ByVal Y As Single)
m_ProjCenterX = X
m_ProjCenterY = Y
End Sub

Private Function Border_Test(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long) As Boolean
Dim T As Long
If X1 > X2 Then
    T = X1
    X1 = X2
    X2 = T
End If
If Y1 > Y2 Then
    T = Y1
    Y1 = Y2
    Y2 = T
End If
If X1 >= m_VoxelCanvasXRes Then Exit Function
If Y1 >= m_VoxelCanvasYRes Then Exit Function
If X2 < 0 Then Exit Function
If Y2 < 0 Then Exit Function
Border_Test = True
End Function

Private Sub Border_Limit(X1 As Long, Y1 As Long, X2 As Long, Y2 As Long)
If X1 < 0 Then X1 = 0
If Y1 < 0 Then Y1 = 0
If X2 > m_VoxelCanvasXRes - 1 Then X2 = m_VoxelCanvasXRes - 1
If Y2 > m_VoxelCanvasYRes - 1 Then Y2 = m_VoxelCanvasYRes - 1
End Sub

Sub Renderer_DrawRect(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Z As Single, ByVal Color As Long)
Dim X As Long, Y As Long
If Border_Test(X1, Y1, X2, Y2) = False Then Exit Sub
Border_Limit X1, Y1, X2, Y2
If m_ZEnabled Then
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 And Z <= m_VoxelCanvasZBuffer(X, Y) Then
                m_VoxelCanvasBuffer(X, Y) = Color
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(X, Y) = Z
            End If
        Next
    Next
Else
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasBuffer(X, Y) = Color
        Next
    Next
    If m_ZWriteEnabled Then
        For Y = Y1 To Y2
            For X = X1 To X2
                If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasZBuffer(X, Y) = Z
            Next
        Next
    End If
End If
End Sub

Sub Renderer_DrawRect_Alpha(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Z As Single, ByVal Color As Long)
Dim X As Long, Y As Long, Alpha As Long
If Border_Test(X1, Y1, X2, Y2) = False Then Exit Sub
Border_Limit X1, Y1, X2, Y2

Alpha = (Color \ &H1000000) And &HFF&
Color = RGBDim(Color, Alpha)

If m_ZEnabled Then
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 And Z <= m_VoxelCanvasZBuffer(X, Y) Then
                m_VoxelCanvasBuffer(X, Y) = RGBDim(m_VoxelCanvasBuffer(X, Y), 255 - Alpha) + Color
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(X, Y) = Z
            End If
        Next
    Next
Else
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasBuffer(X, Y) = RGBDim(m_VoxelCanvasBuffer(X, Y), 255 - Alpha) + Color
        Next
    Next
    If m_ZWriteEnabled Then
        For Y = Y1 To Y2
            For X = X1 To X2
                If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasZBuffer(X, Y) = Z
            Next
        Next
    End If
End If
End Sub

Sub Renderer_DrawRect_Addition(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Z As Single, ByVal Color As Long)
Dim X As Long, Y As Long
If Border_Test(X1, Y1, X2, Y2) = False Then Exit Sub
Border_Limit X1, Y1, X2, Y2

If m_ZEnabled Then
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 And Z <= m_VoxelCanvasZBuffer(X, Y) Then
                m_VoxelCanvasBuffer(X, Y) = RGBAdd(m_VoxelCanvasBuffer(X, Y), Color)
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(X, Y) = Z
            End If
        Next
    Next
Else
    For Y = Y1 To Y2
        For X = X1 To X2
            If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasBuffer(X, Y) = RGBAdd(m_VoxelCanvasBuffer(X, Y), Color)
        Next
    Next
    If m_ZWriteEnabled Then
        For Y = Y1 To Y2
            For X = X1 To X2
                If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasZBuffer(X, Y) = Z
            Next
        Next
    End If
End If
End Sub

Sub Renderer_DrawSprite(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Z As Single, Sprite() As Long, ByVal Sprite_XRes As Long, ByVal Sprite_YRes As Long)
Dim SrcStartX As Long, SrcStartY As Long, X As Long, Y As Long, SrcX As Long, SrcY As Long
Dim DestMaxX As Long, DestMaxY As Long
Dim Sprite_MaxX As Long, Sprite_MaxY As Long
If Border_Test(X1, Y1, X2, Y2) = False Then Exit Sub
Dim DestW As Long, DestH As Long
DestMaxX = max(X2 - X1, 1)
DestMaxY = max(Y2 - Y1, 1)
DestW = DestMaxX + 1
DestH = DestMaxY + 1
If X1 < 0 Then SrcStartX = -X1 * Sprite_XRes \ DestMaxX
If Y1 < 0 Then SrcStartY = -Y1 * Sprite_YRes \ DestMaxY
Border_Limit X1, Y1, X2, Y2
Sprite_MaxX = Sprite_XRes - 1
Sprite_MaxY = Sprite_YRes - 1

If m_ZEnabled Then
    For Y = Y1 To Y2
        SrcY = min(SrcStartY + (Y - Y1) * Sprite_MaxY \ DestMaxY, Sprite_MaxY)
        For X = X1 To X2
            SrcX = min(SrcStartX + (X - X1) * Sprite_MaxX \ DestMaxX, Sprite_MaxX)
            If Renderer_BorderTest(X, Y) <> 0 And Z <= m_VoxelCanvasZBuffer(X, Y) Then
                m_VoxelCanvasBuffer(X, Y) = Sprite(SrcX, SrcY)
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(X, Y) = Z
            End If
        Next
    Next
Else
    For Y = Y1 To Y2
        SrcY = min(SrcStartY + (Y - Y1) * Sprite_MaxY \ DestMaxY, Sprite_MaxY)
        For X = X1 To X2
            SrcX = min(SrcStartX + (X - X1) * Sprite_MaxX \ DestMaxX, Sprite_MaxX)
            If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasBuffer(X, Y) = Sprite(SrcX, SrcY)
        Next
    Next
    If m_ZWriteEnabled Then
        For Y = Y1 To Y2
            For X = X1 To X2
                If Renderer_BorderTest(X, Y) <> 0 Then m_VoxelCanvasZBuffer(X, Y) = Z
            Next
        Next
    End If
End If
End Sub

Sub Renderer_DrawParticle(V As vec4_t, ByVal Size As Single, ByVal Color As Long)
Dim P As vec4_t
P = vec4_mult_matrix(vec4_sub(V, m_CameraPos), m_CameraOrient_Inverse)
If P.Z < m_CameraProjZ Then Exit Sub
P.X = P.X * m_CameraProjZ / (m_CameraProjX * P.Z)
P.Y = P.Y * m_CameraProjZ / (m_CameraProjY * P.Z)

P.X = (m_ProjCenterX + P.X) * m_VoxelCanvasXRes
P.Y = (m_ProjCenterY - P.Y) * m_VoxelCanvasYRes
Dim LSizeX As Single, LSizeY As Single
LSizeX = Size * m_VoxelCanvasXRes * m_CameraProjZ / (m_CameraProjX * P.Z)
LSizeY = Size * m_VoxelCanvasYRes * m_CameraProjZ / (m_CameraProjY * P.Z)

Renderer_DrawRect _
    P.X - LSizeX, P.Y - LSizeY, _
    P.X + LSizeX, P.Y + LSizeY, _
    P.Z, Color
End Sub

Sub Renderer_DrawParticle_Alpha(V As vec4_t, ByVal Size As Single, ByVal Color As Long)
Dim P As vec4_t
P = vec4_mult_matrix(vec4_sub(V, m_CameraPos), m_CameraOrient_Inverse)
If P.Z < m_CameraProjZ Then Exit Sub
P.X = P.X * m_CameraProjZ / (m_CameraProjX * P.Z)
P.Y = P.Y * m_CameraProjZ / (m_CameraProjY * P.Z)

P.X = (m_ProjCenterX + P.X) * m_VoxelCanvasXRes
P.Y = (m_ProjCenterY - P.Y) * m_VoxelCanvasYRes
Dim LSizeX As Single, LSizeY As Single
LSizeX = Size * m_VoxelCanvasXRes * m_CameraProjZ / (m_CameraProjX * P.Z)
LSizeY = Size * m_VoxelCanvasYRes * m_CameraProjZ / (m_CameraProjY * P.Z)

Renderer_DrawRect_Alpha _
    P.X - LSizeX, P.Y - LSizeY, _
    P.X + LSizeX, P.Y + LSizeY, _
    P.Z, Color
End Sub

Sub Renderer_DrawParticle_Addition(V As vec4_t, ByVal Size As Single, ByVal Color As Long)
Dim P As vec4_t
P = vec4_mult_matrix(vec4_sub(V, m_CameraPos), m_CameraOrient_Inverse)
If P.Z < m_CameraProjZ Then Exit Sub
P.X = P.X * m_CameraProjZ / (m_CameraProjX * P.Z)
P.Y = P.Y * m_CameraProjZ / (m_CameraProjY * P.Z)

P.X = (m_ProjCenterX + P.X) * m_VoxelCanvasXRes
P.Y = (m_ProjCenterY - P.Y) * m_VoxelCanvasYRes
Dim LSizeX As Single, LSizeY As Single
LSizeX = Size * m_VoxelCanvasXRes * m_CameraProjZ / (m_CameraProjX * P.Z)
LSizeY = Size * m_VoxelCanvasYRes * m_CameraProjZ / (m_CameraProjY * P.Z)

Renderer_DrawRect_Addition _
    P.X - LSizeX, P.Y - LSizeY, _
    P.X + LSizeX, P.Y + LSizeY, _
    P.Z, Color
End Sub

Sub Renderer_DrawParticle_Textured(V As vec4_t, ByVal Size As Single, Texture() As Long, ByVal Texture_XRes As Long, ByVal Texture_YRes As Long)
Dim P As vec4_t
P = vec4_mult_matrix(vec4_sub(V, m_CameraPos), m_CameraOrient_Inverse)
If P.Z < m_CameraProjZ Then Exit Sub
P.X = P.X * m_CameraProjZ / (m_CameraProjX * P.Z)
P.Y = P.Y * m_CameraProjZ / (m_CameraProjY * P.Z)

P.X = (m_ProjCenterX + P.X) * m_VoxelCanvasXRes
P.Y = (m_ProjCenterY - P.Y) * m_VoxelCanvasYRes
Dim LSizeX As Single, LSizeY As Single
LSizeX = Size * m_VoxelCanvasXRes * m_CameraProjZ / (m_CameraProjX * P.Z)
LSizeY = Size * m_VoxelCanvasYRes * m_CameraProjZ / (m_CameraProjY * P.Z)

Renderer_DrawSprite _
    P.X - LSizeX, P.Y - LSizeY, _
    P.X + LSizeX, P.Y + LSizeY, _
    P.Z, Texture, Texture_XRes, Texture_YRes
End Sub

Sub Renderer_DrawTriangle3D(V1 As vec4_t, V2 As vec4_t, V3 As vec4_t)
Dim P1 As vec4_t, P2 As vec4_t, P3 As vec4_t





End Sub

Private Sub Renderer_DrawProjectedTriangle3D(V1 As vec4_t, V2 As vec4_t, V3 As vec4_t)

End Sub

Sub Renderer_MaskBmp_Setup(ByVal Offset_X As Long, ByVal Offset_Y As Long, ByVal W As Long, ByVal H As Long, MaskBmp() As Long, ByVal MaskBorder As Long, Optional ByVal Invert As Boolean = False)
m_MaskBmpEnabled = True
m_MaskBmpOffsetX = Offset_X
m_MaskBmpOffsetY = Offset_Y
m_MaskBmpWidth = W
m_MaskBmpHeight = H
m_MaskBmpRight = m_MaskBmpOffsetX + W - 1
m_MaskBmpBottom = m_MaskBmpOffsetY + H - 1
m_MaskBmp = MaskBmp
m_MaskBorder = MaskBorder
m_MaskBmpXor = 0
If Invert Then Renderer_MaskBmp_Invert
End Sub

Sub Renderer_MaskBmp_Invert()
m_MaskBmpXor = Not m_MaskBmpXor
m_MaskBorder = Not m_MaskBorder
End Sub

Sub Renderer_MaskBmp_Disable()
m_MaskBmpEnabled = False
End Sub

Private Function Renderer_BorderTest(ByVal X As Long, ByVal Y As Long) As Long
If X >= m_VoxelCanvasXRes Then Exit Function
If Y >= m_VoxelCanvasYRes Then Exit Function
If X < 0 Then Exit Function
If Y < 0 Then Exit Function

If m_DoInterleavedPixelRendering Then
    If ((X + Y + m_FrameCount) And 1) = 1 Then Exit Function
End If

If m_MaskBmpEnabled = False Then
    Renderer_BorderTest = -1
    Exit Function
End If

Renderer_BorderTest = m_MaskBorder
If X < m_MaskBmpOffsetX Then Exit Function
If Y < m_MaskBmpOffsetY Then Exit Function
If X > m_MaskBmpRight Then Exit Function
If Y > m_MaskBmpBottom Then Exit Function

Renderer_BorderTest = m_MaskBmp(X - m_MaskBmpOffsetX, Y - m_MaskBmpOffsetY) Xor m_MaskBmpXor
End Function

Sub Renderer_RenderLandscape(ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, Optional IterCountBias As Long = 0, Optional RenderDistBias As Single = 0)
Dim X As Long, Y As Long, W As Long, H As Long
Dim RR As Long, RB As Long 'Relative right/bottom
Dim MaxX As Long, MaxY As Long
Dim RayStart As vec4_t
Dim RayU As Single, RayV As Single
Dim RayDir As vec4_t
Dim CastPoint As vec4_t, CastDist As Single
Dim CurX As Long, CurY As Long
Dim FogColor As vec4_t

Dim IterCount As Long, RenderDist As Single
IterCount = m_IterCount + IterCountBias
RenderDist = m_RenderDist + RenderDistBias

Dim SkyColor1 As vec4_t
SkyColor1 = vec4_from_rgb(m_SkyColor1)

Dim SkyColor2 As vec4_t
SkyColor2 = vec4_from_rgb(m_SkyColor2)

W = X2 + 1 - X1
H = Y2 + 1 - Y1
RR = X2 - X1
RB = Y2 - Y1
MaxX = m_VoxelCanvasXRes - 1
MaxY = m_VoxelCanvasYRes - 1

Dim CamOrient As mat4x4_t
CamOrient = m_CameraOrient
 
RayStart = m_CameraPos

For Y = 0 To RB
    CurY = Y + Y1
    RayV = (m_ProjCenterY - CurY / MaxY) * m_CameraProjY
#If UseInterleavedPixelDrawing Then
    For X = ((m_FrameCount + Y) And 1) To RR Step 2
#Else 'UseInterleavedPixelDrawing
    For X = 0 To RR
#End If 'UseInterleavedPixelDrawing
        CurX = X + X1
        
        If Renderer_BorderTest(CurX, CurY) = 0 Then GoTo DiscardPixel
        
        RayU = (CurX / MaxX - m_ProjCenterX) * m_CameraProjX
        RayDir = vec4_mult_matrix(vec4_normalize(vec4(RayU, RayV, m_CameraProjZ, 0)), CamOrient)
        
        FogColor = vec4_lerp(SkyColor1, SkyColor2, 1 - RayDir.Y)
        
        If m_ZEnabled Then RenderDist = m_VoxelCanvasZBuffer(CurX, CurY) + 1 + RenderDistBias
        If Map_Raycast(RayStart, RayDir, CastPoint, IterCount, RenderDist, CastDist) Then
            If m_ZEnabled = False Or CastDist <= m_VoxelCanvasZBuffer(CurX, CurY) Then
                If m_TextureInterpolate Then
                    m_VoxelCanvasBuffer(CurX, CurY) = vec4_to_rgb(vec4_lerp(TextureMap_GetVal_Interpolated(CastPoint.X, CastPoint.Z), FogColor, clamp(CastDist / m_RenderDist)))
                Else
                    m_VoxelCanvasBuffer(CurX, CurY) = vec4_to_rgb(vec4_lerp(TextureMap_GetVal(CastPoint.X, CastPoint.Z), FogColor, clamp(CastDist / m_RenderDist)))
                End If
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(CurX, CurY) = CastDist
            End If
        Else
            If m_ZEnabled = False Or RenderDist <= m_VoxelCanvasZBuffer(CurX, CurY) Then
                m_VoxelCanvasBuffer(CurX, CurY) = vec4_to_rgb(FogColor)
                If m_ZWriteEnabled Then m_VoxelCanvasZBuffer(CurX, CurY) = RenderDist
            End If
        End If
DiscardPixel:
    Next
#If DoFrameMoveCheck Then
    If m_LockTick Then
        If (Y And (m_FrameMoveCheckInterval * 2 - 1)) = m_FrameMoveCheckInterval Then Scene_FrameMove_CheckTick m_TickLength, m_TickLengthMax
    End If
#End If 'DoFrameMoveCheck
Next

#If UsePrevAverageColor Then
If m_PrevBlending Then
    Dim CurPix As Long
    For Y = 0 To RB
        CurY = Y + Y1
        For X = 0 To RR
            CurX = X + X1
            If Renderer_BorderTest(CurX, CurY) Then
                CurPix = m_VoxelCanvasBuffer(CurX, CurY)
                m_VoxelCanvasBuffer(CurX, CurY) = ((CurPix And &HFEFEFE) \ 2) + ((m_VoxelCanvasBuffer_Prev(CurX, CurY) And &HFEFEFE) \ 2)
                m_VoxelCanvasBuffer_Prev(CurX, CurY) = CurPix
            End If
        Next
    Next
End If
#End If

End Sub

Sub Renderer_Present(Optional Target As PictureBox = Nothing)
Dim RT As PictureBox
If Target Is Nothing Then
    Set RT = m_RendererTarget
Else
    Set RT = Target
End If

StretchDIBits RT.hDC, 0, 0, RT.ScaleX(RT.ScaleWidth, RT.ScaleMode, vbPixels), RT.ScaleY(RT.ScaleHeight, RT.ScaleMode, vbPixels), 0, 0, m_VoxelCanvasXRes, m_VoxelCanvasYRes, m_VoxelCanvasBuffer(0, 0), m_BMIF, 0, vbSrcCopy

m_FrameCount = (m_FrameCount + 1) And &H3FFFFFFF

#If UseDynamicIterSteps Then
If m_RenderTimeDelta >= 1 / 1000 Then
    If 1 / m_RenderTimeDelta > m_TargetFPS Then
        m_IterCount = m_IterCount + 1
        If m_IterCount > m_IterCountMax Then m_IterCount = m_IterCountMax
    Else
        m_IterCount = m_IterCount - 1
        If m_IterCount < m_IterCountMin Then m_IterCount = m_IterCountMin
    End If
Else
    m_IterCount = m_IterCount + 1
    If m_IterCount > m_IterCountMax Then m_IterCount = m_IterCountMax
End If
#End If

Dim CurTime As Double
CurTime = m_RenderTimer.Value
m_RenderTimeDelta = CurTime - m_RenderTimeUpdate
m_RenderTimeUpdate = CurTime

frmMain.Tag = Format$(1 / m_RenderTimeDelta, "0.00")
End Sub

Sub Renderer_Cleanup()
Erase m_VoxelCanvasBuffer, m_VoxelCanvasZBuffer
m_VoxelCanvasXRes = 0
m_VoxelCanvasYRes = 0
Set m_RendererTarget = Nothing
End Sub
