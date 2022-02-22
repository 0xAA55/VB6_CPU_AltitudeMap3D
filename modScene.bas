Attribute VB_Name = "modScene"
Option Explicit

Global g_SceneTimer As clsTimer
Global g_TimeUpdate As Double
Global g_TimeDelta As Double

Public m_PlayerPos As vec4_t
Public m_PlayerAng As vec4_t
Public m_PlayerVel As vec4_t
Public m_PlayerOrient As mat4x4_t

Public m_CamPos As vec4_t
Public m_CamDir As vec4_t
Public m_CamOrient As mat4x4_t
Public m_CamPosOrient As mat4x4_t
Public Const m_CamFovY As Single = PI * 0.5

Private Const m_MouseSensitivity As Single = 1 / (480 * m_CamFovY * 0.5)
Private Const m_CamMinPitch As Single = -PI * 0.5
Private Const m_CamMaxPitch As Single = PI * 0.5

Private Const m_ScopeSize As Single = 0.4
Private Const m_ScopeXCenter As Single = 0.75
Private Const m_ScopeYCenter As Single = 0.75
Private Const m_ScopeBaseFov As Single = m_CamFovY
Private Const m_ScopeMult As Single = 4

Private Const m_Gravity As Single = 200
Private Const m_Gravity_X As Single = 0
Private Const m_Gravity_Y As Single = -m_Gravity
Private Const m_Gravity_Z As Single = 0

Private Const m_WalkingMaxSpeed As Single = 100
Private Const m_WalkingAccelTime As Single = 0.5
Private Const m_WalkingAcceleration As Single = m_WalkingMaxSpeed * 4 / m_WalkingAccelTime
Private Const m_NeelingAcceleration As Single = m_WalkingAcceleration * 0.2
Private Const m_FloatingAcceleration As Single = m_WalkingAcceleration * 0.05
Private Const m_KneelDownSpeed As Single = 40
Private Const m_StandUpSpeed As Single = 40
Private Const m_StandUpTargetTime As Single = 0.1
Private Const m_StandNeelDamp As Single = 1 / (2 ^ m_StandUpSpeed)

Private Const m_StandingHeight As Single = 16
Private Const m_StandingHeightMin As Single = 10
Private Const m_StandingHeightMax As Single = 17
Private Const m_StandingHeightJumping As Single = 19
Private Const m_StandingStepLength As Single = m_StandingHeight
Private Const m_StandingMovementDamp As Single = 1 / (2 ^ (m_WalkingAcceleration * 2 / m_WalkingMaxSpeed))
Private Const m_StandingMovementDampLimit As Single = m_WalkingAcceleration * 4
Private Const m_NeelingHeight As Single = m_StandingHeight * 0.3
Private Const m_NeelingHeightMin As Single = m_NeelingHeight * 0.8
Private Const m_NeelingHeightMax As Single = m_NeelingHeight * 1.5 'm_StandingHeightMax * 0.4
Private Const m_NeelingStepLength As Single = m_NeelingHeight
Private Const m_NeelingMovementDamp As Single = m_StandingMovementDamp * 0.5
Private Const m_NeelingMovementDampLimit As Single = m_StandingMovementDampLimit * 1.2

Private Const m_JumpAcceleration As Single = 5000
Private Const m_JumpHeight As Single = m_StandingHeightMax * 1
Private Const m_JumpInterval As Single = 0.5
Private Const m_JumpVel = m_Gravity * (m_JumpHeight / m_Gravity) ^ 0.5

Private m_MouseCenter As Point_t
Private m_MouseKeys As Long
Private m_MouseKeysPrev As Long
Private m_TouchingGround As Boolean
Private m_GroundNormal As vec4_t
Private m_GroundPos As Single
Private m_Crouching As Boolean
Private m_PlayerHeight As Single
Private m_PlayerHeightMin As Single
Private m_PlayerHeightMax As Single
Private m_StepPos As Single
Private m_StepLength As Single
Private m_NumSteps As Long
Private m_Walking As Boolean
Private m_WalkingDir As vec4_t
Private m_MovementDamp As Single
Private m_MovementDampLimit As Single
Private m_Jumping As Boolean
Private m_JumpCharge As Single
Private m_JumpCharged As Boolean
Private m_Slippering As Boolean
Private m_ScopeX As Long
Private m_ScopeY As Long
Private m_ScopeXRes As Long
Private m_ScopeYRes As Long
Private m_ScopeMask() As Long
Private m_ScopeOpened As Boolean

Private Const m_BobbingRangeXStanding As Single = m_StandingStepLength * 0.05
Private Const m_BobbingRangeYStanding As Single = m_StandingStepLength * 0.1
Private Const m_BobbingRangeXNeeling As Single = m_NeelingStepLength * 0.05
Private Const m_BobbingRangeYNeeling As Single = m_NeelingStepLength * 0.1
Private Const m_BobbingYawRange As Single = -0.0005
Private Const m_BobbingPitchRange As Single = 0.005
Private Const m_BobbingRollRange As Single = 0.005
Private Const m_BobbingExtentDamp As Single = 0.5
Private Const m_BobbingScopeShakeRangeX As Single = 0.05
Private Const m_BobbingScopeShakeRangeY As Single = -0.01
Private Const m_BobbingPitchFalling As Single = m_BobbingPitchRange * 2 / m_JumpVel
Private Const m_BobbingExtentFactor As Single = 2
Private m_BobbingYaw As Single
Private m_BobbingPitch As Single
Private m_BobbingRoll As Single
Private m_BobbingRangeX As Single
Private m_BobbingRangeY As Single
Private m_BobbingShakeX As Single
Private m_BobbingShakeY As Single
Private m_BobbingExtent As Single

Private m_SceneView As PictureBox

Private Sub GetCursorCenter(ByVal hWnd As Long)
Dim cr As Rect_t

GetClientRect hWnd, cr
m_MouseCenter.X = (cr.X2 - cr.X1) \ 2
m_MouseCenter.Y = (cr.Y2 - cr.Y1) \ 2

ClientToScreen hWnd, m_MouseCenter
End Sub

Private Sub Scope_Init()
m_ScopeOpened = False
m_ScopeX = clamp(m_ScopeXCenter - m_ScopeSize * 0.5) * m_VoxelCanvasXRes
m_ScopeY = clamp(m_ScopeYCenter - m_ScopeSize * 0.5) * m_VoxelCanvasYRes
m_ScopeXRes = m_ScopeSize * m_VoxelCanvasXRes
m_ScopeYRes = m_ScopeSize * m_VoxelCanvasYRes * m_RendererTargetAspect
ReDim m_ScopeMask(m_ScopeXRes - 1, m_ScopeYRes - 1)

Dim X As Long, Y As Long, RX!, RY!
For Y = 0 To m_ScopeYRes - 1
    RY = (Y / (m_ScopeYRes - 1)) * 2 - 1
    For X = 0 To m_ScopeXRes - 1
        RX = (X / (m_ScopeXRes - 1)) * 2 - 1
        If RX * RX + RY * RY <= 1 Then m_ScopeMask(X, Y) = -1 Else m_ScopeMask(X, Y) = 0
    Next
Next
End Sub

Sub Scene_Init(SceneView As PictureBox, ByVal XRes As Long, ByVal YRes As Long, Map_Dir As String, ByVal VoxelCanvasXRes As Long, ByVal VoxelCanvasYRes As Long)
Map_Init Map_Dir

Renderer_Init SceneView, VoxelCanvasXRes, VoxelCanvasYRes, (XRes / VoxelCanvasXRes) / (YRes / VoxelCanvasYRes)
Scope_Init

m_PlayerPos.X = Val(MetaData_Query("[scene]", "initial_x"))
m_PlayerPos.Y = Val(MetaData_Query("[scene]", "initial_y"))
m_PlayerPos.Z = Val(MetaData_Query("[scene]", "initial_z"))

Set m_SceneView = SceneView
m_SceneView.Move 0, 0, m_SceneView.Parent.ScaleX(XRes, vbPixels, m_SceneView.Parent.ScaleMode), m_SceneView.Parent.ScaleY(YRes, vbPixels, m_SceneView.Parent.ScaleMode)
m_SceneView.Visible = True
GetCursorCenter m_SceneView.hWnd
SetCursorPos m_MouseCenter.X, m_MouseCenter.Y

m_PlayerHeight = m_StandingHeight
m_PlayerHeightMin = m_StandingHeightMin
m_PlayerHeightMax = m_StandingHeightMax
m_StepLength = m_StandingStepLength
m_MovementDamp = m_StandingMovementDamp
m_MovementDampLimit = m_StandingMovementDampLimit
m_BobbingRangeX = m_BobbingRangeXStanding
m_BobbingRangeY = m_BobbingRangeYStanding

Set g_SceneTimer = New clsTimer
g_SceneTimer.Start
g_TimeDelta = 1 / 60
End Sub

Private Function WindowIsFront() As Boolean
Dim hWndParent As Long
Dim hWndFront As Long

hWndFront = GetForegroundWindow
hWndParent = GetParent(m_SceneView.hWnd)
Do While hWndParent
    If hWndParent = hWndFront Then
        WindowIsFront = True
        Exit Do
    End If
    hWndParent = GetParent(hWndParent)
Loop
End Function

Private Sub ProcMouseInput()
If WindowIsFront Then
    Dim CurMouse As Point_t
    If GetCursorPos(CurMouse) Then
        GetCursorCenter m_SceneView.hWnd
        
        Dim RotX As Single, RotY As Single
        RotX = m_MouseSensitivity * (CurMouse.X - m_MouseCenter.X) '* Screen.TwipsPerPixelX
        RotY = m_MouseSensitivity * (CurMouse.Y - m_MouseCenter.Y) '* Screen.TwipsPerPixelY
        
        m_PlayerAng.X = m_PlayerAng.X + RotX
        m_PlayerAng.Y = m_PlayerAng.Y + RotY
        
        SetCursorPos m_MouseCenter.X, m_MouseCenter.Y
    End If
    
    m_MouseKeysPrev = m_MouseKeys
    m_MouseKeys = 0
    If frmMain.GetKeyState(1) Then m_MouseKeys = m_MouseKeys Or 1
    If frmMain.GetKeyState(2) Then m_MouseKeys = m_MouseKeys Or 2
    If frmMain.GetKeyState(3) Then m_MouseKeys = m_MouseKeys Or 4
End If

If (m_MouseKeysPrev And 2) = 0 And (m_MouseKeys And 2) = 2 Then m_ScopeOpened = Not m_ScopeOpened

If m_PlayerAng.Y < m_CamMinPitch Then m_PlayerAng.Y = m_CamMinPitch
If m_PlayerAng.Y > m_CamMaxPitch Then m_PlayerAng.Y = m_CamMaxPitch
If m_PlayerAng.X < 0 Then m_PlayerAng.X = m_PlayerAng.X + PI * 2
If m_PlayerAng.X > PI * 2 Then m_PlayerAng.X = m_PlayerAng.X - PI * 2
m_PlayerOrient = mat_rot_y(m_PlayerAng.X)
End Sub

Private Sub UpdatePosition()
m_PlayerPos.X = m_PlayerPos.X + m_PlayerVel.X * g_TimeDelta
m_PlayerPos.Y = m_PlayerPos.Y + m_PlayerVel.Y * g_TimeDelta
m_PlayerPos.Z = m_PlayerPos.Z + m_PlayerVel.Z * g_TimeDelta
End Sub

Private Sub ProcCameraBobbing()
If m_TouchingGround Then
    m_StepPos = m_StepPos + Sqr(m_PlayerVel.X * m_PlayerVel.X + m_PlayerVel.Z * m_PlayerVel.Z) * PI * g_TimeDelta / m_StepLength
    m_BobbingExtent = vec4_length(m_PlayerVel) * m_BobbingExtentFactor / m_WalkingMaxSpeed
Else
    m_BobbingExtent = m_BobbingExtent * (m_BobbingExtentDamp ^ g_TimeDelta)
End If
m_BobbingShakeX = Cos(m_StepPos)
m_BobbingShakeY = Abs(Sin(m_StepPos))
End Sub

Private Sub UpdateCamera()
m_BobbingYaw = m_BobbingShakeX * m_BobbingYawRange * m_BobbingExtent
m_BobbingPitch = m_BobbingShakeY * m_BobbingPitchRange * m_BobbingExtent + m_PlayerVel.Y * m_BobbingPitchFalling
m_BobbingRoll = m_BobbingShakeX * m_BobbingRollRange * m_BobbingExtent
m_CamDir = vec4_add(m_PlayerAng, vec4(m_BobbingYaw, m_BobbingPitch, m_BobbingRoll, 0))
m_CamOrient = mat_rot_euler(m_CamDir.X, m_CamDir.Y, m_CamDir.Z)
m_CamPos = vec4_add(m_PlayerPos, vec4_scale(m_CamOrient.X, m_BobbingShakeX * m_BobbingRangeX * m_BobbingExtent))
m_CamPos = vec4_add(m_CamPos, vec4_scale(vec4(0, 1, 0, 0), m_BobbingShakeY * m_BobbingRangeY * m_BobbingExtent))

If m_StepPos >= PI Then
    m_NumSteps = (m_NumSteps And &HFFFFFFFE) + 1
    If m_StepPos >= PI * 2 Then
        m_NumSteps = m_NumSteps + 1
        m_StepPos = m_StepPos - PI * 2
    End If
End If
m_CamPosOrient = m_CamOrient
m_CamPosOrient.W = vec4(m_CamPos.X, m_CamPos.Y, m_CamPos.Z, 1)
End Sub

Private Sub UpdateWalkingDirection()
Dim AccFront As vec4_t
Dim AccRight As vec4_t
AccFront = m_PlayerOrient.Z
AccRight = m_PlayerOrient.X

m_WalkingDir = vec4(0, 0, 0, 0)

If frmMain.GetKeyState(vbKeyW) Then
    m_WalkingDir = vec4_add(m_WalkingDir, AccFront)
    m_Walking = True
End If
If frmMain.GetKeyState(vbKeyS) Then
    m_WalkingDir = vec4_sub(m_WalkingDir, AccFront)
    m_Walking = True
End If
If frmMain.GetKeyState(vbKeyA) Then
    m_WalkingDir = vec4_sub(m_WalkingDir, AccRight)
    m_Walking = True
End If
If frmMain.GetKeyState(vbKeyD) Then
    m_WalkingDir = vec4_add(m_WalkingDir, AccRight)
    m_Walking = True
End If

Dim TiltNormal As vec4_t
If m_Walking = True And m_TouchingGround = True Then
    If m_GroundNormal.Y >= Sqr2_Half Then
        TiltNormal = m_GroundNormal
    Else
        Dim XZLen As Single
        XZLen = Sqr(m_GroundNormal.X * m_GroundNormal.X + m_GroundNormal.Z * m_GroundNormal.Z)
        If XZLen >= Single_Epsilon Then
            TiltNormal = vec4(m_GroundNormal.X * Sqr2_Half / XZLen, Sqr2_Half, m_GroundNormal.Z * Sqr2_Half / XZLen, 0)
        Else
            TiltNormal = vec4(0, 1, 0, 0)
        End If
    End If
    
    Dim SurfaceNormalMatrix As mat4x4_t
    SurfaceNormalMatrix.Z = vec4(0, 0, 1, 0)
    SurfaceNormalMatrix.Y = TiltNormal
    SurfaceNormalMatrix.X = vec4_cross3(SurfaceNormalMatrix.Y, SurfaceNormalMatrix.Z)
    SurfaceNormalMatrix.Z = vec4_cross3(SurfaceNormalMatrix.X, SurfaceNormalMatrix.Y)
    SurfaceNormalMatrix.W.W = 1
    m_WalkingDir = vec4_mult_matrix(vec4_normalize(m_WalkingDir), SurfaceNormalMatrix)
End If
End Sub

Private Function LimitedDamping(ByVal ScaleVal As Single, ByVal Damp As Single, ByVal ScaleLimit As Single) As Single
If ScaleVal >= Single_Epsilon Then LimitedDamping = 1 - clamp(ScaleVal - ScaleVal * Damp, 0, ScaleLimit) / ScaleVal
End Function

Private Sub ProcWalking()
Dim WalkAcc As Single
If m_TouchingGround Then
    If m_Crouching Then
        WalkAcc = m_NeelingAcceleration
    Else
        WalkAcc = m_WalkingAcceleration * m_GroundNormal.Y
    End If
Else
    WalkAcc = m_FloatingAcceleration
End If

m_PlayerVel = vec4_add(m_PlayerVel, vec4_scale(m_WalkingDir, WalkAcc * g_TimeDelta))
If m_TouchingGround Then
    Dim Vel_Y As Single
    Vel_Y = m_PlayerVel.Y
    m_PlayerVel = vec4_scale(m_PlayerVel, LimitedDamping(vec4_length(m_PlayerVel), m_MovementDamp ^ g_TimeDelta, m_MovementDampLimit * g_TimeDelta))
    If Vel_Y >= 0 Then m_PlayerVel.Y = Vel_Y
End If
End Sub

Private Sub ProcStandingCrouching()
m_Crouching = frmMain.GetKeyState(vbKeyControl)
Dim LerpS As Single
Dim PrevHeight As Single
PrevHeight = m_PlayerHeight
If m_Crouching Then
    m_PlayerHeight = max(m_PlayerHeight + Sgn(m_NeelingHeight - m_PlayerHeight) * m_KneelDownSpeed * g_TimeDelta, m_NeelingHeight)
Else
    m_PlayerHeight = min(m_PlayerHeight + Sgn(m_StandingHeight - m_PlayerHeight) * m_StandUpSpeed * g_TimeDelta, m_StandingHeight)
End If

LerpS = (m_PlayerHeight - m_NeelingHeight) / (m_StandingHeight - m_NeelingHeight)
m_PlayerHeightMin = lerp(m_NeelingHeightMin, m_StandingHeightMin, LerpS)
m_PlayerHeightMax = lerp(m_NeelingHeightMax, m_StandingHeightMax, LerpS)
m_BobbingRangeX = lerp(m_BobbingRangeXNeeling, m_BobbingRangeXStanding, LerpS)
m_BobbingRangeY = lerp(m_BobbingRangeYNeeling, m_BobbingRangeYStanding, LerpS)
m_StepLength = lerp(m_NeelingStepLength, m_StandingStepLength, LerpS)
m_MovementDamp = lerp(m_NeelingMovementDamp, m_StandingMovementDamp, LerpS)
m_MovementDampLimit = lerp(m_NeelingMovementDampLimit, m_StandingMovementDampLimit, LerpS)
m_PlayerPos.Y = m_PlayerPos.Y - (PrevHeight - m_PlayerHeight) * 0.5
End Sub

Private Sub ProcGroundInteraction()
Dim Resist As vec4_t
Dim TouchingGroundThreshold As Single
If m_Jumping Then TouchingGroundThreshold = m_GroundPos + m_StandingHeightJumping Else TouchingGroundThreshold = m_GroundPos + m_PlayerHeightMax
m_GroundPos = WalkMap_GetVal_Interpolated(m_PlayerPos.X, m_PlayerPos.Z)
If m_PlayerPos.Y < TouchingGroundThreshold Then
    m_TouchingGround = True
    m_GroundNormal = WalkNormalMap_GetVal_Interpolated(m_PlayerPos.X, m_PlayerPos.Z)
    Resist = vec4_scale(m_GroundNormal, max(-vec4_dot(m_PlayerVel, m_GroundNormal), 0))
    If m_PlayerPos.Y < m_GroundPos + m_PlayerHeightMin Then
        m_PlayerPos.Y = m_GroundPos + m_PlayerHeightMin
        m_PlayerVel = vec4_add(m_PlayerVel, Resist)
        m_Slippering = True
    Else
        'm_PlayerVel.Y = m_PlayerVel.Y - m_Gravity_Y * g_TimeDelta
        Dim TargetVel As Single
        TargetVel = (m_GroundPos + m_PlayerHeight - m_PlayerPos.Y) / m_StandUpTargetTime
        If m_GroundNormal.Y > 0.7 Then
            m_Slippering = False
            If TargetVel > m_PlayerVel.Y Then
                m_PlayerVel.Y = m_PlayerVel.Y + clamp(TargetVel - m_PlayerVel.Y, 0, m_StandUpSpeed)
            Else
                m_PlayerVel.Y = m_PlayerVel.Y - clamp(m_PlayerVel.Y - TargetVel, 0, m_StandUpSpeed)
            End If
        Else
            m_PlayerVel = vec4_add(m_PlayerVel, vec4_scale(m_GroundNormal, clamp(TargetVel - m_PlayerVel.Y, 0, m_StandUpSpeed)))
            m_Slippering = True
        End If
        If m_Jumping = False Then m_PlayerVel.Y = m_PlayerVel.Y * (m_StandNeelDamp ^ g_TimeDelta)
    End If
Else
    m_TouchingGround = False
    m_GroundNormal = vec4(0, 1, 0, 0)
End If

'If m_Slippering Then
'    frmMain.Caption = "»¬"
'Else
'    frmMain.Caption = "¹Ç"
'End If
'frmMain.Caption = bar_indicator(m_JumpCharge, 10, "ºÙ", "ºÚ", "¿Ú") & " " & m_Jumping & " " & m_JumpCharged

Dim Jmp_Dir As vec4_t

If m_TouchingGround Then
    m_JumpCharge = m_JumpCharge + g_TimeDelta / m_JumpInterval
    If m_JumpCharge < 1 Then
        If m_Jumping = False Then m_JumpCharged = False
    Else
        m_JumpCharge = 1
        m_JumpCharged = True
    End If
    If frmMain.GetKeyState(vbKeySpace) Then
        If m_GroundNormal.Y > 0.7 Then
            Jmp_Dir.Y = 1
        Else
            Jmp_Dir = m_GroundNormal
            Jmp_Dir.Y = Jmp_Dir.Y * 4
            Jmp_Dir = vec4_normalize(Jmp_Dir)
        End If
        Dim CurVel As Single, JumpAcc As Single, MaxAcc As Single
        CurVel = vec4_dot(m_PlayerVel, Jmp_Dir)
        If CurVel < m_JumpVel Then
            If m_JumpCharged Then
                m_Jumping = True
                MaxAcc = m_JumpAcceleration * g_TimeDelta
                JumpAcc = min(m_JumpVel - CurVel, MaxAcc)
                m_PlayerVel = vec4_add(m_PlayerVel, vec4_scale(Jmp_Dir, JumpAcc))
                m_JumpCharge = m_JumpCharge - JumpAcc / MaxAcc
                If m_JumpCharge < 0 Then
                    m_JumpCharge = 0
                    m_JumpCharged = False
                    m_Jumping = False
                End If
            End If
        Else
            m_JumpCharge = 0
            m_JumpCharged = False
            m_Jumping = False
        End If
    Else
        m_Jumping = False
    End If
Else
    m_Jumping = False
    If m_JumpCharge < 1 Then m_JumpCharged = False
End If
End Sub

Private Sub UpdateTimer()
Dim CurTime As Double
CurTime = g_SceneTimer.Value
g_TimeDelta = CurTime - g_TimeUpdate
g_TimeUpdate = CurTime
If g_TimeDelta > 1 Then g_TimeDelta = 1
End Sub

Sub Scene_FrameMove()
'UpdateTimer

ProcMouseInput
ProcWalking
ProcCameraBobbing

m_PlayerVel.X = m_PlayerVel.X + m_Gravity_X * g_TimeDelta
m_PlayerVel.Y = m_PlayerVel.Y + m_Gravity_Y * g_TimeDelta
m_PlayerVel.Z = m_PlayerVel.Z + m_Gravity_Z * g_TimeDelta
If frmMain.GetKeyState(vbKeyQ) Then
    m_PlayerVel.Y = m_PlayerVel.Y - m_Gravity_Y * 1.1 * g_TimeDelta
End If
ProcStandingCrouching
ProcGroundInteraction

UpdateWalkingDirection
UpdatePosition
UpdateCamera
End Sub

Private Sub Render_Objects()
'Dim X As Single, Y As Single
'X = Cos(g_TimeUpdate * PI * 0.1) * 128
'Y = Sin(g_TimeUpdate * PI * 0.1) * 128
'Renderer_DrawParticle_Textured vec4(X, AltMap_GetVal_Interpolated(X, Y), Y, 0), Abs(Sin(g_TimeUpdate * PI) * 32 + 33), m_ScopeMask, m_ScopeXRes, m_ScopeYRes
End Sub

Private Sub Render_Alpha_Objects()
'Renderer_DrawParticle_Alpha vec4(0, AltMap_GetVal(0, 0), 0, 0), 16, Rainbow(g_TimeUpdate * 0.1) Or &H7F000000
'Renderer_DrawParticle_Addition vec4(32, AltMap_GetVal(32, 0), 0, 0), 16, Rainbow(g_TimeUpdate * 0.1)
End Sub

Sub Scene_Render()
Renderer_ClearZ

Renderer_SetZ True, True
Renderer_SetProjCenter 0.5, 0.5
Renderer_SetCamera m_CamPos, m_CamOrient, m_CamFovY

Dim X As Long, Y As Long

If m_ScopeOpened Then
    Dim ScopeXCenter As Single, ScopeYCenter As Single
    Dim ScopeX As Long, ScopeY As Long
    ScopeXCenter = m_ScopeXCenter + m_BobbingYaw * m_BobbingScopeShakeRangeX / m_BobbingYawRange
    ScopeYCenter = m_ScopeYCenter + m_BobbingPitch * m_BobbingScopeShakeRangeY / m_BobbingYawRange
    ScopeX = clamp(ScopeXCenter - m_ScopeSize * 0.5) * m_VoxelCanvasXRes
    ScopeY = clamp(ScopeYCenter - m_ScopeSize * 0.5) * m_VoxelCanvasYRes
    
    Renderer_MaskBmp_Setup ScopeX, ScopeY, m_ScopeXRes, m_ScopeYRes, m_ScopeMask, 0, True
    Render_Objects
    Renderer_RenderLandscape 0, 0, m_VoxelCanvasXRes - 1, m_VoxelCanvasYRes - 1
    Render_Alpha_Objects
    
    Renderer_SetProjCenter ScopeXCenter, ScopeYCenter
    Renderer_SetCamera m_CamPos, m_CamOrient, m_ScopeBaseFov / m_ScopeMult
    Renderer_MaskBmp_Invert
    Render_Objects
    Renderer_RenderLandscape ScopeX, ScopeY, ScopeX + m_ScopeXRes, ScopeY + m_ScopeYRes, 32, m_ScopeMult * 256
    Render_Alpha_Objects
    Renderer_MaskBmp_Disable
Else
    Render_Objects
    Renderer_RenderLandscape 0, 0, m_VoxelCanvasXRes - 1, m_VoxelCanvasYRes - 1
    Render_Alpha_Objects
End If

End Sub

Sub Scene_Present(Target As PictureBox)
Renderer_Present Target
End Sub

Function Scene_FrameMove_CheckTick(Optional ByVal TickLen As Double = 1 / 30, Optional ByVal MaxTickLen As Double = 1) As Boolean
Dim CurTime As Double, Interval As Double
CurTime = g_SceneTimer.Value
Interval = CurTime - g_TimeUpdate

If Interval >= TickLen Then
    If Interval < MaxTickLen Then
        Do While g_TimeUpdate + TickLen <= CurTime
            g_TimeDelta = TickLen
            g_TimeUpdate = g_TimeUpdate + TickLen
            Scene_FrameMove
            Scene_FrameMove_CheckTick = True
        Loop
    Else
        g_TimeDelta = MaxTickLen
        g_TimeUpdate = CurTime
    End If
End If
End Function

Sub Scene_Cleanup()
Set g_SceneTimer = Nothing
Renderer_Cleanup
Map_Cleanup
End Sub
