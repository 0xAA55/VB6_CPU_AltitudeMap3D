VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCharacter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Pos As vec4_t
Private m_Ang As vec4_t
Private m_Vel As vec4_t
Private m_Orient As mat4x4_t

Private m_StandingHeight As Single
Private m_CrouchingHeight As Single

Private m_WalkingSpeed As Single
Private m_WalkingAccelTime As Single
Private m_WalkingAccel As Single
Private m_WalkingDamp As Single
Private m_WalkingMaxDecel As Single
Private m_WalkingStepLength As Single

Private m_NeelingSpeed As Single
Private m_NeelingAccelTime As Single
Private m_NeelingAccel As Single
Private m_NeelingDamp As Single
Private m_NeelingMaxDecel As Single
Private m_NeelingStepLength As Single

Private m_FloatingAccel As Single

Sub SetWalkingSpeed(ByVal Speed As Single, ByVal AccelTime As Single, ByVal MaxDecel As Single, ByVal StepLength As Single)
m_WalkingSpeed = Speed
m_WalkingAccelTime = AccelTime
m_WalkingAccel = m_WalkingSpeed * 4 / m_WalkingAccelTime
m_WalkingDamp = 1 / (2 ^ (m_WalkingAccel * 2 / m_WalkingSpeed))
m_WalkingMaxDecel = MaxDecel
m_WalkingStepLength = StepLength
End Sub

Sub SetNeelingSpeed(ByVal Speed As Single, ByVal AccelTime As Single, ByVal MaxDecel As Single, ByVal StepLength As Single)
m_NeelingSpeed = Speed
m_NeelingAccelTime = AccelTime
m_NeelingAccel = m_NeelingSpeed * 4 / m_NeelingAccelTime
m_NeelingDamp = 1 / (2 ^ (m_NeelingAccel * 2 / m_NeelingSpeed))
m_NeelingMaxDecel = MaxDecel
m_NeelingStepLength = StepLength
End Sub
