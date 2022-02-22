VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function QueryPerformanceCounter Lib "KERNEL32" (lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "KERNEL32" (lpFrequency As Currency) As Long
Private m_Freq As Currency
Private m_BeginTime As Double
Private m_Paused As Boolean
Private m_PauseTime As Double
Private m_PausedTimeVal As Double

Private m_SpeedMod As Double

Private Function GetRawTime() As Double
Dim LI As Currency
QueryPerformanceCounter LI
GetRawTime = CDbl(LI) / m_Freq
End Function

Private Sub Class_Initialize()
QueryPerformanceFrequency m_Freq
m_SpeedMod = 1
m_BeginTime = GetRawTime
m_Paused = True
m_PauseTime = m_BeginTime
End Sub

Public Sub Pause()
If m_Paused Then Exit Sub

m_PauseTime = GetRawTime
m_PausedTimeVal = (m_PauseTime - m_BeginTime) * m_SpeedMod
m_Paused = True
End Sub

Public Sub Start()
If m_Paused = False Then Exit Sub

m_BeginTime = m_BeginTime + GetRawTime - m_PauseTime
m_Paused = False
End Sub

Property Get Paused() As Boolean
Paused = m_Paused
End Property

Property Let Paused(ByVal Value As Boolean)
If Value Then Pause Else Start
End Property

Property Get Value() As Double
If m_Paused Then
    Value = m_PausedTimeVal
Else
    Value = (GetRawTime - m_BeginTime) * m_SpeedMod
End If
End Property

Property Let Value(ByVal TV As Double)
If m_Paused Then
    m_PausedTimeVal = Value
Else
    m_BeginTime = GetRawTime - TV / m_SpeedMod
End If
End Property

Property Get SpeedMod() As Double
SpeedMod = m_SpeedMod
End Property

Property Let SpeedMod(ByVal Value As Double)
Dim TV As Double
TV = Value
m_SpeedMod = Value
Value = TV
End Property

