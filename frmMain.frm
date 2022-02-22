VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   7200
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9600
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7200
   ScaleWidth      =   9600
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.PictureBox picCanvas 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      FillColor       =   &H00C0C0C0&
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   3000
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   161
      TabIndex        =   0
      Top             =   1200
      Visible         =   0   'False
      Width           =   2415
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const m_MinimapWidth As Long = 64
Private Const m_MinimapHeight As Long = 64
Private Const m_MinimapScale As Long = 2

#Const UseMessageKeyDetect = 0

#If UseMessageKeyDetect Then
Private m_Key(255) As Boolean
Private m_KeyDown(255) As Boolean

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next
m_Key(KeyCode) = True
m_KeyDown(KeyCode) = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
On Error Resume Next
m_KeyDown(KeyCode) = False
End Sub

Function GetKeyState(ByVal KeyCode As Long) As Boolean
GetKeyState = m_Key(KeyCode)
m_Key(KeyCode) = m_KeyDown(KeyCode)
End Function
#Else 'UseMessageKeyDetect

Function GetKeyState(ByVal KeyCode As Long) As Boolean
If GetForegroundWindow = hWnd Then GetKeyState = GetAsyncKeyState(KeyCode)
End Function

#End If 'UseMessageKeyDetect

Private Sub ParseCmdArgs()
ParseArgs

g_Arg_LowRes = CBool(Args_Search("-lowres"))
g_Arg_HighRes = CBool(Args_Search("-highres"))
End Sub

Private Sub AdjustWindowPos()
If g_Cfg_Render_Windowed Then
    Dim BorderWidth As Single, BorderHeight As Single
    BorderWidth = Width - ScaleX(ScaleWidth, ScaleMode, vbTwips)
    BorderHeight = Height - ScaleY(ScaleHeight, ScaleMode, vbTwips)
    Dim TargetWidth As Single, TargetHeight As Single
    TargetWidth = g_Cfg_Render_XRes * Screen.TwipsPerPixelX + BorderWidth
    TargetHeight = g_Cfg_Render_YRes * Screen.TwipsPerPixelY + BorderHeight
    
    Move _
        Screen.Width / 2 - TargetWidth / 2, _
        Screen.Height / 2 - TargetHeight / 2, _
        TargetWidth, TargetHeight
    
Else
    
End If
End Sub

Private Sub Form_Load()
ParseCmdArgs
Config_Load
AdjustWindowPos

Show
Randomize Timer
Set g_ProgressCallback = Me

Dim VXRes As Long, VYRes As Long
VXRes = g_Cfg_Render_XRes \ g_Cfg_Render_VoxelPix_W
VYRes = g_Cfg_Render_YRes \ g_Cfg_Render_VoxelPix_H

If g_Arg_LowRes Then
    VXRes = VXRes \ 2
    VYRes = VYRes \ 2
End If
If g_Arg_HighRes Then
    'VXRes = VXRes * 2
    VYRes = VYRes * 2
End If

Scene_Init picCanvas, g_Cfg_Render_XRes, g_Cfg_Render_YRes, App.Path & "\maps\" & CStr(Int(Rnd * 8) + 1), VXRes, VYRes

GetKeyState vbKeyEscape
Do
    Scene_FrameMove_CheckTick
    
    picCanvas.Cls
    Scene_Render
    
    Scene_Present picCanvas

    picCanvas.PSet (0, 0)
    picCanvas.Print Tag
    picCanvas.Refresh
    If GetKeyState(vbKeyEscape) Then Unload Me
Loop While DoEvents
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
#If UseMessageKeyDetect Then
If (Button And 1) = 1 Then
    m_Key(1) = True
    m_KeyDown(1) = True
End If
If (Button And 2) = 2 Then
    m_Key(2) = True
    m_KeyDown(2) = True
End If
If (Button And 4) = 4 Then
    m_Key(3) = True
    m_KeyDown(3) = True
End If
#End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
#If UseMessageKeyDetect Then
If (Button And 1) = 1 Then m_KeyDown(1) = False
If (Button And 2) = 2 Then m_KeyDown(2) = False
If (Button And 4) = 4 Then m_KeyDown(3) = False
#End If
End Sub

Private Sub Form_Resize()
picCanvas.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
Scene_Cleanup
End Sub

Public Sub OnStartProgress(ProgressName As String)
Const ProgBarWidth As Single = 2400
Const ProgBarHeight As Single = 120
Const ProgBarHalfWidth As Single = ProgBarWidth / 2
Const ProgBarHalfHeight As Single = ProgBarHeight / 2
Dim CenX As Single, CenY As Single
CenX = ScaleWidth / 2
CenY = ScaleHeight / 2

Dim MsgW As Single, MsgH As Single
MsgW = TextWidth(ProgressName)
MsgH = TextHeight(ProgressName)

Line (0, 0)-(ScaleWidth, ScaleHeight), vbBlack, BF
CurrentX = CenX - MsgW / 2
CurrentY = CenY - ProgBarHalfHeight - MsgH - 2 * Screen.TwipsPerPixelY
ForeColor = vbGreen
Print ProgressName

Line (CenX - ProgBarHalfWidth - 2 * Screen.TwipsPerPixelX, CenY - ProgBarHalfHeight - 2 * Screen.TwipsPerPixelY)-(CenX + ProgBarHalfWidth + 2 * Screen.TwipsPerPixelX, CenY + ProgBarHalfHeight + 2 * Screen.TwipsPerPixelY), vbGreen, B
End Sub

Public Sub OnProgress(ByVal Progress As Single, ByVal TotalProgress As Single)
Const ProgBarWidth As Single = 2400
Const ProgBarHeight As Single = 120
Const ProgBarHalfWidth As Single = ProgBarWidth / 2
Const ProgBarHalfHeight As Single = ProgBarHeight / 2

Static LastTime As Single
Dim CurTime As Single
CurTime = Timer
If CurTime - LastTime > 0.1 Then
    LastTime = CurTime
    
    Dim CenX As Single, CenY As Single
    CenX = ScaleWidth / 2
    CenY = ScaleHeight / 2
    
    Line (CenX - ProgBarHalfWidth, CenY - ProgBarHalfHeight)-(CenX - ProgBarHalfWidth + Progress * ProgBarWidth / TotalProgress, CenY + ProgBarHalfHeight), vbGreen, BF
    
    If DoEvents = 0 Then End
End If
End Sub

