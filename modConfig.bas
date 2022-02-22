Attribute VB_Name = "modConfig"
Option Explicit

Global g_Arg_LowRes As Boolean
Global g_Arg_HighRes As Boolean
Global g_Arg_CfgFilePath As String

Global g_Cfg_Render_XRes As Long
Global g_Cfg_Render_YRes As Long
Global g_Cfg_Render_VoxelPix_W As Long
Global g_Cfg_Render_VoxelPix_H As Long
Global g_Cfg_Render_Interpolate As Long
Global g_Cfg_Render_Interleaved As Long
Global g_Cfg_Render_Blending As Long
Global g_Cfg_Render_Windowed As Long
Global g_Cfg_Game_Tick_Lock As Long
Global g_Cfg_Game_Tick_Freq As Long
Global g_Cfg_Game_Tick_Skip As Long

Sub Config_Load()
g_Arg_CfgFilePath = Args_Query("-cfg")
If Len(g_Arg_CfgFilePath) = 0 Then g_Arg_CfgFilePath = App.Path & "\cfg.ini"

MetaData_Load g_Arg_CfgFilePath
g_Cfg_Render_XRes = MetaData_Query("[render]", "xres", "640")
g_Cfg_Render_YRes = MetaData_Query("[render]", "yres", "480")
g_Cfg_Render_VoxelPix_W = MetaData_Query("[render]", "voxelpix_w", "4")
g_Cfg_Render_VoxelPix_H = MetaData_Query("[render]", "voxelpix_h", "2")
g_Cfg_Render_Interpolate = MetaData_Query("[render]", "interpolate", "1")
g_Cfg_Render_Interleaved = MetaData_Query("[render]", "interleaved", "0")
g_Cfg_Render_Blending = MetaData_Query("[render]", "blending", "1")
g_Cfg_Render_Windowed = MetaData_Query("[render]", "windowed", "1")

g_Cfg_Game_Tick_Lock = MetaData_Query("[game]", "tick_lock", "1")
g_Cfg_Game_Tick_Freq = MetaData_Query("[game]", "tick_freq", "30")
g_Cfg_Game_Tick_Skip = MetaData_Query("[game]", "tick_skip", "30")

End Sub
