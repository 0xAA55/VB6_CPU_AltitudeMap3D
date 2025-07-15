VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAVIWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type AVIMainHeader_t
    FourCC As Long
    cb As Long
    MicroSecPerFrame As Long
    MaxBytesPerSec As Long
    PaddingGranularity As Long
    Flags As Long
    TotalFrames As Long
    InitialFrames As Long
    Streams As Long
    SuggestedBufferSize As Long
    Width As Long
    Height As Long
    Reserved(3) As Long
End Type

Private Type AVIStreamHeader_t
    FourCC As Long
    Handler As Long
    Flags As Long
    Priority As Integer
    Language As Integer
    InitialFrames As Long
    Scale_ As Long
    Rate As Long
    Start As Long
    Length As Long
    SuggestedBufferSize As Long
    Quality As Long
    SampleSize As Long
    Frame(3) As Integer
End Type

Private Type VideoFormat_t
    BMIF As BitmapInfoHeader_t
    PaletteOrBitfields(255) As Long
End Type

Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, pBitmapInfo As Any, ByVal Usage As Long, PointerToReceive As Long, ByVal hSection As Long, ByVal Offset As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hWnd As Long, ByVal hDC As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private m_FileNumber As Integer
Private m_FilePath As String
Private m_Width As Long
Private m_Height As Long
Private m_VideoFormat As VideoFormat_t
Private m_VideoFormatIsSet As Boolean
Private m_VideoFormatLength As Long
Private m_AudioFormat As WaveFormatEx_t
Private m_AudioFormatIsSet As Boolean
Private m_OffsetToAVIH As Long
Private m_OffsetToSTRH_Video As Long
Private m_OffsetToSTRH_Audio As Long
Private m_OffsetToFirstFrame As Long
Private m_FrameRate As Long
Private m_FrameSize As Long
Private m_VideoWritten As Long
Private m_AudioWritten As Long
Private m_IsBegun As Boolean
Private m_CaptureSourceControl As Long
Private m_CaptureSourceHDC As Long
Private m_CapturerHDC As Long
Private m_CapturerFrameBufferPtr As Long

Public Property Get FrameSize() As Long
FrameSize = m_FrameSize
End Property

Public Property Get FrameRate() As Long
FrameRate = m_FrameRate
End Property

Public Property Let FrameRate(ByVal NewFrameRate As Long)
If m_IsBegun Then Err.Raise 5, "Property Let FrameRate()", "Must set framerate before calling `BeginWrite()`"

m_FrameRate = NewFrameRate
End Property

Public Property Get VideoWidth() As Long
VideoWidth = m_Width
End Property

Public Property Get VideoHeight() As Long
VideoHeight = m_Height
End Property

Public Property Get CaptureSourceHWND() As Long
CaptureSourceHWND = m_CaptureSourceControl
End Property

Public Property Get CaptureSourceHDC() As Long
CaptureSourceHDC = m_CaptureSourceHDC
End Property

Public Sub SetVideoFormat(ByVal VideoFormatPtr As Long, ByVal VideoFormatSize As Long)
If VideoFormatSize > Len(m_VideoFormat) Then Err.Raise 5, "SetVideoFormat()", "`VideoFormatSize` is too long"
VideoFormatSize = min(VideoFormatSize, Len(m_VideoFormat))

If VideoFormatPtr > 0 And VideoFormatSize > 0 Then
    CopyMemory m_VideoFormat, ByVal VideoFormatPtr, VideoFormatSize
    m_VideoFormatLength = VideoFormatSize
    
    '进行一些预计算，算出一些参数
    m_Width = m_VideoFormat.BMIF.biWidth
    m_Height = m_VideoFormat.BMIF.biHeight
    If m_Height < 0 Then m_Height = -m_Height
    Dim Pitch As Long
    Pitch = ((m_Width * 24 - 1) \ 32 + 1) * 4
    m_FrameSize = Pitch * m_Height
    m_VideoFormatIsSet = True
Else
    m_VideoFormatIsSet = False
End If
End Sub

Public Sub SetAudioFormat(ByVal AudioFormatPtr As Long, ByVal AudioFormatSize As Long)
If AudioFormatSize > Len(m_AudioFormat) Then Err.Raise 5, "SetAudioFormat()", "`AudioFormatSize` is too long"
AudioFormatSize = min(AudioFormatSize, Len(m_AudioFormat))

If AudioFormatPtr > 0 And AudioFormatSize > 0 Then
    CopyMemory m_AudioFormat, ByVal AudioFormatPtr, AudioFormatSize
    m_AudioFormatIsSet = True
Else
    m_AudioFormatIsSet = False
End If
End Sub

Public Sub SetCaptureSource(hWnd As Long, hDC As Long)
'这个函数可以重复调用。
'hDC 可以是与 hWnd 无关的 HDC，也可以是 0，如果是 0 则会长期占用针对 hWnd 的 HDC 引用。
If m_VideoFormatIsSet = False Then Err.Raise 5, "SetCaptureSource()", "Must call `SetVideoFormat()` first"

If hWnd = 0 Then
    '调用者传递 hWnd = 0 意思是想要清除捕获来源。
    m_CaptureSourceControl = 0
    m_CaptureSourceHDC = 0
    If m_CapturerHDC Then
        DeleteDC m_CapturerHDC
        m_CapturerHDC = 0
    End If
    m_CapturerFrameBufferPtr = 0
    Exit Sub
End If

m_CaptureSourceControl = hWnd

Dim CaptureSourceGetDC As Long
CaptureSourceGetDC = GetDC(hWnd)
m_CaptureSourceHDC = hDC

If m_CapturerHDC = 0 Then m_CapturerHDC = CreateCompatibleDC(CaptureSourceGetDC)

Dim CaptureFrameBufferBitmap As Long
CaptureFrameBufferBitmap = CreateDIBSection(CaptureSourceGetDC, m_VideoFormat, 0, m_CapturerFrameBufferPtr, 0, 0)
ReleaseDC hWnd, CaptureSourceGetDC

SelectObject m_CapturerHDC, CaptureFrameBufferBitmap
DeleteObject CaptureFrameBufferBitmap
End Sub

Public Sub BeginWrite(AVIFilePath As String)
If m_VideoFormatIsSet = False Then Err.Raise 5, "BeginWrite()", "Must call `SetVideoFormat()` before calling `BeginWrite()`"
If m_FrameRate = 0 Then Err.Raise 5, "BeginWrite()", "Must set `FrameRate` before calling `BeginWrite()`"

m_FileNumber = FreeFile
m_FilePath = AVIFilePath

Dim ListSize_HDRL As Long
Dim ListSize_STRL_V As Long
Dim ListSize_STRL_A As Long

Dim AVIHeader As AVIMainHeader_t
Dim StreamHeader As AVIStreamHeader_t

ListSize_STRL_V = 4 + 8 + Len(StreamHeader) + 8 + m_VideoFormatLength
ListSize_HDRL = 4 + Len(AVIHeader) + 8 + 4 + ListSize_STRL_V
If m_AudioFormatIsSet Then
    ListSize_STRL_A = 4 + 8 + Len(StreamHeader) + 8 + Len(m_AudioFormat)
    ListSize_HDRL = ListSize_HDRL + ListSize_STRL_A
End If

'先用文字输出的方式打开文件，会导致文件被清空
Open m_FilePath For Output As #m_FileNumber
Close #m_FileNumber

'然后再用二进制的方式打开文件
Open m_FilePath For Binary Access Write As #m_FileNumber
Put #m_FileNumber, , "RIFF"
Put #m_FileNumber, , 0&
Put #m_FileNumber, , "AVI LIST"
Put #m_FileNumber, , ListSize_HDRL
Put #m_FileNumber, , "hdrl"
m_OffsetToAVIH = Seek(m_FileNumber)
Put #m_FileNumber, , AVIHeader
Put #m_FileNumber, , "LIST"
Put #m_FileNumber, , ListSize_STRL_V
Put #m_FileNumber, , "strlstrh"
Put #m_FileNumber, , CLng(Len(StreamHeader))
m_OffsetToSTRH_Video = Seek(m_FileNumber)
Put #m_FileNumber, , StreamHeader
Put #m_FileNumber, , "strf"
Put #m_FileNumber, , m_VideoFormatLength
Dim Buffer() As Byte
ReDim Buffer(m_VideoFormatLength - 1)
CopyMemory Buffer(0), m_VideoFormat, m_VideoFormatLength
Put #m_FileNumber, , Buffer
If m_AudioFormatIsSet Then
    Put #m_FileNumber, , "LIST"
    Put #m_FileNumber, , ListSize_STRL_A
    Put #m_FileNumber, , "strlstrh"
    Put #m_FileNumber, , CLng(Len(StreamHeader))
    m_OffsetToSTRH_Audio = Seek(m_FileNumber)
    Put #m_FileNumber, , StreamHeader
    Put #m_FileNumber, , "strf"
    Put #m_FileNumber, , CLng(Len(m_AudioFormat))
    Put #m_FileNumber, , m_AudioFormat
End If
Put #m_FileNumber, , "LIST"
Put #m_FileNumber, , 0&
Put #m_FileNumber, , "movi"
m_OffsetToFirstFrame = Seek(m_FileNumber)

m_IsBegun = True
End Sub

Public Sub EndWrite()
If m_IsBegun = False Then Err.Raise 5, "EndWrite()", "Must call `BeginWrite()` first"

Dim AVIH As AVIMainHeader_t
Dim STRH_Video As AVIStreamHeader_t
Dim STRH_Audio As AVIStreamHeader_t

AVIH.FourCC = &H68697661
AVIH.cb = Len(AVIH) - 8
AVIH.MicroSecPerFrame = 1000000 \ m_FrameRate
AVIH.MaxBytesPerSec = m_FrameSize * m_FrameRate + m_AudioFormat.AvgBytesPerSec
AVIH.Flags = IIf(m_AudioFormatIsSet, &H100&, 0)
AVIH.Streams = IIf(m_AudioFormatIsSet, 2, 1)
AVIH.SuggestedBufferSize = m_FrameSize
AVIH.Width = m_Width
AVIH.Height = m_Height

STRH_Video.FourCC = &H73646976
STRH_Video.InitialFrames = AVIH.InitialFrames
STRH_Video.Scale_ = 1
STRH_Video.Rate = m_FrameRate
STRH_Video.Length = m_VideoWritten
STRH_Video.SuggestedBufferSize = m_FrameSize
STRH_Video.Quality = -1
STRH_Video.Frame(0) = 0
STRH_Video.Frame(1) = 0
STRH_Video.Frame(2) = m_Width
STRH_Video.Frame(3) = m_Height

If m_AudioFormatIsSet Then
    STRH_Audio.FourCC = &H73647561
    STRH_Audio.Handler = m_AudioFormat.FormatTag
    STRH_Audio.Scale_ = 1
    STRH_Audio.Rate = m_AudioFormat.SamplesPerSec
    STRH_Audio.Length = m_AudioWritten \ m_AudioFormat.BlockAlign
    STRH_Audio.SuggestedBufferSize = m_AudioFormat.AvgBytesPerSec \ 10
    STRH_Audio.Quality = -1
    STRH_Audio.SampleSize = m_AudioFormat.BlockAlign
End If

Dim CurPos As Long
CurPos = Seek(m_FileNumber)

Dim FileSize As Long
FileSize = CurPos - 1

Put #m_FileNumber, 5, FileSize - 8
Put #m_FileNumber, m_OffsetToAVIH, AVIH
Put #m_FileNumber, m_OffsetToSTRH_Video, STRH_Video
If m_AudioFormatIsSet Then Put #m_FileNumber, m_OffsetToSTRH_Audio, STRH_Audio
Put #m_FileNumber, m_OffsetToFirstFrame - 8, FileSize - m_OffsetToFirstFrame
Close #m_FileNumber

m_IsBegun = False
End Sub

Public Sub WriteVideo(Optional ByVal VideoDataPtr As Long = 0)
Dim DoCapture As Boolean

If VideoDataPtr = 0 Then
    VideoDataPtr = m_CapturerFrameBufferPtr
    DoCapture = True
    If VideoDataPtr = 0 Then Err.Raise 5, "WriteVideo()", "Must give a valid video data pointer or setup a capture source control"
End If

If DoCapture Then
    Dim SourceHDC As Long
    If m_CaptureSourceHDC Then SourceHDC = m_CaptureSourceHDC Else SourceHDC = GetDC(m_CaptureSourceControl)
    BitBlt m_CapturerHDC, 0, 0, m_Width, m_Height, SourceHDC, 0, 0, vbSrcCopy
    If m_CaptureSourceHDC = 0 Then ReleaseDC m_CaptureSourceControl, SourceHDC
End If

Dim VideoData() As Byte
ReDim VideoData(m_FrameSize - 1)
CopyMemory VideoData(0), ByVal VideoDataPtr, m_FrameSize

Put #m_FileNumber, , "00dc"
Put #m_FileNumber, , m_FrameSize
Put #m_FileNumber, , VideoData

m_VideoWritten = m_VideoWritten + 1
If m_VideoWritten Mod m_FrameRate = 0 Then
    '通过关闭文件再重新打开的方式来把内存缓存内容写入文件。
    Dim CurPos As Long
    CurPos = Seek(m_FileNumber)
    Close #m_FileNumber
    Open m_FilePath For Binary Access Write As #m_FileNumber
    Seek #m_FileNumber, CurPos
End If
End Sub

Public Sub WriteAudio(ByVal AudioDataPtr As Long, ByVal AudioDataLength As Long)
If AudioDataPtr = 0 Or AudioDataLength <= 0 Then Err.Raise 5, "WriteAudio()", "Invalid parameters for `WriteAudio()`"

Dim AudioData() As Byte
ReDim AudioData(AudioDataLength - 1)

CopyMemory AudioData(0), ByVal AudioDataPtr, AudioDataLength

Put #m_FileNumber, , "01wb"
Put #m_FileNumber, , AudioDataLength
Put #m_FileNumber, , AudioData

m_AudioWritten = m_AudioWritten + AudioDataLength
End Sub

Public Property Get VideoTimestamp() As Double
VideoTimestamp = m_VideoWritten / m_FrameRate
End Property

Public Property Get AudioTimestamp() As Double
AudioTimestamp = m_AudioWritten / m_AudioFormat.AvgBytesPerSec
End Property

Public Property Get HaveAudio() As Boolean
HaveAudio = m_AudioFormatIsSet
End Property

Private Sub Class_Terminate()
If m_IsBegun Then EndWrite
If m_CapturerHDC Then DeleteDC m_CapturerHDC
End Sub
