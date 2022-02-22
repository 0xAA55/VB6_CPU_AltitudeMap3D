Attribute VB_Name = "modBmpFile"
Option Explicit

Function Bmp_ReadColors(BmpFilePath As String, XRes As Long, YRes As Long, Pixels() As Long) As Boolean
Dim BMFH As BitmapFileHeader_t
Dim BMIF As BitmapInfoHeader_t
Dim X As Long, Y As Long, PalColors As Long, I As Long

Open BmpFilePath For Binary Access Read As #1

Get #1, 1, BMFH
If BMFH.bfType <> &H4D42 Then GoTo End_Return

Get #1, , BMIF
If BMIF.biCompression <> 0 Then GoTo End_Return

If XRes <> 0 And XRes <> BMIF.biWidth Then GoTo End_Return
If YRes <> 0 And YRes <> Abs(BMIF.biHeight) Then GoTo End_Return
XRes = BMIF.biWidth
YRes = Abs(BMIF.biHeight)
Dim MaxX As Long, MaxY As Long
MaxX = XRes - 1
MaxY = YRes - 1
ReDim Pixels(MaxX, MaxY)

Dim BPP As Long
Dim Pal() As Long
BPP = BMIF.biBitCount
If BPP <> 8 And BPP <> 24 And BPP <> 32 Then GoTo End_Return
If BPP <= 8 Then
    PalColors = BMIF.biClrUsed
    If PalColors = 0 Then PalColors = 2 ^ BPP
    ReDim Pal(PalColors - 1)
    Get #1, , Pal
    For X = 0 To PalColors - 1
        Pal(X) = ((Pal(X) And &HFF&) * &H10000) Or (Pal(X) And &HFF00&) Or (Pal(X) And &HFF0000) \ &H10000
    Next
End If

Dim LineRead() As Byte
Dim Y_Store As Long
Dim LinePtr As Long
LinePtr = 1 + BMFH.bfOffBits

If BPP = 8 Then
    ReDim LineRead(BMIF.biWidth - 1)
    For Y = 0 To MaxY
        Y_Store = IIf(BMIF.biHeight < 0, Y, MaxY - Y)
        Get #1, LinePtr, LineRead
        LinePtr = LinePtr + ((XRes - 1) \ 4 + 1) * 4
        For X = 0 To MaxX
            Pixels(X, Y_Store) = Pal(LineRead(X))
        Next
    Next
    Bmp_ReadColors = True
ElseIf BPP = 24 Then
    ReDim LineRead(BMIF.biWidth * 3 - 1)
    For Y = 0 To MaxY
        Y_Store = IIf(BMIF.biHeight < 0, Y, MaxY - Y)
        Get #1, LinePtr, LineRead
        LinePtr = LinePtr + ((XRes * 3 - 1) \ 4 + 1) * 4
        I = 0
        For X = 0 To MaxX
            Pixels(X, Y_Store) = RGB(LineRead(I + 2), LineRead(I + 1), LineRead(I + 0))
            I = I + 3
        Next
    Next
    Bmp_ReadColors = True
ElseIf BPP = 32 Then
    ReDim LineRead(BMIF.biWidth * 4 - 1)
    For Y = 0 To MaxY
        Y_Store = IIf(BMIF.biHeight < 0, Y, MaxY - Y)
        Get #1, LinePtr, LineRead
        LinePtr = LinePtr + XRes * 4
        CopyMemory Pixels(0, Y_Store), LineRead(0), XRes * 4
    Next
    Bmp_ReadColors = True
Else
    '
End If
End_Return:
Close #1
End Function

Function Bmp_WriteColors(BmpFilePath As String, ByVal XRes As Long, ByVal YRes As Long, Pixels() As Long) As Boolean
On Error GoTo ErrReturn
Dim BMFH As BitmapFileHeader_t
Dim BMIF As BitmapInfoHeader_t
Dim X As Long, Y As Long, C As Long, I As Long

BMFH.bfType = &H4D42
BMFH.bfSize = Len(BMFH)
BMFH.bfOffBits = Len(BMFH) + Len(BMIF)
BMIF.biSize = Len(BMIF)
BMIF.biWidth = XRes
BMIF.biHeight = YRes
BMIF.biPlanes = 1
BMIF.biBitCount = 24

Open BmpFilePath For Binary Access Write As #1
Put #1, 1, BMFH
Put #1, , BMIF

Dim LineWrite() As Byte
ReDim LineWrite(((XRes * 3 - 1) \ 4 + 1) * 4 - 1)
Dim MaxX As Long, MaxY As Long
MaxX = XRes - 1
MaxY = YRes - 1
For Y = 0 To MaxY
    I = 0
    For X = 0 To MaxX
        C = Pixels(X, MaxY - Y)
        LineWrite(I + 0) = (C And &HFF0000) \ &H10000
        LineWrite(I + 1) = (C And &HFF00&) \ &H100&
        LineWrite(I + 2) = (C And &HFF&)
        I = I + 3
    Next
    Put #1, , LineWrite
Next
Bmp_WriteColors = True
ErrReturn:
Close #1
End Function

