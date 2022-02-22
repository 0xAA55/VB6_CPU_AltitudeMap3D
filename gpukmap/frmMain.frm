VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "GPU K Map Generator"
   ClientHeight    =   4785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10920
   LinkTopic       =   "Form1"
   ScaleHeight     =   319
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   728
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox txtOutput 
      Height          =   4335
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.PictureBox picAltPreview 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2895
      Left            =   7320
      ScaleHeight     =   193
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   1
      Top             =   1320
      Width           =   2895
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_MaxK As Single

Private Sub Form_Load()
If Len(Dir$(App.Path & "\gpukmap.log")) Then Kill App.Path & "\gpukmap.log"
If GPUK_ParseArgs = False Then
    Output_WriteToLog
    End
End If
If g_HideWindow = False Then Show Else Hide
GL_Init hWnd, hDC
If GPUK_Init = False Then Unload Me
If g_HideWindow = False Then DrawAltPreview
DoCalcK
DoWriteK
Output "Nothing to do."
If g_HideWindow Then Unload Me
End Sub

Sub DrawAltPreview()
Output "Drawing preview"
picAltPreview.Move 0, 0, g_Map_Size_X, g_Map_Size_Y
Dim X As Long, Y As Long, Col As Long
For Y = 0 To g_Map_Size_Y - 1
    For X = 0 To g_Map_Size_X - 1
        Col = g_Map(X, Y) * 255 / g_MaxAlt
        Col = RGB(Col, Col, Col)
        picAltPreview.PSet (X, Y), Col
    Next
Next
End Sub

Sub DoCalcK()
On Error GoTo ErrHandler
If GPUK_CalcKVal = False Then
    Output "K map generate failed."
    Exit Sub
End If
Dim X As Long, Y As Long, Col As Long
For Y = 0 To g_Map_Size_Y - 1
    For X = 0 To g_Map_Size_X - 1
        If m_MaxK < g_KMap(X, Y) Then m_MaxK = g_KMap(X, Y)
    Next
Next
If m_MaxK < 0.000001 Then
    Output "Max K value is zero?"
    Exit Sub
End If
If g_HideWindow = False Then
    For Y = 0 To g_Map_Size_Y - 1
        For X = 0 To g_Map_Size_X - 1
            Col = g_KMap(X, Y) * 255 / m_MaxK
            Col = RGB(Col, Col, Col)
            picAltPreview.PSet (X, Y), Col
        Next
    Next
End If
Exit Sub
ErrHandler:
Output "Error: " & Err.Number
Output Err.Description
End Sub

Sub DoWriteK()
Output "Writing K file"
If Len(Dir$(g_KFile)) Then
    Output "Old K file detected, delete it."
    Kill g_KFile
End If
Open g_KFile For Binary As #1
Put #1, 1, g_KMap
Close #1
Output "K file saved."
If g_WriteBmp Then
    Output "Writing Bmp preview: " & g_BmpPath
    Dim BmpData() As Long, X As Long, Y As Long, Col As Long
    ReDim BmpData(g_Map_Size_X - 1, g_Map_Size_Y - 1)
    If m_MaxK <= 0.000001 Then m_MaxK = 1
    For Y = 0 To g_Map_Size_Y - 1
        For X = 0 To g_Map_Size_X - 1
            Col = g_KMap(X, Y) * 255 / m_MaxK
            BmpData(X, Y) = RGB(Col, Col, Col)
        Next
    Next
    If Bmp_WriteColors(g_BmpPath, g_Map_Size_X, g_Map_Size_Y, BmpData) Then
        Output "Bmp preview saved."
    Else
        Output "Could not save Bmp preview."
    End If
End If
End Sub

Private Sub Form_Resize()
'txtOutput.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Form_Unload(Cancel As Integer)
GL_Term
Output_WriteToLog
End
End Sub

Sub Output(Text As String)
If Len(txtOutput.Text) > 4096 Then Output_WriteToLog

txtOutput.SelStart = Len(txtOutput.Text)
txtOutput.SelText = Text
txtOutput.SelText = vbCrLf

If g_HideWindow Then Debug.Print Text
End Sub

Sub Output_WriteToLog()
Dim txtContent As String
txtContent = txtOutput.Text
txtOutput.Text = ""
If Right$(txtContent, 2) = vbCrLf Then txtContent = Left$(txtContent, Len(txtContent) - 2)
Open App.Path & "\gpukmap.log" For Append As #1
Print #1, txtContent
Close #1
End Sub

Private Sub picAltPreview_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Static OldX!, OldY!
If Button And 1 Then
    picAltPreview.Move picAltPreview.Left + X - OldX, picAltPreview.Top + Y - OldY
Else
    OldX = X
    OldY = Y
End If
End Sub
