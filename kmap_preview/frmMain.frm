VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   202
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private KMap() As Single
Private KMap_Size_X As Long
Private KMap_Size_Y As Long

Private Sub Form_Load()
Dim MapFile As String
MapFile = Command
If Len(MapFile) = 0 Then MapFile = App.Path & "\kmap.bin"

MetaData_Load App.Path & "\meta.ini"

KMap_Size_X = MetaData_Query("[map]", "size_x")
KMap_Size_Y = MetaData_Query("[map]", "size_y")

ReDim KMap(KMap_Size_X - 1, KMap_Size_Y - 1)

Open MapFile For Binary As #1
Get #1, 1, KMap
Close #1

Dim X As Long, Y As Long
Dim MaxK As Single
For Y = 0 To KMap_Size_Y - 1
    For X = 0 To KMap_Size_X - 1
        If MaxK < KMap(X, Y) Then MaxK = KMap(X, Y)
    Next
Next

Caption = Format$(MaxK, "0.0000000")

Dim Col As Long
For Y = 0 To KMap_Size_Y - 1
    For X = 0 To KMap_Size_X - 1
        Col = KMap(X, Y) * 255 / MaxK
        PSet (X, Y), RGB(Col, Col, Col)
    Next
Next

End Sub
