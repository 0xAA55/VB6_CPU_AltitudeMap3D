VERSION 5.00
Begin VB.Form frmMain 
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   FontTransparent =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Show

Const BatchBench As Long = 1024# ^ 2

Dim Tmr As New clsTimer
Dim I As Long
Dim Volatile As Single

Do
    Tmr.Start
    Volatile = Tmr.Value
    Tmr.Pause
    Tmr.Value = 0
    Tmr.Start
    Tmr.Value = 0
    For I = 1 To BatchBench
        Volatile = ColorLerp(vbWhite, vbBlack, Rnd)
    Next
    Tmr.Pause
    PSet (0, 0)
    Print Format$(Tmr.Value, "0.00000000000"); "s", "per "; BatchBench, " "
    Print Format$(Tmr.Value / BatchBench, "0.00000000000"); "s", "once", " "
Loop While DoEvents
End Sub
