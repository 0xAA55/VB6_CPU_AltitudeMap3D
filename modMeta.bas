Attribute VB_Name = "modMeta"
Option Explicit

Type MetaData_t
    MetaName As String
    NumKeyPairs As Long
    Keys() As String
    Values() As String
End Type

Global g_MetaDataArr() As MetaData_t
Global g_NumMetaData As Long
Global Const g_DefMetaName As String = "[global]"

Sub MetaData_Load(Path As String)
Erase g_MetaDataArr

Dim MetaArr() As MetaData_t
Dim NumMeta As Long
Dim MaxMeta As Long
Dim CurMeta As Long

Const MetaAlloc As Long = 16

Dim Keys() As String
Dim Values() As String
Dim NumKeyPairs As Long
Dim MaxKeyPairs As Long
Const KeyPairsAlloc As Long = 32

Dim LineRead As String
Dim EqualMark As Long

Open Path For Input As #1
CurMeta = -1
GoSub AllocMeta
MetaArr(0).MetaName = g_DefMetaName
Do While Not EOF(1)

    Line Input #1, LineRead

    If Len(LineRead) = 0 Then GoTo Continue
    LineRead = Trim$(Split(LineRead, ";")(0))
    If Len(LineRead) = 0 Then GoTo Continue
    
    If Left$(LineRead, 1) = "[" Then
        If CurMeta < 0 Then
            MetaArr(0).MetaName = LineRead
            CurMeta = 0
        Else
            CurMeta = NumMeta - 1
            GoSub StoreMetaKeyPairs
            GoSub AllocMeta
            CurMeta = NumMeta - 1
            MetaArr(CurMeta).MetaName = LineRead
        End If
    Else
        EqualMark = InStr(LineRead, "=")
        If EqualMark Then
            
            If NumKeyPairs >= MaxKeyPairs Then
                If MaxKeyPairs Then
                    MaxKeyPairs = NumKeyPairs + KeyPairsAlloc
                    ReDim Preserve Keys(MaxKeyPairs - 1)
                    ReDim Preserve Values(MaxKeyPairs - 1)
                Else
                    MaxKeyPairs = NumKeyPairs + KeyPairsAlloc
                    ReDim Keys(MaxKeyPairs - 1)
                    ReDim Values(MaxKeyPairs - 1)
                End If
            End If
            
            Keys(NumKeyPairs) = Left$(LineRead, EqualMark - 1)
            Values(NumKeyPairs) = Mid$(LineRead, EqualMark + 1)
            NumKeyPairs = NumKeyPairs + 1
        End If
    End If

Continue:
Loop
Close #1

If NumMeta Then
    GoSub StoreMetaKeyPairs
    ReDim Preserve MetaArr(NumMeta - 1)
    g_MetaDataArr = MetaArr
    Erase MetaArr
    g_NumMetaData = NumMeta
End If

Exit Sub

AllocMeta:
    If NumMeta >= MaxMeta Then
        If MaxMeta Then
            MaxMeta = NumMeta + MetaAlloc
            ReDim Preserve MetaArr(MaxMeta - 1)
        Else
            MaxMeta = NumMeta + MetaAlloc
            ReDim MetaArr(MaxMeta - 1)
        End If
    End If
    NumMeta = NumMeta + 1
Return

StoreMetaKeyPairs:
    MetaArr(CurMeta).NumKeyPairs = NumKeyPairs
    If NumKeyPairs Then
        ReDim Preserve Keys(NumKeyPairs - 1)
        ReDim Preserve Values(NumKeyPairs - 1)
        MetaArr(CurMeta).Keys = Keys
        MetaArr(CurMeta).Values = Values
        NumKeyPairs = 0
        MaxKeyPairs = 0
        Erase Keys, Values
    End If
Return

End Sub

Function MetaData_Query(MetaName As String, Key As String, Optional Default As String) As String
Dim I As Long, J As Long

MetaData_Query = Default
For I = 0 To g_NumMetaData - 1
    If UCase$(g_MetaDataArr(I).MetaName) = UCase$(MetaName) Then
        For J = 0 To g_MetaDataArr(I).NumKeyPairs - 1
            If UCase$(g_MetaDataArr(I).Keys(J)) = UCase$(Key) Then
                MetaData_Query = g_MetaDataArr(I).Values(J)
                Exit Function
            End If
        Next
        Exit Function
    End If
Next
End Function
