Attribute VB_Name = "modMTPool"
Option Explicit

Private Type ThreadParams
    FuncPtr As Long
    Param1 As Long
    Param2 As Long
    Param3 As Long
    Param4 As Long
End Type

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type SYSTEM_INFO
    dwOemID As Long
    dwPageSize As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask As Long
    dwNumberOrfProcessors As Long
    dwProcessorType As Long
    dwAllocationGranularity As Long
    dwReserved As Long
End Type

Private Declare Sub GetSystemInfo Lib "KERNEL32" (lpSystemInfo As SYSTEM_INFO)

Private MT_Works() As ThreadParams
Private MT_WorksToDo As Long
Private MT_WorksFinished As Long
Private MT_Status As Long
Private MT_NumWorkers As Long
Private MT_ThreadHandles() As Long

Function GetNumProcessors() As Long
Dim SI As SYSTEM_INFO
GetSystemInfo SI
GetNumProcessors = SI.dwNumberOrfProcessors
End Function

Private Function ThreadPoolProc(ByVal Params As Long) As Long
ThreadInit

Do
    If MT_Status = 1 Then
        Dim Work As Long
        If MT_WorksToDo > 0 Then
            Work = InterlockedDecrement(MT_WorksToDo)
            If Work >= 0 Then
                CallWindowProc MT_Works(Work).FuncPtr, MT_Works(Work).Param1, MT_Works(Work).Param2, MT_Works(Work).Param3, MT_Works(Work).Param4
                InterlockedIncrement MT_WorksFinished
            Else
                Sleep 0
            End If
        Else
            Sleep 0
        End If
    Else
        Sleep 0
    End If
Loop While MT_Status >= 0

ThreadQuit
ExitThread 0
End Function

Sub MT_Init()
'ʹʱ��Ƭ���Ⱦ������
timeBeginPeriod 0

'�ж���Ҫ���߳���
MT_NumWorkers = GetNumProcessors

'׼�������߳�
ReDim MT_ThreadHandles(MT_NumWorkers - 1)

'���������߳�
Dim I As Long, TID As Long
For I = 0 To MT_NumWorkers - 1
    MT_ThreadHandles(I) = CreateThread(ByVal 0, 0, AddressOf ThreadPoolProc, I, 0, TID)
Next
End Sub

Sub MT_Terminate()
InterlockedExchange MT_Status, -1
WaitForMultipleObjects MT_NumWorkers, MT_ThreadHandles(0), 1, -1

timeEndPeriod 0
End Sub

'�������ö��߳�����ѭ��
'������
'FunctionPointer������ָ�룬�ú��������ĸ� Long ����������Լ���� stdcall
'  ѭ��������������һ�����������������ָ��
'WorkSetSize���������Ĵ�С��Ҳ����ѭ���Ĵ���
Sub MT_RunForLoop(ByVal FunctionPointer As Long, ByVal WorkSetSize As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long)
If WorkSetSize = 0 Then Exit Sub

ReDim Preserve MT_Works(WorkSetSize - 1)

Dim I As Long
For I = 0 To WorkSetSize - 1
    MT_Works(I).FuncPtr = FunctionPointer
    MT_Works(I).Param1 = I
    MT_Works(I).Param2 = Param2
    MT_Works(I).Param3 = Param3
    MT_Works(I).Param4 = Param4
Next

InterlockedExchange MT_WorksToDo, WorkSetSize
InterlockedExchange MT_WorksFinished, 0
InterlockedExchange MT_Status, 1

Do While MT_WorksFinished < WorkSetSize
    Sleep 1
Loop

InterlockedExchange MT_Status, 0
End Sub
