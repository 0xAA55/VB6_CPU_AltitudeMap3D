Attribute VB_Name = "modMTPool"
Option Explicit

Private Type ThreadParams
    StartI As Long
    EndI As Long
    FuncPtr As Long
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

Private Declare Sub GetSystemInfo Lib "kernel32" (lpSystemInfo As SYSTEM_INFO)

Function GetNumProcessors() As Long
Dim SI As SYSTEM_INFO
GetSystemInfo SI
GetNumProcessors = SI.dwNumberOrfProcessors
End Function

Private Function ThreadPoolProc(TP As ThreadParams) As Long
ThreadInit

Dim I As Long
For I = TP.StartI To TP.EndI
    CallWindowProc TP.FuncPtr, I, Param2, Param3, Param4
Next

ThreadQuit
End Function

'�������ö��߳�����ѭ��
'������
'FunctionPointer������ָ�룬�ú��������ĸ� Long ����������Լ���� stdcall
'  ѭ��������������һ�����������������ָ��
'
'WorkSetSize���������Ĵ�С��Ҳ����ѭ���Ĵ���
Sub RunMTForLoop(ByVal FunctionPointer As Long, ByVal WorkSetSize As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long)



End Sub
