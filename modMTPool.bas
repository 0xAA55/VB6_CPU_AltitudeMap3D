Attribute VB_Name = "modMTPool"
Option Explicit

Private Type ThreadParams
    StartI As Long
    EndI As Long
    Param2 As Long
    Param3 As Long
    Param4 As Long
End Type

Private Function ThreadPoolProc(TP As ThreadParams) As Long
ThreadInit






ThreadQuit
End Function

'函数：用多线程运行循环
'参数：
'FunctionPointer：函数指针，该函数具有四个 Long 参数，调用约定是 stdcall
'  循环索引被当作第一个参数传入这个函数指针
'
'WorkSetSize：工作集的大小，也就是循环的次数
Sub RunMTForLoop(ByVal FunctionPointer As Long, ByVal WorkSetSize As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long)



End Sub
