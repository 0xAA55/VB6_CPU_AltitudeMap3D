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

'�������ö��߳�����ѭ��
'������
'FunctionPointer������ָ�룬�ú��������ĸ� Long ����������Լ���� stdcall
'  ѭ��������������һ�����������������ָ��
'
'WorkSetSize���������Ĵ�С��Ҳ����ѭ���Ĵ���
Sub RunMTForLoop(ByVal FunctionPointer As Long, ByVal WorkSetSize As Long, ByVal Param2 As Long, ByVal Param3 As Long, ByVal Param4 As Long)



End Sub
