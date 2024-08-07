Attribute VB_Name = "modMTMain"
Option Explicit

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Private Declare Function UserDllMain Lib "msvbvm60" (OutInstance As Long, ByVal Unused As Long, ByVal hInstDll As Long, ByVal dwReason As Long, ByVal lpReserved As Long) As Long
Private Declare Function CreateIExprSrvObj Lib "msvbvm60" (Optional ByVal Reserved As Long, Optional ByVal Size As Long = 4, Optional ByVal Fail As Boolean) As IUnknown
Private Declare Function VBDllGetClassObject Lib "msvbvm60" (lpHInstDll As Long, ByVal Reserved As Long, lpVBHeader As Any, CLSID As Any, IID As Any, lpOutObject As IUnknown) As Long

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Declare Sub GetMem4 Lib "msvbvm60" (ByVal Addr As Long, Target As Any)
Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Any, dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long

Private MTInit As Long
Private MTInst As Long
Private hInst As Long
Private MT_CLSID(15) As Byte
Private MT_IID_IUnknown(15) As Byte
Private Const DLL_PROCESS_ATTACH As Long = 1
Private Const DLL_THREAD_ATTACH As Long = 2
Private Const DLL_THREAD_DETACH As Long = 3
Private Const DLL_PROCESS_DETACH As Long = 0

Sub Main()
On Error Resume Next
If MTInit = 1 Then Exit Sub

hInst = App.hInstance
MT_IID_IUnknown(8) = &HC0
MT_IID_IUnknown(15) = &H46

frmMain.Show

MTInit = 1
End Sub

Sub StartNewThread(ByVal ThreadEntry As Long, Optional ByVal ThreadParam As Long, Optional Out_ThreadId As Long)
CloseHandle CreateThread(ByVal 0, 0, ThreadEntry, ThreadParam, 0, Out_ThreadId)
End Sub

Private Function GetVBHeaderPtr() As Long
Dim Ptr As Long
' Get e_lfanew
GetMem4 ByVal hInst + &H3C, Ptr
' Get AddressOfEntryPoint
GetMem4 ByVal Ptr + &H28 + hInst, Ptr
' Get VBHeader
GetMem4 ByVal Ptr + hInst + 1, GetVBHeaderPtr
End Function

Sub ThreadInit()
'初始化线程
Dim ESO As IUnknown, ClassObj As IUnknown, VBHPtr As Long
Set ESO = CreateIExprSrvObj()
UserDllMain MTInst, 0, hInst, DLL_THREAD_ATTACH, 0
VBHPtr = GetVBHeaderPtr
If VBHPtr > 0 Then VBDllGetClassObject MTInst, 0, ByVal VBHPtr, MT_CLSID(0), MT_IID_IUnknown(0), ClassObj
End Sub

Sub ThreadQuit()
'线程退出
UserDllMain MTInst, 0, hInst, DLL_THREAD_DETACH, 0
End Sub

Private Function ThreadEntryTest(ByVal ThreadParam As Long) As Long
ThreadInit

'在这里写你的多线程内容
Sleep 100
MsgBox "线程函数：测试" & ThreadParam

ThreadQuit
End Function
