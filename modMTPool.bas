Attribute VB_Name = "modMTPool"
Option Explicit

Private NumProcessors As Long

Private Sub ThreadProc()

End Sub

Sub RunMTForLoop(ByVal FunctionPointer As Long, ByVal WorkSetSize As Long)

If NumProcessors = 0 Then
    NumProcessors = CLng(Environ("NUMBER_OF_PROCESSORS"))
    
    Dim cpuSet As Object, cpu As Object
    Set cpuSet = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Processor")
    For Each cpu In cpuSet
        NumProcessors = NumProcessors + 1
    Next
End If


End Sub
