Attribute VB_Name = "mDEP"
Option Explicit

Public Declare Function SetProcessDEPPolicy _
                         Lib "kernel32.dll" (ByVal dwFlags As Long) As Boolean

Public Sub SetDEPDisable()

Dim mbCallback                          As Boolean

    DebugMode "Disable DEP: Try to Disable DEP for this Process"

    If APIFunctionPresent("SetProcessDEPPolicy", "kernel32.dll") Then
        mbCallback = SetProcessDEPPolicy(0)
        DebugMode "Disable DEP: Result: " & mbCallback & " - Err ¹" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
    Else
        DebugMode "Disable DEP: ApiFunction not Supported"

    End If

End Sub
