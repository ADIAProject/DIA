Attribute VB_Name = "mDEP"
Option Explicit

Public Declare Function SetProcessDEPPolicy Lib "kernel32.dll" (ByVal dwFlags As Long) As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub SetDEPDisable
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub SetDEPDisable()

    Dim mbCallback As Boolean

    If mbDebugStandart Then DebugMode "Disable DEP: Try to Disable DEP for this Process"

    If APIFunctionPresent("SetProcessDEPPolicy", "kernel32.dll") Then
        mbCallback = SetProcessDEPPolicy(0)
        If mbDebugStandart Then DebugMode "Disable DEP: Result: " & mbCallback & " - Err �" & err.LastDllError & " - " & ApiErrorText(err.LastDllError)
    Else
        If mbDebugStandart Then DebugMode "Disable DEP: ApiFunction not Supported"
    End If

End Sub
