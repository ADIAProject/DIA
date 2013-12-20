Attribute VB_Name = "mDEP"
Option Explicit

Public Declare Function SetProcessDEPPolicy Lib "kernel32.dll" (ByVal dwFlags As Long) As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetDEPDisable
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub SetDEPDisable()

    Dim mbCallback As Boolean

    DebugMode "Disable DEP: Try to Disable DEP for this Process"

    If APIFunctionPresent("SetProcessDEPPolicy", "kernel32.dll") Then
        mbCallback = SetProcessDEPPolicy(0)
        DebugMode "Disable DEP: Result: " & mbCallback & " - Err №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
    Else
        DebugMode "Disable DEP: ApiFunction not Supported"
    End If

End Sub
