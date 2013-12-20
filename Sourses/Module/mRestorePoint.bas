Attribute VB_Name = "mRestorePoint"
Option Explicit

Private Const DEVICE_DRIVER_INSTALL As Integer = 10
Private Const BEGIN_SYSTEM_CHANGE   As Integer = 100

' Проверка реестра на опцию SystemRestore - включена или нет
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckRestorePoint
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function CheckRestorePoint() As Boolean

    If OsCurrVersionStruct.VerFull <> "5.0" And OsCurrVersionStruct.VerFull <> "5.2" Or (OsCurrVersionStruct.VerFull <> "6.2" And OsCurrVersionStruct.ClientOrServer) Or (OsCurrVersionStruct.VerFull <> "6.3" And OsCurrVersionStruct.ClientOrServer) _
                                Then
        regParam = GetKeyValue(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows NT\CurrentVersion\SystemRestore", "DisableSR")

        'HKLM\\SOFTWARE\\Microsoft\\Windows NT\\CurrentVersion\\SystemRestore\\DisableSR
        'HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\sr
        If LenB(regParam) = 0 Then
            CheckRestorePoint = True
        Else
            CheckRestorePoint = regParam = "0"
            DebugMode "CheckRestorePoint: Enable in Operation System: " & CheckRestorePoint
        End If

    Else
        DebugMode "CheckRestorePoint: Not Supported by Operation System"
        CheckRestorePoint = False
    End If

    DebugMode "CheckRestorePoint: " & regParam & "(" & CheckRestorePoint & ")"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateRestorePoint
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CreateRestorePoint()

    Dim strComputer   As String
    Dim objWMIService As Object
    Dim objRP         As Object
    Dim errResults    As Long

    DebugMode vbTab & "CreateRestorePoint-Start"
    ChangeStatusTextAndDebug strMessages(118)
    strComputer = "."
    ' Блокируем форму при создании точки восстановления
    frmMain.BlockControl False

    On Error GoTo HandErr

    'http://msdn.microsoft.com/en-us/library/aa378951%28v=VS.85%29.aspx
    'http://www.kellys-korner-xp.com/xp_restore.htm
    If CheckRestorePoint Then
        Set objWMIService = CreateObject("winmgmts:" & "{impersonationLevel=impersonate}!\\" & strComputer & "\root\default")
        Set objRP = objWMIService.Get("SystemRestore")
        errResults = objRP.CreateRestorePoint(strProductName & " v" & strProductVersion, DEVICE_DRIVER_INSTALL, BEGIN_SYSTEM_CHANGE)

        If errResults = 0 Then
            DebugMode vbTab & "CreateRestorePoint-Success"
            ChangeStatusTextAndDebug strMessages(119) & " " & strProductName & " v" & strProductVersion

            If Not mbSilentRun Then
                MsgBox strMessages(119) & " " & strProductName & " v" & strProductVersion, vbInformation, strProductName
            End If

        Else
            DebugMode vbTab & "CreateRestorePoint-Failed: err=" & errResults
            ChangeStatusTextAndDebug strMessages(117)

            If Not mbSilentRun Then
                MsgBox strMessages(117), vbCritical, strProductName
            End If
        End If

    Else
        ChangeStatusTextAndDebug strMessages(116)

        If Not mbSilentRun Then
            MsgBox strMessages(116), vbInformation, strProductName
        End If
    End If

    Set objWMIService = Nothing
    Set objRP = Nothing
    ' РазБлокируем форму при создании точки восстановления
    frmMain.BlockControl True
ExitFromSub:

    Exit Sub

HandErr:
    DebugMode "CreateRestorePoint:  Err.Number: " & Err.Number & " Err.Description: " & Err.Description

    If Err.Number = -2147217389 Then
        MsgBox "Error №: " & Err.Number & vbNewLine & "Description: " & Err.Description & str2vbNewLine & "This Error in Function 'CreateRestorePoint'. Probably trouble with WMI.", vbCritical, strProductName
    ElseIf Err.Number = -2147217406 Then
        MsgBox "Error №: " & Err.Number & vbNewLine & "Description: " & Err.Description & str2vbNewLine & "This Error in Function 'CreateRestorePoint'. Maybe this Function not Supported this operation system.", vbCritical, strProductName
    ElseIf Err.Number <> 0 Then
        GoTo ExitFromSub
    End If

    ' РазБлокируем форму при создании точки восстановления
    frmMain.BlockControl True
End Sub
