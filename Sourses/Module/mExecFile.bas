Attribute VB_Name = "mExecFile"
Option Explicit

Private Type STARTUPINFO
    cb                                      As Long
    lpReserved                          As String
    lpDesktop                           As String
    lpTitle                             As String
    dwX                                 As Long
    dwY                                 As Long
    dwXSize                             As Long
    dwYSize                             As Long
    dwXCountChars                       As Long
    dwYCountChars                       As Long
    dwFillAttribute                     As Long
    dwFlags                             As Long
    wShowWindow                         As Integer
    cbReserved2                         As Integer
    lpReserved2                         As Long
    hStdInput                           As Long
    hStdOutput                          As Long
    hStdError                           As Long

End Type

Private Type PROCESS_INFORMATION
    hProcess                                As Long
    hThread                             As Long
    dwProcessId                         As Long
    dwThreadId                          As Long

End Type

Private Const STARTF_USESHOWWINDOW      As Long = &H1
Private Const INFINITE                  As Long = -1&

Public Const SW_SHOWNORMAL              As Long = 1

Private Const NORMAL_PRIORITY_CLASS     As Long = &H20

Public lngExitProc                      As Long

'Декларация функции для запуска файла.
Public Enum EShellShowConstants
    essSW_HIDE = 0
    essSW_MAXIMIZE = 3
    essSW_MINIMIZE = 6
    essSW_SHOWMAXIMIZED = 3
    essSW_SHOWMINIMIZED = 2
    essSW_SHOWNORMAL = 1
    essSW_SHOWNOACTIVATE = 4
    essSW_SHOWNA = 8
    essSW_SHOWMINNOACTIVE = 7
    essSW_SHOWDEFAULT = 10
    essSW_RESTORE = 9
    essSW_SHOW = 5
End Enum

#If False Then

    Private essSW_HIDE, essSW_MAXIMIZE, essSW_MINIMIZE, essSW_SHOWMAXIMIZED, essSW_SHOWMINIMIZED, essSW_SHOWNORMAL, essSW_SHOWNOACTIVATE
    Private essSW_SHOWNA, essSW_SHOWMINNOACTIVE, essSW_SHOWDEFAULT, essSW_RESTORE, essSW_SHOW
#End If

Private Const ERROR_FILE_NOT_FOUND      As Long = 2
Private Const ERROR_PATH_NOT_FOUND      As Long = 3
Private Const ERROR_BAD_FORMAT          As Long = 11
Private Const SE_ERR_ACCESSDENIED       As Integer = 5    ' access denied
Private Const SE_ERR_ASSOCINCOMPLETE    As Integer = 27
Private Const SE_ERR_DDEBUSY            As Integer = 30
Private Const SE_ERR_DDEFAIL            As Integer = 29
Private Const SE_ERR_DDETIMEOUT         As Integer = 28
Private Const SE_ERR_DLLNOTFOUND        As Integer = 32
Private Const SE_ERR_FNF                As Integer = 2    ' file not found
Private Const SE_ERR_NOASSOC            As Integer = 31
Private Const SE_ERR_PNF                As Integer = 3    ' path not found
Private Const SE_ERR_OOM                As Integer = 8    ' out of memory
Private Const SE_ERR_SHARE              As Integer = 26

Private Declare Function CreateProcess _
                          Lib "kernel32.dll" _
                              Alias "CreateProcessA" (ByVal lpApplicationName As String, _
                                                      ByVal lpCommandLine As String, _
                                                      ByVal lpProcessAttributes As Long, _
                                                      ByVal lpThreadAttributes As Long, _
                                                      ByVal bInheritHandles As Long, _
                                                      ByVal dwCreationFlags As Long, _
                                                      ByVal lpEnvironment As Long, _
                                                      ByVal lpCurrentDirectory As String, _
                                                      lpStartupInfo As STARTUPINFO, _
                                                      lpProcessInformation As PROCESS_INFORMATION) As Long

Private Declare Function GetExitCodeProcess _
                          Lib "kernel32.dll" (ByVal hProcess As Long, _
                                              lpExitCode As Long) As Long

Public Declare Function ShellExecute _
                         Lib "shell32.dll" _
                             Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                    ByVal lpOperation As String, _
                                                    ByVal lpFile As String, _
                                                    ByVal lpParameters As String, _
                                                    ByVal lpDirectory As String, _
                                                    ByVal nShowCmd As Long) As Long

Private Declare Function WaitForSingleObject _
                          Lib "kernel32.dll" (ByVal hHandle As Long, _
                                              ByVal dwMilliseconds As Long) As Long

Private Declare Function ShellExecuteForExplore _
                          Lib "shell32.dll" _
                              Alias "ShellExecuteA" (ByVal hWnd As Long, _
                                                     ByVal lpOperation As String, _
                                                     ByVal lpFile As String, _
                                                     lpParameters As Any, _
                                                     lpDirectory As Any, _
                                                     ByVal nShowCmd As Long) As Long


'! -----------------------------------------------------------
'!  Функция     :  RunAndWaitNewNew
'!  Переменные  :  ComLine As String, DefaultDir As String, ShowFlag As VbAppWinStyle
'!  Возвр. знач.:  As Boolean
'!  Описание    :  'запустить приложение с ожиданием завершения.
'! -----------------------------------------------------------
Public Function RunAndWaitNew(ComLine As String, _
                              DefaultDir As String, _
                              ShowFlag As VbAppWinStyle) As Boolean

Dim SI                                  As STARTUPINFO
Dim PI                                  As PROCESS_INFORMATION
Dim nRet                                As Long

    DebugMode vbTab & "RunAndWait-Start"
    DoEvents
    lngExitProc = 0

    If ShowFlag = vbHide Then
        If Not mbHideOtherProcess Then
            ShowFlag = vbNormalFocus

        End If

    End If

    DebugMode str2VbTab & "RunString: " & ComLine
    DebugMode str2VbTab & "StartDir: " & DefaultDir
    nRet = ShellW(ComLine, ShowFlag, INFINITE)
    'WaitForSingleObject PI.hProcess, INFINITE
    'GetExitCodeProcess PI.hProcess, nRet
    DebugMode str2VbTab & "ReturnCode: " & CStr(nRet) & " - " & ApiErrorText(err.LastDllError)
    'CloseHandle PI.hProcess
    lngExitProc = nRet
    RunAndWaitNew = True
    DebugMode vbTab & "RunAndWaitNew-End"
    DoEvents

End Function

'! -----------------------------------------------------------
'!  Функция     :  RunAndWait
'!  Переменные  :  ComLine As String, DefaultDir As String, ShowFlag As VbAppWinStyle
'!  Возвр. знач.:  As Boolean
'!  Описание    :  'запустить приложение с ожиданием завершения.
'! -----------------------------------------------------------
Public Function RunAndWait(ComLine As String, _
                           DefaultDir As String, _
                           ShowFlag As VbAppWinStyle) As Boolean

Dim SI                                  As STARTUPINFO
Dim PI                                  As PROCESS_INFORMATION
Dim nRet                                As Long

    DebugMode vbTab & "RunAndWait-Start"
    DoEvents
    lngExitProc = 0

    If ShowFlag = vbHide Then
        If Not mbHideOtherProcess Then
            ShowFlag = vbNormalFocus

        End If

    End If

    SI.wShowWindow = ShowFlag
    SI.dwFlags = STARTF_USESHOWWINDOW
    DebugMode str2VbTab & "RunString: " & ComLine
    DebugMode str2VbTab & "StartDir: " & DefaultDir
    nRet = CreateProcess(vbNullString, ComLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, DefaultDir, SI, PI)
    WaitForSingleObject PI.hProcess, INFINITE
    GetExitCodeProcess PI.hProcess, nRet
    DebugMode str2VbTab & "ReturnCode: " & CStr(nRet) & " - " & ApiErrorText(err.LastDllError)
    CloseHandle PI.hProcess
    lngExitProc = nRet
    RunAndWait = True
    DebugMode vbTab & "RunAndWait-End"
    DoEvents

End Function

Public Sub RunUtilsShell(ByVal strPathUtils As String, _
                         Optional ByVal mbCollectPath As Boolean = True, _
                         Optional ByVal mbStartPathAsPathExe As Boolean = False)

Dim nRetShellEx                         As Boolean
Dim cmdString                           As String
Dim strStartPath                        As String

    If mbCollectPath Then
        cmdString = PathCollect(strPathUtils)

        If mbStartPathAsPathExe Then
            strStartPath = PathNameFromPath(cmdString)

        End If

    Else
        cmdString = strPathUtils

    End If

    DebugMode "cmdString: " & cmdString

    If mbStartPathAsPathExe Then
        nRetShellEx = ShellEx(cmdString, essSW_SHOWDEFAULT, vbNullString, strStartPath, "open")
    Else
        nRetShellEx = ShellEx(cmdString, essSW_SHOWNORMAL)

    End If

    DebugMode "cmdString: " & nRetShellEx

End Sub

Public Function ShellEx(ByVal sFile As String, _
                        Optional ByVal eShowCmd As EShellShowConstants = essSW_SHOWDEFAULT, _
                        Optional ByVal sParameters As String = vbNullString, _
                        Optional ByVal sDefaultDir As String = vbNullString, _
                        Optional sOperation As String = "open", _
                        Optional Owner As Long = 0) As Boolean

Dim lR                                  As Long
Dim lErr                                As Long
Dim sErr                                As String

    If InStr(1, sFile, ".exe", vbTextCompare) Then
        eShowCmd = 0
    End If

    On Error Resume Next

    If LenB(sParameters) = 0 Then
        If LenB(sDefaultDir) = 0 Then
            lR = ShellExecuteForExplore(Owner, sOperation, sFile, 0, 0, essSW_SHOWNORMAL)
        Else
            lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)

        End If

    Else
        lR = ShellExecute(Owner, sOperation, sFile, sParameters, sDefaultDir, eShowCmd)

    End If

    If (lR < 0) Or (lR > 32) Then
        ShellEx = True
        DebugMode "ShellExecute: True - and result API ShellExecute:" & ApiErrorText(lR)
    Else
        ' raise an appropriate error:
        lErr = vbObjectError + 1048 + lR

        Select Case lR

            Case 0
                lErr = 7
                sErr = "Out of memory"

            Case ERROR_FILE_NOT_FOUND
                lErr = 53
                sErr = "File not found"

            Case ERROR_PATH_NOT_FOUND
                lErr = 76
                sErr = "Path not found"

            Case ERROR_BAD_FORMAT
                sErr = "The executable file is invalid or corrupt"

            Case SE_ERR_ACCESSDENIED
                lErr = 75
                sErr = "Path/file access error"

            Case SE_ERR_ASSOCINCOMPLETE
                sErr = "This file type does not have a valid file association."

            Case SE_ERR_DDEBUSY
                lErr = 285
                sErr = "The file could not be opened because the target application is busy. Please try again in a moment."

            Case SE_ERR_DDEFAIL
                lErr = 285
                sErr = "The file could not be opened because the DDE transaction failed. Please try again in a moment."

            Case SE_ERR_DDETIMEOUT
                lErr = 286
                sErr = "The file could not be opened due to time out. Please try again in a moment."

            Case SE_ERR_DLLNOTFOUND
                lErr = 48
                sErr = "The specified dynamic-link library was not found."

            Case SE_ERR_FNF
                lErr = 53
                sErr = "File not found"

            Case SE_ERR_NOASSOC
                sErr = "No application is associated with this file type."

            Case SE_ERR_OOM
                lErr = 7
                sErr = "Out of memory"

            Case SE_ERR_PNF
                lErr = 76
                sErr = "Path not found"

            Case SE_ERR_SHARE
                lErr = 75
                sErr = "A sharing violation occurred."

            Case Else
                sErr = "An error occurred occurred whilst trying to open or print the selected file."

        End Select

        DebugMode "ShellExecute: " & lErr & " - " & sErr & " - ErrAPI: " & ApiErrorText(lR)
        ShellEx = False

    End If

    On Error GoTo 0

End Function

Sub RestartProgram()

Dim S_A                                 As SECURITY_ATTRIBUTES
    ' атрибуты защиты описателя и наследования
Dim SI                                  As STARTUPINFO
    ' доп информация для запуска
Dim PI                                  As PROCESS_INFORMATION
    ' информация о процессе
Dim strError                            As String
Dim strExecute                          As String
Dim strCaption                          As String
Dim Flags                               As Long
Dim Directory                           As String
Dim lngResult                           As Long

    ' Задаем флаг приоритета
    Flags = NORMAL_PRIORITY_CLASS

    'формируем строку пути

    strExecute = Kavichki & App.EXEName & ".exe" & Kavichki
    MsgBox Kavichki & App.EXEName & ".exe" & Kavichki

    'формируем строку папки
    Directory = Space$(Len(App.Path))
    Directory = App.Path

    'инициализация структуры желательна
    S_A.bInheritHandle = 0&
    S_A.lpSecurityDescriptor = 0&
    S_A.nLength = Len(S_A)
    'особенно этого параметра

    SI.cb = Len(SI)
    'и этого


    lngResult = CreateProcess(vbNullString, strExecute, 0&, 0&, 0&, NORMAL_PRIORITY_CLASS, 0&, Directory, SI, PI)
    'nRet = CreateProcess(vbNullString, ComLine, 0&, 0&, 1&, NORMAL_PRIORITY_CLASS, 0&, DefaultDir, SI, PI)
    strError = ApiErrorText(err.LastDllError)
    If lngResult <> 0 Then
        CloseHandle PI.hThread
        'этот описатель не понадобится

        'hProcess = PI.hProcess              'для последующего завершения
        'lProcess = PI.dwProcessId

    Else
        'hProcess = 0
        'lProcess = 0

        strExecute = Space$(32)
        strCaption = Space$(32)

        strExecute = "Ошибка запуска процесса: " & App.EXEName & " Код ошибки: " & err.LastDllError
        strCaption = "Error"
        Call MessageBox(frmMain.hWnd, ByVal strExecute, ByVal strCaption, 16)

        'Label1 = "Ошибка: " & Error
    End If

End Sub

