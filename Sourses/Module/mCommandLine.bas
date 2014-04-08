Attribute VB_Name = "mCommandLine"
Option Explicit

' Запуск с коммандной строкой
Public mbRunWithParam                    As Boolean

' Работаем в тихом режиме
Public mbSilentRun                       As Boolean
Public miSilentRunTimer                  As Integer
Public mbSilentDLL                       As Boolean
Public strSilentSelectMode               As String

'Api-Declare for Unicode Support for CommandLine. Writen by  Krool - http://www.vbforums.com/showthread.php?748607-Command%28%29-Unicode-aware
Private Declare Function GetCommandLine Lib "kernel32" Alias "GetCommandLineW" () As Long
Private Declare Function PathGetArgs Lib "shlwapi" Alias "PathGetArgsW" (ByVal lpszPath As Long) As Long
Private Declare Function SysReAllocString Lib "oleaut32" (ByVal pbString As Long, ByVal pszStrPtr As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CmdLineParsing
'! Description (Описание)  :   [Функция анализа коммандной строки и присвоение переменных на основании передеваемых комманд]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CmdLineParsing()

    Dim argRetCMD    As Collection
    Dim i            As Integer
    Dim intArgCount  As Integer
    Dim strArg       As String
    Dim strArg_x()   As String
    Dim iArgRavno    As Integer
    Dim iArgDvoetoch As Integer
    Dim strArgParam  As String

    With New cCMDArguments
        .CommandLine = "CMDLineParams " & Command
        Set argRetCMD = .Arguments
        intArgCount = argRetCMD.Count
    End With

    For i = 2 To intArgCount
        strArg = argRetCMD(i)
        iArgRavno = InStr(strArg, strRavno)
        iArgDvoetoch = InStr(strArg, strDvoetochie)

        If iArgRavno Then
            strArg_x = Split(strArg, strRavno)
            strArg = strArg_x(0)
            strArgParam = strArg_x(1)
        ElseIf iArgDvoetoch Then
            'strArg_x = Split(strArg, strDvoetochie)
            strArg = Left$(argRetCMD(i), iArgDvoetoch - 1)
            strArgParam = Right$(argRetCMD(i), Len(argRetCMD(i)) - iArgDvoetoch)
        End If

        Select Case LCase$(strArg)

            Case "/?", "/h", "-help", "/help", "-h", "--h", "--help"
                ShowHelpMsg

                End

            Case "/extractdll", "-extractdll", "--extractdll"
                ExtractrResToFolder strArgParam

                End

            Case "/regdll", "-regdll", "--regdll"
                RegisterAddComponent

                End

            Case "/t", "-t", "--t"

                If IsNumeric(strArgParam) Then
                    miSilentRunTimer = CInt(strArgParam)
                Else
                    miSilentRunTimer = 10
                End If

                mbDebugStandart = True
                mbUpdateCheck = False

            Case "/s", "-s", "--s"

                Select Case LCase$(strArgParam)

                    Case "n"
                        'новые
                        strSilentSelectMode = "n"

                    Case "q"
                        'неустановленные
                        strSilentSelectMode = "q"

                    Case "a"
                        'Все на вкладке
                        strSilentSelectMode = "a"

                    Case "n2"
                        'новые
                        strSilentSelectMode = "n2"

                    Case "q2"
                        'неустановленные
                        strSilentSelectMode = "q2"

                    Case "a2"
                        'Все на вкладке
                        strSilentSelectMode = "a2"

                    Case Else
                        'по умолчанию
                        strSilentSelectMode = "n"
                End Select

                ' на случай если не указано время ожидания запуска
                If miSilentRunTimer <= 0 Then
                    miSilentRunTimer = 10
                End If

                mbDebugStandart = True
                mbUpdateCheck = False

            'todo SaveHWIDs2File
            Case "/savereport", "-savereport", "--savereport"
                'SaveHWIDs2File
                'strArgParam
                
                End
                
            Case Else
                ShowHelpMsg

                End

        End Select

    Next i

End Sub

' (VB-Overwrite)
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CmdLineParsing
'! Description (Описание)  :   [Функция получения строки запуска программы, заменяет встроенную функцию VBA.Command$()]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Command() As String
    If InIDE() = False Then
        SysReAllocString VarPtr(Command), PathGetArgs(GetCommandLine())
        Command = LTrim$(Command)
    Else
        Command = VBA.Command$()
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveReport
'! Description (Описание)  :   [Показ окна с параметрами запуска]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveReport(ByVal strFilePathTo As String)

    If LenB(strFilePathTo) Then
        If PathExists(strResultHwidsExtTxtPath) Then
            CopyFileTo strResultHwidsExtTxtPath, strFilePathTo
        Else

            If SaveHwidsArray2File(strResultHwidsExtTxtPath, arrHwidsLocal) Then
                If PathExists(strResultHwidsExtTxtPath) Then
                    CopyFileTo strResultHwidsExtTxtPath, strFilePathTo
                Else
                    If mbDebugStandart Then DebugMode strMessages(45) & vbNewLine & strFilePathTo
                End If
            Else
                If mbDebugStandart Then DebugMode strMessages(45) & vbNewLine & strFilePathTo
            End If
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowHelpMsg
'! Description (Описание)  :   [Показ окна с параметрами запуска]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ShowHelpMsg()
    MsgBox strMessages(137), vbInformation & vbOKOnly, strProductName
End Sub
