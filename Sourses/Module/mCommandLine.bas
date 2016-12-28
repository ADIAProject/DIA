Attribute VB_Name = "mCommandLine"
Option Explicit

#Const mbIDE_DBSProject = False

' Запуск с коммандной строкой
Public mbRunWithParam                    As Boolean

' Работаем в тихом режиме
Public mbSilentRun                       As Boolean
Public miSilentRunTimer                  As Integer
Public mbSilentDLL                       As Boolean
Public strSilentSelectMode               As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CmdLineParsing
'! Description (Описание)  :   [Функция анализа коммандной строки и присвоение переменных на основании передеваемых комманд, возвращаяет True если трубуется выход]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function CmdLineParsing() As Boolean

    Dim argRetCMD    As Collection
    Dim ii           As Integer
    Dim intArgCount  As Integer
    Dim strArg       As String
    Dim strArg_x()   As String
    Dim iArgRavno    As Integer
    Dim iArgDvoetoch As Integer
    Dim strArgParam  As String

    With New cCMDArguments
        .CommandLine = "CMDLineParams " & Command
        Set argRetCMD = .Arguments
        intArgCount = argRetCMD.count
    End With

    For ii = 2 To intArgCount
        strArg = argRetCMD(ii)
        iArgRavno = InStr(strArg, strRavno)
        iArgDvoetoch = InStr(strArg, strColon)

        If iArgRavno Then
            strArg_x = Split(strArg, strRavno)
            strArg = strArg_x(0)
            strArgParam = strArg_x(1)
        ElseIf iArgDvoetoch Then
            strArg = Left$(argRetCMD(ii), iArgDvoetoch - 1)
            strArgParam = Right$(argRetCMD(ii), Len(argRetCMD(ii)) - iArgDvoetoch)
        End If

        mbRunWithParam = True
        
        Select Case LCase$(strArg)

            Case "/?", "/h", "-help", "/help", "-h", "--h", "--help"
                
                ShowHelpMsg
                CmdLineParsing = True

            Case "/extractdll", "-extractdll", "--extractdll"
                
                mbSilentRun = True
                ExtractrResToFolder strArgParam
                CmdLineParsing = True

            Case "/regdll", "-regdll", "--regdll"
                
                RegisterAddComponent
                CmdLineParsing = True

#If Not mbIDE_DBSProject Then
            Case "/t", "-t", "--t"

                If IsNumeric(strArgParam) Then
                    miSilentRunTimer = CInt(strArgParam)
                Else
                    miSilentRunTimer = 10
                End If

                mbDebugStandart = True
                mbUpdateCheck = False

            Case "/s", "-s", "--s"

                mbSilentRun = True
                
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
                
            ' SaveSnapReport - сохранение снимка системы в файл
            Case "/savereport", "-savereport", "--savereport"
                
                mbSilentRun = True
                ' Получаем данные от devcon.exe и Обновляем данные из реестра
                If RunDevcon Then
                    DevParserLocalHwids2
                    CollectHwidFromReestr
                    strCompModel = GetMBInfo()
                    ' Сохраняем снимок
                    SaveSnapReport strArgParam
                End If
                CmdLineParsing = True
                
#End If
            Case Else
            
                ShowHelpMsg
                CmdLineParsing = True

        End Select

    Next ii

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowHelpMsg
'! Description (Описание)  :   [Показ окна с параметрами запуска]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub ShowHelpMsg()
    MsgBoxEx strMessages(137), vbInformation & vbOKOnly, strProductName, 25
End Sub
