Attribute VB_Name = "mDevParser"
Option Explicit

' Сравнение версий баз драйверов, с константой в программе
Public Function CompareDevDBVersion(strDevDBFullFileName As String, _
                                    Optional ByVal strPathDRP As String) As Boolean

Dim LngValue                            As Long
Dim strFilePath_woExt                   As String

    strFilePath_woExt = FileName_woExt(strDevDBFullFileName)
    LngValue = IniLongPrivate(FileNameFromPath(strFilePath_woExt), "Version", BackslashAdd2Path(PathNameFromPath(strFilePath_woExt)) & "DevDBVersions.ini")

    If LngValue = 9999 Then
        CompareDevDBVersion = False
    Else
        CompareDevDBVersion = Not (LngValue <> lngDevDBVersion)
    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  DevParserLocalHwids2
'!  Переменные  :
'!  Описание    :  Парсинг выходного файла devcon для локальных устройств
'! -----------------------------------------------------------
Public Sub DevParserLocalHwids2()

Dim objInfFile                          As TextStream
Dim str                                 As String
Dim i                                   As Long
Dim strCnt                              As Long
Dim miStatus                            As Long
Dim strID                               As String
Dim strIDOrig                           As String
Dim strIDCutting                        As String
Dim strName                             As String
Dim strName_x()                         As String
Dim miMaxCountArr                       As Long
Dim RecCountArr                         As Long
Dim strID_x()                           As String
Dim RegExpDevcon                        As RegExp
Dim MatchesDevcon                       As MatchCollection
Dim objMatch                            As Match
Dim strStatus                           As String

    Set RegExpDevcon = New RegExp
    With RegExpDevcon
        .Pattern = "(^[^\n\r\s][^\n\r]+)\r\n(\s+[^\n\r]+\r\n)*[^\n\r]*((?:DEVICE IS|DEVICE HAS|DRIVER IS|DRIVER HAS)[^\r]+)"
        .MultiLine = True
        .IgnoreCase = True
        .Global = True

    End With

    DebugMode "DevParserLocalHwids2-Start"

    If PathFileExists(strHwidsTxtPath) = 1 Then
        If Not IsPathAFolder(strHwidsTxtPath) Then
            Set objInfFile = objFSO.OpenTextFile(strHwidsTxtPath, ForReading, False, TristateUseDefault)
            str = objInfFile.ReadAll
            Set MatchesDevcon = RegExpDevcon.Execute(str)
            miMaxCountArr = 100
            ' максимальное кол-во элементов в массиве
            ReDim arrHwidsLocal(miMaxCountArr)
            strCnt = MatchesDevcon.Count
            RecCountArr = 0

            'For i = 0 To MatchesDevcon.Count - 1
            For i = 0 To strCnt - 1
                Set objMatch = MatchesDevcon.Item(i)

                ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                If RecCountArr = miMaxCountArr Then
                    miMaxCountArr = miMaxCountArr + miMaxCountArr
                    ReDim Preserve arrHwidsLocal(miMaxCountArr)

                End If

                ' получаем данные
                With objMatch
                    strID = .SubMatches(0)
                    strName = .SubMatches(1)
                    strStatus = .SubMatches(2)
                End With

                'objMatch
                strID = UCase$(Trim$(Replace$(strID, vbNewLine, vbNullString)))
                ' разбиваем по "\"
                strIDOrig = strID

                If InStr(strID, "\") Then
                    strID_x = Split(strID, "\")
                    strID = strID_x(0) & "\" & strID_x(1)
                End If

                strIDCutting = ParseDoubleHwid(strID)

                'Если не входит в список исключений, то продолжаем
                If strExcludeHWID <> "*" Then
                    If Not MatchSpec(strID, strExcludeHWID) Then
                        '"Ура: " & strID & " present in " & strExcludeHWID
                        miStatus = 0

                        ' устройство активно
                        If InStr(1, strStatus, "running", vbTextCompare) Then
                            miStatus = 1

                        End If

                        If LenB(strName) > 0 Then
                            strName_x = Split(strName, vbNewLine)
                            strName = strName_x(0)
                        End If

                        strName = Replace$(strName, vbNewLine, vbNullString)
                        strName = Replace$(strName, "Name:", vbNullString, , , vbTextCompare)
                        strName = Trim$(strName)

                        If Len(strID) > 3 Then
                            arrHwidsLocal(RecCountArr).HWID = strID
                            arrHwidsLocal(RecCountArr).DevName = strName
                            arrHwidsLocal(RecCountArr).Status = miStatus
                            arrHwidsLocal(RecCountArr).HWIDOrig = strIDOrig
                            arrHwidsLocal(RecCountArr).HWIDCutting = strIDCutting
                            RecCountArr = RecCountArr + 1

                        End If

                    End If

                Else
                    miStatus = 0

                    ' устройство активно
                    If InStr(1, strStatus, "running", vbTextCompare) Then
                        miStatus = 1
                    End If

                    If LenB(strName) > 0 Then
                        strName_x = Split(strName, vbNewLine)
                        strName = strName_x(0)
                    End If

                    strName = Replace$(strName, vbNewLine, vbNullString)
                    strName = Replace$(strName, "Name:", vbNullString, , , vbTextCompare)
                    strName = Trim$(strName)

                    If Len(strID) > 3 Then
                        ' ID оборудования
                        arrHwidsLocal(RecCountArr).HWID = strID
                        ' Наименование оборудования
                        arrHwidsLocal(RecCountArr).DevName = strName
                        ' Статус оборудования
                        arrHwidsLocal(RecCountArr).Status = miStatus
                        arrHwidsLocal(RecCountArr).HWIDOrig = UCase$(strIDOrig)
                        arrHwidsLocal(RecCountArr).HWIDCutting = UCase$(strIDCutting)
                        RecCountArr = RecCountArr + 1

                    End If

                End If

            Next

            ' Переобъявляем массив на реальное кол-во записей
            If RecCountArr > 0 Then
                ReDim Preserve arrHwidsLocal(RecCountArr - 1)
            Else
                ReDim Preserve arrHwidsLocal(0)

            End If

            If SaveHwidsArray2File(strResultHwidsTxtPath, arrHwidsLocal) = False Then
                MsgBox strMessages(45) & vbNewLine & strResultHwidsTxtPath, vbCritical + vbInformation, strProductName
            End If

        End If

    Else
        MsgBox strHwidsTxtPath & vbNewLine & strMessages(46), vbInformation, strProductName
        Unload frmMain

    End If

    DebugMode "DevParserLocalHwids2-End"

End Sub

Public Function ParseDoubleHwid(ByVal strValuer As String) As String

Dim strValuer_x()                       As String
Dim miSubSys                            As Long
Dim miREV                               As Long
Dim miMI                                As Long
Dim miCC                                As Long

    If LenB(strValuer) > 0 Then

        ' разбиваем по "\" - оставляем только xxx\yyy
        If InStr(strValuer, "\") Then
            strValuer_x = Split(strValuer, "\")

            If UBound(strValuer_x) >= 1 Then
                strValuer = strValuer_x(0) & "\" & strValuer_x(1)
            Else
                strValuer = strValuer_x(0)
            End If

        End If

    End If

    ParseDoubleHwid = strValuer

End Function

'! -----------------------------------------------------------
'!  Функция     :  RunDevcon
'!  Переменные  :
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Запуск программы Devcon
'! -----------------------------------------------------------
Public Function RunDevcon() As Boolean

Dim cmdString                           As String

    DebugMode "RunDevcon-Start"
    cmdString = Kavichki & strDevconCmdPath & Kavichki & " " & Kavichki & strDevConExePath & Kavichki & " " & Kavichki & strHwidsTxtPath & Kavichki & " 1"

    CreateIfNotExistPath strWorkTemp

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        RunDevcon = False
    Else
        RunDevcon = True

    End If

    If GetFileSizeByPath(strHwidsTxtPath) > 0 Then
        PrintFileInDebugLog strHwidsTxtPath
    Else
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        RunDevcon = False
    End If
    DebugMode vbTab & "Run Devcon: " & RunDevcon
    DebugMode "RunDevcon-End"

End Function

'! -----------------------------------------------------------
'!  Функция     :  RunDevconRescan
'!  Переменные  :
'!  Возвр. знач.:  As Boolean
'!  Описание    :  'Поиск новых устройств + Запуск программы Devcon
'! -----------------------------------------------------------
Public Function RunDevconRescan(Optional ByVal lngPause As Long = 1) As Boolean

Dim cmdString                           As String

    DebugMode "RunDevconRescan-Start"
    cmdString = Kavichki & strDevConExePath & Kavichki & " rescan"
    ChangeStatusTextAndDebug strMessages(96) & " " & cmdString

    CreateIfNotExistPath strWorkTemp

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        RunDevconRescan = False
    Else
        RunDevconRescan = True

    End If

    DebugMode vbTab & "Run RunDevconRescan: " & RunDevconRescan
    DebugMode "RunDevconRescan-End"

End Function

'! -----------------------------------------------------------
'!  Функция     :  RunDevconView
'!  Переменные  :
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Запуск программы Devcon
'! -----------------------------------------------------------
Public Function RunDevconView() As Boolean

Dim cmdString                           As String

    DebugMode "RunDevconView-Start"
    cmdString = Kavichki & strDevconCmdPath & Kavichki & " " & Kavichki & strDevConExePath & Kavichki & " " & Kavichki & strHwidsTxtPathView & Kavichki & " 3"

    If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        RunDevconView = False
    Else
        RunDevconView = True

    End If

    DebugMode vbTab & "Run DevconView: " & RunDevconView
    DebugMode "RunDevconView-End"

End Function

'! -----------------------------------------------------------
'!  Функция     :  CollectCmdString
'!  Переменные  :
'!  Описание    :  Создание коммандной строки запуска программы DPInst
'! -----------------------------------------------------------
Public Function CollectCmdString() As String

Dim strCmdStringDPInstTemp              As String

    If mbDpInstLegacyMode Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/LM "

    End If

    If mbDpInstPromptIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/P "

    End If

    If mbDpInstForceIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/F "

    End If

    If mbDpInstSuppressAddRemovePrograms Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SA "

    End If

    If mbDpInstSuppressWizard Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SW "

    End If

    If mbDpInstQuietInstall Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/Q "

    End If

    If mbDpInstScanHardware Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SH "

    End If

    ' Результирующая строка
    CollectCmdString = strCmdStringDPInstTemp

End Function

