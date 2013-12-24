Attribute VB_Name = "mDevParserByReqExp"
Option Explicit

Private objInfFile        As TextStream
Private objRezultFile     As TextStream
Private objRezultFileHWID As TextStream
Private objCatFile        As TextStream
Private RegExpStrSect     As RegExp
Private RegExpStrDefs     As RegExp
Private RegExpVerSect     As RegExp
Private RegExpCatSect     As RegExp
Private RegExpManSect     As RegExp
Private RegExpManDef      As RegExp
Private RegManID          As RegExp
Private RegExpDevDef      As RegExp
Private RegExpDevID       As RegExp
Private RegExpDevSect     As RegExp
Private objHashOutput     As Scripting.Dictionary
Private objStringHash     As Scripting.Dictionary

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DevParserByRegExp
'! Description (Описание)  :   [Парсинг ID и названий устройст из inf-файла и построение БД]
'! Parameters  (Переменные):   strPackFileName (String)
'                              strPathDRP (String)
'                              strPathDevDB (String)
'!--------------------------------------------------------------------------------
Public Sub DevParserByRegExp(strPackFileName As String, ByVal strPathDRP As String, ByVal strPathDevDB As String)

    Dim objMatch                  As Match
    Dim objMatch1                 As Match
    Dim objMatchesStrSect         As MatchCollection
    Dim objMatchesVerSect         As MatchCollection
    Dim objMatchesCatSect         As MatchCollection
    Dim objMatchesManSect         As MatchCollection
    Dim objMatchesManDef          As MatchCollection
    Dim objMatchesManID           As MatchCollection
    Dim objMatchesDevSect         As MatchCollection
    Dim objMatchesDevDef          As MatchCollection
    Dim objMatchesDevID           As MatchCollection
    Dim objMatchesStrDefs         As MatchCollection
    Dim TimeScriptRun             As Long
    Dim TimeScriptFinish          As Long
    Dim strWorkDir                As String
    Dim strInfFullname            As String
    Dim strInfPath                As String
    Dim strInfFileName            As String
    Dim strInfPathTemp            As String
    Dim cmdString                 As String
    Dim i                         As Long
    Dim InfCount                  As Long
    Dim strValuer                 As String
    Dim strDevName                As String
    Dim strPackFileName_woExt     As String
    Dim strRezultTxt              As String
    Dim strRezultTxtTo            As String
    Dim strRezultTxtHwid          As String
    Dim strRezultTxtHwidTo        As String
    Dim strInfPathTempList_x()    As String
    Dim strDevID                  As String
    Dim strDrvDate                As String
    Dim strDrvVersion             As String
    Dim strDrvCatFileName         As String
    Dim strCatFileExists          As String
    Dim strValval                 As String
    Dim Strings                   As String
    Dim strRegEx_mansect          As String
    Dim strRegEx_strsect          As String
    Dim strRegEx_version          As String
    Dim strRegEx_devs_l           As String
    Dim strRegEx_devs_r           As String
    Dim strRegEx_devid            As String
    Dim strRegEx_mandef           As String
    Dim strRegEx_devdef           As String
    Dim strRegEx_strings          As String
    Dim strRegEx_sectnames        As String
    Dim FileContent               As String
    Dim strLinesArr()             As String
    Dim strLinesArrHwid()         As String
    Dim lngNumLines               As Long
    Dim strManufSection           As String
    Dim strKey                    As String
    Dim strValue                  As String
    Dim R                         As Boolean
    Dim strVarname                As String
    Dim strSections               As String
    Dim strSectlist               As String
    Dim ss                        As String
    Dim strBaseName               As String
    Dim j                         As Long
    Dim sB                        As String
    Dim K                         As Long
    Dim K2                        As Long
    Dim strK2Sectlist()           As String
    Dim strThisSection            As String
    Dim strDevIDs                 As String
    Dim Pos                       As Long
    Dim PosRev                    As Long
    Dim strVer                    As String
    Dim k1                        As Long
    Dim t                         As Long
    Dim a                         As String
    Dim strFileDBSize             As String
    Dim strSectNoCompatVerOSList  As String
    Dim strRegEx_catFile          As String
    Dim strArchCatFileList        As String
    Dim strArchCatFileListContent As String
    Dim strVarname_x()            As String
    Dim ii                        As Long

    DebugMode "DevParserByRegExp-Start"
    
    TimeScriptRun = GetTickCount
    Set RegExpStrSect = New RegExp
    Set RegExpStrDefs = New RegExp
    Set RegExpVerSect = New RegExp
    Set RegExpCatSect = New RegExp
    Set RegExpManSect = New RegExp
    Set RegExpManDef = New RegExp
    Set RegManID = New RegExp
    Set RegExpDevDef = New RegExp
    Set RegExpDevID = New RegExp
    Set RegExpDevSect = New RegExp
    Set objHashOutput = New Scripting.Dictionary
    
    'Имя папки с распакованными драйверами
    strPackFileName_woExt = FileName_woExt(FileNameFromPath(strPackFileName))
    'Рабочий каталог
    strWorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackFileName_woExt)
    'Если рабочий каталог уже есть, то удаляем его
    DoEvents

    If PathExists(strWorkDir) Then
        ChangeStatusTextAndDebug strMessages(81)
        DelRecursiveFolder (strWorkDir)
        DoEvents
    End If

    ' Каталог для распаковки inf файлов
    strInfPathTemp = strWorkTempBackSL & strPackFileName_woExt

    If PathExists(strInfPathTemp) = False Then
        CreateNewDirectory strInfPathTemp
    End If

    If Not mbDP_Is_aFolder Then
        ' Запуск распаковки
        cmdString = Kavichki & strArh7zExePATH & Kavichki & " x -yo" & Kavichki & strInfPathTemp & Kavichki & " -r " & Kavichki & strPathDRP & strPackFileName & Kavichki & " *.inf DriverPack*.ini"
        ChangeStatusTextAndDebug strMessages(72) & " " & strPackFileName

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else

            ' Архиватор отработал на все 100%? Если нет то сообщаем
            If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            ' Создаем спсиок файлов *.cat в архиве
            strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
            cmdString = "cmd.exe /c " & Kavichki & Kavichki & strArh7zExePATH & Kavichki & " l " & Kavichki & strPathDRP & strPackFileName & Kavichki & " -yr *.cat >" & Kavichki & strArchCatFileList & Kavichki

            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                DebugMode strMessages(13) & str2vbNewLine & cmdString
            End If
        End If

        ChangeStatusTextAndDebug strMessages(73) & " " & strPackFileName
        'Построение списка inf файлов в рабочем каталоге
        strInfPathTempList_x = SearchFilesInRoot(strInfPathTemp, "*.inf", True, False)
    Else
        ' Создаем спсиок файлов *.cat в архиве
        strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
        cmdString = "cmd.exe /c Dir " & Kavichki & strPathDRP & strPackFileName & "\*.cat" & Kavichki & " /B /S >" & Kavichki & strArchCatFileList & Kavichki

        'dir c:\windows\temp\*.tmp /S /B
        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            DebugMode strMessages(33) & str2vbNewLine & cmdString
        End If

        ChangeStatusTextAndDebug strMessages(148) & " " & strPackFileName
        'Построение списка inf файлов в рабочем каталоге
        strInfPathTempList_x = SearchFilesInRoot(strPathDRP & strPackFileName, "*.inf", True, False)
    End If

    If UBound(strInfPathTempList_x, 2) = 0 Then
        If LenB(strInfPathTempList_x(0, 0)) = 0 Then

            Exit Sub

        End If
    End If

    TimeScriptFinish = GetTickCount
    DebugMode "DevParserByRegExp-Time to Unpack Inf-file: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True), 1
    DoEvents
    ' sections
    strRegEx_mansect = "^[ ]*\[Manufacturer\](?:([\s\S]*?)^[ #]*(?=\[)|([\s\S]*))"
    strRegEx_strsect = "^[ ]*\[strings\](?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
    strRegEx_version = "^[ ]*\[Version\][\s\S]*?^[ ]*DriverVer[ ]*=[ ]*(%[^%]*%|(?:[\w/ ])+)(?:[ ]*,[ ]*(%[^%]*%|(?:[\w/ .])+))?"
    strRegEx_catFile = "^[ ]*\[Version\][\s\S]*?^[ ]*CatalogFile[.nt|.ntamd64|.ntx86]*[ ]*=[ ]*([^;\r\n]*)"
    'sections "Devices"
    strRegEx_devs_l = "^[ ]*\[("
    strRegEx_devs_r = ")\](?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
    ' ,ID,ID, ...
    strRegEx_devid = ",[ ]*([^ \r\n,][^ \r\n&,]+(?:&(?:[\w/]+))*)"
    ' name = sectname,suffix,suffix,...
    strRegEx_mandef = "^[ ]*[^;\r\n=]*=[ ]*([^;\r\n]*)"
    ' name = driver,ID,ID,...
    'strRegEx_devdef = "^[ ]*(?:%([^%\r\n]+)%|([^;=\r\n]+))[^=\r\n]*=[^\r\n,]*([^;\r\n]*)"
    strRegEx_devdef = "^[ ]*((?:[^;=\r\n]*(?:%[^%\r\n]+)*%[^;=\r\n]*)|(?:[^;=\r\n]+))[^=\r\n]*=[^\r\n,]*([^;\r\n]*)"
    ' variable = "str"
    strRegEx_strings = "^[ ]*([^; \r\n][^;=\r\n]*[^; \r\n])[ ]*=[ ]*(?:([^\r\n;]*))"
    ' variable = "sectnames"
    strRegEx_sectnames = "(?:,?[ ]*,?[ ]*([^,\r\n;]+[^,\r\n ;]))"
    objHashOutput.CompareMode = TextCompare

    ' Init regexps
    With RegExpStrSect
        .Pattern = strRegEx_strsect
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
        ' Note: "XP Alternative (by Greg)\D\3\M\A\12\prime.inf" has two [strings] sections
    End With

    With RegExpStrDefs
        .Pattern = strRegEx_strings
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpVerSect
        .Pattern = strRegEx_version
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With

    With RegExpCatSect
        .Pattern = strRegEx_catFile
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With

    With RegExpManSect
        .Pattern = strRegEx_mansect
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With

    With RegExpManDef
        .Pattern = strRegEx_mandef
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegManID
        .Pattern = strRegEx_sectnames
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpDevDef
        .Pattern = strRegEx_devdef
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpDevID
        .Pattern = strRegEx_devid
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpDevSect
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    ReDim strLinesArr(200000) As String
    ReDim strLinesArrHwid(200000) As String

    For InfCount = LBound(strInfPathTempList_x, 2) To UBound(strInfPathTempList_x, 2)
        strInfFullname = strInfPathTempList_x(0, InfCount)
        ' полный путь к файлу inf
        strInfPath = PathNameFromPath(strInfFullname)
        'If InIDE() Then
        'If InStr(1, strInfFullname, "FORCED\7x64\HP\E1D62x64.INF", vbTextCompare) Then Stop ' Debug .Assert strInfPath
        'End If
        ' Имя inf файла
        strInfFileName = LCase$(FileNameFromPath(strInfFullname))
        'Debug.Print strInfFileName
        ChangeStatusTextAndDebug strMessages(73) & " " & strPackFileName & " " & strMessages(124) & " (" & InfCount & " " & strMessages(124) & " " & (UBound(strInfPathTempList_x, 2) - LBound(strInfPathTempList_x, 2) + 1) & ": " & strInfFileName & ")"

        ' путь к файлу inf для записи в параметры
        If Not mbDP_Is_aFolder Then
            strInfPath = Replace$(strInfPath, strWorkDir, vbNullString, , , vbTextCompare)
        Else
            strInfPath = Replace$(strInfPath, BackslashAdd2Path(strPathDRP & strPackFileName), vbNullString, , , vbTextCompare)
        End If

        ' Read INF file
        FileContent = vbNullString
        strFileDBSize = strInfPathTempList_x(1, InfCount)

        If InStr(strFileDBSize, "0 ") = 1 Then
            DebugMode str2VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfFullname
        Else
            Set objInfFile = objFSO.OpenTextFile(strInfFullname, ForReading, False, TristateUseDefault)
            FileContent = objInfFile.ReadAll()
            objInfFile.Close

            ' Убираем символ """
            If InStr(FileContent, Kavichki) Then
                FileContent = Replace$(FileContent, Kavichki, vbNullString)
            End If

            If InStr(FileContent, vbTab) Then
                FileContent = Replace$(FileContent, vbTab, vbNullString)
            End If

            ' Чтение списка содержимого архива *.Cat
            strArchCatFileListContent = vbNullString

            If PathExists(strArchCatFileList) Then
                If GetFileSizeByPath(strArchCatFileList) > 0 Then
                    Set objCatFile = objFSO.OpenTextFile(strArchCatFileList, ForReading, False, TristateUseDefault)
                    strArchCatFileListContent = objCatFile.ReadAll()
                    objCatFile.Close
                Else
                    DebugMode str2VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strArchCatFileList
                End If
            End If
        End If

        ' Find [strings] section
        Strings = vbNullString
        Set objStringHash = New Scripting.Dictionary
        objStringHash.CompareMode = TextCompare
        Set objMatchesStrSect = RegExpStrSect.Execute(FileContent)

        If objMatchesStrSect.Count >= 1 Then
            Set objMatch = objMatchesStrSect.Item(0)
            Strings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
            'Debug.Print RegExpStrDefs.Pattern
            Set objMatchesStrDefs = RegExpStrDefs.Execute(Strings)

            For i = 0 To objMatchesStrDefs.Count - 1
                Set objMatch = objMatchesStrDefs.Item(i)
                strKey = Trim$(objMatch.SubMatches(0))
                strValue = Trim$(objMatch.SubMatches(1))
                R = objStringHash.Exists(strKey)

                If Not R Then
                    objStringHash.Add strKey, strValue
                    'Debug.Print strRegEx_strings
                    'Debug.Print strRegEx_strsect
                    objStringHash.Add Percentage & strKey & Percentage, strValue
                End If

            Next

        End If

        ' Find [version] section
        Set objMatchesVerSect = RegExpVerSect.Execute(FileContent)

        If objMatchesVerSect.Count >= 1 Then
            Set objMatch = objMatchesVerSect.Item(0)
            strDrvDate = objMatch.SubMatches(0)

            If InStr(strDrvDate, Percentage) Then
                strVarname = Left$(strDrvDate, InStrRev(strDrvDate, Percentage))
                strValval = objStringHash.Item(strVarname)

                If LenB(strValval) = 0 Then
                    DebugMode "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                Else
                    strDrvDate = Replace$(strDrvDate, strVarname, strValval)
                End If
            End If

            strDrvDate = Trim$(strDrvDate)
            strDrvVersion = objMatch.SubMatches(1)

            If InStr(strDrvVersion, Percentage) Then
                strVarname = Left$(strDrvVersion, InStrRev(strDrvVersion, Percentage))
                strValval = objStringHash.Item(strVarname)

                If LenB(strValval) = 0 Then
                    DebugMode "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                Else
                    strDrvVersion = Replace$(strDrvVersion, strVarname, strValval)
                End If
            End If

            If LenB(strDrvVersion) > 0 Then
                strVer = strDrvDate & "," & strDrvVersion
            Else

                If LenB(strDrvDate) > 0 Then
                    strVer = strDrvDate
                Else
                    strVer = "unknown"
                End If
            End If

        Else
            DebugMode "DevParserbyRegExp: Error in inf: Section [version] not found: " & strInfFullname
            strDrvDate = vbNullString
            strDrvVersion = vbNullString
            strVer = "unknown"
        End If

        ' Find [CatalogeFile] section
        Set objMatchesCatSect = RegExpCatSect.Execute(FileContent)

        If objMatchesCatSect.Count >= 1 Then
            Set objMatch = objMatchesCatSect.Item(0)
            strDrvCatFileName = objMatch.SubMatches(0)

            If InStr(strDrvCatFileName, Percentage) Then
                strVarname = Left$(strDrvCatFileName, InStrRev(strDrvCatFileName, Percentage))
                strValval = objStringHash.Item(strVarname)

                If LenB(strValval) = 0 Then
                    DebugMode "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                Else
                    strDrvCatFileName = Replace$(strDrvCatFileName, strVarname, strValval)
                End If
            End If

            ' Если ли файл *.cat в списке файлов архива?
            If InStr(1, strDrvCatFileName, ".cat", vbTextCompare) Then
                If InStr(1, strArchCatFileListContent, strInfPath & strDrvCatFileName, vbTextCompare) Then
                    strCatFileExists = 1
                Else
                    strCatFileExists = 0
                End If

            Else
                strCatFileExists = 0
            End If

        Else
            DebugMode "DevParserbyRegExp: Error in inf: Section [CatalogeFile] not found: " & strInfFullname
            strDrvCatFileName = vbNullString
            strCatFileExists = 0
        End If

        ' Find [manufacturer] section
        Set objMatchesManSect = RegExpManSect.Execute(FileContent)

        If objMatchesManSect.Count >= 1 Then
            Set objMatch = objMatchesManSect.Item(0)
            strSections = objMatch.SubMatches(0) & objMatch.SubMatches(1)
            strSectlist = vbNullString
            Set objMatchesManDef = RegExpManDef.Execute(strSections)

            If objMatchesManDef.Count = 0 Then
                DebugMode "DevParserbyRegExp: Error in inf: Parametr in section [Manufacturer] not match 'name = sectname,suffix,suffix'. Inf-File=" & strInfFullname, 1

                If InStr(strSectlist, vbNewLine) Then
                    strSectlist = Replace$(strSectlist, vbNewLine, vbNullString)
                End If

                DebugMode "DevParserbyRegExp: Try seek in section [Manufacturer] parametr: " & strSectlist, 2
            Else

                For i = 0 To objMatchesManDef.Count - 1
                    Set objMatch = objMatchesManDef.Item(i)
                    ss = objMatch.SubMatches(0)
                    Set objMatchesManID = RegManID.Execute(ss)
                    strBaseName = vbNullString

                    'found =0
                    For j = 0 To objMatchesManID.Count - 1
                        Set objMatch1 = objMatchesManID.Item(j)
                        sB = objMatch1.SubMatches(0)
                        sB = RTrim$(sB)

                        If i <> 0 Or j <> 0 Then
                            strSectlist = strSectlist & "|"
                        End If

                        If j = 0 Then
                            strBaseName = sB
                            strSectlist = strSectlist & sB
                        Else
                            strSectlist = strSectlist & (strBaseName & "." & sB)
                        End If

                    Next
                Next

            End If

            'Debug.Print strSectlist
            strK2Sectlist = Split(strSectlist, "|")
            ' Переменная несовместымых ОС для данного инфа
            strSectNoCompatVerOSList = vbNullString

            If InStr(strSectlist, "|") Then
                strSectNoCompatVerOSList = GetIniEmptySectionFromList(strSectlist, strInfFullname)
            Else
                strSectNoCompatVerOSList = "-"
            End If

            'Debug.Print strSectlist & vbNewLine & strSectNoCompatVerOSList
            For K2 = 0 To UBound(strK2Sectlist)
                RegExpDevSect.Pattern = strRegEx_devs_l & strK2Sectlist(K2) & strRegEx_devs_r
                'Debug.Print RegExpDevSect.Pattern
                Set objMatchesDevSect = RegExpDevSect.Execute(FileContent)

                For K = 0 To objMatchesDevSect.Count - 1
                    Set objMatch = objMatchesDevSect.Item(K)
                    strThisSection = objMatch.SubMatches(1) & objMatch.SubMatches(2)
                    strManufSection = UCase$(objMatch.SubMatches(0))
                    ' Find device definitions
                    Set objMatchesDevDef = RegExpDevDef.Execute(strThisSection)

                    'Debug.Print RegExpDevDef.Pattern
                    ' Если секция пустая, то установка из данного файла запрещена на данной системе
                    If objMatchesDevDef.Count = 0 Then
                        DebugMode "DevParserByRegExp: Section [" & strManufSection & "] is Empty -> this OS not Supported by inf: " & strInfPath & strInfFileName
                    End If

                    ' Handle definition
                    For i = 0 To objMatchesDevDef.Count - 1
                        Set objMatch = objMatchesDevDef.Item(i)
                        strDevIDs = objMatch.SubMatches(1)
                        strDevName = Trim$(objMatch.SubMatches(0))

                        'Debug.Print strDevName
                        If LenB(strDevName) > 0 Then
                            Pos = InStr(strDevName, Percentage)
                            strValval = vbNullString

                            If Pos > 0 Then
                                PosRev = InStrRev(strDevName, Percentage)

                                If Pos <> PosRev Then
                                    strVarname = Mid$(strDevName, Pos + 1, PosRev - 2)

                                    If InStr(strVarname, Percentage) Then
                                        strVarname_x = Split(strVarname, Percentage)

                                        For ii = LBound(strVarname_x) To UBound(strVarname_x)
                                            strValval = AppendStr(strValval, objStringHash.Item(strVarname_x(ii)))
                                        Next ii

                                    Else
                                        strValval = objStringHash.Item(strVarname)
                                    End If

                                Else
                                    strVarname = Replace$(strDevName, Percentage, vbNullString)
                                    strValval = objStringHash.Item(strVarname)
                                End If

                                If LenB(strValval) = 0 Then
                                    DebugMode "DevParserByRegExp: Cannot find '" & strVarname & "'"
                                    strDevName = strVarname
                                Else
                                    strDevName = Replace$(strDevName, "%" & strVarname & "%", strValval)
                                    ' Если все таки есть процент, то есть не определился из cекции String
                                End If
                            End If

                            If InStr(strDevName, Percentage) Then
                                '                            strDevName = "not defined in the inf"
                                strDevName = Replace$(strDevName, Percentage, vbNullString)
                                strDevName = objStringHash.Item(strDevName)
                            End If

                            ' На случай если есть юникодовые символы в имени устройства
                            a = vbNullString

                            For k1 = 1 To Len(strDevName)
                                t = Asc(Mid$(strDevName, k1, 1))
                                a = a + Chr$(t)
                            Next

                            '
                            ' Если требуется то удаление лишних символов
                            strDevName = ReplaceBadSymbol(a)
                        Else
                            DebugMode "Error in inf: " & strInfFullname & " (Variable not defined: " & objMatch.SubMatches(0) & ")"
                            strDevName = "not defined in the inf"
                            '                                If mbIsDesignMode Then
                            '                                    Debug.Print "Not defined variable in [Strings] - " & strPackFileName & vbTab & strInfPath & strInfFileName & vbTab & objMatch.SubMatches(0) & vbTab & objMatchesDevDef.item(i)
                            '                                End If
                        End If

                        ' add IDs
                        Set objMatchesDevID = RegExpDevID.Execute(strDevIDs)

                        For j = 0 To objMatchesDevID.Count - 1
                            Set objMatch = objMatchesDevID.Item(j)
                            strValuer = objMatch.SubMatches(0)

                            If InStr(strValuer, Percentage) Then
                                strVarname = Left$(strValuer, InStrRev(strValuer, Percentage))

                                If InStr(strVarname, Percentage) > 1 Then
                                    strVarname = Right$(strVarname, Len(strVarname) - InStr(strValuer, Percentage) + 1)
                                End If

                                strValval = objStringHash.Item(strVarname)

                                If LenB(strValval) = 0 Then
                                    DebugMode "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                Else
                                    strValuer = Replace$(strValuer, strVarname, strValval)
                                End If
                            End If

                            strValuer = UCase$(Trim$(strValuer))
                            strDevID = ParseDoubleHwid(strValuer)
                            ss = strDevID & vbTab & strInfPath & vbTab & strInfFileName & vbTab & strManufSection
                            R = objHashOutput.Exists(ss)

                            If Not R Then
                                objHashOutput.Item(ss) = "+"

                                'Итоговая строка
                                If InStr(strVer, " ") Then
                                    strVer = Replace$(strVer, " ", vbNullString)
                                End If

                                strLinesArr(lngNumLines) = ss & (vbTab & strVer & vbTab & strSectNoCompatVerOSList & vbTab & strCatFileExists & vbTab & strDevName)
                                strLinesArrHwid(lngNumLines) = strDevID
                                lngNumLines = lngNumLines + 1
                            End If

                        Next

                        ' strDevIDs'
                    Next

                    ' dev_defs'
                Next

                ' dev_Sub_sects'
            Next

            ' dev_sects'
        End If

        ' sect_list
        objHashOutput.RemoveAll
        objStringHash.RemoveAll
    Next

    ChangeStatusTextAndDebug strMessages(121) & " " & strPackFileName
    strRezultTxt = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".txt"
    strRezultTxtHwid = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".hwid"
    strRezultTxtTo = Replace$(PathCombine(strPathDevDB, FileNameFromPath(strRezultTxt)), "rezult", vbNullString, , , vbTextCompare)
    strRezultTxtHwidTo = Replace$(PathCombine(strPathDevDB, FileNameFromPath(strRezultTxtHwid)), "rezult", vbNullString, , , vbTextCompare)
    TimeScriptFinish = GetTickCount
    DebugMode "DevParserByRegExp-Time to Create Index Data: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True), 1

    ' Если данные найдены, то выводим итог в файл
    If lngNumLines > 0 Then

        ReDim Preserve strLinesArr(lngNumLines - 1)
        ReDim Preserve strLinesArrHwid(lngNumLines - 1)

        ' сортируем массивы
        TimeScriptRun = GetTickCount
        ShellSortAny VarPtr(strLinesArr(0)), lngNumLines, 4&, AddressOf CompareString
        ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLines, 4&, AddressOf CompareString
        TimeScriptFinish = GetTickCount
        DebugMode "DevParserByRegExp-Time to Sort Index: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True) & vbNewLine & _
                  "DevParserByRegExp--Finished reading inf-files"
        DoEvents
        TimeScriptRun = GetTickCount

        '---------------------------------------------
        '---------------Выводим итог в файл-----
        If PathExists(strPathDevDB) = False Then
            CreateNewDirectory strPathDevDB
        End If

        'Debug.Print strLinesArr(lngNumLines)
        Set objRezultFile = objFSO.CreateTextFile(strRezultTxt, True)
        objRezultFile.Write Join(strLinesArr(), vbNewLine)
        Set objRezultFileHWID = objFSO.CreateTextFile(strRezultTxtHwid, True)
        objRezultFileHWID.Write Join(strLinesArrHwid(), vbNewLine)
        TimeScriptFinish = GetTickCount
        DebugMode "DevParserByRegExp-Time to Save Index 2 File: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True), 1
        ' Удаление массива, т.е освобождение памяти
        Erase strLinesArr
        Erase strLinesArrHwid

        If CopyFileTo(strRezultTxt, strRezultTxtTo) = False Then
            MsgBox strMessages(31), vbInformation, strProductName
            DebugMode "DevParserByRegExp--Error of the saving file in directory database driver: " & strRezultTxtTo
        Else

            'Копируем файл HWID
            If CopyFileTo(strRezultTxtHwid, strRezultTxtHwidTo) = False Then
                MsgBox strMessages(31), vbInformation, strProductName
                DebugMode "DevParserByRegExp--Error of the saving file in directory database driver: " & strRezultTxtHwidTo
            Else
                ' Записываем версию базы драйверов в ini-файл
                IniWriteStrPrivate strPackFileName_woExt, "Version", lngDevDBVersion, PathCombine(strPathDevDB, "DevDBVersions.ini")
                'IniWriteStrPrivate strPackFileName_woExt, "FullHwid", CStr(Abs(Not mbDelDouble)), strPathDevDB & "DevDBVersions.ini"
                'Ищем файл DriverPack*.ini
                strRezultTxt = vbNullString
                strRezultTxt = SearchFilesInRoot(strInfPathTemp, "DriverPack*.ini", False, True)

                ' Копируем DriverPack*.ini в каталог базы данных
                If PathExists(strRezultTxt) Then
                    If PathExists(strRezultTxtTo) Then
                        strRezultTxtTo = Replace$(strRezultTxtTo, ".txt", ".ini", , , vbTextCompare)

                        If CopyFileTo(strRezultTxt, strRezultTxtTo) = False Then
                            DebugMode "DevParserByRegExp--Error of the saving file in directory database driver: " & strRezultTxtTo
                        Else
                            DebugMode "DevParserByRegExp--Save DPFinish file: " & strRezultTxtTo
                        End If

                    Else
                        DebugMode "DevParserByRegExp--Error of the saving file in directory database driver: " & strRezultTxtTo
                    End If
                End If
            End If
        End If
    End If

    If Not (objRezultFile Is Nothing) Then
        objRezultFile.Close
    End If

    If Not (objRezultFileHWID Is Nothing) Then
        objRezultFileHWID.Close
    End If

    DebugMode "DevParserByRegExp-End"
End Sub
