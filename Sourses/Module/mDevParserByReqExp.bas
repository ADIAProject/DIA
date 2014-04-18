Attribute VB_Name = "mDevParserByReqExp"
Option Explicit

' Текущая версия базы данных
Public Const lngDevDBVersion        As Long = 6

' Рабочие переменные
Private RegExpStrSect       As RegExp
Private RegExpStrDefs       As RegExp
Private RegExpVerSect       As RegExp
Private RegExpVerParam      As RegExp
Private RegExpCatParam      As RegExp
Private RegExpManSect       As RegExp
Private RegExpManDef        As RegExp
Private RegManID            As RegExp
Private RegExpDevDef        As RegExp
Private RegExpDevSect       As RegExp
Private RegExpReplace       As RegExp
Private objHashOutput       As Scripting.Dictionary
Private objStringHash       As Scripting.Dictionary
Private objHWIDOutput       As Scripting.Dictionary
Private cSortHWID           As cBlizzard

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DevParserByRegExp
'! Description (Описание)  :   [Парсинг ID и названий устройст из inf-файла и построение БД]
'! Parameters  (Переменные):   strPackFileName (String)
'                              strPathDRP (String)
'                              strPathDevDB (String)
'!--------------------------------------------------------------------------------
Public Sub DevParserByRegExp(ByVal strPackFileName As String, ByVal strPathDRP As String, ByVal strPathDevDB As String)

    Dim objMatch                  As Match
    Dim objMatch1                 As Match
    Dim objMatchesStrSect         As MatchCollection
    Dim objMatchesVerSect         As MatchCollection
    Dim objMatchesVerParam        As MatchCollection
    Dim objMatchesCatParam        As MatchCollection
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
    Dim infNum                    As Long
    Dim infCount                  As Long
    Dim strValuer                 As String
    Dim strValuer_x()             As String
    Dim strDevName                As String
    Dim strPackFileName_woExt     As String
    Dim strRezultTxt_x()          As FindListStruct
    Dim strInfPathTempList_x()    As FindListStruct
    Dim strRezultTxt              As String
    Dim strRezultTxtTo            As String
    Dim strRezultTxtHwid          As String
    Dim strRezultTxtHwidTo        As String
    Dim strDevID                  As String
    Dim strDrvDate                As String
    Dim strDrvVersion             As String
    Dim strDrvCatFileName         As String
    Dim lngCatFileExists          As Long
    Dim strValval                 As String
    Dim sStrings                  As String
    Dim strRegEx_devs_l           As String
    Dim strRegEx_devs_r           As String
    Dim sFileContent              As String
    Dim sVerSectContent           As String
    Dim strLinesArr()             As String
    Dim strLinesArrHwid()         As String
    Dim lngNumLines               As Long
    Dim lngNumLinesHwid           As Long
    Dim strManufSection           As String
    Dim strKey                    As String
    Dim strKeyPercent             As String
    Dim strValue                  As String
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
    Dim strSectEmptyList          As String
    Dim strSectEmptyList4Check    As String
    Dim strArchCatFileList        As String
    Dim strArchCatFileListContent As String
    Dim strVarname_x()            As String
    Dim ii                        As Long
    Dim strUnpackMask             As String
    Dim strInfPathRelative        As String
    Dim strInfPathTabQuoted       As String
    Dim strPartString2Index       As String
    Dim strDevIDs_x()             As String
    Dim mbDevNameIsCollected      As Boolean
    Dim strVerTemp                As String
    Dim strVerTemp_x()            As String
    
    If mbDebugStandart Then DebugMode vbTab & "DevParserByRegExp-Start"
    
    TimeScriptRun = GetTickCount
    
    ' Hash-таблица уникальности значения strDevID & strManufSection в рамках inf-файла
    Set objHashOutput = New Scripting.Dictionary
    objHashOutput.CompareMode = BinaryCompare
    ' Hash-таблица уникальности значений секции String в рамках inf-файла
    Set objStringHash = New Scripting.Dictionary
    objStringHash.CompareMode = BinaryCompare
    ' Hash-таблица уникальности значений HWID в рамках пакета-драйверов
    Set objHWIDOutput = New Scripting.Dictionary
    objHWIDOutput.CompareMode = BinaryCompare
    
    ' Должно ускорить распаковку, если выключено чтение файла finish.ini
    If Not mbLoadFinishFile Then
        strUnpackMask = " *.inf"
    Else
        strUnpackMask = " *.inf DriverPack*.ini"
    End If
    
    'Имя папки с распакованными драйверами
    strPackFileName_woExt = GetFileName_woExt(GetFileNameFromPath(strPackFileName))
    'Рабочий каталог
    strWorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackFileName_woExt)
    'Если рабочий каталог уже есть, то удаляем его
    DoEvents

    If PathExists(strWorkDir) Then
        ChangeStatusBarText strMessages(81)
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
        cmdString = strQuotes & strArh7zExePATH & strQuotes & " x -yo" & strQuotes & strInfPathTemp & strQuotes & " -r " & strQuotes & strPathDRP & strPackFileName & strQuotes & strUnpackMask
        ChangeStatusBarText strMessages(72) & strSpace & strPackFileName

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else

            ' Архиватор отработал на все 100%? Если нет то сообщаем
            If lngExitProc = 2 Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            ElseIf lngExitProc = 7 Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            ElseIf lngExitProc = 255 Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            ' Создаем спсиок файлов *.cat в архиве
            strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
            cmdString = "cmd.exe /c " & strQuotes & strQuotes & strArh7zExePATH & strQuotes & " l " & strQuotes & strPathDRP & strPackFileName & strQuotes & " -yr *.cat >" & strQuotes & strArchCatFileList & strQuotes
            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                If mbDebugStandart Then DebugMode strMessages(13) & str2vbNewLine & cmdString
            End If
        End If

        ChangeStatusBarText strMessages(73) & strSpace & strPackFileName
        'Построение списка inf файлов в рабочем каталоге
        strInfPathTempList_x = SearchFilesInRoot(strInfPathTemp, "*.inf", True, False)
    Else
        ' Создаем список файлов *.cat в архиве
        strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
        cmdString = "cmd.exe /c Dir " & strQuotes & strPathDRP & strPackFileName & "\*.cat" & strQuotes & " /A- /B /S >" & strQuotes & strArchCatFileList & strQuotes

        'dir c:\windows\temp\*.tmp /S /B
        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            If mbDebugStandart Then DebugMode strMessages(33) & str2vbNewLine & cmdString
        End If

        ChangeStatusBarText strMessages(148) & strSpace & strPackFileName
        'Построение списка inf файлов в рабочем каталоге
        strInfPathTempList_x = SearchFilesInRoot(strPathDRP & strPackFileName, "*.inf", True, False)
    End If

    If UBound(strInfPathTempList_x) = 0 Then
        If LenB(strInfPathTempList_x(0).FullPath) = 0 Then
            Exit Sub
        End If
    End If

    TimeScriptFinish = GetTickCount
    If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Unpack Inf-file: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
    DoEvents
            
    ' sections [Strings]
    Set RegExpStrSect = New RegExp
    With RegExpStrSect
        .Pattern = "^[ ]*\[strings\][ ]*[ \S]*$(?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With

    ' sections [Version]
    Set RegExpVerSect = New RegExp
    With RegExpVerSect
        .Pattern = "^[ ]*\[version\][ ]*[ \S]*$(?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With
    
    ' sections [Manufacturer]
    Set RegExpManSect = New RegExp
    With RegExpManSect
        .Pattern = "^[ ]*\[manufacturer\][ ]*[ \S]*$(?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
    End With
    
    'sections "Devices"
    Set RegExpDevSect = New RegExp
    With RegExpDevSect
        strRegEx_devs_l = "^[ ]*\["
        strRegEx_devs_r = "\][ ]*[ \S]*$(?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
        '.Pattern = strRegEx_devs_l & strManufSection & strRegEx_devs_r
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With
    
    ' sections [Strings] - variable = param
    Set RegExpStrDefs = New RegExp
    With RegExpStrDefs
        .Pattern = "^[ ]*([^ \r\n][^=\r\n]*[^ \r\n])[ ]*=[ ]*(?:([^\r\n;]*))"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    
    ' sections [Strings] - parametr driverver=param
    Set RegExpVerParam = New RegExp
    With RegExpVerParam
        .Pattern = "^[ ]*driverver[ ]*=[ ]*(?:([^\r\n;]*))"
        .MultiLine = True
        .IgnoreCase = False
        .Global = False
    End With

    ' sections [Strings] - parametr catalogfile=param
    Set RegExpCatParam = New RegExp
    With RegExpCatParam
        .Pattern = "^[ ]*catalogfile[.nt|.ntamd64|.ntx86|.ntia64]*[ ]*=[ ]*(?:([^\r\n;]*))"
        .MultiLine = True
        .IgnoreCase = False
        .Global = False
    End With

    ' sections [Manufacturer] - name = sectname,suffix,suffix,...
    Set RegExpManDef = New RegExp
    With RegExpManDef
        .Pattern = "^[ ]*[^\r\n=]*=[ ]*([^;\r\n]*)"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With

    ' variable = "sectnames"
    Set RegManID = New RegExp
    With RegManID
        .Pattern = "(?:,?[ ]*,?[ ]*([^,\r\n;]+[^,\r\n ;]))"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    
    ' name = driver,ID,ID,...
    Set RegExpDevDef = New RegExp
    With RegExpDevDef
        .Pattern = "^[ ]*((?:(?:%[^%\r\n,]+)*%[^ ;=]*)|(?:[^;=\r\n]+))[^=\r\n]*=[^\r\n,]*[, ]*[ ]*((?:[^;\r\n]*))"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With
    
    ' Чистка файла inf от строк коментариев начинаются с ";#"
    Set RegExpReplace = New RegExp
    With RegExpReplace
        .Pattern = "^([ ]*[;#]+[ \S]*)$"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With

    ReDim strLinesArr(150000)
    ReDim strLinesArrHwid(50000)
    
    ' Чтение списка содержимого архива *.Cat
    strArchCatFileListContent = vbNullString

    If PathExists(strArchCatFileList) Then
        If GetFileSizeByPath(strArchCatFileList) Then
            FileReadData strArchCatFileList, strArchCatFileListContent
            strArchCatFileListContent = LCase$(strArchCatFileListContent)
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strArchCatFileList
        End If
    End If
        
    infCount = UBound(strInfPathTempList_x) + 1
    ChangeStatusBarText strMessages(73) & strSpace & strPackFileName & " (" & infCount & " inf-files)"
    
    ' Запускаем цикл обработки inf-файлов
    For infNum = 0 To UBound(strInfPathTempList_x)
        
            If strInfPathTempList_x(infNum).Size Then
            
            ' полный путь к файлу inf
            strInfFullname = strInfPathTempList_x(infNum).FullPath
            ' Имя inf файла
            strInfFileName = strInfPathTempList_x(infNum).NameLcase
            
            If (infNum Mod 20) = 0 Then
                ChangeStatusBarText strMessages(73) & strSpace & strPackFileName & " (" & infNum & strSpace & strMessages(124) & strSpace & infCount & ": " & strInfFileName & ")"
            Else
                If GetInputState Then
                    DoEvents
                End If
            End If
        
            ' Очистка буфера значений уникальных строк HWID
            Set objHashOutput = New Scripting.Dictionary
            ' Очистка буфера значений секции strings
            Set objStringHash = New Scripting.Dictionary
            
            ' путь к файлу inf для записи в параметры - Каталог где лежит inf-файл
            strInfPath = strInfPathTempList_x(infNum).RelativePath
            strInfPathRelative = strInfPathTempList_x(infNum).RelativePath & strInfFileName
            strInfPathTabQuoted = vbTab & strInfPathRelative & vbTab
            
            ' Read INF file
            FileReadData strInfFullname, sFileContent

            ' Убираем символ """
            If InStr(sFileContent, strQuotes) Then
                sFileContent = Replace$(sFileContent, strQuotes, vbNullString)
            End If

            If InStr(sFileContent, vbTab) Then
                sFileContent = Replace$(sFileContent, vbTab, vbNullString)
            End If
            
            ' Удаляем строки с ; или # в начале и пустые строки
            sFileContent = RegExpReplace.Replace(sFileContent, vbNewLine)
                        
            ' Find [strings] section
            sStrings = vbNullString
            Set objMatchesStrSect = RegExpStrSect.Execute(sFileContent)
    
            If objMatchesStrSect.Count Then
                Set objMatch = objMatchesStrSect.Item(0)
                
                sStrings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
                Set objMatchesStrDefs = RegExpStrDefs.Execute(sStrings)
    
                For i = 0 To objMatchesStrDefs.Count - 1
                    Set objMatch = objMatchesStrDefs.Item(i)
                    strKey = Trim$(LCase$(objMatch.SubMatches(0)))
                    strValue = Trim$(objMatch.SubMatches(1))
    
                    If Not objStringHash.Exists(strKey) Then
                        objStringHash.Add strKey, strValue
                        strKeyPercent = strPercent & strKey & strPercent
                        If Not objStringHash.Exists(strKeyPercent) Then
                            objStringHash.Add strKeyPercent, strValue
                        End If
                    End If
    
                Next
    
            End If
    
            ' Find [version] section
            Set objMatchesVerSect = RegExpVerSect.Execute(sFileContent)
            If objMatchesVerSect.Count Then
                sVerSectContent = LCase$(objMatchesVerSect.Item(0))
            
                ' Find DriverVer parametr
                Set objMatchesVerParam = RegExpVerParam.Execute(sVerSectContent)
        
                If objMatchesVerParam.Count Then
                    Set objMatch = objMatchesVerParam.Item(0)
                    strVerTemp = objMatch.SubMatches(0)
                    'strDrvDate = objMatch.SubMatches(0)
                    
                    If InStr(strVerTemp, strPercent) Then
                        If InStr(strVerTemp, strComma) Then
                            strVerTemp_x = Split(strVerTemp, strComma)
                            strDrvDate = Trim$(strVerTemp_x(0))
                            strDrvVersion = Trim$(strVerTemp_x(1))
                        Else
                            strDrvDate = Trim$(strVerTemp)
                        End If
                        
                        If InStr(strDrvDate, strPercent) Then
                            strVarname = Left$(strDrvDate, InStrRev(strDrvDate, strPercent))
                            strValval = objStringHash.Item(strVarname)
            
                            If LenB(strValval) Then
                                strDrvDate = Replace$(strDrvDate, strVarname, strValval)
                            Else
                                If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                            End If
                        End If
            
                        'strDrvVersion = objMatch.SubMatches(1)
            
                        If InStr(strDrvVersion, strPercent) Then
                            strVarname = Left$(strDrvVersion, InStrRev(strDrvVersion, strPercent))
                            strValval = objStringHash.Item(strVarname)
            
                            If LenB(strValval) Then
                                strDrvVersion = Replace$(strDrvVersion, strVarname, strValval)
                            Else
                                If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                            End If
                        End If
                        
                        If LenB(strDrvVersion) Then
                            strVer = strDrvDate & strComma & strDrvVersion
                        Else
            
                            If LenB(strDrvDate) Then
                                strVer = strDrvDate
                            Else
                                strVer = strUnknownLCase
                            End If
                        End If
            
                        If InStr(strVer, strSpace) Then
                            strVer = Replace$(strVer, strSpace, vbNullString)
                        End If
                    Else
                        If InStr(strVerTemp, strComma) Then
                            strVerTemp_x = Split(strVerTemp, strComma)
                            strDrvDate = Trim$(strVerTemp_x(0))
                            strDrvVersion = Trim$(strVerTemp_x(1))
                        Else
                            strDrvDate = Trim$(strVerTemp)
                        End If
                        
                        If LenB(strDrvVersion) Then
                            strVer = strDrvDate & strComma & strDrvVersion
                        Else
            
                            If LenB(strDrvDate) Then
                                strVer = strDrvDate
                            Else
                                strVer = strUnknownLCase
                            End If
                        End If
                        
                        If InStr(strVer, strSpace) Then
                            strVer = Replace$(strVer, strSpace, vbNullString)
                        End If
                    End If
                Else
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr 'DriverVer' not found: " & strInfFullname
                    strDrvDate = vbNullString
                    strDrvVersion = vbNullString
                    strVer = strUnknownLCase
                End If
        
                ' Find CatalogFile parametr
                Set objMatchesCatParam = RegExpCatParam.Execute(sVerSectContent)
        
                If objMatchesCatParam.Count Then
                    Set objMatch = objMatchesCatParam.Item(0)
                    strDrvCatFileName = objMatch.SubMatches(0)
                    
                    If InStr(strDrvCatFileName, strPercent) Then
                        strVarname = Left$(strDrvCatFileName, InStrRev(strDrvCatFileName, strPercent))
                        strValval = objStringHash.Item(strVarname)
        
                        If LenB(strValval) Then
                            strDrvCatFileName = Replace$(strDrvCatFileName, strVarname, strValval)
                        Else
                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                        End If
                    End If
                    strDrvCatFileName = LCase$(strDrvCatFileName)
        
                    ' Если ли файл *.cat в списке файлов архива?
                    If InStr(strDrvCatFileName, ".cat") Then
                        If InStr(1, strArchCatFileListContent, LCase$(strInfPath) & strDrvCatFileName) Then
                            lngCatFileExists = 1
                        Else
                            lngCatFileExists = 0
                        End If
        
                    Else
                        lngCatFileExists = 0
                    End If
        
                Else
                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr 'CatalogeFile' not found: " & strInfFullname
                    strDrvCatFileName = vbNullString
                    lngCatFileExists = 0
                End If
            Else
                If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Section [version] not found: " & strInfFullname
                strDrvDate = vbNullString
                strDrvVersion = vbNullString
                strVer = strUnknownLCase
                lngCatFileExists = 0
            End If
            
            ' Find [manufacturer] section
            Set objMatchesManSect = RegExpManSect.Execute(sFileContent)
    
            If objMatchesManSect.Count Then
                Set objMatch = objMatchesManSect.Item(0)
                strSections = objMatch.SubMatches(0) & objMatch.SubMatches(1)
                strSectlist = vbNullString
                Set objMatchesManDef = RegExpManDef.Execute(strSections)
    
                If objMatchesManDef.Count Then
                
                    For i = 0 To objMatchesManDef.Count - 1
                        Set objMatch = objMatchesManDef.Item(i)
                        ss = objMatch.SubMatches(0)
                        Set objMatchesManID = RegManID.Execute(ss)
                        strBaseName = vbNullString
    
                        For j = 0 To objMatchesManID.Count - 1
                            Set objMatch1 = objMatchesManID.Item(j)
                            sB = RTrim$(objMatch1.SubMatches(0))
    
                            If i <> 0 Then
                                strSectlist = strSectlist & "|"
                            ElseIf j <> 0 Then
                                strSectlist = strSectlist & "|"
                            End If
    
                            If j = 0 Then
                                strBaseName = sB
                                strSectlist = strSectlist & sB
                            Else
                                strSectlist = strSectlist & (strBaseName & strDot & sB)
                            End If
    
                        Next
                        strSectlist = UCase$(strSectlist)
                    Next
    
                Else
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr in section [Manufacturer] not match 'name = sectname,suffix,suffix'. Inf-File=" & strInfFullname
    
                    If InStr(strSectlist, vbNewLine) Then
                        strSectlist = Replace$(strSectlist, vbNewLine, vbNullString)
                    End If
    
                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Try seek in section [Manufacturer] parametr: " & strSectlist
                End If
    
                ' Переменная несовместымых ОС для данного инфа
                If InStr(strSectlist, "|") Then
                    strK2Sectlist = Split(strSectlist, "|")
                    strSectEmptyList = GetIniEmptySectionFromList(strSectlist, strInfFullname)
                Else
                    ReDim strK2Sectlist(0)
                    strK2Sectlist(0) = strSectlist
                    strSectEmptyList = strDash
                End If
                
                ' Часть строки которая будет позже добавлена в индекс
                strPartString2Index = vbTab & strVer & vbTab & strSectEmptyList & vbTab & lngCatFileExists & vbTab
                            
                strSectEmptyList4Check = strSectEmptyList & strComma
    
                For K2 = 0 To UBound(strK2Sectlist)
                    ' Если секция пустая, то пропускаем ее обработку (список пустых секций получен ранее)
                    strManufSection = strK2Sectlist(K2)
                    If InStr(strSectEmptyList4Check, strManufSection & strComma) = 0 Then
                        RegExpDevSect.Pattern = strRegEx_devs_l & strManufSection & strRegEx_devs_r
                        Set objMatchesDevSect = RegExpDevSect.Execute(sFileContent)
                    
                        ' Если совпадения найдены
                        If objMatchesDevSect.Count Then
                            For K = 0 To objMatchesDevSect.Count - 1
                                Set objMatch = objMatchesDevSect.Item(K)
                                strThisSection = objMatch.SubMatches(0) & objMatch.SubMatches(1)
                                
                                ' Find device definitions
                                Set objMatchesDevDef = RegExpDevDef.Execute(strThisSection)
            
                                ' Если секция не пустая, то
                                If objMatchesDevDef.Count Then
                                    ' Handle definition
                                    For i = 0 To objMatchesDevDef.Count - 1
                                        Set objMatch = objMatchesDevDef.Item(i)
                                        strDevIDs = objMatch.SubMatches(1)
                                        If InStr(strDevIDs, vbCr) Then
                                            strDevIDs = Replace$(strDevIDs, vbCr, vbNullString)
                                        End If
                                        strDevName = Trim$(objMatch.SubMatches(0))
                                        mbDevNameIsCollected = False
                            
                                        ' add IDs
                                        If InStr(strDevIDs, strComma) Then
                    
                                            strDevIDs_x = Split(strDevIDs, strComma)
                                            For j = 0 To UBound(strDevIDs_x)
    
                                                strValuer = strDevIDs_x(j)
                                                
                                                If InStr(strValuer, strSpace) Then
                                                    strValuer = Trim$(strValuer)
                                                    Pos = InStr(strValuer, strSpace)
                                                    If Pos Then
                                                        strValuer = Left$(strValuer, Pos - 1)
                                                    End If
                                                End If
                                                                
                                                If LenB(strValuer) > 8 Then
                                                    If InStr(strValuer, strPercent) Then
                                                        strVarname = Left$(strValuer, InStrRev(strValuer, strPercent))
                        
                                                        If InStr(strVarname, strPercent) > 1 Then
                                                            strVarname = Right$(strVarname, Len(strVarname) - InStr(strValuer, strPercent) + 1)
                                                        End If
                        
                                                        strValval = objStringHash.Item(LCase$(strVarname))
                        
                                                        If LenB(strValval) Then
                                                            strValuer = Replace$(strValuer, strVarname, strValval)
                                                        Else
                                                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                                        End If
                                                    End If
                        
                                                    strDevID = UCase$(strValuer)
                                                    
                                                    ' разбиваем Hwid по "\" - оставляем только xxx\yyy
                                                    If InStr(strDevID, vbBackslash) Then
                                                        strValuer_x = Split(strDevID, vbBackslash)
                                                        strDevID = strValuer_x(0) & vbBackslash & strValuer_x(1)
                                                    End If
                    
                                                    If InStr(strDevID, strSpace) Then
                                                        strDevID = Trim$(strDevID)
                                                    End If
                                                    
                                                    ss = strDevID & strManufSection
                        
                                                    If Not objHashOutput.Exists(ss) Then
                                                        objHashOutput.Item(ss) = 1
                                                        
                                                        ' Обработаем имя устройства
                                                        If Not mbDevNameIsCollected Then
                                                            If LenB(strDevName) Then
                                                                Pos = InStr(strDevName, strPercent)
                                                                strValval = vbNullString
                                    
                                                                If Pos Then
                                                                    PosRev = InStrRev(strDevName, strPercent)
                                    
                                                                    If Pos <> PosRev Then
                                                                        strVarname = Mid$(strDevName, Pos + 1, PosRev - 2)
                                    
                                                                        If InStr(strVarname, strPercent) = 0 Then
                                                                            strValval = objStringHash.Item(LCase$(strVarname))
                                                                        Else
                                                                            strVarname_x = Split(strVarname, strPercent)
                                    
                                                                            For ii = 0 To UBound(strVarname_x)
                                                                                
                                                                                If LenB(strValval) Then
                                                                                    strValval = strValval & strSpace & objStringHash.Item(LCase$(strVarname_x(ii)))
                                                                                Else
                                                                                    strValval = objStringHash.Item(LCase$(strVarname_x(ii)))
                                                                                End If
                        '
                                                                            Next ii
                                                                                                                                        
                                                                        End If
                                                                        
                                                                        If LenB(strValval) Then
                                                                            strDevName = Replace$(strDevName, "%" & strVarname & "%", strValval)
                                                                            ' Если все таки есть процент, то есть не определился из cекции String
                                                                        Else
                                                                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                                                            strDevName = strVarname
                                                                        End If
                                    
                                                                    Else
                                                                        strVarname = Replace$(strDevName, strPercent, vbNullString)
                                                                        strValval = objStringHash.Item(LCase$(strVarname))
                                                                        If LenB(strValval) Then
                                                                            strDevName = strValval
                                                                        Else
                                                                            strDevName = strVarname
                                                                        End If
                                                                    End If
                                                                End If
                                                
                                                                ' На случай если есть юникодовые символы в имени устройства
                                                                RemoveUni strDevName
                                    
                                                                ' Если требуется то удаление лишних символов
                                                                ReplaceBadSymbol strDevName
                                                            Else
                                                                If mbDebugDetail Then DebugMode "Error in inf: " & strInfFullname & " (Variable Name of Device is Empty) for HWID: " & strDevIDs
                                                                strDevName = "not defined in the inf"
                                                                '                                If mbIsDesignMode Then
                                                                '                                    Debug.Print "Not defined variable in [Strings] - " & strPackFileName & vbTab & strInfPath & strInfFileName & vbTab & objMatch.SubMatches(0) & vbTab & objMatchesDevDef.item(i)
                                                                '                                End If
                                                            End If
                                                        End If
                                                        
                                                        'Итоговая строка
                                                        'strDevID & vbTab & strInfFileName & vbTab & strManufSection & vbTab & strVer & vbTab & strSectEmptyList & vbTab & lngCatFileExists & vbTab & strDevName
                                                        strLinesArr(lngNumLines) = (strDevID & strInfPathTabQuoted & strManufSection) & (strPartString2Index & strDevName)
                                                        lngNumLines = lngNumLines + 1
                                                        
                                                        If Not objHWIDOutput.Exists(strDevID) Then
                                                            objHWIDOutput.Item(strDevID) = 1
                                                            strLinesArrHwid(lngNumLinesHwid) = strDevID
                                                            lngNumLinesHwid = lngNumLinesHwid + 1
                                                        End If
                                                        
                                                    End If
                                                End If
                                            ' strDevIDs'
                                            Next
                                        Else
                                        
                                            strValuer = strDevIDs
                                            
                                            If InStr(strValuer, strSpace) Then
                                                strValuer = Trim$(strValuer)
                                                Pos = InStr(strValuer, strSpace)
                                                If Pos Then
                                                    strValuer = Left$(strValuer, Pos - 1)
                                                End If
                                            End If
                                            
                                            If LenB(strValuer) > 8 Then
                                                If InStr(strValuer, strPercent) Then
                                                    strVarname = Left$(strValuer, InStrRev(strValuer, strPercent))
                    
                                                    If InStr(strVarname, strPercent) > 1 Then
                                                        strVarname = Right$(strVarname, Len(strVarname) - InStr(strValuer, strPercent) + 1)
                                                    End If
                    
                                                    strValval = objStringHash.Item(LCase$(strVarname))
                    
                                                    If LenB(strValval) Then
                                                        strValuer = Replace$(strValuer, strVarname, strValval)
                                                    Else
                                                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                                    End If
                                                End If
                    
                                                strDevID = UCase$(strValuer)
                                                
                                                ' разбиваем Hwid по "\" - оставляем только xxx\yyy
                                                If InStr(strDevID, vbBackslash) Then
                                                    strValuer_x = Split(strDevID, vbBackslash)
                                                    strDevID = strValuer_x(0) & vbBackslash & strValuer_x(1)
                                                End If
                
                                                If InStr(strDevID, strSpace) Then
                                                    strDevID = Trim$(strDevID)
                                                End If
                                                
                                                ss = strDevID & strManufSection
                    
                                                ' Если такая строка раньше не обрабаотывалась, то добавляем ее
                                                If Not objHashOutput.Exists(ss) Then
                                                    objHashOutput.Item(ss) = 1
                                                    
                                                    ' Обработаем имя устройства
                                                    If LenB(strDevName) Then
                                                        Pos = InStr(strDevName, strPercent)
                                                        strValval = vbNullString
                            
                                                        If Pos Then
                                                            PosRev = InStrRev(strDevName, strPercent)
                            
                                                            If Pos <> PosRev Then
                                                                strVarname = Mid$(strDevName, Pos + 1, PosRev - 2)
                            
                                                                If InStr(strVarname, strPercent) = 0 Then
                                                                    strValval = objStringHash.Item(LCase$(strVarname))
                                                                Else
                                                                    strVarname_x = Split(strVarname, strPercent)
                            
                                                                    For ii = 0 To UBound(strVarname_x)
                                                                        
                                                                        If LenB(strValval) Then
                                                                            strValval = strValval & strSpace & objStringHash.Item(LCase$(strVarname_x(ii)))
                                                                        Else
                                                                            strValval = objStringHash.Item(LCase$(strVarname_x(ii)))
                                                                        End If
                '
                                                                    Next ii
                                                                                                                                
                                                                End If
                                                                
                                                                If LenB(strValval) Then
                                                                    strDevName = Replace$(strDevName, "%" & strVarname & "%", strValval)
                                                                    ' Если все таки есть процент, то есть не определился из cекции String
                                                                Else
                                                                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                                                    strDevName = strVarname
                                                                End If
                            
                                                            Else
                                                                strVarname = Replace$(strDevName, strPercent, vbNullString)
                                                                strValval = objStringHash.Item(LCase$(strVarname))
                                                                If LenB(strValval) Then
                                                                    strDevName = strValval
                                                                Else
                                                                    strDevName = strVarname
                                                                End If
                                                            End If
                                                        End If
                                        
                                                        ' На случай если есть юникодовые символы в имени устройства
                                                        RemoveUni strDevName
                            
                                                        ' Если требуется то удаление лишних символов
                                                        ReplaceBadSymbol strDevName
                                                    Else
                                                        If mbDebugDetail Then DebugMode "Error in inf: " & strInfFullname & " (Variable Name of Device is Empty) for HWID: " & strDevIDs
                                                        strDevName = "not defined in the inf"
                                                        '                                If mbIsDesignMode Then
                                                        '                                    Debug.Print "Not defined variable in [Strings] - " & strPackFileName & vbTab & strInfPath & strInfFileName & vbTab & objMatch.SubMatches(0) & vbTab & objMatchesDevDef.item(i)
                                                        '                                End If
                                                    End If
                                                    
                                                    
                                                    'Итоговая строка
                                                    'strDevID & vbTab & strInfFileName & vbTab & strManufSection & vbTab & strVer & vbTab & strSectEmptyList & vbTab & lngCatFileExists & vbTab & strDevName
                                                    strLinesArr(lngNumLines) = (strDevID & strInfPathTabQuoted & strManufSection) & (strPartString2Index & strDevName)
                                                    lngNumLines = lngNumLines + 1
                                                    
                                                    If Not objHWIDOutput.Exists(strDevID) Then
                                                        objHWIDOutput.Item(strDevID) = 1
                                                        strLinesArrHwid(lngNumLinesHwid) = strDevID
                                                        lngNumLinesHwid = lngNumLinesHwid + 1
                                                    End If
                                                
                                                End If
                                            End If
                                        End If
                                    ' dev_defs'
                                    Next i
                                Else
                                    ' Если секция непустая, то установка из данного файла запрещена на данной системе
                                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManufSection & "] is Empty -> this OS not Supported by inf: " & strInfPathRelative
                                End If
                            Next K
                        ' dev_Sub_sects
                        Else
                            ' Если секция c HWID не найдена
                            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManufSection & "] Not Find in inf-file: " & strInfPathRelative
                        End If
                        'Next
                            
                    Else
                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strK2Sectlist(K2) & "] is Empty -> this OS not Supported by inf: " & strInfPathRelative
                    '  dev_Sub_sects not empty
                    End If
                    
                ' dev_sects'
                Next
            
            ' sect_list
            End If
        
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfPathTempList_x(infNum).FullPath
        End If

    Next

    ChangeStatusBarText strMessages(121) & strSpace & strPackFileName
    
    strRezultTxt = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".txt"
    strRezultTxtHwid = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".hwid"
    strRezultTxtTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxt)), "rezult", vbNullString, , , vbTextCompare)
    strRezultTxtHwidTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxtHwid)), "rezult", vbNullString, , , vbTextCompare)
    TimeScriptFinish = GetTickCount
    If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Create Index Data: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)

    ' Если данные найдены, то выводим итог в файл
    If lngNumLines Then

        ReDim Preserve strLinesArr(lngNumLines - 1)
        ReDim Preserve strLinesArrHwid(lngNumLinesHwid - 1)

        ' сортируем массивы
        TimeScriptRun = GetTickCount
    
        If lngSortMethodShell = 0 Then
            
            Set cSortHWID = New cBlizzard
            cSortHWID.SortMethod = BinaryCompare
            cSortHWID.SortOrder = Ascending
        
            cSortHWID.BlizzardStringSort strLinesArrHwid, 0&, lngNumLinesHwid - 1, False
            
            Set cSortHWID = Nothing
            
        ElseIf lngSortMethodShell = 1 Then
        
            ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLinesHwid, 4&, AddressOf CompareString
            
        ElseIf lngSortMethodShell = 2 Then
        
            Set cSortHWID = New cBlizzard
            cSortHWID.SortMethod = BinaryCompare
            cSortHWID.SortOrder = Ascending
            
            cSortHWID.TwisterStringSort strLinesArrHwid, 0&, lngNumLinesHwid - 1
            
            Set cSortHWID = Nothing
        Else
            ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLinesHwid, 4&, AddressOf CompareString
        End If
        
        TimeScriptFinish = GetTickCount
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Sort Index: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
        DoEvents
        TimeScriptRun = GetTickCount

        '---------------------------------------------
        '---------------Выводим итог в файл-----
        If PathExists(strPathDevDB) = False Then
            CreateNewDirectory strPathDevDB
        End If
        ' Запись в файл индекса
        FileWriteData strRezultTxt, Join(strLinesArr(), vbNewLine)
        ' Запись в файл короткого индекса
        FileWriteData strRezultTxtHwid, Join(strLinesArrHwid(), vbNewLine)
        
        TimeScriptFinish = GetTickCount
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Save Index Files: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
        
        ' Удаление массива, т.е освобождение памяти
        Erase strLinesArr
        Erase strLinesArrHwid

        If CopyFileTo(strRezultTxt, strRezultTxtTo) Then
            'Копируем файл HWID
            If CopyFileTo(strRezultTxtHwid, strRezultTxtHwidTo) Then
                ' Записываем версию базы драйверов в ini-файл
                IniWriteStrPrivate strPackFileName_woExt, "Version", lngDevDBVersion, PathCombine(strPathDevDB, "DevDBVersions.ini")
                'Ищем файл DriverPack*.ini
                strRezultTxt = vbNullString
                strRezultTxt_x = SearchFilesInRoot(strInfPathTemp, "DriverPack*.ini", False, True)
                strRezultTxt = strRezultTxt_x(0).FullPath

                ' Копируем DriverPack*.ini в каталог базы данных
                If PathExists(strRezultTxt) Then
                    If PathExists(strRezultTxtTo) Then
                        strRezultTxtTo = Replace$(strRezultTxtTo, ".txt", ".ini", , , vbTextCompare)

                        If CopyFileTo(strRezultTxt, strRezultTxtTo) = False Then
                            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Error of the saving file in directory database driver: " & strRezultTxtTo
                        Else
                            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Save DPFinish file: " & strRezultTxtTo
                        End If

                    Else
                        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Error of the saving file in directory database driver: " & strRezultTxtTo
                    End If
                End If
            Else
                MsgBox strMessages(31), vbInformation, strProductName
                If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Error of the saving file in directory database driver: " & strRezultTxtHwidTo
            End If
        Else
            MsgBox strMessages(31), vbInformation, strProductName
            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Error of the saving file in directory database driver: " & strRezultTxtTo
        End If
    End If

    If mbDebugStandart Then DebugMode vbTab & "DevParserByRegExp-End"
End Sub

