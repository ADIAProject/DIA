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
Private RegExpDevID         As RegExp
Private RegExpDevSect       As RegExp
Private objHashOutput       As Scripting.Dictionary
Private objStringHash       As Scripting.Dictionary

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
    Dim strDevName                As String
    Dim strPackGetFileName_woExt     As String
    Dim strRezultTxt_x()          As FindListStruct
    Dim strRezultTxt              As String
    Dim strRezultTxtTo            As String
    Dim strRezultTxtHwid          As String
    Dim strRezultTxtHwidTo        As String
    Dim strInfPathTempList_x()    As FindListStruct
    Dim strDevID                  As String
    Dim strDrvDate                As String
    Dim strDrvVersion             As String
    Dim strDrvCatFileName         As String
    Dim lngCatFileExists          As Long
    Dim strValval                 As String
    Dim sStrings                  As String
    Dim strRegEx_mansect          As String
    Dim strRegEx_strsect          As String
    Dim strRegEx_versect          As String
    Dim strRegEx_version          As String
    Dim strRegEx_devs_l           As String
    Dim strRegEx_devs_r           As String
    Dim strRegEx_devid            As String
    Dim strRegEx_mandef           As String
    Dim strRegEx_devdef           As String
    Dim strRegEx_strings          As String
    Dim strRegEx_sectnames        As String
    Dim sFileContent              As String
    Dim sVerSectContent           As String
    Dim strLinesArr()             As String
    Dim strLinesArrHwid()         As String
    Dim lngNumLines               As Long
    Dim strManufSection           As String
    Dim strKey                    As String
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
    Dim strSectNoCompatVerOSList  As String
    Dim strRegEx_catFile          As String
    Dim strArchCatFileList        As String
    Dim strArchCatFileListContent As String
    Dim strVarname_x()            As String
    Dim ii                        As Long
    Dim strUnpackMask             As String

    If mbDebugStandart Then DebugMode vbTab & "DevParserByRegExp-Start"
    
    TimeScriptRun = GetTickCount
    Set RegExpStrSect = New RegExp
    Set RegExpStrDefs = New RegExp
    Set RegExpVerSect = New RegExp
    Set RegExpVerParam = New RegExp
    Set RegExpCatParam = New RegExp
    Set RegExpManSect = New RegExp
    Set RegExpManDef = New RegExp
    Set RegManID = New RegExp
    Set RegExpDevDef = New RegExp
    Set RegExpDevID = New RegExp
    Set RegExpDevSect = New RegExp
    Set objHashOutput = New Scripting.Dictionary
    objHashOutput.CompareMode = TextCompare
    
    ' Должно ускорить распаковку, если выключено чтение файла finish.ini
    If Not mbLoadFinishFile Then
        strUnpackMask = " *.inf"
    Else
        strUnpackMask = " *.inf DriverPack*.ini"
    End If
    
    'Имя папки с распакованными драйверами
    strPackGetFileName_woExt = GetFileName_woExt(GetFileNameFromPath(strPackFileName))
    'Рабочий каталог
    strWorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackGetFileName_woExt)
    'Если рабочий каталог уже есть, то удаляем его
    DoEvents

    If PathExists(strWorkDir) Then
        ChangeStatusTextAndDebug strMessages(81)
        DelRecursiveFolder (strWorkDir)
        DoEvents
    End If

    ' Каталог для распаковки inf файлов
    strInfPathTemp = strWorkTempBackSL & strPackGetFileName_woExt

    If PathExists(strInfPathTemp) = False Then
        CreateNewDirectory strInfPathTemp
    End If

    If Not mbDP_Is_aFolder Then
        ' Запуск распаковки
        cmdString = strKavichki & strArh7zExePATH & strKavichki & " x -yo" & strKavichki & strInfPathTemp & strKavichki & " -r " & strKavichki & strPathDRP & strPackFileName & strKavichki & strUnpackMask
        ChangeStatusTextAndDebug strMessages(72) & " " & strPackFileName

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else

            ' Архиватор отработал на все 100%? Если нет то сообщаем
            If lngExitProc = 2 Or lngExitProc = 7 Or lngExitProc = 255 Then
                MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
            End If

            ' Создаем спсиок файлов *.cat в архиве
            strArchCatFileList = strWorkTempBackSL & "list_" & strPackGetFileName_woExt & ".txt"
            cmdString = "cmd.exe /c " & strKavichki & strKavichki & strArh7zExePATH & strKavichki & " l " & strKavichki & strPathDRP & strPackFileName & strKavichki & " -yr *.cat >" & strKavichki & strArchCatFileList & strKavichki
            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                If mbDebugStandart Then DebugMode strMessages(13) & str2vbNewLine & cmdString
            End If
        End If

        ChangeStatusTextAndDebug strMessages(73) & " " & strPackFileName
        'Построение списка inf файлов в рабочем каталоге
        strInfPathTempList_x = SearchFilesInRoot(strInfPathTemp, "*.inf", True, False)
    Else
        ' Создаем список файлов *.cat в архиве
        strArchCatFileList = strWorkTempBackSL & "list_" & strPackGetFileName_woExt & ".txt"
        cmdString = "cmd.exe /c Dir " & strKavichki & strPathDRP & strPackFileName & "\*.cat" & strKavichki & " /A- /B /S >" & strKavichki & strArchCatFileList & strKavichki

        'dir c:\windows\temp\*.tmp /S /B
        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            If mbDebugStandart Then DebugMode strMessages(33) & str2vbNewLine & cmdString
        End If

        ChangeStatusTextAndDebug strMessages(148) & " " & strPackFileName
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
    ' sections
    strRegEx_mansect = "^[ ]*\[Manufacturer\](?:([\s\S]*?)^[ #]*(?=\[)|([\s\S]*))"
    strRegEx_strsect = "^[ ]*\[Strings\](?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
    strRegEx_versect = "^[ ]*\[Version\](?:([\s\S]*?)^[ ]*(?=\[)|([\s\S]*))"
    strRegEx_version = "^[ ]*DriverVer[ ]*=[ ]*(%[^%]*%|(?:[\w/ ])+)(?:[ ]*,[ ]*(%[^%]*%|(?:[\w/ .])+))?"
    strRegEx_catFile = "^[ ]*CatalogFile[.nt|.ntamd64|.ntx86]*[ ]*=[ ]*([^;\r\n]*)"
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
        .Pattern = strRegEx_versect
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With
    
    With RegExpVerParam
        .Pattern = strRegEx_version
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpCatParam
        .Pattern = strRegEx_catFile
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    With RegExpManSect
        .Pattern = strRegEx_mansect
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
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

    ReDim strLinesArr(200000)
    ReDim strLinesArrHwid(200000)
    
    ' Чтение списка содержимого архива *.Cat
    strArchCatFileListContent = vbNullString

    If PathExists(strArchCatFileList) Then
        If GetFileSizeByPath(strArchCatFileList) Then
            strArchCatFileListContent = LCase$(FileReadData(strArchCatFileList))
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strArchCatFileList
        End If
    End If
        
    infCount = UBound(strInfPathTempList_x) + 1
    ChangeStatusTextAndDebug strMessages(73) & " " & strPackFileName & " (" & infCount & " inf-files)"
    
    ' Запускаем цикл обработки inf-файлов
    For infNum = 0 To UBound(strInfPathTempList_x)
        ' полный путь к файлу inf
        strInfFullname = strInfPathTempList_x(infNum).FullPath
        ' Имя inf файла
        strInfFileName = strInfPathTempList_x(infNum).NameLcase
        
        If (infNum Mod 20) = 0 Then
            ChangeStatusTextAndDebug strMessages(73) & " " & strPackFileName & " (" & infNum & " " & strMessages(124) & " " & infCount & ": " & strInfFileName & ")"
        Else
            If GetInputState Then
                DoEvents
            End If
        End If

        ' путь к файлу inf для записи в параметры - Каталог где лежит inf-файл
        strInfPath = strInfPathTempList_x(infNum).RelativePath

        If strInfPathTempList_x(infNum).Size Then
            ' Read INF file
            sFileContent = FileReadData(strInfFullname)

            ' Убираем символ """
            If InStr(sFileContent, strKavichki) Then
                sFileContent = Replace$(sFileContent, strKavichki, vbNullString)
            End If

            If InStr(sFileContent, vbTab) Then
                sFileContent = Replace$(sFileContent, vbTab, vbNullString)
            End If
                        
            ' Find [strings] section
            sStrings = vbNullString
            Set objStringHash = New Scripting.Dictionary
            objStringHash.CompareMode = TextCompare
            Set objMatchesStrSect = RegExpStrSect.Execute(sFileContent)
    
            If objMatchesStrSect.Count Then
                Set objMatch = objMatchesStrSect.Item(0)
                sStrings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
                'Debug.Print RegExpStrDefs.Pattern
                Set objMatchesStrDefs = RegExpStrDefs.Execute(sStrings)
    
                For i = 0 To objMatchesStrDefs.Count - 1
                    Set objMatch = objMatchesStrDefs.Item(i)
                    strKey = Trim$(objMatch.SubMatches(0))
                    strValue = Trim$(objMatch.SubMatches(1))
    
                    If Not objStringHash.Exists(strKey) Then
                        objStringHash.Add strKey, strValue
                        'Debug.Print strRegEx_strings
                        'Debug.Print strRegEx_strsect
                        objStringHash.Add strPercentage & strKey & strPercentage, strValue
                    End If
    
                Next
    
            End If
    
            ' Find [version] section
            Set objMatchesVerSect = RegExpVerSect.Execute(sFileContent)
            If objMatchesVerSect.Count Then
                sVerSectContent = objMatchesVerSect.Item(0)
            
                ' Find DriverVer parametr
                Set objMatchesVerParam = RegExpVerParam.Execute(sVerSectContent)
        
                If objMatchesVerParam.Count Then
                    Set objMatch = objMatchesVerParam.Item(0)
                    strDrvDate = objMatch.SubMatches(0)
        
                    If InStr(strDrvDate, strPercentage) Then
                        strVarname = Left$(strDrvDate, InStrRev(strDrvDate, strPercentage))
                        strValval = objStringHash.Item(strVarname)
        
                        If LenB(strValval) Then
                            strDrvDate = Replace$(strDrvDate, strVarname, strValval)
                        Else
                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                        End If
                    End If
        
                    strDrvDate = Trim$(strDrvDate)
                    strDrvVersion = objMatch.SubMatches(1)
        
                    If InStr(strDrvVersion, strPercentage) Then
                        strVarname = Left$(strDrvVersion, InStrRev(strDrvVersion, strPercentage))
                        strValval = objStringHash.Item(strVarname)
        
                        If LenB(strValval) Then
                            strDrvVersion = Replace$(strDrvVersion, strVarname, strValval)
                        Else
                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                        End If
                    End If
        
                    If LenB(strDrvVersion) Then
                        strVer = strDrvDate & "," & strDrvVersion
                    Else
        
                        If LenB(strDrvDate) Then
                            strVer = strDrvDate
                        Else
                            strVer = "unknown"
                        End If
                    End If
        
                Else
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr 'DriverVer' not found: " & strInfFullname
                    strDrvDate = vbNullString
                    strDrvVersion = vbNullString
                    strVer = "unknown"
                End If
        
                ' Find CatalogFile parametr
                Set objMatchesCatParam = RegExpCatParam.Execute(sVerSectContent)
        
                If objMatchesCatParam.Count Then
                    Set objMatch = objMatchesCatParam.Item(0)
                    strDrvCatFileName = objMatch.SubMatches(0)
                    
                    If InStr(strDrvCatFileName, strPercentage) Then
                        strVarname = Left$(strDrvCatFileName, InStrRev(strDrvCatFileName, strPercentage))
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
                strVer = "unknown"
                lngCatFileExists = 0
            End If
            
            ' Find [manufacturer] section
            Set objMatchesManSect = RegExpManSect.Execute(sFileContent)
    
            If objMatchesManSect.Count Then
                Set objMatch = objMatchesManSect.Item(0)
                strSections = objMatch.SubMatches(0) & objMatch.SubMatches(1)
                strSectlist = vbNullString
                Set objMatchesManDef = RegExpManDef.Execute(strSections)
    
                If objMatchesManDef.Count <> 0 Then
                
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
    
                            If i <> 0 Then
                                strSectlist = strSectlist & "|"
                            ElseIf j <> 0 Then
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
    
                Else
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr in section [Manufacturer] not match 'name = sectname,suffix,suffix'. Inf-File=" & strInfFullname
    
                    If InStr(strSectlist, vbNewLine) Then
                        strSectlist = Replace$(strSectlist, vbNewLine, vbNullString)
                    End If
    
                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Try seek in section [Manufacturer] parametr: " & strSectlist
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
                    Set objMatchesDevSect = RegExpDevSect.Execute(sFileContent)
    
                    For K = 0 To objMatchesDevSect.Count - 1
                        Set objMatch = objMatchesDevSect.Item(K)
                        strThisSection = objMatch.SubMatches(1) & objMatch.SubMatches(2)
                        strManufSection = UCase$(objMatch.SubMatches(0))
                        ' Find device definitions
                        Set objMatchesDevDef = RegExpDevDef.Execute(strThisSection)
    
                        'Debug.Print RegExpDevDef.Pattern
                        ' Если секция пустая, то установка из данного файла запрещена на данной системе
                        If objMatchesDevDef.Count = 0 Then
                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManufSection & "] is Empty -> this OS not Supported by inf: " & strInfPath & strInfFileName
                        End If
    
                        ' Handle definition
                        For i = 0 To objMatchesDevDef.Count - 1
                            Set objMatch = objMatchesDevDef.Item(i)
                            strDevIDs = objMatch.SubMatches(1)
                            strDevName = Trim$(objMatch.SubMatches(0))
    
                            'Debug.Print strDevName
                            If LenB(strDevName) Then
                                Pos = InStr(strDevName, strPercentage)
                                strValval = vbNullString
    
                                If Pos Then
                                    PosRev = InStrRev(strDevName, strPercentage)
    
                                    If Pos <> PosRev Then
                                        strVarname = Mid$(strDevName, Pos + 1, PosRev - 2)
    
                                        If InStr(strVarname, strPercentage) Then
                                            strVarname_x = Split(strVarname, strPercentage)
    
                                            For ii = 0 To UBound(strVarname_x)
                                                AppendStr strValval, objStringHash.Item(strVarname_x(ii))
                                            Next ii
    
                                        Else
                                            strValval = objStringHash.Item(strVarname)
                                        End If
    
                                    Else
                                        strVarname = Replace$(strDevName, strPercentage, vbNullString)
                                        strValval = objStringHash.Item(strVarname)
                                    End If
    
                                    If LenB(strValval) Then
                                        strDevName = Replace$(strDevName, "%" & strVarname & "%", strValval)
                                        ' Если все таки есть процент, то есть не определился из cекции String
                                    Else
                                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                        strDevName = strVarname
                                    End If
                                End If
    
                                If InStr(strDevName, strPercentage) Then
                                    '                            strDevName = "not defined in the inf"
                                    strDevName = Replace$(strDevName, strPercentage, vbNullString)
                                    strDevName = objStringHash.Item(strDevName)
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
    
                            ' add IDs
                            Set objMatchesDevID = RegExpDevID.Execute(strDevIDs)
    
                            For j = 0 To objMatchesDevID.Count - 1
                                Set objMatch = objMatchesDevID.Item(j)
                                strValuer = objMatch.SubMatches(0)
    
                                If InStr(strValuer, strPercentage) Then
                                    strVarname = Left$(strValuer, InStrRev(strValuer, strPercentage))
    
                                    If InStr(strVarname, strPercentage) > 1 Then
                                        strVarname = Right$(strVarname, Len(strVarname) - InStr(strValuer, strPercentage) + 1)
                                    End If
    
                                    strValval = objStringHash.Item(strVarname)
    
                                    If LenB(strValval) Then
                                        strValuer = Replace$(strValuer, strVarname, strValval)
                                    Else
                                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarname & "'"
                                    End If
                                End If
    
                                strValuer = UCase$(Trim$(strValuer))
                                strDevID = ParseDoubleHwid(strValuer)
                                ss = strDevID & vbTab & strInfPath & strInfFileName & vbTab & strManufSection
    
                                If Not objHashOutput.Exists(ss) Then
                                    objHashOutput.Item(ss) = "+"
    
                                    'Итоговая строка
                                    If InStr(strVer, " ") Then
                                        strVer = Replace$(strVer, " ", vbNullString)
                                    End If
    
                                    strLinesArr(lngNumLines) = ss & (vbTab & strVer & vbTab & strSectNoCompatVerOSList & vbTab & lngCatFileExists & vbTab & strDevName)
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
        
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfFullname
        End If

    Next

    ChangeStatusTextAndDebug strMessages(121) & " " & strPackFileName
    strRezultTxt = strWorkTempBackSL & "rezult" & strPackGetFileName_woExt & ".txt"
    strRezultTxtHwid = strWorkTempBackSL & "rezult" & strPackGetFileName_woExt & ".hwid"
    strRezultTxtTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxt)), "rezult", vbNullString, , , vbTextCompare)
    strRezultTxtHwidTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxtHwid)), "rezult", vbNullString, , , vbTextCompare)
    TimeScriptFinish = GetTickCount
    If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Create Index Data: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)

    ' Если данные найдены, то выводим итог в файл
    If lngNumLines Then

        ReDim Preserve strLinesArr(lngNumLines - 1)
        ReDim Preserve strLinesArrHwid(lngNumLines - 1)

        ' сортируем массивы
        TimeScriptRun = GetTickCount
        ShellSortAny VarPtr(strLinesArr(0)), lngNumLines, 4&, AddressOf CompareString
        ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLines, 4&, AddressOf CompareString
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
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Save Index 2 File: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
        ' Удаление массива, т.е освобождение памяти
        Erase strLinesArr
        Erase strLinesArrHwid

        If CopyFileTo(strRezultTxt, strRezultTxtTo) Then
            'Копируем файл HWID
            If CopyFileTo(strRezultTxtHwid, strRezultTxtHwidTo) Then
                ' Записываем версию базы драйверов в ini-файл
                IniWriteStrPrivate strPackGetFileName_woExt, "Version", lngDevDBVersion, PathCombine(strPathDevDB, "DevDBVersions.ini")
                'IniWriteStrPrivate strPackGetFileName_woExt, "FullHwid", CStr(Abs(Not mbDelDouble)), strPathDevDB & "DevDBVersions.ini"
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

