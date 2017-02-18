Attribute VB_Name = "mDevParserByReqExp"
Option Explicit

' ������� ������ ���� ������
Public Const lngDevDBVersion        As Long = 8

' ������� ����������
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
'������� ������ ���������� ��������� ��������
Private cSortHWID           As cAsmShell
Private cSortHWID2          As cBlizzard

'The GetInputState() API call will First check if there are any events and what-not that your application may have queued up waiting to be processed. Below is the declare for that function�
Private Declare Function GetInputState Lib "user32.dll" () As Long

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DevParserByRegExp
'! Description (��������)  :   [������� ID � �������� �������� �� inf-����� � ���������� ��]
'! Parameters  (����������):   strPackFileName (String)
'                              strPathDRP (String)
'                              strPathDevDB (String)
'!--------------------------------------------------------------------------------
Public Sub DevParserByRegExp(ByVal strPackFileName As String, ByVal strPathDRP As String, ByVal strPathDevDB As String)

    Dim objMatchStrSect           As Match
    Dim objMatchStrDefs           As Match
    Dim objMatchVer               As Match
    Dim objMatchCat               As Match
    Dim objMatchManSect           As Match
    Dim objMatchDevSect           As Match
    Dim objMatchManDef            As Match
    Dim objMatchManID             As Match
    Dim objMatchDevDef            As Match
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
    Dim lngTimeScriptRun          As Currency
    Dim lngTimeScriptFinish       As Currency
    Dim cmdString                 As String
    Dim strInfFullName            As String
    Dim strInfFileName            As String
    Dim strInfPath                As String
    Dim strInfPathRelative        As String
    Dim strInfPathRelativeDRP     As String
    Dim strInfPathTabQuoted       As String
    Dim strWorkDir                As String
    Dim strWorkDirInfList_x()     As FindListStruct
    Dim ii                        As Long
    Dim iii                       As Long
    Dim jj                        As Long
    Dim k1                        As Long
    Dim k2                        As Long
    Dim lngInfN                   As Long
    Dim lngInfCount               As Long
    Dim strValueID                As String
    Dim strValueID_x()            As String
    Dim strDevName                As String
    Dim strPackFileName_woExt     As String
    Dim strRezultTxt_x()          As FindListStruct
    Dim strRezultTxt              As String
    Dim strRezultTxtTo            As String
    Dim strRezultTxtHwid          As String
    Dim strRezultTxtHwidTo        As String
    Dim strRezultTxtTemp          As String
    Dim strDevID                  As String
    Dim strDrvDate                As String
    Dim strDrvVersion             As String
    Dim strDrvCatFileName         As String
    Dim lngCatFileExists          As Long
    Dim strValueHash              As String
    Dim strRegEx_devs_l           As String
    Dim strRegEx_devs_r           As String
    Dim sFileContent              As String
    Dim sVerSectContent           As String
    Dim strLinesArr()             As String
    Dim strLinesArrHwid()         As String
    Dim lngNumLines               As Long
    Dim lngNumLinesHwid           As Long
    Dim strKey                    As String
    Dim strKeyPercent             As String
    Dim strValue                  As String
    Dim strVarName                As String
    Dim strVarName_Orig           As String
    Dim strVarname_x()            As String
    Dim strManSectList            As String
    Dim strManSectEmptyList       As String
    Dim strManSectEmptyList4Check As String
    Dim strManSectBaseName        As String
    Dim strManSectSubString       As String
    Dim strManSection             As String
    Dim strManSectList_x()        As String
    Dim strSeekString             As String
    Dim strDevIDs                 As String
    Dim strDevIDs_x()             As String
    Dim lngPos                    As Long
    Dim lngPosRev                 As Long
    Dim strArchCatFileList        As String
    Dim strArchCatFileListContent As String
    Dim strUnpackMask             As String
    Dim strPartString2Index       As String
    Dim strVer                    As String
    Dim strVerTemp                As String
    Dim strVerTemp_x()            As String
    Dim mbParseInfDrp             As Boolean
    Dim lngLinesArrMax            As Long
    Dim lngLinesArrHwidMax        As Long

    If mbDebugStandart Then DebugMode vbTab & "DevParserByRegExp-Start"
    
    lngTimeScriptRun = GetTimeStart
    
    ' Hash-������� ������������ �������� strDevID & strManSection � ������ inf-�����
    Set objHashOutput = New Scripting.Dictionary
    objHashOutput.CompareMode = BinaryCompare
    ' Hash-������� ������������ �������� ������ String � ������ inf-�����
    Set objStringHash = New Scripting.Dictionary
    objStringHash.CompareMode = BinaryCompare
    ' Hash-������� ������������ �������� HWID � ������ ������-���������
    Set objHWIDOutput = New Scripting.Dictionary
    objHWIDOutput.CompareMode = BinaryCompare
    
    ' ������ �������� ����������, ���� ��������� ������ ����� finish.ini
    If Not mbParseHwidByInfDrpFile Then

        If Not mbLoadFinishFile Then
            strUnpackMask = " *.inf"
        Else
            strUnpackMask = " *.inf DriverPack*.ini"
        End If
    ' ��������� � ���������� ����� *.infdrp
    Else
        If Not mbLoadFinishFile Then
            strUnpackMask = " *.inf *.infdrp"
        Else
            strUnpackMask = " *.inf DriverPack*.ini *.infdrp"
        End If
    End If
    
    '��� ����� � �������������� ����������
    strPackFileName_woExt = GetFileNameOnly_woExt(strPackFileName)
    '������� ������� - ������� ��� ���������� inf ������
    strWorkDir = BackslashAdd2Path(strWorkTempBackSL & strPackFileName_woExt)
    
    '���� ������� ������� ��� ����, �� ������� ���
    If PathExists(strWorkDir) Then
        ChangeStatusBarText strMessages(81)
        DelRecursiveFolder strWorkDir
        DoEvents
    Else
        CreateNewDirectory strWorkDir
    End If

    If Not mbDP_Is_aFolder Then
        ' ������ ����������
        cmdString = strQuotes & strArh7zExePath & strQuotes & " x -y -o" & strQuotes & strWorkDir & strQuotes & " -r " & strQuotes & strPathDRP & strPackFileName & strQuotes & strUnpackMask
        ChangeStatusBarText strMessages(72) & strSpace & strPackFileName

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else

            ' ��������� ��������� �� ��� 100%? ���� ��� �� ��������
            If lngExitProc <> 0 Then
                If lngExitProc = 2 Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                ElseIf lngExitProc = 7 Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                ElseIf lngExitProc = 255 Then
                    MsgBox strMessages(13) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If
            End If
            
            ' ������� ������ ������ *.cat � ������
            strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
            cmdString = "cmd.exe /c " & strQuotes & strQuotes & strArh7zExePath & strQuotes & " l " & strQuotes & strPathDRP & strPackFileName & strQuotes & " -y -r *.cat >" & strQuotes & strArchCatFileList & strQuotes
            If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
                If mbDebugStandart Then DebugMode strMessages(13) & str2vbNewLine & cmdString
            End If
        End If

        ChangeStatusBarText strMessages(73) & strSpace & strPackFileName
        '���������� ������ inf ������ � ������� ��������
        If Not mbParseHwidByInfDrpFile Then
            strWorkDirInfList_x = SearchFilesInRoot(strWorkDir, "*.inf", True, False)
        Else
            strWorkDirInfList_x = SearchFilesInRoot(strWorkDir, "*.inf;*.infdrp", True, False)
        End If
    Else
        ' ������� ������ ������ *.cat � ������
        strArchCatFileList = strWorkTempBackSL & "list_" & strPackFileName_woExt & ".txt"
        cmdString = "cmd.exe /c Dir " & strQuotes & strPathDRP & strPackFileName & "\*.cat" & strQuotes & " /A- /B /S >" & strQuotes & strArchCatFileList & strQuotes

        'dir c:\windows\temp\*.tmp /S /B
        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            If mbDebugStandart Then DebugMode strMessages(33) & str2vbNewLine & cmdString
        End If

        ChangeStatusBarText strMessages(148) & strSpace & strPackFileName
        '���������� ������ inf ������ � ������� ��������
        If Not mbParseHwidByInfDrpFile Then
            strWorkDirInfList_x = SearchFilesInRoot(strPathDRP & strPackFileName, "*.inf", True, False)
        Else
            strWorkDirInfList_x = SearchFilesInRoot(strPathDRP & strPackFileName, "*.inf;*.infdrp", True, False)
        End If
    End If

    If UBound(strWorkDirInfList_x) = 0 Then
        If LenB(strWorkDirInfList_x(0).FullPath) = 0 Then
            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Error to Unpack Inf-file: no files in DP or extracting error"
            Exit Sub
        End If
    End If

    lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
    If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Unpack Inf-file: " & CalculateTime(lngTimeScriptFinish, True)
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
        '.Pattern = strRegEx_devs_l & strManSection & strRegEx_devs_r
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
    
    ' ������ ����� inf �� ����� ����������� ���������� � ";#"
    Set RegExpReplace = New RegExp
    With RegExpReplace
        .Pattern = "^([ ]*[;#]+[ \S]*)$"
        .MultiLine = True
        .IgnoreCase = False
        .Global = True
    End With

    lngLinesArrMax = 100000
    lngLinesArrHwidMax = 20000
    ReDim strLinesArr(lngLinesArrMax)
    ReDim strLinesArrHwid(lngLinesArrHwidMax)
    
    ' ������ ������ ����������� ������ *.Cat
    If FileExists(strArchCatFileList) Then
        If GetFileSizeByPath(strArchCatFileList) Then
            FileReadData strArchCatFileList, strArchCatFileListContent
            strArchCatFileListContent = LCase$(strArchCatFileListContent)
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strArchCatFileList
        End If
    End If
        
    lngInfCount = UBound(strWorkDirInfList_x) + 1
    ChangeStatusBarText strMessages(73) & strSpace & strPackFileName & " (" & lngInfCount & " inf-files)"
    
    ' ��������� ���� ��������� inf-������
    For lngInfN = 0 To UBound(strWorkDirInfList_x)
        
        mbParseInfDrp = False
        ' ����� ��������� ������ infdrp �� �������, �����...
        If Not mbParseHwidByInfDrpFile Then
            ' ���� ���� infdrp, �� ���������� ���
            If StrComp(strWorkDirInfList_x(lngInfN).Extension, "infdrp") = 0 Then
                'SkipFile - �� ������������ ����� *.InfDrp
                GoTo SkipFileInfDrp
            End If
        Else
            '���� �� ������������ - ����� �� �������
            If StrComp(strWorkDirInfList_x(lngInfN).Extension, "infdrp") = 0 Then
                'SkipFile - �� ������������ ����� *.InfDrp
                'GoTo SkipFileInfDrp
                mbParseInfDrp = True
                GoTo StartParseInfFile
            Else
                ' ���� � ����������� inf, ��������� ������� ����� infdrp, ���� ���� �� ���������� ������������ inf
                If FileExists(strWorkDirInfList_x(lngInfN).FullPath & "drp") Then
                    GoTo SkipFileInfDrp
                End If
            End If
        End If

StartParseInfFile:
        If strWorkDirInfList_x(lngInfN).Size Then
            
            ' ������ ���� � ����� inf
            strInfFullName = strWorkDirInfList_x(lngInfN).FullPath
            ' ��� inf �����
            strInfFileName = strWorkDirInfList_x(lngInfN).NameLCase
            
            If (lngInfN Mod 20) = 0 Then
                ChangeStatusBarText strMessages(73) & strSpace & strPackFileName & " (" & lngInfN & strSpace & strMessages(124) & strSpace & lngInfCount & ": " & strInfFileName & ")"
            Else
                If GetInputState Then
                    DoEvents
                End If
            End If
        
            ' ������� ������ �������� ���������� ����� HWID
            Set objHashOutput = New Scripting.Dictionary
            objHashOutput.CompareMode = BinaryCompare
            ' ������� ������ �������� ������ strings
            Set objStringHash = New Scripting.Dictionary
            objStringHash.CompareMode = BinaryCompare
            
            ' ���� � ����� inf ��� ������ � ��������� - ������� ��� ����� inf-����
            strInfPath = strWorkDirInfList_x(lngInfN).RelativePath
            strInfPathRelative = strInfPath & strInfFileName
            strInfPathRelativeDRP = Replace$(strInfPathRelative, ".infdrp", ".inf")
            If Not mbParseInfDrp Then
                strInfPathTabQuoted = vbTab & strInfPathRelative & vbTab
            Else
                strInfPathTabQuoted = vbTab & strInfPathRelativeDRP & vbTab
            End If
            
            ' Read INF file
            FileReadData strInfFullName, sFileContent

            ' ������� ������ """
            If InStr(sFileContent, strQuotes) Then
                sFileContent = Replace$(sFileContent, strQuotes, vbNullString)
            End If

            ' ������� ������ "tab"
            If InStr(sFileContent, vbTab) Then
                sFileContent = Replace$(sFileContent, vbTab, vbNullString)
            End If
            
            ' ������� ������ � ; ��� # � ������ � ������ ������
            sFileContent = RegExpReplace.Replace(sFileContent, vbNewLine)
            
            
            ' Find [strings] section
            Set objMatchesStrSect = RegExpStrSect.Execute(sFileContent)
    
            If objMatchesStrSect.count Then
                Set objMatchStrSect = objMatchesStrSect.item(0)
                
                Set objMatchesStrDefs = RegExpStrDefs.Execute(objMatchStrSect.SubMatches(0) & objMatchStrSect.SubMatches(1))
    
                For ii = 0 To objMatchesStrDefs.count - 1
                    Set objMatchStrDefs = objMatchesStrDefs.item(ii)
                    strKey = Trim$(LCase$(objMatchStrDefs.SubMatches(0)))
                    strValue = Trim$(objMatchStrDefs.SubMatches(1))
    
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
            If objMatchesVerSect.count Then
                sVerSectContent = LCase$(objMatchesVerSect.item(0))
            
                ' Find DriverVer parametr
                Set objMatchesVerParam = RegExpVerParam.Execute(sVerSectContent)
        
                If objMatchesVerParam.count Then
                    Set objMatchVer = objMatchesVerParam.item(0)
                    strVerTemp = objMatchVer.SubMatches(0)
                    
                    If InStr(strVerTemp, strPercent) Then
                        If InStr(strVerTemp, strComma) Then
                            strVerTemp_x = Split(strVerTemp, strComma)
                            strDrvDate = Trim$(strVerTemp_x(0))
                            strDrvVersion = Trim$(strVerTemp_x(1))
                        Else
                            strDrvDate = Trim$(strVerTemp)
                        End If
                        
                        If InStr(strDrvDate, strPercent) Then
                            strVarName = Left$(strDrvDate, InStrRev(strDrvDate, strPercent))
                            strValueHash = objStringHash.item(strVarName)
            
                            If LenB(strValueHash) Then
                                strDrvDate = Replace$(strDrvDate, strVarName, strValueHash)
                            Else
                                If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarName & "'"
                            End If
                        End If
            
                        If InStr(strDrvVersion, strPercent) Then
                            strVarName = Left$(strDrvVersion, InStrRev(strDrvVersion, strPercent))
                            strValueHash = objStringHash.item(strVarName)
            
                            If LenB(strValueHash) Then
                                strDrvVersion = Replace$(strDrvVersion, strVarName, strValueHash)
                            Else
                                If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarName & "'"
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
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr 'DriverVer' not found: " & strInfFullName
                    strDrvDate = vbNullString
                    strDrvVersion = vbNullString
                    strVer = strUnknownLCase
                End If
        
                ' Find CatalogFile parametr
                Set objMatchesCatParam = RegExpCatParam.Execute(sVerSectContent)
        
                If objMatchesCatParam.count Then
                    Set objMatchCat = objMatchesCatParam.item(0)
                    strDrvCatFileName = objMatchCat.SubMatches(0)
                    
                    If InStr(strDrvCatFileName, strPercent) Then
                        strVarName = Left$(strDrvCatFileName, InStrRev(strDrvCatFileName, strPercent))
                        strValueHash = objStringHash.item(strVarName)
        
                        If LenB(strValueHash) Then
                            strDrvCatFileName = Replace$(strDrvCatFileName, strVarName, strValueHash)
                        Else
                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Cannot find '" & strVarName & "'"
                        End If
                    End If
                    strDrvCatFileName = LCase$(strDrvCatFileName)
        
                    ' ���� �� ���� *.cat � ������ ������ ������?
                    If InStr(strDrvCatFileName, ".cat") Then
                        If InStr(strArchCatFileListContent, LCase$(strInfPath) & strDrvCatFileName) Then
                            lngCatFileExists = 1
                        Else
                            lngCatFileExists = 0
                        End If
        
                    Else
                        lngCatFileExists = 0
                    End If
        
                Else
                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr 'CatalogeFile' not found: " & strInfFullName
                    strDrvCatFileName = vbNullString
                    lngCatFileExists = 0
                End If
            Else
                If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Section [version] not found: " & strInfFullName
                strDrvDate = vbNullString
                strDrvVersion = vbNullString
                strVer = strUnknownLCase
                lngCatFileExists = 0
            End If
            
            ' Find [manufacturer] section
            Set objMatchesManSect = RegExpManSect.Execute(sFileContent)
    
            If objMatchesManSect.count Then
                Set objMatchManSect = objMatchesManSect.item(0)
                strManSectList = vbNullString
                Set objMatchesManDef = RegExpManDef.Execute(objMatchManSect.SubMatches(0) & objMatchManSect.SubMatches(1))
    
                If objMatchesManDef.count Then
                
                    For ii = 0 To objMatchesManDef.count - 1
                        Set objMatchManDef = objMatchesManDef.item(ii)
                        strSeekString = objMatchManDef.SubMatches(0)
                        Set objMatchesManID = RegManID.Execute(strSeekString)
                        strManSectBaseName = vbNullString
    
                        For jj = 0 To objMatchesManID.count - 1
                            Set objMatchManID = objMatchesManID.item(jj)
                            strManSectSubString = RTrim$(objMatchManID.SubMatches(0))
    
                            If ii <> 0 Then
                                strManSectList = strManSectList & "|"
                            ElseIf jj <> 0 Then
                                strManSectList = strManSectList & "|"
                            End If
    
                            If jj = 0 Then
                                strManSectBaseName = strManSectSubString
                                strManSectList = strManSectList & strManSectBaseName
                            Else
                                strManSectList = strManSectList & (strManSectBaseName & strDot & strManSectSubString)
                            End If
    
                        Next
                        strManSectList = UCase$(strManSectList)
                    Next
    
                Else
                    If mbDebugStandart Then DebugMode str2VbTab & "DevParserbyRegExp: Error in inf: Parametr in section [Manufacturer] not match 'name = sectname,suffix,suffix'. Inf-File=" & strInfFullName
    
                    If InStr(strManSectList, vbNewLine) Then
                        strManSectList = Replace$(strManSectList, vbNewLine, vbNullString)
                    End If
    
                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserbyRegExp: Try seek in section [Manufacturer] parametr: " & strManSectList
                End If
    
                ' ���������� ������������� �� ��� ������� ����
                If InStr(strManSectList, "|") Then
                    strManSectList_x = Split(strManSectList, "|")
                    strManSectEmptyList = GetIniEmptySectionFromList(strManSectList, strInfFullName)
                Else
                    ReDim strManSectList_x(0)
                    strManSectList_x(0) = strManSectList
                    strManSectEmptyList = strDash
                End If
                
                ' ����� ������ ������� ����� ����� ��������� � ������
                strPartString2Index = vbTab & (strVer & vbTab & strManSectEmptyList) & (vbTab & lngCatFileExists & vbTab)
                            
                strManSectEmptyList4Check = strManSectEmptyList & strComma
                    
                For k2 = 0 To UBound(strManSectList_x)
                    
                    strManSection = strManSectList_x(k2)
                    ' ���� ������ "������", �� ���������� �� ��������� (������ ������ ������ ������� �����)
                    If InStr(strManSectEmptyList4Check, strManSection & strComma) = 0 Then
                        RegExpDevSect.Pattern = strRegEx_devs_l & strManSection & strRegEx_devs_r
                        Set objMatchesDevSect = RegExpDevSect.Execute(sFileContent)
                    
                        ' ���� ���������� �������
                        If objMatchesDevSect.count Then
                            For k1 = 0 To objMatchesDevSect.count - 1
                                Set objMatchDevSect = objMatchesDevSect.item(k1)
                                
                                ' Find device definitions
                                Set objMatchesDevDef = RegExpDevDef.Execute(objMatchDevSect.SubMatches(0) & objMatchDevSect.SubMatches(1))
            
                                ' ���� ������ �� ������, ��
                                If objMatchesDevDef.count Then
                                    ' Handle definition
                                    For ii = 0 To objMatchesDevDef.count - 1
                                        Set objMatchDevDef = objMatchesDevDef.item(ii)
                                        strDevIDs = objMatchDevDef.SubMatches(1)
                                        If InStr(strDevIDs, vbCr) Then
                                            strDevIDs = Replace$(strDevIDs, vbCr, vbNullString)
                                        End If
                                        strDevName = Trim$(objMatchDevDef.SubMatches(0))
                            
                                        ' add IDs
                                        If InStr(strDevIDs, strComma) Then
                    
                                            strDevIDs_x = Split(strDevIDs, strComma)
                                            For jj = 0 To UBound(strDevIDs_x)
    
                                                strValueID = strDevIDs_x(jj)
                                                
                                                If InStr(strValueID, strSpace) Then
                                                    strValueID = Trim$(strValueID)
                                                    lngPos = InStr(strValueID, strSpace)
                                                    If lngPos Then
                                                        strValueID = Left$(strValueID, lngPos - 1)
                                                    End If
                                                End If
                                                                
                                                If LenB(strValueID) > 8 Then
                                                    lngPos = InStr(strValueID, strPercent)
                                                    If lngPos Then
                                                        strValueID = LCase$(strValueID)
                                                        lngPosRev = InStrRev(strValueID, strPercent)
                                                        strValueHash = vbNullString
                        
                                                        If lngPos <> lngPosRev Then
                                                            strVarName = Mid$(strValueID, lngPos + 1, lngPosRev - lngPos - 1)
                        
                                                            If InStr(strVarName, strPercent) = 0 Then
                                                                strValueHash = objStringHash.item(strVarName)
                                                            Else
                                                                strVarname_x = Split(strVarName, strPercent)
                        
                                                                For iii = 0 To UBound(strVarname_x)
                                                                    
                                                                    If LenB(strValueHash) Then
                                                                        strValueHash = strValueHash & strSpace & objStringHash.item(strVarname_x(iii))
                                                                    Else
                                                                        strValueHash = objStringHash.item(strVarname_x(iii))
                                                                    End If
                                                                
                                                                Next iii
                                                                strValueHash = Trim$(strValueHash)
                                                                strValueHash = Replace$(strValueHash, str2Space, "&")
                                                                strValueHash = Replace$(strValueHash, strSpace, "&")
                                                                                                                            
                                                            End If
                                                            
                                                            If LenB(strValueHash) Then
                                                                strValueID = Replace$(strValueID, strPercent & strVarName & strPercent, strValueHash)
                                                                ' ���� ��� ���� ���� �������, �� ���� �� ����������� �� c����� String
                                                            Else
                                                                If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarName & "'"
                                                                strValueID = strVarName
                                                            End If
                        
                                                        Else
                                                            strVarName = Replace$(strValueID, strPercent, vbNullString)
                                                            strValueHash = objStringHash.item(strVarName)
                                                            If LenB(strValueHash) Then
                                                                strValueID = strValueHash
                                                            Else
                                                                strValueID = strVarName
                                                            End If
                                                        End If
                                                    End If
                        
                                                    strDevID = UCase$(strValueID)
                                                    
                                                    ' ��������� Hwid �� "\" - ��������� ������ xxx\yyy
                                                    If InStr(strDevID, vbBackslash) Then
                                                        strValueID_x = Split(strDevID, vbBackslash)
                                                        strDevID = strValueID_x(0) & vbBackslash & strValueID_x(1)
                                                    End If
                    
                                                    If InStr(strDevID, strSpace) Then
                                                        strDevID = Trim$(strDevID)
                                                    End If
                                                    
                                                    strSeekString = strDevID & strManSection
                        
                                                    If Not objHashOutput.Exists(strSeekString) Then
                                                        objHashOutput.Add strSeekString, 1
                                                        
                                                        ' ���������� ��� ����������
                                                        If LenB(strDevName) Then
                                                            lngPos = InStr(strDevName, strPercent)
                                                            strValueHash = vbNullString
                                
                                                            If lngPos Then
                                                                lngPosRev = InStrRev(strDevName, strPercent)
                                
                                                                If lngPos <> lngPosRev Then
                                                                    strVarName_Orig = Mid$(strDevName, lngPos + 1, lngPosRev - lngPos - 1)
                                                                    strVarName = LCase$(strVarName_Orig)
                                
                                                                    If InStr(strVarName, strPercent) = 0 Then
                                                                        If objStringHash.Exists(strVarName) Then
                                                                            strValueHash = objStringHash.item(strVarName)
                                                                        Else
                                                                            strValueHash = vbNullString
                                                                        End If
                                                                    Else
                                                                        strVarname_x = Split(strVarName, strPercent)
                                
                                                                        For iii = 0 To UBound(strVarname_x)
                                                                            
                                                                            If LenB(strValueHash) Then
                                                                                strVarname_x(iii) = Trim$(strVarname_x(iii))
                                                                                If LenB(strVarname_x(iii)) Then
                                                                                    strValueHash = strValueHash & strSpace & objStringHash.item(strVarname_x(iii))
                                                                                End If
                                                                            Else
                                                                                strVarname_x(iii) = Trim$(strVarname_x(iii))
                                                                                strValueHash = objStringHash.item(strVarname_x(iii))
                                                                            End If
                    
                                                                        Next iii
                                                                                                                                    
                                                                    End If
                                                                                                                                                                                                            
                                                                    If LenB(strValueHash) Then
                                                                        strDevName = Replace$(strDevName, strPercent & strVarName_Orig & strPercent, strValueHash)
                                                                        ' ���� ��� ���� ���� �������, �� ���� �� ����������� �� c����� String
                                                                    Else
                                                                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarName & "'"
                                                                        strDevName = strVarName
                                                                    End If
                                
                                                                Else
                                                                    strVarName = Replace$(strDevName, strPercent, vbNullString)
                                                                    strValueHash = objStringHash.item(LCase$(strVarName))
                                                                    If LenB(strValueHash) Then
                                                                        strDevName = strValueHash
                                                                    Else
                                                                        strDevName = strVarName
                                                                    End If
                                                                End If
                                                            End If
                                            
                                                            ' �� ������ ���� ���� ���������� ������� � ����� ����������
                                                            RemoveUni strDevName
                                
                                                            ' ���� ��������� �� �������� ������ ��������
                                                            ReplaceBadSymbol strDevName
                                                        Else
                                                            If mbDebugDetail Then DebugMode "Error in inf: " & strInfFullName & " (Variable Name of Device is Empty) for HWID: " & strDevIDs
                                                            strDevName = "not defined in the inf"
                                                            'If mbIsDesignMode Then
                                                            '    Debug.Print "Not defined variable in [Strings] - " & strPackFileName & vbTab & strInfPath & strInfFileName & vbTab & objMatch.SubMatches(0) & vbTab & objMatchesDevDef.Item(i)
                                                            'End If
                                                        End If
                                                        
                                                        '�������� ������
                                                        'strDevID & vbTab & strInfFileName & vbTab & strManSection & vbTab & strVer & vbTab & strManSectEmptyList & vbTab & lngCatFileExists & vbTab & strDevName
                                                        ' ��������������� ������� ���� ��������� �������� �����������
                                                        If lngNumLines >= lngLinesArrMax Then
                                                            lngLinesArrMax = 2 * lngLinesArrMax
                                                            ReDim Preserve strLinesArr(lngLinesArrMax)
                                                        End If
                                                        strLinesArr(lngNumLines) = (strDevID & strInfPathTabQuoted & strManSection) & (strPartString2Index & strDevName)
                                                        lngNumLines = lngNumLines + 1
                                                        
                                                        If Not objHWIDOutput.Exists(strDevID) Then
                                                            objHWIDOutput.Add strDevID, 1
                                                            If lngNumLinesHwid >= lngLinesArrHwidMax Then
                                                                lngLinesArrHwidMax = 2 * lngLinesArrHwidMax
                                                                ReDim Preserve strLinesArrHwid(lngLinesArrHwidMax)
                                                            End If
                                                            strLinesArrHwid(lngNumLinesHwid) = strDevID
                                                            lngNumLinesHwid = lngNumLinesHwid + 1
                                                        End If
                                                        
                                                    End If
                                                End If
                                            ' strDevIDs'
                                            Next
                                        Else
                                        
                                            strValueID = strDevIDs
                                            
                                            If InStr(strValueID, strSpace) Then
                                                strValueID = Trim$(strValueID)
                                                lngPos = InStr(strValueID, strSpace)
                                                If lngPos Then
                                                    strValueID = Left$(strValueID, lngPos - 1)
                                                End If
                                            End If
                                            
                                            If LenB(strValueID) > 8 Then
                                                lngPos = InStr(strValueID, strPercent)
                                                If lngPos Then
                                                    lngPosRev = InStrRev(strValueID, strPercent)
                                                    strValueHash = vbNullString
                    
                                                    If lngPos <> lngPosRev Then
                                                        strVarName = Mid$(strValueID, lngPos + 1, lngPosRev - lngPos - 1)
                    
                                                        If InStr(strVarName, strPercent) = 0 Then
                                                            strValueHash = objStringHash.item(LCase$(strVarName))
                                                        Else
                                                            strVarname_x = Split(LCase$(strVarName), strPercent)
                    
                                                            For iii = 0 To UBound(strVarname_x)
                                                                
                                                                If LenB(strVarname_x(iii)) > 2 Then
                                                                    If LenB(strValueHash) Then
                                                                        strValueHash = strValueHash & strSpace & objStringHash.item(strVarname_x(iii))
                                                                    Else
                                                                        strValueHash = objStringHash.item(strVarname_x(iii))
                                                                    End If
                                                                End If

                                                            Next iii
                                                            strValueHash = Trim$(strValueHash)
                                                            strValueHash = Replace$(strValueHash, str2Space, "&")
                                                            strValueHash = Replace$(strValueHash, strSpace, "&")
                                                                                                                        
                                                        End If
                                                        
                                                        If LenB(strValueHash) Then
                                                            strValueID = Replace$(strValueID, strPercent & strVarName & strPercent, strValueHash)
                                                            ' ���� ��� ���� ���� �������, �� ���� �� ����������� �� c����� String
                                                        Else
                                                            If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarName & "'"
                                                            strValueID = strVarName
                                                        End If
                    
                                                    Else
                                                        strVarName = Replace$(strValueID, strPercent, vbNullString)
                                                        strValueHash = objStringHash.item(LCase$(strVarName))
                                                        If LenB(strValueHash) Then
                                                            strValueID = strValueHash
                                                        Else
                                                            strValueID = strVarName
                                                        End If
                                                    End If
                                                End If
                    
                                                strDevID = UCase$(strValueID)
                                                
                                                ' ��������� Hwid �� "\" - ��������� ������ xxx\yyy
                                                If InStr(strDevID, vbBackslash) Then
                                                    strValueID_x = Split(strDevID, vbBackslash)
                                                    strDevID = strValueID_x(0) & vbBackslash & strValueID_x(1)
                                                End If
                
                                                If InStr(strDevID, strSpace) Then
                                                    strDevID = Trim$(strDevID)
                                                End If
                                                
                                                strSeekString = strDevID & strManSection
                    
                                                ' ���� ����� ������ ������ �� ��������������, �� ��������� ��
                                                If Not objHashOutput.Exists(strSeekString) Then
                                                    objHashOutput.Add strSeekString, 1
                                                    
                                                    ' ���������� ��� ����������
                                                    If LenB(strDevName) Then
                                                        lngPos = InStr(strDevName, strPercent)
                                                        strValueHash = vbNullString
                            
                                                        If lngPos Then
                                                            lngPosRev = InStrRev(strDevName, strPercent)
                            
                                                            If lngPos <> lngPosRev Then
                                                                strVarName_Orig = Mid$(strDevName, lngPos + 1, lngPosRev - 2)
                                                                strVarName = LCase$(strVarName_Orig)
                            
                                                                If InStr(strVarName, strPercent) = 0 Then
                                                                    If objStringHash.Exists(strVarName) Then
                                                                        strValueHash = objStringHash.item(strVarName)
                                                                    Else
                                                                        strValueHash = vbNullString
                                                                    End If
                                                                Else
                                                                    strVarname_x = Split(strVarName, strPercent)
                            
                                                                    For iii = 0 To UBound(strVarname_x)
                                                                        
                                                                        If LenB(strValueHash) Then
                                                                            strVarname_x(iii) = Trim$(strVarname_x(iii))
                                                                            If LenB(strVarname_x(iii)) Then
                                                                                strValueHash = strValueHash & strSpace & objStringHash.item(strVarname_x(iii))
                                                                            End If
                                                                        Else
                                                                            strVarname_x(iii) = Trim$(strVarname_x(iii))
                                                                            strValueHash = objStringHash.item(strVarname_x(iii))
                                                                        End If
                
                                                                    Next iii
                                                                                                                                
                                                                End If
                                                                
                                                                If LenB(strValueHash) Then
                                                                    strDevName = Replace$(strDevName, strPercent & strVarName_Orig & strPercent, strValueHash)
                                                                    ' ���� ��� ���� ���� �������, �� ���� �� ����������� �� c����� String
                                                                Else
                                                                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Error in inf: Cannot find '" & strVarName & "'"
                                                                    strDevName = strVarName
                                                                End If
                            
                                                            Else
                                                                strVarName = Replace$(strDevName, strPercent, vbNullString)
                                                                If InStr(strVarName, "$") Then
                                                                    strVarName = Replace$(strVarName, "$", vbNullString)
                                                                End If
                                                                strValueHash = objStringHash.item(LCase$(strVarName))
                                                                
                                                                If LenB(strValueHash) Then
                                                                    strDevName = strValueHash
                                                                Else
                                                                    strDevName = strVarName
                                                                End If
                                                            End If
                                                        End If
                                        
                                                        ' �� ������ ���� ���� ���������� ������� � ����� ����������
                                                        RemoveUni strDevName
                            
                                                        ' ���� ��������� �� �������� ������ ��������
                                                        ReplaceBadSymbol strDevName
                                                    Else
                                                        If mbDebugDetail Then DebugMode "Error in inf: " & strInfFullName & " (Variable Name of Device is Empty) for HWID: " & strDevIDs
                                                        strDevName = "not defined in the inf"
                                                        'If mbIsDesignMode Then
                                                        '    Debug.Print "Not defined variable in [Strings] - " & strPackFileName & vbTab & strInfPath & strInfFileName & vbTab & objMatch.SubMatches(0) & vbTab & objMatchesDevDef.Item(i)
                                                        'End If
                                                    End If
                                                                                                        
                                                    '�������� ������
                                                    'strDevID & vbTab & strInfFileName & vbTab & strManSection & vbTab & strVer & vbTab & strManSectEmptyList & vbTab & lngCatFileExists & vbTab & strDevName
                                                    ' ��������������� ������� ���� ��������� �������� �����������
                                                    If lngNumLines >= lngLinesArrMax Then
                                                        lngLinesArrMax = 2 * lngLinesArrMax
                                                        ReDim Preserve strLinesArr(lngLinesArrMax)
                                                    End If
                                                    strLinesArr(lngNumLines) = (strDevID & strInfPathTabQuoted & strManSection) & (strPartString2Index & strDevName)
                                                    lngNumLines = lngNumLines + 1
                                                    
                                                    If Not objHWIDOutput.Exists(strDevID) Then
                                                        objHWIDOutput.Add strDevID, 1
                                                        If lngNumLinesHwid >= lngLinesArrHwidMax Then
                                                            lngLinesArrHwidMax = 2 * lngLinesArrHwidMax
                                                            ReDim Preserve strLinesArrHwid(lngLinesArrHwidMax)
                                                        End If
                                                        strLinesArrHwid(lngNumLinesHwid) = strDevID
                                                        lngNumLinesHwid = lngNumLinesHwid + 1
                                                    End If
                                                End If
                                            End If
                                        End If
                                    
                                    ' dev_defs'
                                    Next ii
                                
                                Else
                                    ' ���� ������ ��������, �� ��������� �� ������� ����� ��������� �� ������ �������
                                    If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManSection & "] is Empty -> this OS not Supported by inf: " & strInfPathRelative
                                End If
                            
                            Next k1
                        ' dev_sub_sects
                        
                        Else
                            ' ���� ������ c HWID �� �������
                            If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManSection & "] Not Find in inf-file: " & strInfPathRelative
                        End If
                            
                    Else
                        If mbDebugDetail Then DebugMode str2VbTab & "DevParserByRegExp: Section [" & strManSectList_x(k2) & "] is Empty -> this OS not Supported by inf: " & strInfPathRelative
                    '  dev_Sub_sects not empty
                    End If
                    
                ' dev_sects
                Next k2
            
            ' sect_list
            End If
        
        Else
            If mbDebugStandart Then DebugMode str3VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strWorkDirInfList_x(lngInfN).FullPath
        End If

SkipFileInfDrp:

    Next

    ChangeStatusBarText strMessages(121) & strSpace & strPackFileName
    
    strRezultTxt = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".txt"
    strRezultTxtHwid = strWorkTempBackSL & "rezult" & strPackFileName_woExt & ".hwid"
    strRezultTxtTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxt)), "rezult", vbNullString, , , vbTextCompare)
    strRezultTxtHwidTo = Replace$(PathCombine(strPathDevDB, GetFileNameFromPath(strRezultTxtHwid)), "rezult", vbNullString, , , vbTextCompare)
    lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
    If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Create Index Data: " & CalculateTime(lngTimeScriptFinish, True)

    ' ���� ������ �������, �� ������� ���� � ����
    If lngNumLines Then

        ReDim Preserve strLinesArr(lngNumLines - 1)
        ReDim Preserve strLinesArrHwid(lngNumLinesHwid - 1)

        ' ��������� �������
        lngTimeScriptRun = GetTimeStart
    
        If lngSortMethodShell = 0 Then
            
            Set cSortHWID = New cAsmShell
            cSortHWID.SortMethod = BinaryCompare
            cSortHWID.SortOrder = Ascending
        
            cSortHWID.sShell strLinesArrHwid, False
            
            '���� ���������, �� ��������� �������� ��������� ���� - ������ ��� �������� ������
            If mbSortDBTxtFileByHWID Then
                cSortHWID.sShell strLinesArr, False
            End If
            
            Set cSortHWID = Nothing
            
        ElseIf lngSortMethodShell = 1 Then
        
            ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLinesHwid, 4&, AddressOf CompareString
            
            '���� ���������, �� ��������� �������� ��������� ���� - ������ ��� �������� ������
            If mbSortDBTxtFileByHWID Then
                ShellSortAny VarPtr(strLinesArr(0)), lngNumLines, 4&, AddressOf CompareString
            End If
            
        ElseIf lngSortMethodShell = 2 Then
        
            Set cSortHWID2 = New cBlizzard
            cSortHWID2.SortMethod = BinaryCompare
            cSortHWID2.SortOrder = Ascending
            
            cSortHWID2.TwisterStringSort strLinesArrHwid, 0&, lngNumLinesHwid - 1
            
            '���� ���������, �� ��������� �������� ��������� ���� - ������ ��� �������� ������
            If mbSortDBTxtFileByHWID Then
                cSortHWID2.TwisterStringSort strLinesArr, 0&, lngNumLines - 1
            End If
            
            Set cSortHWID2 = Nothing
        Else
            ShellSortAny VarPtr(strLinesArrHwid(0)), lngNumLinesHwid, 4&, AddressOf CompareString
            
            '���� ���������, �� ��������� �������� ��������� ���� - ������ ��� �������� ������
            If mbSortDBTxtFileByHWID Then
                ShellSortAny VarPtr(strLinesArr(0)), lngNumLines, 4&, AddressOf CompareString
            End If
        End If
        
        lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Sort Index: " & CalculateTime(lngTimeScriptFinish, True)
        DoEvents
        lngTimeScriptRun = GetTimeStart

        '---------------------------------------------
        '---------------������� ���� � ����-----
        If PathExists(strPathDevDB) = False Then
            CreateNewDirectory strPathDevDB
        End If
        strRezultTxtTemp = Replace$(strRezultTxt, "rezult", vbNullString)
        ' ������ � ���� �������
        FileWriteData strRezultTxtTemp, Join(strLinesArr(), vbNewLine)
        ' ������ � ���� ��������� �������
        FileWriteData strRezultTxtHwid, Join(strLinesArrHwid(), vbNewLine)
        
        lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp-Time to Save Index Files: " & CalculateTime(lngTimeScriptFinish, True)
        
        ' �������� �������, �.� ������������ ������
        Erase strLinesArr
        Erase strLinesArrHwid

        '7z a -t7z -mx=1 archive.7z filename.txt
        ' ������ ������ �����-�������
        
        cmdString = strQuotes & strArh7zExePath & strQuotes & " a -y -t7z -mx=1 " & strQuotes & strRezultTxt & strQuotes & strSpace & strQuotes & strRezultTxtTemp & strQuotes & " -sdel"

        If RunAndWaitNew(cmdString, strWorkTemp, vbHide) = False Then
            MsgBox strMessages(45) & str2vbNewLine & cmdString, vbInformation, strProductName
        Else

            ' ��������� ��������� �� ��� 100%? ���� ��� �� ��������
            If lngExitProc <> 0 Then
                If lngExitProc = 2 Then
                    MsgBox strMessages(45) & str2vbNewLine & cmdString, vbInformation, strProductName
                ElseIf lngExitProc = 7 Then
                    MsgBox strMessages(45) & str2vbNewLine & cmdString, vbInformation, strProductName
                ElseIf lngExitProc = 255 Then
                    MsgBox strMessages(45) & str2vbNewLine & cmdString, vbInformation, strProductName
                End If
            End If
        End If
        
        If CopyFileTo(strRezultTxt, strRezultTxtTo) Then
            '�������� ���� HWID
            If CopyFileTo(strRezultTxtHwid, strRezultTxtHwidTo) Then
                ' ���������� ������ ���� ��������� � ini-����
                IniWriteStrPrivate strPackFileName_woExt, "Version", lngDevDBVersion, PathCombine(strPathDevDB, "DevDBVersions.ini")
                '���� ���� DriverPack*.ini
                strRezultTxt = vbNullString
                strRezultTxt_x = SearchFilesInRoot(strWorkDir, "DriverPack*.ini", False, True)
                strRezultTxt = strRezultTxt_x(0).FullPath

                ' �������� DriverPack*.ini � ������� ���� ������
                If FileExists(strRezultTxt) Then
                    If FileExists(strRezultTxtTo) Then
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

    Set objMatchStrSect = Nothing
    Set objMatchStrDefs = Nothing
    Set objMatchVer = Nothing
    Set objMatchCat = Nothing
    Set objMatchManSect = Nothing
    Set objMatchDevSect = Nothing
    Set objMatchManDef = Nothing
    Set objMatchManID = Nothing
    Set objMatchDevDef = Nothing
    Set objMatchesStrSect = Nothing
    Set objMatchesVerSect = Nothing
    Set objMatchesVerParam = Nothing
    Set objMatchesCatParam = Nothing
    Set objMatchesManSect = Nothing
    Set objMatchesManDef = Nothing
    Set objMatchesManID = Nothing
    Set objMatchesDevSect = Nothing
    Set objMatchesDevDef = Nothing
    Set objMatchesDevID = Nothing
    Set objMatchesStrDefs = Nothing
    Set RegExpStrSect = Nothing
    Set RegExpStrDefs = Nothing
    Set RegExpVerSect = Nothing
    Set RegExpVerParam = Nothing
    Set RegExpCatParam = Nothing
    Set RegExpManSect = Nothing
    Set RegExpManDef = Nothing
    Set RegManID = Nothing
    Set RegExpDevDef = Nothing
    Set RegExpDevSect = Nothing
    Set RegExpReplace = Nothing
    Set objHashOutput = Nothing
    Set objStringHash = Nothing
    Set objHWIDOutput = Nothing

    If mbDebugStandart Then DebugMode vbTab & "DevParserByRegExp-End"
End Sub
