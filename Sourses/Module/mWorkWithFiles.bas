Attribute VB_Name = "mWorkWithFiles"
Option Explicit

' Переменные для работы с файловой системой
Public objFSO              As Scripting.FileSystemObject

'Константы для FSO
Public Const ForWriting    As Long = 2
Public Const ForAppending  As Long = 8
Public Const ForReading    As Long = 1

Private Root               As String
Private xFOL               As Folder
Private xFile              As File

' Переменная
Public strFileListInFolder As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BacklashDelFromPath
'! Description (Описание)  :   [Удаление слэша на конце]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function BacklashDelFromPath(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathRemoveBackslash strPath
    BacklashDelFromPath = TrimNull(strPath)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BackslashAdd2Path
'! Description (Описание)  :   [Добавление слэша на конце]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function BackslashAdd2Path(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathAddBackslash strPath
    BackslashAdd2Path = TrimNull(strPath)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompareFilesByHashCAPICOM
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strFirstFile (String)
'                              strSecondFile (String)
'!--------------------------------------------------------------------------------
Public Function CompareFilesByHashCAPICOM(ByVal strFirstFile As String, ByVal strSecondFile As String) As Boolean

    Dim strDataSHAFirst  As String
    Dim strDataSHASecond As String
    Dim lngResult        As Long

    If PathExists(strFirstFile) Then
        strDataSHAFirst = CalcHashFile(strFirstFile, CAPICOM_HASH_ALGORITHM_SHA1)
    End If

    If PathExists(strSecondFile) Then
        strDataSHASecond = CalcHashFile(strSecondFile, CAPICOM_HASH_ALGORITHM_SHA1)
    End If

    lngResult = StrComp(strDataSHAFirst, strDataSHASecond, vbTextCompare)

    If lngResult = 0 Then
        CompareFilesByHashCAPICOM = True
    Else
        CompareFilesByHashCAPICOM = False
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CopyFileTo
'! Description (Описание)  :   [Скопирует файл 'PathFrom' в директорию 'CopyFileTo', Если файл существует, то он будет перезаписан новым файлом.]
'! Parameters  (Переменные):   PathFrom (String)
'                              PathTo (String)
'!--------------------------------------------------------------------------------
Public Function CopyFileTo(ByVal PathFrom As String, ByVal PathTo As String) As Boolean

    Dim ret As Long

    If PathExists(PathFrom) Then
        ' Для всех файлов, сброс атрибута только для чтения, и системный если есть
        ResetReadOnly4File PathTo
        ' Собственно копирование
        'Если вы хотите, чтобы новый файл не записывался на место старого, замените 'False' на 'True'
        ret = CopyFile(PathFrom, PathTo, False)

        If ret <> 0 Then
            CopyFileTo = True
            ' Сброс атрибута только для чтения, если есть
            ResetReadOnly4File PathTo
        Else
            CopyFileTo = False
            MsgBox strMessages(42) & vbNewLine & "From: " & PathFrom & vbNewLine & "To:" & PathTo & vbNewLine & "Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError), vbExclamation, strProductName
            DebugMode vbTab & "Copy file: False: " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        CopyFileTo = False
        DebugMode vbTab & "Copy file: False : " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CreateNewDirectory
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   NewDirectory (String)
'!--------------------------------------------------------------------------------
Public Sub CreateNewDirectory(ByVal NewDirectory As String)

    Dim SecAttrib  As SECURITY_ATTRIBUTES
    Dim sPath      As String
    Dim iCounter   As Integer
    Dim sTempDir   As String
    Dim ret        As Long
    Dim retLasrErr As Long

    sPath = BackslashAdd2Path(NewDirectory)
    iCounter = 1

    Do Until InStr(iCounter, sPath, vbBackslash) = 0
        iCounter = InStr(iCounter, sPath, vbBackslash)
        sTempDir = Left$(sPath, iCounter)
        iCounter = iCounter + 1

        'create directory
        With SecAttrib
            .lpSecurityDescriptor = &O0
            .bInheritHandle = False
            .nLength = Len(SecAttrib)
        End With

        ret = CreateDirectory(sTempDir, SecAttrib)

        If ret = 0 Then
            retLasrErr = Err.LastDllError

            If PathExists(sTempDir) = False Then
                DebugMode vbTab & "CreateDirectory: False : " & sTempDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)
            End If
        End If

    Loop

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DeleteFiles
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PathFile (String)
'!--------------------------------------------------------------------------------
Public Function DeleteFiles(ByVal PathFile As String) As Boolean

    Dim ret As Long

    If PathIsValidUNC(PathFile) = False Then
        ret = DeleteFile(StrPtr("\\?\" & PathFile & vbNullChar))
    Else
        '\\?\UNC\
        ret = DeleteFile(StrPtr("\\?\UNC\" & Right$(PathFile, Len(PathFile) - 2) & vbNullChar))
    End If

    DeleteFiles = CBool(ret)

    If ret = 0 Then
        If PathExists(PathFile) Then

            On Error GoTo errhandler

            objFSO.DeleteFile PathFile, True
        End If

        If PathExists(PathFile) Then
            DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    Exit Function

errhandler:
    DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.Number & ": " & Err.Description & vbNewLine & _
              vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
    Err.Clear

    Resume Next

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelFolderBackUp
'! Description (Описание)  :   [Удаление временного каталога, если включена опция]
'! Parameters  (Переменные):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Sub DelFolderBackUp(ByVal strFolderPath As String)

    Dim ret As Long

    On Error Resume Next

    DebugMode "DelFolder-Start: " & strFolderPath

    If PathExists(strFolderPath) Then
        DelRecursiveFolder strFolderPath
    End If

    If PathExists(strFolderPath) Then
        ret = RemoveDirectory(strFolderPath)

        If ret = 0 Then
            DebugMode vbTab & "RemoveDirectory: False : " & strFolderPath & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    On Error GoTo 0

    DebugMode "DelFolder-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelRecursiveFolder
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Folder (String)
'!--------------------------------------------------------------------------------
Public Sub DelRecursiveFolder(ByVal Folder As String)

    Dim retDelete As Long
    Dim retStrMsg As String

    Root = BacklashDelFromPath(Folder)
    DebugMode vbTab & "DeleteFolder: " & Root

    If PathExists(Root) Then
        SearchFilesInRoot Root, ALL_FILES, True, False, True
        Set xFOL = objFSO.GetFolder(Root)

        If xFOL.Files.Count > 0 Then

            For Each xFile In xFOL.Files
                DeleteFiles xFile.Path
            Next

        End If

        ' Удаление пустых каталогов
        If PathExists(Root) Then
            retDelete = DelTree(Root)

            If mbDebugEnable Then

                Select Case retDelete

                    Case 0
                        retStrMsg = "Deleted"

                    Case -1
                        retStrMsg = "Invalid Directory"

                    Case Else
                        retStrMsg = "An Error was occured"
                End Select

                DebugMode vbTab & "DeleteFolder: " & " Result: " & retStrMsg
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelTemp
'! Description (Описание)  :   [Удаление временного каталога, если включена опция]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub DelTemp()

    On Error Resume Next

    DebugMode "DelTemp-Start"

    If PathExists(strWorkTemp) Then
        DelRecursiveFolder strWorkTemp
    End If

    If PathExists(strWorkTemp) Then
        RemoveDirectory strWorkTemp
    End If

    On Error GoTo 0

    DebugMode "DelTemp-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DelTree
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strDir (String)
'!--------------------------------------------------------------------------------
Private Function DelTree(ByVal strDir As String) As Long

    Dim X          As Long
    Dim intAttr    As Integer
    Dim strAllDirs As String
    Dim strFile    As String
    Dim ret        As Long
    Dim retLasrErr As Long

    DelTree = -1

    On Error Resume Next

    strDir = Trim$(strDir)

    If LenB(strDir) > 0 Then
        If Right$(strDir, 1) = vbBackslash Then
            strDir = Left$(strDir, Len(strDir) - 1)
        End If

        If InStr(strDir, vbBackslash) Then
            intAttr = GetAttr(strDir)

            If (intAttr And vbDirectory) Then
                strDir = BackslashAdd2Path(strDir)
                strFile = Dir$(strDir & ALL_FILES, vbSystem Or vbDirectory Or vbHidden)

                Do While Len(strFile)

                    If strFile <> "." Then
                        If strFile <> ".." Then
                            intAttr = GetAttr(strDir & strFile)

                            If (intAttr And vbDirectory) Then
                                strAllDirs = strAllDirs & strFile & vbNullChar
                            Else

                                If intAttr <> vbNormal Then
                                    SetAttr strDir & strFile, vbNormal

                                    If Err Then
                                        DelTree = Err.Number
                                    End If

                                    Exit Function

                                End If

                                DeleteFiles strDir & strFile

                                If Err Then
                                    DelTree = Err.Number
                                End If

                                Exit Function

                            End If
                        End If
                    End If

                    strFile = Dir
                Loop

                Do While Len(strAllDirs)
                    X = InStr(strAllDirs, vbNullChar)
                    strFile = Left$(strAllDirs, X - 1)
                    strAllDirs = Mid$(strAllDirs, X + 1)
                    X = DelTree(strDir & strFile)

                    If X Then
                        DelTree = X
                    End If

                Loop

                ret = RemoveDirectory(strDir)

                If ret = 0 Then
                    retLasrErr = Err.LastDllError

                    If PathExists(strDir) = False Then
                        DebugMode vbTab & "RemoveDirectory: False : " & strDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)
                    End If

                    DelTree = retLasrErr
                Else
                    DelTree = 0
                End If

                If Err Then
                    DelTree = Err.Number
                Else
                    DelTree = 0
                End If

                On Error GoTo 0

            End If
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ExtFromFileName
'! Description (Описание)  :   [Получить расширение файла из пути или имени файла]
'! Parameters  (Переменные):   FileName (String)
'!--------------------------------------------------------------------------------
Public Function ExtFromFileName(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FileName, ".")

    If intLastSeparator > 0 Then
        ExtFromFileName = Right$(FileName, Len(FileName) - intLastSeparator)
    Else
        ExtFromFileName = vbNullString
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileisReadOnly
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PathFile (String)
'!--------------------------------------------------------------------------------
Public Function FileisReadOnly(ByVal PathFile As String) As Boolean
    FileisReadOnly = GetAttr(PathFile) And vbReadOnly
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileisSystemAttr
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   PathFile (String)
'!--------------------------------------------------------------------------------
Public Function FileisSystemAttr(PathFile As String) As Boolean
    FileisSystemAttr = GetAttr(PathFile) And vbSystem
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileName_woExt
'! Description (Описание)  :   [Получить имя файла без расширения, зная имя файла]
'! Parameters  (Переменные):   FileName (String)
'!--------------------------------------------------------------------------------
Public Function FileName_woExt(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    FileName_woExt = FileName

    If LenB(FileName) > 0 Then
        intLastSeparator = InStrRev(FileName, ".")

        If intLastSeparator > 0 Then
            FileName_woExt = Left$(FileName, intLastSeparator - 1)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileNameFromPath
'! Description (Описание)  :   [Получить имя файла из полного пути]
'! Parameters  (Переменные):   FilePath (String)
'!--------------------------------------------------------------------------------
Public Function FileNameFromPath(ByVal FilePath As String) As String

    Dim intLastSeparator As Long

    FileNameFromPath = FilePath

    If LenB(FilePath) > 0 Then
        intLastSeparator = InStrRev(FilePath, vbBackslash)

        If intLastSeparator >= 0 Then
            FileNameFromPath = Right$(FilePath, Len(FilePath) - intLastSeparator)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetEnviron
'! Description (Описание)  :   [Получение переменной системного окружения]
'! Parameters  (Переменные):   strEnv (String)
'                              mbCollectFull (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function GetEnviron(ByVal strEnv As String, Optional ByVal mbCollectFull As Boolean = False) As String

    Dim strTemp        As String
    Dim strTempEnv     As String
    Dim strNumPosition As Long

    strNumPosition = InStr(strEnv, Percentage)

    If strNumPosition > 0 Then
        strTemp = Mid$(strEnv, strNumPosition + 1, Len(strEnv) - strNumPosition)
        strNumPosition = InStr(strTemp, Percentage)

        If strNumPosition > 0 Then
            strTemp = Left$(strTemp, strNumPosition - 1)
        End If
    End If

    strTempEnv = Environ$(strTemp)

    If mbCollectFull Then
        GetEnviron = Replace$(strEnv, Percentage & strTemp & Percentage, strTempEnv, , , vbTextCompare)
    Else
        GetEnviron = strTempEnv
    End If

    DebugMode str2VbTab & "GetEnviron: %" & strTemp & "%=" & strTempEnv & vbNewLine & _
              str2VbTab & "GetEnviron-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetUniqueTempFile
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetUniqueTempFile() As String

    Dim ll_Buffer       As Long
    Dim ls_TempFileName As String

    ll_Buffer = 255
    ls_TempFileName = Space$(255)
    ll_Buffer = GetTempFileName(strWinTemp, "xdia", 0, ls_TempFileName)

    'xxx is a three letter prefix - can be anything you want.
    '3rd parameter (0 above) is uUnique...If uUnique is nonzero, the function appends the hexadecimal string to lpPrefixString to form the temporary filename. In this case, the function does not create the specified file, and does not test whether the filename is unique.
    'If uUnique is zero, the function uses a hexadecimal string derived from the current system time. In this case, the function uses different values until it finds a unique filename, and then it creates the file in the lpPathName directory.
    If ll_Buffer = 0 Then
        MsgBox strMessages(44) & vbNewLine & strWinTemp, vbCritical, strProductName
    Else
        ls_TempFileName = Left$(ls_TempFileName, ll_Buffer)
        GetUniqueTempFile = ls_TempFileName
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsDriveCDRoom
'! Description (Описание)  :   [Проверка на запск программы с CD\DVD]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsDriveCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    IsDriveCDRoom = False
    strDriveName = Left$(strAppPath, 2)

    ' Проверяем на запуск из сети
    If InStr(strDriveName, vbBackslash) = 0 Then
        'получаем тип диска
        Set xDrv = objFSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            IsDriveCDRoom = True
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathIsAFolder
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sPath (String)
'!--------------------------------------------------------------------------------
Public Function PathIsAFolder(ByVal sPath As String) As Boolean

    'Verifies that a path is a valid
    'directory, and returns True (1) if
    'the path is a valid directory,
    'or False otherwise. The path must
    'exist.
    'If the path is a directory on the
    'local machine, PathIsDirectory returns
    '16 (the file attribute for a folder).
    'If the path is a directory on a server
    'share, PathIsDirectory returns 1.
    'If it is neither PathIsDirectory returns 0.
    PathIsAFolder = PathIsDirectory(StrPtr(sPath & vbNullChar))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function MoveFileTo
'! Description (Описание)  :   [Скопирует файл 'PathFrom' в директорию 'PathTo', Если файл существует, то он будет перезаписан новым файлом.]
'! Parameters  (Переменные):   PathFrom (String)
'                              PathTo (String)
'!--------------------------------------------------------------------------------
Public Function MoveFileTo(PathFrom As String, PathTo As String) As Boolean

    Dim ret As Long

    If StrComp(PathFrom, PathTo, vbTextCompare) <> 0 Then
        If PathExists(PathFrom) Then
            ' Для всех файлов, сброс атрибута только для чтения, и системный если есть
            ResetReadOnly4File PathTo
            ' Собственно копирование
            'Если вы хотите, чтобы новый файл не записывался на место старого, замените 'False' на 'True'
            ret = MoveFile(PathFrom, PathTo)

            If ret <> 0 Then
                MoveFileTo = True
                ' Сброс атрибута только для чтения, если есть
                ResetReadOnly4File PathTo
            Else
                MoveFileTo = False
                MsgBox strMessages(42) & vbNewLine & "From: " & PathFrom & vbNewLine & "To:" & PathTo & vbNewLine & "Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError), vbExclamation, strProductName
                DebugMode vbTab & "Move file: False: " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If

        Else
            MoveFileTo = False
            DebugMode vbTab & "Move file: False : " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        DebugMode vbTab & "Move file: Source and Destination are identicaly (" & PathFrom & " ; " & PathTo & ")"
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ParserInf4Strings
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strInfFilePath (String)
'                              strSearchString (String)
'!--------------------------------------------------------------------------------
Public Function ParserInf4Strings(ByVal strInfFilePath As String, ByVal strSearchString As String) As String

    Dim StringHash     As Scripting.Dictionary
    Dim objInfFile     As TextStream
    Dim RegExpStrSect  As RegExp
    Dim RegExpStrDefs  As RegExp
    Dim MatchesStrSect As MatchCollection
    Dim MatchesStrDefs As MatchCollection
    Dim objMatch       As Match
    Dim objMatch1      As Match
    Dim regex_strsect  As String
    Dim regex_strings  As String
    Dim r_beg          As String
    Dim r_identS       As String
    Dim r_str          As String
    Dim FileContent    As String
    Dim Key            As String
    Dim Value          As String
    Dim R              As Boolean
    Dim i              As Long
    Dim Strings        As String
    Dim valval         As String
    Dim varname        As String
    Dim lngFileDBSize  As Long
    Dim Pos            As Long

    r_beg = "^[ \t]*"
    r_identS = "([^; \t\r\n][^;\t\r\n]*[^; \t\r\n])"
    r_str = "(?:""([^\r\n""]*)""|([^\r\n;]*))"
    regex_strsect = r_beg & "\[strings\](?:([\s\S]*?)" & r_beg & "(?=\[)|([\s\S]*))"
    ' variable = "str"
    regex_strings = r_beg & r_identS & "[ \t]*=[ \t]*" & r_str
    ' Init regexps
    Set RegExpStrSect = New RegExp

    With RegExpStrSect
        .Pattern = regex_strsect
        .MultiLine = True
        .IgnoreCase = True
        .Global = False
        ' Note: "XP Alternative (by Greg)\D\3\M\A\12\prime.inf" has two [strings] sections
    End With

    Set RegExpStrDefs = New RegExp

    With RegExpStrDefs
        .Pattern = regex_strings
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    ' Read INF file
    FileContent = vbNullString
    lngFileDBSize = GetFileSizeByPath(strInfFilePath)

    If lngFileDBSize > 0 Then
        Set objInfFile = objFSO.OpenTextFile(strInfFilePath, ForReading, False, TristateUseDefault)
        FileContent = objInfFile.ReadAll()
        objInfFile.Close
    Else
        DebugMode str2VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfFilePath
    End If

    ' Find [strings] section
    Strings = vbNullString
    StringHash.CompareMode = TextCompare
    Set MatchesStrSect = RegExpStrSect.Execute(FileContent)

    If MatchesStrSect.Count >= 1 Then
        Set objMatch = MatchesStrSect.Item(0)
        Strings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
        Set MatchesStrDefs = RegExpStrDefs.Execute(Strings)

        For i = 0 To MatchesStrDefs.Count - 1
            Set objMatch1 = MatchesStrDefs.Item(i)
            Key = objMatch1.SubMatches(0)
            Value = objMatch1.SubMatches(1)

            If LenB(Value) = 0 Then
                Value = objMatch1.SubMatches(2)
            End If

            R = StringHash.Exists(Key)

            If Not R Then
                StringHash.Add Key, Value
                StringHash.Add Percentage & Key & Percentage, Value
            End If

        Next

    End If

    ' Собственно ищем саму переменную
    Pos = InStr(strSearchString, Percentage)

    If Pos > 0 Then
        varname = Mid$(strSearchString, Pos, InStrRev(strSearchString, Percentage))
        valval = StringHash.Item(varname)

        If LenB(valval) = 0 Then
            DebugMode "ParserInf4Strings: Error in inf: Cannot find '" & strSearchString & "'"
        Else
            ParserInf4Strings = Replace$(strSearchString, varname, valval)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathNameFromPath
'! Description (Описание)  :   [Получить путь к файлу из полного пути]
'! Parameters  (Переменные):   FilePath (String)
'!--------------------------------------------------------------------------------
Public Function PathNameFromPath(FilePath As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FilePath, vbBackslash)

    If intLastSeparator > 0 Then
        If intLastSeparator < Len(FilePath) Then
            PathNameFromPath = Left$(FilePath, intLastSeparator)
        Else
            PathNameFromPath = FilePath
        End If

    Else
        PathNameFromPath = FilePath
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ResetReadOnly4File
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Public Sub ResetReadOnly4File(ByVal StrPathFile As String)

    If PathExists(StrPathFile) Then
        If FileisReadOnly(StrPathFile) Then
            SetAttr StrPathFile, vbNormal
        End If

        If FileisSystemAttr(StrPathFile) Then
            SetAttr StrPathFile, vbNormal
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SafeDir
'! Description (Описание)  :   [function to replace special chars to create dirs correctly]
'! Parameters  (Переменные):   str (String)
'!--------------------------------------------------------------------------------
Public Function SafeDir(ByVal str As String) As String

    Dim R As String

    R = str
    R = Replace$(R, vbBackslash, "_")
    R = Replace$(R, "/", "-")
    R = Replace$(R, "*", "_")
    R = Replace$(R, ":", "_")
    R = Replace$(R, ";", "_")
    R = Replace$(R, "?", "_")
    R = Replace$(R, ">", "_")
    R = Replace$(R, "<", "_")
    R = Replace$(R, "|", "_")
    R = Replace$(R, "@", "_")
    R = Replace$(R, "'", "")
    R = Replace$(R, " ", "_")
    R = Replace$(R, "_-_", "_")
    R = Replace$(R, "(R)", "_")
    R = Replace$(R, "___", "_")
    R = Replace$(R, "__", "_")
    R = Trim$(R)
    SafeDir = R
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SafeFileName
'! Description (Описание)  :   [function to replace special chars to create files correctly]
'! Parameters  (Переменные):   strString (Variant)
'!--------------------------------------------------------------------------------
Public Function SafeFileName(ByVal strString As String) As String
    ' Заменяем VbTab
    strString = Replace$(strString, vbTab, vbNullString)
    strString = TrimNull(strString)

    ' Отбрасываем все после ","
    If InStr(strString, ",") Then
        strString = Left$(strString, InStr(strString, ",") - 1)
    End If

    ' Отбрасываем все после ";"
    If InStr(strString, ";") Then
        strString = Left$(strString, InStr(strString, ";") - 1)
    End If

    strString = Trim$(TrimNull(strString))
    SafeFileName = strString
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function WhereIsDir
'! Description (Описание)  :   [function to discover dirs with inf code]
'! Parameters  (Переменные):   str (String)
'                              strInfFilePath (String)
'!--------------------------------------------------------------------------------
Public Function WhereIsDir(ByVal str As String, ByVal strInfFilePath As String) As String

    Dim cDir             As String
    Dim Str_x()          As String
    Dim mbAdditionalPath As Boolean

    If InStr(str, ";") Then
        Str_x = Split(str, ";")
        str = Trim$(Str_x(0))
    End If

    If InStr(str, ",") Then
        Str_x = Split(str, ",")
        mbAdditionalPath = True
        str = Str_x(0)
    End If

    If InStr(str, vbNullChar) Then
        str = TrimNull(str)
    End If

    If InStr(str, vbTab) Then
        str = Replace$(str, vbTab, vbNullString)
    End If

    'http://msdn.microsoft.com/en-us/library/ff553598.aspx
    Select Case str

        Case "01"
            cDir = strSysDrive

        Case "10"
            cDir = strWinDir

            'system32 независимо от винды
        Case "11"
            cDir = strSysDir86

        Case "12"
            cDir = strSysDir86 & "Drivers"

        Case "17"
            cDir = strInfDir

        Case "18"
            cDir = strWinDir & "Help"

        Case "20"
            cDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "21"
            cDir = vbNullString

            'viewer dir
        Case "23"
            cDir = strSysDir86 & "spool\drivers\color"

        Case "24"
            cDir = strSysDrive

        Case "25"
            cDir = vbNullString

            'shared dir
        Case "30"
            cDir = strSysDrive

        Case "50"
            cDir = strWinDir & "system"

        Case "51"
            cDir = strSysDir86 & "Spool"

        Case "52"
            cDir = strSysDir86 & "Spool\Drivers"

        Case "53"
            cDir = vbNullString

            'user profile dir
        Case "54"
            cDir = vbNullString

            ' ntldr.exe dir
        Case "55"
            cDir = strSysDir86 & "spool\prtprocs"

        Case "16384"
            cDir = GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)

        Case "16386"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAMS)

        Case "16389"
            cDir = GetSpecialFolderPath(CSIDL_MYDOCUMENTS)

        Case "16391"
            cDir = GetSpecialFolderPath(CSIDL_STARTUP)

        Case "16392"
            cDir = GetSpecialFolderPath(CSIDL_RECENT)

        Case "16393"
            cDir = GetSpecialFolderPath(CSIDL_SENDTO)

        Case "16395"
            cDir = GetSpecialFolderPath(CSIDL_STARTMENU)

        Case "16397"
            cDir = GetSpecialFolderPath(CSIDL_MYMUSIC)

        Case "16397"
            cDir = GetSpecialFolderPath(CSIDL_MYVIDEO)

        Case "16400"
            cDir = GetSpecialFolderPath(CSIDL_DESKTOP)

        Case "16403"
            cDir = GetSpecialFolderPath(CSIDL_NETHOOD)

        Case "16404"
            cDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "16405"
            cDir = GetSpecialFolderPath(CSIDL_TEMPLATES)

        Case "16406"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_STARTMENU)

        Case "16407"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)

        Case "16408"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_STARTUP)

        Case "16409"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)

        Case "16410"
            cDir = GetSpecialFolderPath(CSIDL_APPDATA)

        Case "16411"
            cDir = GetSpecialFolderPath(CSIDL_PRINTHOOD)

        Case "16412"
            cDir = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)

        Case "16415"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_FAVORITES)

        Case "16416"
            cDir = GetSpecialFolderPath(CSIDL_INTERNET_CACHE)

        Case "16417"
            cDir = GetSpecialFolderPath(CSIDL_COOKIES)

        Case "16418"
            cDir = GetSpecialFolderPath(CSIDL_HISTORY)

        Case "16419"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_APPDATA)

        Case "16420"
            cDir = strWinDir

        Case "16421"
            cDir = strSysDir86

        Case "16422"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES)

        Case "16423"
            cDir = GetSpecialFolderPath(CSIDL_MYPICTURES)

        Case "16424"
            cDir = GetSpecialFolderPath(CSIDL_PROFILE)

        Case "16425"
            cDir = strSysDir64

        Case "16426"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16427"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMON)

        Case "16428"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMONX86)

        Case "16429"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_TEMPLATES)

        Case "16430"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_DOCUMENTS)

        Case "16432"
            cDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16437"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_MUSIC)

        Case "16438"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_PICTURES)

        Case "16439"
            cDir = GetSpecialFolderPath(CSIDL_COMMON_VIDEO)

        Case "16440"
            cDir = strWinDir & "resources"

        Case "16441"
            cDir = strWinDir & "resources\0409"

        Case "-1"
            cDir = vbNullString

            ' absolute path
            'http://msdn.microsoft.com/en-us/library/ff560821.aspx
        Case "66000"
            cDir = Getpath_PrinterDriverDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\Drivers\w32x86"
            End If

        Case "66001"
            cDir = Getpath_PrintProcessorDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\prtprocs\w32x86"
            End If

        Case "66002"
            cDir = strSysDir86

        Case "66003"
            cDir = Getpath_PrinterColorDirectory

            If LenB(cDir) = 0 Then
                cDir = strSysDir86 & "spool\drivers\color"
            End If

        Case "66004"
            cDir = strSysDir86 & "spool\Drivers\w32x86"

        Case Else
            cDir = vbNullString
    End Select

    If InStr(cDir, vbNullChar) Then
        cDir = TrimNull(cDir)
    End If

    If mbAdditionalPath Then
        cDir = BackslashAdd2Path(cDir) & Trim$(Str_x(1))

        If InStr(cDir, Percentage) Then
            cDir = ParserInf4Strings(strInfFilePath, cDir)
        End If
    End If

    cDir = Replace$(cDir, vbTab, vbNullString)
    cDir = Replace$(cDir, Kavichki, vbNullString)
    cDir = BackslashAdd2Path(cDir)
    WhereIsDir = TrimNull(cDir)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathCollect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Path (String)
'!--------------------------------------------------------------------------------
Public Function PathCollect(Path As String) As String

    If InStr(Path, ":") = 2 Then
        PathCollect = Path
    ElseIf Left$(Path, 2) = vbBackslash And PathIsValidUNC(Path) Then
        PathCollect = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect = PathCombine(strAppPath, Path)
        Else

            If InStr(Path, vbBackslash) = 1 Then
                PathCollect = strAppPath & Path
            Else

                If Left$(Path, 3) = "..\" Then
                    PathCollect = PathCombine(strAppPath, Path)
                Else

                    If InStr(Path, Percentage) Then
                        PathCollect = GetEnviron(Path, True)
                    Else

                        If LenB(ExtFromFileName(Path)) > 0 Then
                            If FileNameFromPath(Path) = Path Then
                                PathCollect = Path
                            Else
                                PathCollect = strAppPathBackSL & Path
                            End If

                        Else
                            PathCollect = strAppPathBackSL & Path
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect, vbBackslash) Then
        If Left$(strAppPath, 2) <> vbBackslash Then
            PathCollect = Replace$(PathCollect, vbBackslash, vbBackslash)
        End If
    End If

    If PathIsAFolder(PathCollect) Then
        PathCollect = BackslashAdd2Path(PathCollect)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathCollect4Dest
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Path (String)
'                              strDest (String)
'!--------------------------------------------------------------------------------
Public Function PathCollect4Dest(ByVal Path As String, ByVal strDest As String) As String

    If InStr(Path, ":") = 2 Then
        PathCollect4Dest = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect4Dest = strDest & Mid$(Path, 2, Len(Path) - 1)
        Else

            If InStr(Path, vbBackslash) = 1 Then
                PathCollect4Dest = strDest & Path
            Else

                If Left$(Path, 3) = "..\" Then
                    PathCollect4Dest = PathNameFromPath(strDest) & Mid$(Path, 4, Len(Path) - 1)
                Else

                    If InStr(Path, Percentage) Then
                        PathCollect4Dest = GetEnviron(Path, True)
                    Else

                        If LenB(ExtFromFileName(Path)) > 0 Then
                            If FileNameFromPath(Path) = Path Then
                                PathCollect4Dest = Path
                            Else
                                PathCollect4Dest = BackslashAdd2Path(strDest) & Path
                            End If

                        Else
                            PathCollect4Dest = BackslashAdd2Path(strDest) & Path
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect4Dest, vbBackslash) Then
        PathCollect4Dest = Replace$(PathCollect4Dest, vbBackslash, vbBackslash)

        If Left$(strDest, 2) = vbBackslash Then
            If InStr(PathCollect4Dest, vbBackslash) = 1 Then
                PathCollect4Dest = vbBackslash & PathCollect4Dest
            Else
                PathCollect4Dest = vbBackslash & PathCollect4Dest
            End If
        End If
    End If

    PathCollect4Dest = BackslashAdd2Path(PathCollect4Dest)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathIsValidUNC
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sPath (String)
'!--------------------------------------------------------------------------------
Public Function PathIsValidUNC(ByVal sPath As String) As Boolean
    ' Returns True if the string is a valid UNC path.
    PathIsValidUNC = PathIsUNC(StrPtr(sPath))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ExpandFileNamebyEnvironment
'! Description (Описание)  :   [Расширить имя файла - использование переменных %%]
'! Parameters  (Переменные):   strFileName (String)
'!--------------------------------------------------------------------------------
Public Function ExpandFileNamebyEnvironment(ByVal strFileName As String) As String

    Dim R         As String
    Dim str_OSVer As String
    Dim str_OSBit As String
    Dim str_DATE  As String

    If InStr(strFileName, Percentage) Then
        ' Макроподстановка версия ОС %OSVer%
        str_OSVer = "wnt" & Left$(strOsCurrentVersion, 1)

        ' Макроподстановка битность ОС %OSBit%
        If mbIsWin64 Then
            str_OSBit = "x64"
        Else
            str_OSBit = "x32"
        End If

        ' Макроподстановка ДАТА %DATE%
        str_DATE = Replace$(CStr(Now()), ".", "-")
        str_DATE = SafeDir(str_DATE)
        ' Замена макросов значениями
        R = strFileName
        R = Replace$(R, "%PCNAME%", strCompModel, , , vbTextCompare)
        R = Replace$(R, "%PCMODEL%", Replace$(strCompModel, " ", "_"))
        R = Replace$(R, "%OSVer%", str_OSVer, , , vbTextCompare)
        R = Replace$(R, "%OSBit%", str_OSBit, , , vbTextCompare)
        R = Replace$(R, "%DATE%", str_DATE, , , vbTextCompare)
        R = Trim$(R)
        ExpandFileNamebyEnvironment = R
    Else
        ExpandFileNamebyEnvironment = strFileName
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CreateIfNotExistPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Function CreateIfNotExistPath(strFolderPath As String) As Boolean

    If LenB(strFolderPath) > 0 Then

        ' Если нет, то создаем каталог
        If PathExists(strFolderPath) = False Then
            CreateNewDirectory strFolderPath
            CreateIfNotExistPath = PathIsAFolder(strFolderPath)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CopyFolderByShell
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sSource (String)
'                              sDestination (String)
'!--------------------------------------------------------------------------------
Public Function CopyFolderByShell(sSource As String, sDestination As String) As Long

    Dim FOF_FLAGS As Long
    Dim SHFileOp  As SHFILEOPSTRUCT

    'terminate the folder string with a pair of nulls
    sSource = BacklashDelFromPath(sSource) & str2vbNullChar

    If PathExists(sDestination) = False Then
        CreateIfNotExistPath sDestination
    End If

    sDestination = BacklashDelFromPath(sDestination) & str2vbNullChar
    'determine the user's options selected
    FOF_FLAGS = FOF_FLAGS Or FOF_RENAMEONCOLLISION Or FOF_NOCONFIRMATION

    'set up the options
    With SHFileOp
        .wFunc = FO_COPY
        .pFrom = sSource
        .pTo = sDestination
        .fFlags = FOF_FLAGS
    End With

    'and perform the chosen copy or move operation
    CopyFolderByShell = SHFileOperation(SHFileOp)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathCombine
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strDirectory (String)
'                              strFile (String)
'!--------------------------------------------------------------------------------
Public Function PathCombine(ByVal strDirectory As String, ByVal strFile As String) As String

    Dim strBuffer As String

    ' Concatenates two strings that represent properly formed
    ' paths into one path, as well as any relative path pieces.
    strBuffer = String$(MAX_PATH_UNICODE, vbNullChar)

    If PathCombineW(StrPtr(strBuffer), StrPtr(strDirectory & vbNullChar), StrPtr(strFile & vbNullChar)) Then
        PathCombine = TrimNull(strBuffer)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathExists
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function PathExists(ByVal strPath As String) As Boolean
    PathExists = PathFileExists(StrPtr(strPath & vbNullChar))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileSizeByPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileSizeByPath(ByVal strPath As String) As Long

    Dim lHandle As Long

    GetFileSizeByPath = -1
    lHandle = CreateFile(StrPtr(strPath & vbNullChar), GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If lHandle <> INVALID_HANDLE_VALUE Then
        GetFileSizeByPath = GetFileSize(lHandle, 0&)
        CloseHandle lHandle
    End If

End Function
