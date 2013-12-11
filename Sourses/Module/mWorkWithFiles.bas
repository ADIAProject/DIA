Attribute VB_Name = "mWorkWithFiles"
Option Explicit

' Переменные для работы с файловой системой
Public objFSO                           As Scripting.FileSystemObject

Private Root                            As String
Private xFOL                            As Folder
Private xFile                           As File

' Переменная
Public strFileListInFolder              As String

'Удаление слэша на конце
Public Function BacklashDelFromPath(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathRemoveBackslash strPath
    BacklashDelFromPath = TrimNull(strPath)
End Function

'Добавление слэша на конце
Public Function BackslashAdd2Path(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathAddBackslash strPath
    BackslashAdd2Path = TrimNull(strPath)
End Function

Public Function CompareFilesByHashCAPICOM(ByVal strFirstFile As String, _
                                          ByVal strSecondFile As String) As Boolean

Dim strDataSHAFirst                     As String
Dim strDataSHASecond                    As String
Dim lngResult                           As Long

    If PathFileExists(strFirstFile) = 1 Then
        strDataSHAFirst = CalcHashFile(strFirstFile, CAPICOM_HASH_ALGORITHM_SHA1)

    End If

    If PathFileExists(strSecondFile) = 1 Then
        strDataSHASecond = CalcHashFile(strSecondFile, CAPICOM_HASH_ALGORITHM_SHA1)

    End If

    lngResult = StrComp(strDataSHAFirst, strDataSHASecond, vbTextCompare)

    If lngResult = 0 Then
        CompareFilesByHashCAPICOM = True
    Else
        CompareFilesByHashCAPICOM = False

    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  CopyFileTo
'!  Переменные  :  PathFrom As String, PathTo As String
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Скопирует файл 'PathFrom' в директорию 'CopyFileTo', Если файл существует, то он будет перезаписан новым файлом.
'! -----------------------------------------------------------
Public Function CopyFileTo(ByVal PathFrom As String, ByVal PathTo As String) As Boolean

Dim ret                                 As Long

    If PathFileExists(PathFrom) = 1 Then
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

Public Sub CreateNewDirectory(ByVal NewDirectory As String)

Dim SecAttrib                           As SECURITY_ATTRIBUTES
Dim sPath                               As String
Dim iCounter                            As Integer
Dim sTempDir                            As String
Dim ret                                 As Long
Dim retLasrErr                          As Long

    sPath = BackslashAdd2Path(NewDirectory)
    iCounter = 1

    Do Until InStr(iCounter, sPath, "\") = 0
        iCounter = InStr(iCounter, sPath, "\")
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

            If PathFileExists(sTempDir) = 0 Then
                DebugMode vbTab & "CreateDirectory: False : " & sTempDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)

            End If

        End If

    Loop

End Sub

Public Function DeleteFiles(ByVal PathFile As String) As Boolean

Dim ret                                 As Long
Dim retDllerr                           As Long

    'ret = DeleteFile(PathFile)
    If PathIsUNC(PathFile) = 0 Then
        ret = DeleteFileW(StrPtr("\\?\" & PathFile & vbNullChar))
    Else
        '\\?\UNC\
        ret = DeleteFileW(StrPtr("\\?\UNC\" & Right$(PathFile, Len(PathFile) - 2) & vbNullChar))
    End If

    DeleteFiles = CBool(ret)

    If ret = 0 Then
        If PathFileExists(PathFile) = 1 Then

            On Error GoTo errhandler

            objFSO.DeleteFile PathFile, True

        End If

        retDllerr = Err.LastDllError

        If PathFileExists(PathFile) = 1 Then
            DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & retDllerr & " - " & ApiErrorText(retDllerr)

        End If

    End If

    Exit Function
errhandler:
    retDllerr = Err.LastDllError
    DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.Number & ": " & Err.Description
    DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & retDllerr & " - " & ApiErrorText(retDllerr)
    Err.Clear

    Resume Next

End Function

'! -----------------------------------------------------------
'!  Функция     :  DelFolderBackUp
'!  Переменные  :
'!  Описание    :  Удаление временного каталога, если включена опция
'! -----------------------------------------------------------
Public Sub DelFolderBackUp(ByVal strFolderPath As String)

Dim ret                                 As Long

    On Error Resume Next

    DebugMode "DelFolder-Start: " & strFolderPath

    If PathFileExists(strFolderPath) = 1 Then
        DelRecursiveFolder strFolderPath

    End If

    If PathFileExists(strFolderPath) = 1 Then
        ret = RemoveDirectory(strFolderPath)

        If ret = 0 Then
            DebugMode vbTab & "RemoveDirectory: False : " & strFolderPath & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)

        End If

    End If

    On Error GoTo 0

    DebugMode "DelFolder-End"

End Sub

'! -----------------------------------------------------------
'!  Функция     :  DelRecursiveFolder
'!  Переменные  :  Folder As String
'!  Описание    :
'! -----------------------------------------------------------
Public Sub DelRecursiveFolder(ByVal Folder As String)

Dim retDelete                           As Long
Dim retStrMsg                           As String

    Root = BacklashDelFromPath(Folder)
    DebugMode vbTab & "DeleteFolder: " & Root

    If PathFileExists(Root) = 1 Then
        SearchFilesInRoot Root, ALL_FILES, True, False, True
        Set xFOL = objFSO.GetFolder(Root)

        If xFOL.Files.Count > 0 Then

            For Each xFile In xFOL.Files

                DeleteFiles xFile.Path
            Next

        End If

        ' Получение списка каталогов подлежащих удалению
        If PathFileExists(Root) = 1 Then
            GetAllFolderInRoot Root, True

        End If

        If PathFileExists(Root) = 1 Then
            GetAllFolderInRoot Root, True
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

'! -----------------------------------------------------------
'!  Функция     :  DelTemp
'!  Переменные  :
'!  Описание    :  Удаление временного каталога, если включена опция
'! -----------------------------------------------------------
Public Sub DelTemp()

    On Error Resume Next

    DebugMode "DelTemp-Start"

    If PathFileExists(strWorkTemp) = 1 Then
        DelRecursiveFolder strWorkTemp
    End If

    If PathFileExists(strWorkTemp) = 1 Then
        RemoveDirectory strWorkTemp
    End If

    On Error GoTo 0

    DebugMode "DelTemp-End"

End Sub

Private Function DelTree(ByVal strDir As String) As Long

Dim X                                   As Long
Dim intAttr                             As Integer
Dim strAllDirs                          As String
Dim strFile                             As String
Dim ret                                 As Long
Dim retLasrErr                          As Long

    DelTree = -1

    On Error Resume Next

    strDir = Trim$(strDir)

    If LenB(strDir) > 0 Then
        If Right$(strDir, 1) = vbBackslash Then
            strDir = Left$(strDir, Len(strDir) - 1)
        End If

        If InStr(strDir, "\") Then
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

                    If PathFileExists(strDir) = 0 Then
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

'! -----------------------------------------------------------
'!  Функция     :  ExtFromFileName
'!  Переменные  :  FileName As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить расширение файла из пути или имени файла
'! -----------------------------------------------------------
Public Function ExtFromFileName(ByVal FileName As String) As String

Dim intLastSeparator                    As Long

    intLastSeparator = InStrRev(FileName, ".")

    If intLastSeparator > 0 Then
        ExtFromFileName = Right$(FileName, Len(FileName) - intLastSeparator)
    Else
        ExtFromFileName = vbNullString

    End If

End Function

Public Function FileisReadOnly(ByVal PathFile As String) As Boolean
    FileisReadOnly = GetAttr(PathFile) And vbReadOnly

End Function

Public Function FileisSystemAttr(PathFile As String) As Boolean
    FileisSystemAttr = GetAttr(PathFile) And vbSystem

End Function

'! -----------------------------------------------------------
'!  Функция     :  FileName_woExt
'!  Переменные  :  FileName As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить имя файла без расширения, зная имя файла
'! -----------------------------------------------------------
Public Function FileName_woExt(ByVal FileName As String) As String

Dim intLastSeparator                    As Long

    FileName_woExt = FileName

    If LenB(FileName) > 0 Then
        intLastSeparator = InStrRev(FileName, ".")

        If intLastSeparator > 0 Then
            FileName_woExt = Left$(FileName, intLastSeparator - 1)

        End If

    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  FileNameFromPath
'!  Переменные  :  FilePath As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить имя файла из полного пути
'! -----------------------------------------------------------
Public Function FileNameFromPath(ByVal FilePath As String) As String

Dim intLastSeparator                    As Long

    FileNameFromPath = FilePath

    If LenB(FilePath) > 0 Then
        intLastSeparator = InStrRev(FilePath, "\")

        If intLastSeparator >= 0 Then
            FileNameFromPath = Right$(FilePath, Len(FilePath) - intLastSeparator)

        End If

    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  GetAllFileInFolder
'!  Переменные  :  xFolder As String, RealDelete As Boolean, Optional ExtFile As String
'!  Описание    :  Получение всех файлов в выбранном каталоге
'! -----------------------------------------------------------
Public Sub GetAllFileInFolder(ByVal xFolder As String, _
                              RealDelete As Boolean, _
                              Optional ExtFile As String, _
                              Optional ByVal mbRecursFolder As Boolean = True)

Dim strExtFile_x()                      As String
Dim strExtFile                          As String
Dim strExtFileReal                      As String
Dim strTemp                             As String
Dim strTempAll                          As String
Dim i                                   As Long

    DebugMode str2VbTab & "GetAllFileInFolder-Start: " & xFolder, 2

    If Not PathFileExists(xFolder) = 0 Then
        Set xFOL = objFSO.GetFolder(xFolder)
        strExtFile_x = Split(ExtFile, ";")

        For Each xFile In xFOL.Files

            ' Если требуется удаление файла, то удалаем
            If RealDelete Then

                On Error GoTo errhandler

                xFile.Delete True
            Else
                ' Если расширение файл INF, то добавляем путь файла в массив
                strTemp = vbNullString

                For i = LBound(strExtFile_x) To UBound(strExtFile_x)
                    strExtFile = UCase$(strExtFile_x(i))
                    strExtFileReal = UCase$(ExtFromFileName(xFile.Path))

                    If strExtFile = "INF" Then
                        If strExtFileReal = strExtFile Then
                            InfTempPathListCount = InfTempPathListCount + 1
                            InfTempPathList(InfTempPathListCount) = xFile.Path

                        End If

                    Else

                        If strExtFile = strExtFileReal Then
                            strTemp = xFile.Path

                        End If

                    End If

                Next

                If LenB(strTemp) > 0 Then
                    strTempAll = AppendStr(strTempAll, strTemp, ";")

                End If

            End If

        Next

        ' Если требуется удаление каталога, то удалаем
        If RealDelete Then

            With xFOL

                If .Files.Count = 0 Then
                    If .SubFolders.Count = 0 Then
                        .Delete True

                    End If

                End If

            End With

        Else

            If LenB(strTempAll) > 0 Then
                DebugMode str2VbTab & "ListFiles in Folder '" & xFOL.Name & "': " & vbNewLine & "*****************************************" & vbNewLine & strTempAll & vbNewLine & "*****************************************"
                strFileListInFolder = AppendStr(strFileListInFolder, strTempAll, ";")

            End If

        End If

        If mbRecursFolder Or RealDelete Then
            ' Проверяем есть ли подкаталоги в каталоге
            GetAllFolderInRoot xFolder, RealDelete, ExtFile

        End If

    End If

    DebugMode str2VbTab & "GetAllFileInFolder-End", 2
    Exit Sub
errhandler:
    DebugMode vbTab & "GetAllFileInFolder: False : " & xFolder & " Error: №" & Err.Number & ": " & Err.Description
    Err.Clear

    Resume Next

End Sub

'! -----------------------------------------------------------
'!  Функция     :  GetAllFolderInFolder
'!  Переменные  :  rootFolder As String
'!  Описание    :  Получение всех подкаталогов в выбранном каталоге
'! -----------------------------------------------------------
Public Function GetAllFolderInFolder(ByVal RootFolder As String) As Variant

Dim xFolder                             As Folder
Dim strListFolder                       As String

    DebugMode str2VbTab & "GetAllFolderInFolder-Start: "

    If PathFileExists(RootFolder) = 1 Then
        Set xFOL = objFSO.GetFolder(RootFolder)

        If xFOL.SubFolders.Count > 0 Then

            For Each xFolder In xFOL.SubFolders

                strListFolder = AppendStr(strListFolder, xFolder.Name, ";")
            Next

        End If

        GetAllFolderInFolder = Split(strListFolder, ";")

    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  GetAllFolderInRoot
'!  Переменные  :  rootFolder As String, RealDelete As Boolean, Optional ExtFile As String
'!  Описание    :  Получение всех подкаталогов в выбранном каталоге
'! -----------------------------------------------------------
Private Sub GetAllFolderInRoot(ByVal RootFolder As String, _
                               ByVal RealDelete As Boolean, _
                               Optional ExtFile As String)

Dim xFolder                             As Folder

    If PathFileExists(RootFolder) = 1 Then
        Set xFOL = objFSO.GetFolder(RootFolder)

        If xFOL.SubFolders.Count > 0 Then

            For Each xFolder In xFOL.SubFolders

                DebugMode str2VbTab & "Analize Subfolder: " & xFolder.Path, 2
                GetAllFileInFolder xFolder.Path, RealDelete, ExtFile
            Next

        End If

    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  GetEnviron
'!  Переменные  :  strEnv As String, Optional mbCollectFull As Boolean = False
'!  Возвр. знач.:  As String
'!  Описание    :  Получение переменной системного окружения
'! -----------------------------------------------------------
Public Function GetEnviron(ByVal strEnv As String, _
                           Optional ByVal mbCollectFull As Boolean = False) As String

Dim strTemp                             As String
Dim strTempEnv                          As String
Dim strNumPosition                      As Long

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

    DebugMode str2VbTab & "GetEnviron: %" & strTemp & "%=" & strTempEnv
    DebugMode str2VbTab & "GetEnviron-End"

End Function

Public Function GetUniqueTempFile() As String

Dim ll_Buffer                           As Long
Dim ls_TempFileName                     As String

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

'! -----------------------------------------------------------
'!  Функция     :  IsDriveCDRoom
'!  Переменные  :
'!  Описание    :  Проверка на запск программы с CD\DVD
'! -----------------------------------------------------------
Public Function IsDriveCDRoom() As Boolean

Dim strDriveName                        As String
Dim xDrv                                As Drive

    IsDriveCDRoom = False
    strDriveName = Left$(strAppPath, 2)

    ' Проверяем на запуск из сети
    If InStr(strDriveName, "\\") = 0 Then
        'получаем тип диска
        Set xDrv = objFSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            IsDriveCDRoom = True

        End If

    End If

End Function

Public Function IsPathAFolder(ByVal sPath As String) As Boolean

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
Dim Result                              As Long

    Result = PathIsDirectory(StrPtr(sPath & vbNullChar))
    IsPathAFolder = (Result = vbDirectory) Or (Result = 1)

End Function

'! -----------------------------------------------------------
'!  Функция     :  MoveFileTo
'!  Переменные  :  PathFrom As String, PathTo As String
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Скопирует файл 'PathFrom' в директорию 'PathTo', Если файл существует, то он будет перезаписан новым файлом.
'! -----------------------------------------------------------
Public Function MoveFileTo(PathFrom As String, PathTo As String) As Boolean

Dim ret                                 As Long

    If StrComp(PathFrom, PathTo, vbTextCompare) <> 0 Then
        If PathFileExists(PathFrom) = 1 Then
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

Public Function ParserInf4Strings(ByVal strInfFilePath As String, _
                                  ByVal strSearchString As String) As String

Dim StringHash                          As Scripting.Dictionary
Dim objInfFile                          As TextStream
Dim RegExpStrSect                       As RegExp
Dim RegExpStrDefs                       As RegExp
Dim MatchesStrSect                      As MatchCollection
Dim MatchesStrDefs                      As MatchCollection
Dim objMatch                            As Match
Dim objMatch1                           As Match
Dim regex_strsect                       As String
Dim regex_strings                       As String
Dim r_beg                               As String
Dim r_identS                            As String
Dim r_str                               As String
Dim FileContent                         As String
Dim Key                                 As String
Dim Value                               As String
Dim R                                   As Boolean
Dim i                                   As Long
Dim Strings                             As String
Dim valval                              As String
Dim varname                             As String
Dim lngFileDBSize                       As Long
Dim Pos                                 As Long

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

'! -----------------------------------------------------------
'!  Функция     :  PathNameFromPath
'!  Переменные  :  FilePath As String
'!  Возвр. знач.:  As String
'!  Описание    :  Получить путь к файлу из полного пути
'! -----------------------------------------------------------
Public Function PathNameFromPath(FilePath As String) As String

Dim intLastSeparator                    As Long
    intLastSeparator = InStrRev(FilePath, "\")
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

Public Sub ResetReadOnly4File(ByVal StrPathFile As String)

    If PathFileExists(StrPathFile) = 1 Then
        If FileisReadOnly(StrPathFile) Then
            SetAttr StrPathFile, vbNormal

        End If

        If FileisSystemAttr(StrPathFile) Then
            SetAttr StrPathFile, vbNormal

        End If

    End If

End Sub

'# function to replace special chars to create dirs correctly #
Public Function SafeDir(ByVal str As String) As String

Dim R                                   As String

    R = str
    R = Replace$(R, "\", "_")
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

'# function to replace special chars to create dirs correctly #
Public Function SafeFileName(ByVal strString) As String
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

'# function to discover dirs with inf code #
Public Function WhereIsDir(ByVal str As String, ByVal strInfFilePath As String) As String

Dim cDir                                As String
Dim Str_x()                             As String
Dim mbAdditionalPath                    As Boolean

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

'! -----------------------------------------------------------
'!  Функция     :  PathCollect
'!  Переменные  :  Path As String
'!  Возвр. знач.:  As String
'!  Описание    :
'! -----------------------------------------------------------
Public Function PathCollect(Path As String) As String

    If InStr(Path, ":") = 2 Then
        PathCollect = Path
    ElseIf Left$(Path, 2) = "\\" And IsUNCPathValid(Path) Then
        PathCollect = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect = PathCombine(strAppPath, Path)
        Else

            If InStr(Path, "\") = 1 Then
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

    If InStr(PathCollect, "\\") Then
        If Left$(strAppPath, 2) <> "\\" Then
            PathCollect = Replace$(PathCollect, "\\", "\")
        End If
    End If

    If IsPathAFolder(PathCollect) Then
        PathCollect = BackslashAdd2Path(PathCollect)
    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  PathCollect4Dest
'!  Переменные  :  Path As String
'!  Возвр. знач.:  As String
'!  Описание    :
'! -----------------------------------------------------------
Public Function PathCollect4Dest(ByVal Path As String, ByVal strDest As String) As String

    If InStr(Path, ":") = 2 Then
        PathCollect4Dest = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect4Dest = strDest & Mid$(Path, 2, Len(Path) - 1)
        Else

            If InStr(Path, "\") = 1 Then
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

    If InStr(PathCollect4Dest, "\\") Then
        PathCollect4Dest = Replace$(PathCollect4Dest, "\\", "\")

        If Left$(strDest, 2) = "\\" Then
            If InStr(PathCollect4Dest, "\") = 1 Then
                PathCollect4Dest = vbBackslash & PathCollect4Dest
            Else
                PathCollect4Dest = "\\" & PathCollect4Dest

            End If

        End If

    End If

    PathCollect4Dest = BackslashAdd2Path(PathCollect4Dest)

End Function

Public Function IsUNCPathValid(ByVal sPath As String) As Boolean
'Determines if the string is a valid UNC
'(universal naming convention) for a server
'and share path. Returns True (1) if the string
'is a valid UNC path, or False otherwise.
    IsUNCPathValid = PathIsUNC(sPath) = 1

End Function

' Расширить имя файла - использование переменных %%
Public Function ExpandFileNamebyEnvironment(ByVal strFileName As String) As String

Dim R                                   As String
Dim str_OSVer                           As String
Dim str_OSBit                           As String
Dim str_DATE                            As String

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

Public Function CreateIfNotExistPath(strFolderPath As String) As Boolean
    If LenB(strFolderPath) > 0 Then
        ' Если нет, то создаем каталог
        If PathFileExists(strFolderPath) = 0 Then
            CreateNewDirectory strFolderPath
            CreateIfNotExistPath = IsPathAFolder(strFolderPath)
        End If
    End If
End Function

Public Function CopyFolderByShell(sSource As String, sDestination As String) As Long

Dim FOF_FLAGS                           As Long
Dim SHFileOp                            As SHFILEOPSTRUCT

    'terminate the folder string with a pair of nulls
    sSource = BacklashDelFromPath(sSource) & str2vbNullChar
    If PathFileExists(sDestination) = 0 Then
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

Public Function PathCombine(ByVal strDirectory As String, ByVal strFile As String) As String
Dim strBuffer                           As String
    ' Concatenates two strings that represent properly formed
    ' paths into one path, as well as any relative path pieces.
    strBuffer = String$(MAX_PATH_UNICODE, vbNullChar)
    If PathCombineW(StrPtr(strBuffer), StrPtr(strDirectory & vbNullChar), StrPtr(strFile & vbNullChar)) Then
        PathCombine = TrimNull(strBuffer)
    End If
End Function

Public Function PathFileExists(ByVal strPath As String) As Long
    PathFileExists = PathFileExistsW(StrPtr(strPath & vbNullChar))
End Function

Public Function GetFileSizeByPath(ByVal strPath As String) As Long
Dim lHandle                             As Long
    GetFileSizeByPath = -1
    lHandle = CreateFile(StrPtr(strPath & vbNullChar), GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)
    If lHandle <> INVALID_HANDLE_VALUE Then
        GetFileSizeByPath = GetFileSize(lHandle, 0&)
        CloseHandle lHandle
    End If
End Function

