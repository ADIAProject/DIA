Attribute VB_Name = "mWorkWithFiles"
Option Explicit

' Not add to project (if not DBS) - option for compile
#Const mbIDE_DBSProject = False
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompareFilesByHashCAPICOM
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strFirstFile (String)
'                              strSecondFile (String)
'!--------------------------------------------------------------------------------
#If mbIDE_DBSProject Then
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
    
        CompareFilesByHashCAPICOM = lngResult = 0

    End Function
#End If

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
'! Procedure   (Функция)   :   Function BackslashDelFromPath
'! Description (Описание)  :   [Удаление слэша на конце]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function BackslashDelFromPath(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathRemoveBackslash strPath
    BackslashDelFromPath = TrimNull(strPath)
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
            If mbDebugStandart Then DebugMode vbTab & "Copy file: False: " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        CopyFileTo = False
        If mbDebugStandart Then DebugMode vbTab & "Copy file: False : " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
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
    sSource = BackslashDelFromPath(sSource) & str2vbNullChar

    If PathExists(sDestination) = False Then
        CreateIfNotExistPath sDestination
    End If

    sDestination = BackslashDelFromPath(sDestination) & str2vbNullChar
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
'! Procedure   (Функция)   :   Function CreateIfNotExistPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Function CreateIfNotExistPath(ByVal strFolderPath As String) As Boolean

    If LenB(strFolderPath) Then

        ' Если нет, то создаем каталог
        If PathExists(strFolderPath) = False Then
            CreateNewDirectory strFolderPath
            CreateIfNotExistPath = PathIsAFolder(strFolderPath)
        End If
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

        If PathExists(sTempDir) = False Then
            ret = CreateDirectory(sTempDir, SecAttrib)

            If ret = 0 Then
                If mbDebugStandart Then DebugMode str2VbTab & "CreateNewDirectory: False : " & sTempDir & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
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

    Dim ret             As Long
    Dim lngFilePathPtr  As Long
    
    If PathIsValidUNC(PathFile) = False Then
        lngFilePathPtr = StrPtr("\\?\" & PathFile)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(PathFile, Len(PathFile) - 2))
    End If
    ret = DeleteFile(lngFilePathPtr)

    If ret = 0 Then
        ' Если нет доступа, то возможно атрибут только для чтения, пытаемся снять и снова удалить файл
        If Err.LastDllError = 5 Then
            ResetReadOnly4File PathFile
            ret = DeleteFile(lngFilePathPtr)
            If ret = 0 Then
                If mbDebugStandart Then DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If
        Else
            If mbDebugStandart Then DebugMode vbTab & "DeleteFiles: False : " & PathFile & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
        
    End If

    DeleteFiles = CBool(ret)
        
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DeleteFolder
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Function DeleteFolder(ByVal strFolderPath As String) As Boolean

    Dim ret As Long
    Dim lngFilePathPtr As Long
    
    If PathExists(strFolderPath) Then
        If PathIsValidUNC(strFolderPath) = False Then
            lngFilePathPtr = StrPtr("\\?\" & strFolderPath)
        Else
            '\\?\UNC\
            lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strFolderPath, Len(strFolderPath) - 2))
        End If
        ret = RemoveDirectory(lngFilePathPtr)
    End If

    If ret = 0 Then
        ' Папка не пуста
        If Err.LastDllError = 145 Then
            If mbDebugDetail Then DebugMode vbTab & "DeleteFiles: False : " & strFolderPath & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        Else
            If mbDebugStandart Then DebugMode vbTab & "DeleteFolder: False : " & strFolderPath & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    DeleteFolder = CBool(ret)
    
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelFolderBackUp
'! Description (Описание)  :   [Удаление временного каталога, если включена опция]
'! Parameters  (Переменные):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Sub DelFolderBackUp(ByVal strFolderPath As String)

    Dim ret As Long
    Dim lngFilePathPtr As Long

    On Error Resume Next

    If mbDebugStandart Then DebugMode "DelFolderBackUp-Start: " & strFolderPath

    If PathExists(strFolderPath) Then
        DelRecursiveFolder strFolderPath
    End If

    If PathExists(strFolderPath) Then
                
        If PathIsValidUNC(strFolderPath) = False Then
            lngFilePathPtr = StrPtr("\\?\" & strFolderPath)
        Else
            '\\?\UNC\
            lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strFolderPath, Len(strFolderPath) - 2))
        End If
        ret = RemoveDirectory(lngFilePathPtr)
        
        If ret = 0 Then
            If mbDebugStandart Then DebugMode vbTab & "DelFolderBackUp: False : " & strFolderPath & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    On Error GoTo 0

    If mbDebugStandart Then DebugMode "DelFolderBackUp-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelRecursiveFolder
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Folder (String)
'!--------------------------------------------------------------------------------
Public Sub DelRecursiveFolder(ByVal Folder As String)

    Dim retDelete   As Long
    Dim retStrMsg   As String
    Dim Root        As String

    Root = BackslashDelFromPath(Folder)
    If mbDebugStandart Then DebugMode vbTab & "DeleteFolder: " & Root

    If PathExists(Root) Then
        SearchFilesInRoot Root, ALL_FILES, True, False, True
        SearchFoldersInRoot Root, ALL_FILES, True, True

        ' Удаление пустых каталогов, если остались
        If PathExists(Root) Then
            retDelete = DelTree(Root)

            If mbDebugStandart Then

                If retDelete = 0 Then
                    retStrMsg = "Deleted"
                ElseIf retDelete = -1 Then
                    retStrMsg = "Invalid Directory"
                Else
                    retStrMsg = "An Error was occured"
                End If

                If mbDebugStandart Then DebugMode vbTab & "DelRecursiveFolder: " & " Result: " & retStrMsg
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

    Dim TimeScriptRun       As Long
    Dim TimeScriptFinish    As Long
    Dim ret                 As Long
    Dim lngFilePathPtr      As Long

    On Error Resume Next

    If mbDebugDetail Then DebugMode "DelTemp-Start"
    TimeScriptRun = GetTickCount

    If PathExists(strWorkTemp) Then
        DelRecursiveFolder strWorkTemp
    End If

    If PathExists(strWorkTemp) Then
        
        If PathIsValidUNC(strWorkTemp) = False Then
            lngFilePathPtr = StrPtr("\\?\" & strWorkTemp)
        Else
            '\\?\UNC\
            lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strWorkTemp, Len(strWorkTemp) - 2))
        End If
        ret = RemoveDirectory(lngFilePathPtr)
        
        If ret = 0 Then
            If mbDebugStandart Then DebugMode vbTab & "DelTemp: False : " & strWorkTemp & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
        
    End If

    TimeScriptFinish = GetTickCount
    If mbDebugStandart Then DebugMode "DelTemp-End: Time to Delete: " & CalculateTime(TimeScriptRun, TimeScriptFinish, True)
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

    If LenB(strDir) Then
        If Right$(strDir, 1) = vbBackslash Then
            strDir = Left$(strDir, Len(strDir) - 1)
        End If

        If InStr(strDir, vbBackslash) Then
            intAttr = GetAttr(strDir)

            If (intAttr And vbDirectory) Then
                strDir = BackslashAdd2Path(strDir)
                strFile = Dir$(strDir & ALL_FILES, vbSystem Or vbDirectory Or vbHidden)

                Do While Len(strFile)

                    If strFile <> strDot Then
                        If strFile <> str2Dot Then
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

                If PathIsValidUNC(strDir) = False Then
                    ret = RemoveDirectory(StrPtr("\\?\" & strDir))
                Else
                    '\\?\UNC\
                    ret = RemoveDirectory(StrPtr("\\?\UNC\" & Right$(strDir, Len(strDir) - 2)))
                End If

                If ret = 0 Then
                    retLasrErr = Err.LastDllError
                    If mbDebugStandart Then DebugMode vbTab & "DelTree: False : " & strDir & " Error: №" & retLasrErr & " - " & ApiErrorText(retLasrErr)
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
'! Procedure   (Функция)   :   Function ExpandFileNamebyEnvironment
'! Description (Описание)  :   [Расширить имя файла - использование переменных %%]
'! Parameters  (Переменные):   strFileName (String)
'!--------------------------------------------------------------------------------
Public Function ExpandFileNamebyEnvironment(ByVal strFileName As String) As String

    Dim R            As String
    Dim str_OSVer    As String
    Dim str_OSBit    As String
    Dim str_DATE     As String
    Dim str_PCMODEL  As String

    If InStr(strFileName, strPercentage) Then
        ' Макроподстановка версия ОС %OSVer%
        str_OSVer = "wnt" & Left$(strOSCurrentVersion, 1)

        ' Макроподстановка битность ОС %OSBit%
        If mbIsWin64 Then
            str_OSBit = "x64"
        Else
            str_OSBit = "x32"
        End If

        ' Макроподстановка ДАТА %DATE%
        str_DATE = SafeDir(Replace$(CStr(Now()), strDot, "-"))
        ' Макроподстановка %PCMODEL%
        str_PCMODEL = SafeDir(Replace$(strCompModel, "_", "-"))
        
        ' Замена макросов значениями
        R = strFileName
        R = Replace$(R, "%PCNAME%", strCompModel, , , vbTextCompare)
        R = Replace$(R, "%PCMODEL%", str_PCMODEL, , , vbTextCompare)
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
Public Function FileIsSystemAttr(ByVal PathFile As String) As Boolean
    FileIsSystemAttr = GetAttr(PathFile) And vbSystem
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileReadData
'! Description (Описание)  :   [Read data from file with check for unicode yes/no]
'! Parameters  (Переменные):   sFileName (String)
'                              LocaleID (Long)
'!--------------------------------------------------------------------------------
Public Function FileReadData(ByVal sFileName As String, Optional ByVal LocaleID As Long = 1033) As String

    Dim sText As String
    Dim fNum As Long
    Dim B1(0 To 1) As Byte
    
    fNum = FreeFile

    Open sFileName For Binary Access Read Lock Write As fNum
    ' read first 2 byte, for check on Unicode
    Get #fNum, 1, B1()
    
    ' Если Unicode &HFF and &HFE 255-254
    If B1(0) = &HFF And B1(1) = &HFE Then
        'sText = Space$(LOF(fNum) - 2)
        sText = MemAPIs.AllocStr(vbNullString, LOF(fNum) - 2)
        Seek #fNum, 3
        Get #fNum, , sText
        FileReadData = StrConv(sText, vbFromUnicode, LocaleID)
    ' Если ANSI
    Else
        'sText = Space$(LOF(fNum))
        sText = MemAPIs.AllocStr(vbNullString, LOF(fNum))
        Seek #fNum, 1
        Get #fNum, , sText
        FileReadData = sText
    End If
    
    Close #fNum

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileWriteData
'! Description (Описание)  :   [Write data to file with check]
'! Parameters  (Переменные):   sFileName (String)
'                              sStringOut (String)
'!--------------------------------------------------------------------------------
Public Sub FileWriteData(ByVal sFileName As String, Optional ByVal sStringOut As String)

    Dim fNum As Integer
    
    fNum = FreeFile

    Open sFileName For Binary Access Write Lock Write As fNum
    Put #fNum, , sStringOut
    Close #fNum

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileWriteDataAPI
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFilePath (String)
'                              strData (String)
'!--------------------------------------------------------------------------------
Private Sub FileWriteDataAPI(ByVal sFilePath As String, ByVal strData As String)
    Dim fHandle         As Long
    Dim fSuccess        As Long
    Dim lBytesWritten   As Long
    Dim anArray()       As Byte
    Dim lngFilePathPtr  As Long
    
    ' Convert to byte
    anArray = StrConv(strData, vbFromUnicode)
    
    'Get a pointer to a string with file name.
    If PathIsValidUNC(sFilePath) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sFilePath)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sFilePath, Len(sFilePath) - 2))
    End If
    'Get a handle to a file Fname.
    fHandle = CreateFile(lngFilePathPtr, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0, CREATE_ALWAYS, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    
    'CreateFile returns INVALID_HANDLE_VALUE if it fails.
    If fHandle <> INVALID_HANDLE_VALUE Then
        fSuccess = WriteFile(fHandle, VarPtr(anArray(0)), UBound(anArray) + 1, lBytesWritten, 0)
        'Check to see if you were successful writing the data
        If fSuccess <> 0 Then
            'Flush the file buffers to force writing of the data.
            FlushFileBuffers fHandle
            'Close the file.
            CloseHandle fHandle
        Else
            If mbDebugStandart Then DebugMode str2VbTab & "FileWriteDataAPI: WriteFile - ReturnCode: " & ApiErrorText(Err.LastDllError)
        End If
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "FileWriteDataAPI: CreateFile - ReturnCode: " & ApiErrorText(Err.LastDllError)
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileWriteDataAPIUni
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sFilePath (String)
'                              strData (String)
'!--------------------------------------------------------------------------------
Private Sub FileWriteDataAPIUni(ByVal sFilePath As String, ByVal strData As String)
    Dim fHandle As Long
    Dim fSuccess As Long
    Dim lBytesWritten As Long
    Dim BytesToWrite As Long
    Dim anArray() As Byte
    Dim lngFilePathPtr As Long
    Dim lngStringSize As Long
    
    lngStringSize = LenB(strData)
    ReDim anArray(0 To lngStringSize)
    CopyMemory anArray(0), ByVal StrPtr(strData), lngStringSize
    'Get the length of data to write
    BytesToWrite = (UBound(anArray) + 1) * LenB(anArray(0))
    
    'Get a pointer to a string with file name.
    If PathIsValidUNC(sFilePath) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sFilePath)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sFilePath, Len(sFilePath) - 2))
    End If
    'Get a handle to a file Fname.
    fHandle = CreateFile(lngFilePathPtr, GENERIC_WRITE Or GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE Or FILE_SHARE_DELETE, 0, CREATE_ALWAYS, FILE_FLAG_SEQUENTIAL_SCAN, 0)
    
    'CreateFile returns INVALID_HANDLE_VALUE if it fails.
    If fHandle <> INVALID_HANDLE_VALUE Then
        fSuccess = WriteFile(fHandle, VarPtr(anArray(0)), BytesToWrite, lBytesWritten, 0)
        'Check to see if you were successful writing the data
        If fSuccess <> 0 Then
            'Flush the file buffers to force writing of the data.
            FlushFileBuffers fHandle
            'Close the file.
            CloseHandle fHandle
        Else
            If mbDebugStandart Then DebugMode str2VbTab & "FileWriteDataAPIUni: WriteFile - ReturnCode: " & ApiErrorText(Err.LastDllError)
        End If
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "FileWriteDataAPIUni: CreateFile - ReturnCode: " & ApiErrorText(Err.LastDllError)
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileWriteDataAppend
'! Description (Описание)  :   [Read data from file with check for unicode yes/no]
'! Parameters  (Переменные):   sFileName (String)
'                              sStringOut (String)
'!--------------------------------------------------------------------------------
Private Sub FileWriteDataAppend(ByVal sFileName As String, Optional ByVal sStringOut As String)

    Dim fNum As Integer
    
    fNum = FreeFile

    Open sFileName For Binary Access Write Lock Write As fNum
    Put #fNum, LOF(fNum), sStringOut
    Close #fNum

End Sub

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

    strNumPosition = InStr(strEnv, strPercentage)

    If strNumPosition Then
        strTemp = Mid$(strEnv, strNumPosition + 1, Len(strEnv) - strNumPosition)
        strNumPosition = InStr(strTemp, strPercentage)

        If strNumPosition Then
            strTemp = Left$(strTemp, strNumPosition - 1)
        End If
    End If

    strTempEnv = Environ$(strTemp)

    If mbCollectFull Then
        GetEnviron = Replace$(strEnv, strPercentage & strTemp & strPercentage, strTempEnv, , , vbTextCompare)
    Else
        GetEnviron = strTempEnv
    End If

    If mbDebugStandart Then DebugMode str2VbTab & "GetEnviron: %" & strTemp & "%=" & strTempEnv & vbNewLine & _
              str2VbTab & "GetEnviron-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileName_woExt
'! Description (Описание)  :   [Получить имя файла без расширения, зная имя файла]
'! Parameters  (Переменные):   FileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileName_woExt(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    GetFileName_woExt = FileName

    If LenB(FileName) Then
        intLastSeparator = InStrRev(FileName, strDot)

        If intLastSeparator Then
            GetFileName_woExt = Left$(FileName, intLastSeparator - 1)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileNameExtension
'! Description (Описание)  :   [Получить расширение файла из пути или имени файла]
'! Parameters  (Переменные):   FileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileNameExtension(ByVal FileName As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FileName, strDot)

    If intLastSeparator Then
        GetFileNameExtension = Right$(FileName, Len(FileName) - intLastSeparator)
    Else
        GetFileNameExtension = vbNullString
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileNameFromPath
'! Description (Описание)  :   [Получить имя файла из полного пути]
'! Parameters  (Переменные):   FilePath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileNameFromPath(ByVal FilePath As String) As String

    Dim intLastSeparator As Long

    GetFileNameFromPath = FilePath

    If LenB(FilePath) Then
        intLastSeparator = InStrRev(FilePath, vbBackslash)

        If intLastSeparator >= 0 Then
            GetFileNameFromPath = Right$(FilePath, Len(FilePath) - intLastSeparator)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileSizeByPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileSizeByPath(ByVal strPath As String) As Long

    Dim lHandle As Long
    Dim lngFilePathPtr  As Long
    
    If PathIsValidUNC(strPath) = False Then
        lngFilePathPtr = StrPtr("\\?\" & strPath)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strPath, Len(strPath) - 2))
    End If
    
    lHandle = CreateFile(lngFilePathPtr, GENERIC_READ, 0, 0, OPEN_EXISTING, FILE_ATTRIBUTE_NORMAL, 0)

    If lHandle <> INVALID_HANDLE_VALUE Then
        GetFileSizeByPath = GetFileSize(lHandle, 0&)
        CloseHandle lHandle
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetFileVersionOnly
'! Description (Описание)  :   [Return file version information string.]
'! Parameters  (Переменные):   sFileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileVersionOnly(ByVal sFileName As String) As String
    Dim nUnused As Long
    Dim sBuffer() As Byte
    Dim nBufferSize As Long
    Dim lpBuffer As Long
    Dim FFI As VS_FIXEDFILEINFO
    Dim nVerSize As Long
    Dim sResult As String

    ' Get the version information buffer size.
    nBufferSize = GetFileVersionInfoSize(sFileName, nUnused)
    If nBufferSize Then
        ' Load the fixed file information into a buffer.
        ReDim sBuffer(0 To nBufferSize)
        If GetFileVersionInfo(sFileName, 0&, nBufferSize, sBuffer(0)) Then
            'VerQueryValue function returns selected version info
            'from the specified version-information resource. Grab
            'the file info and copy it into the  VS_FIXEDFILEINFO structure.
            If VerQueryValue(sBuffer(0), "\", lpBuffer, nVerSize) Then
            
                ' Copy the information from the buffer into a usable structure.
                CopyMemory FFI, ByVal lpBuffer, Len(FFI)
            
                ' Get the version information.
                With FFI
                    ' File version number.
                    sResult = Format$(.dwFileVersionMSh) & strDot & Format$(.dwFileVersionMSl) & strDot & Format$(.dwFileVersionLSh) & strDot & Format$(.dwFileVersionLSl)
                    'sResult = Format$(.dwFileVersionMSh) & strDot & Format$(.dwFileVersionMSl) & strDot & Format$(.dwFileVersionLSl)
                End With
            
                GetFileVersionOnly = sResult
            End If
            ' Else MsgBox "Error getting fixed file version information"
        End If
        'Else MsgBox "Error getting version information"
    End If
    'Else MsgBox "No version information available"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetPathNameFromPath
'! Description (Описание)  :   [Получить путь к файлу из полного пути]
'! Parameters  (Переменные):   FilePath (String)
'!--------------------------------------------------------------------------------
Public Function GetPathNameFromPath(ByVal FilePath As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(FilePath, vbBackslash)

    If intLastSeparator Then
        If intLastSeparator < Len(FilePath) Then
            GetPathNameFromPath = Left$(FilePath, intLastSeparator)
        Else
            GetPathNameFromPath = FilePath
        End If

    Else
        GetPathNameFromPath = FilePath
    End If

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
    Dim xDrv         As Long

    strDriveName = Left$(strAppPath, 3)

    ' Проверяем на запуск из сети
    If InStr(strDriveName, vbBackslashDouble) = 0 Then
        'получаем тип диска
        If PathIsRoot(strDriveName) Then
            xDrv = GetDriveType(strDriveName)
    
            If xDrv = DRIVE_CDROM Then
                IsDriveCDRoom = True
            End If
        End If
    End If

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
                If mbDebugStandart Then DebugMode vbTab & "Move file: False: " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If

        Else
            MoveFileTo = False
            If mbDebugStandart Then DebugMode vbTab & "Move file: False : " & PathFrom & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        If mbDebugStandart Then DebugMode vbTab & "Move file: Source and Destination are identicaly (" & PathFrom & " ; " & PathTo & ")"
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

    If lngFileDBSize Then
        FileContent = FileReadData(strInfFilePath)
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfFilePath
    End If

    ' Find [strings] section
    Strings = vbNullString
    StringHash.CompareMode = TextCompare
    Set MatchesStrSect = RegExpStrSect.Execute(FileContent)

    If MatchesStrSect.Count Then
        Set objMatch = MatchesStrSect.item(0)
        Strings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
        Set MatchesStrDefs = RegExpStrDefs.Execute(Strings)

        For i = 0 To MatchesStrDefs.Count - 1
            Set objMatch1 = MatchesStrDefs.item(i)
            Key = objMatch1.SubMatches(0)
            Value = objMatch1.SubMatches(1)

            If LenB(Value) = 0 Then
                Value = objMatch1.SubMatches(2)
            End If

            If Not StringHash.Exists(Key) Then
                StringHash.Add Key, Value
                StringHash.Add strPercentage & Key & strPercentage, Value
            End If

        Next

    End If

    ' Собственно ищем саму переменную
    Pos = InStr(strSearchString, strPercentage)

    If Pos Then
        varname = Mid$(strSearchString, Pos, InStrRev(strSearchString, strPercentage))
        valval = StringHash.item(varname)

        If LenB(valval) = 0 Then
            If mbDebugDetail Then DebugMode "ParserInf4Strings: Error in inf: Cannot find '" & strSearchString & "'"
        Else
            ParserInf4Strings = Replace$(strSearchString, varname, valval)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function PathCollect
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Path (String)
'!--------------------------------------------------------------------------------
Public Function PathCollect(Path As String) As String

    If InStr(Path, strDvoetochie) = 2 Then
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

                    If InStr(Path, strPercentage) Then
                        PathCollect = GetEnviron(Path, True)
                    Else

                        If LenB(GetFileNameExtension(Path)) Then
                            If GetFileNameFromPath(Path) = Path Then
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

    If InStr(Path, strDvoetochie) = 2 Then
        PathCollect4Dest = Path
    Else

        If Left$(Path, 2) = ".\" Then
            PathCollect4Dest = strDest & Mid$(Path, 2, Len(Path) - 1)
        Else

            If InStr(Path, vbBackslash) = 1 Then
                PathCollect4Dest = strDest & Path
            Else

                If Left$(Path, 3) = "..\" Then
                    PathCollect4Dest = GetPathNameFromPath(strDest) & Mid$(Path, 4, Len(Path) - 1)
                Else

                    If InStr(Path, strPercentage) Then
                        PathCollect4Dest = GetEnviron(Path, True)
                    Else

                        If LenB(GetFileNameExtension(Path)) Then
                            If GetFileNameFromPath(Path) = Path Then
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
'! Procedure   (Функция)   :   Function PathCombine
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strDirectory (String)
'                              strFile (String)
'!--------------------------------------------------------------------------------
Public Function PathCombine(ByVal strDirectory As String, ByVal strFile As String) As String

    Dim strBuffer As String

    ' Concatenates two strings that represent properly formed
    ' paths into one path, as well as any relative path pieces.
    strBuffer = FillNullChar(MAX_PATH_UNICODE)

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
'! Procedure   (Функция)   :   Function PathIsValidUNC
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sPath (String)
'!--------------------------------------------------------------------------------
Public Function PathIsValidUNC(ByVal sPath As String) As Boolean
    ' Returns True if the string is a valid UNC path.
    PathIsValidUNC = PathIsUNC(StrPtr(sPath))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ResetReadOnly4File
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Public Sub ResetReadOnly4File(ByVal strPathFile As String)

    If PathExists(strPathFile) Then
        If (GetAttr(strPathFile) And vbReadOnly) Then
            SetAttr strPathFile, vbNormal
        End If

        If (GetAttr(strPathFile) And vbSystem) Then
            SetAttr strPathFile, vbNormal
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SafeDir
'! Description (Описание)  :   [function to replace special chars to create dirs correctly]
'! Parameters  (Переменные):   str (String)
'!--------------------------------------------------------------------------------
Public Function SafeDir(ByVal str As String) As String

    If InStr(str, vbBackslash) Then
        str = Replace$(str, vbBackslash, "-")
    End If
    
    If InStr(str, "/") Then
        str = Replace$(str, "/", "-")
    End If
    
    If InStr(str, "*") Then
        str = Replace$(str, "*", "-")
    End If
    
    If InStr(str, strDvoetochie) Then
        str = Replace$(str, strDvoetochie, "-")
    End If
    
    If InStr(str, strVopros) Then
        str = Replace$(str, strVopros, "-")
    End If
    
    If InStr(str, ">") Then
       str = Replace$(str, ">", "-")
    End If
    
    If InStr(str, "<") Then
        str = Replace$(str, "<", "-")
    End If
    
    If InStr(str, "|") Then
        str = Replace$(str, "|", "-")
    End If
    
    If InStr(str, "@") Then
        str = Replace$(str, "@", "-")
    End If
    
    If InStr(str, "'") Then
        str = Replace$(str, "'", vbNullString)
    End If
        
    If InStr(str, strSpace) Then
        str = Replace$(str, strSpace, "-")
    End If
    
    If InStr(str, "(R)") Then
        str = Replace$(str, "(R)", "-")
    End If
        
    If InStr(str, "---") Then
        str = Replace$(str, "---", "-")
    End If
    
    If InStr(str, "--") Then
        str = Replace$(str, "--", "-")
    End If
        
    SafeDir = Trim$(str)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SafeFileName
'! Description (Описание)  :   [function to replace special chars to create files correctly]
'! Parameters  (Переменные):   strString (Variant)
'!--------------------------------------------------------------------------------
Public Function SafeFileName(ByVal strString As String) As String
    ' Отбрасываем vbNullChar и все что после
    If InStr(strString, vbNullChar) Then
        strString = TrimNull(strString)
    End If
    
    ' Заменяем VbTab
    If InStr(strString, vbTab) Then
        strString = Replace$(strString, vbTab, vbNullString)
    End If

    ' Отбрасываем все после ","
    If InStr(strString, strComma) Then
        strString = Left$(strString, InStr(strString, strComma) - 1)
    End If

    ' Отбрасываем все после ";"
    If InStr(strString, strCommaDot) Then
        strString = Left$(strString, InStr(strString, strCommaDot) - 1)
    End If

    SafeFileName = Trim$(strString)
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

    If InStr(str, strCommaDot) Then
        Str_x = Split(str, strCommaDot)
        str = Trim$(Str_x(0))
    End If

    If InStr(str, strComma) Then
        Str_x = Split(str, strComma)
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

        If InStr(cDir, strPercentage) Then
            cDir = ParserInf4Strings(strInfFilePath, cDir)
        End If
    End If

    cDir = Replace$(cDir, vbTab, vbNullString)
    cDir = Replace$(cDir, strKavichki, vbNullString)
    cDir = BackslashAdd2Path(cDir)
    WhereIsDir = TrimNull(cDir)
End Function

