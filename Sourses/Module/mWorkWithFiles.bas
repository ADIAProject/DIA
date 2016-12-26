Attribute VB_Name = "mWorkWithFiles"
Option Explicit

' Not add to project (if not DBS) - option for compile
#Const mbIDE_DBSProject = False

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CompareFilesByHashCAPICOM
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strFirstFile (String)
'                              strSecondFile (String)
'!--------------------------------------------------------------------------------
#If mbIDE_DBSProject Then
    Public Function CompareFilesByHashCAPICOM(ByVal strFirstFile As String, ByVal strSecondFile As String) As Boolean
    
        Dim strDataSHAFirst  As String
        Dim strDataSHASecond As String
        Dim lngResult        As Long
    
        If FileExists(strFirstFile) Then
            strDataSHAFirst = CalcHashFile(strFirstFile, CAPICOM_HASH_ALGORITHM_SHA1)
        End If
    
        If FileExists(strSecondFile) Then
            strDataSHASecond = CalcHashFile(strSecondFile, CAPICOM_HASH_ALGORITHM_SHA1)
        End If
    
        lngResult = StrComp(strDataSHAFirst, strDataSHASecond, vbTextCompare)
    
        CompareFilesByHashCAPICOM = lngResult = 0

    End Function
#End If

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function BackslashAdd2Path
'! Description (��������)  :   [���������� ����� �� �����]
'! Parameters  (����������):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function BackslashAdd2Path(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathAddBackslash strPath
    BackslashAdd2Path = TrimNull(strPath)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function BackslashDelFromPath
'! Description (��������)  :   [�������� ����� �� �����]
'! Parameters  (����������):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function BackslashDelFromPath(ByVal strPath As String) As String
    strPath = strPath & str2vbNullChar
    PathRemoveBackslash strPath
    BackslashDelFromPath = TrimNull(strPath)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CopyFileTo
'! Description (��������)  :   [��������� ���� 'strPathFrom' � ���������� 'CopyFileTo', ���� ���� ����������, �� �� ����� ����������� ����� ������.]
'! Parameters  (����������):   strPathFrom (String)
'                              strPathTo (String)
'!--------------------------------------------------------------------------------
Public Function CopyFileTo(ByVal strPathFrom As String, ByVal strPathTo As String) As Boolean

    Dim ret As Long

    If FileExists(strPathFrom) Then
        ' ��� ���� ������, ����� �������� ������ ��� ������, � ��������� ���� ����
        ResetReadOnly4File strPathTo
        ' ���������� �����������
        '���� �� ������, ����� ����� ���� �� ����������� �� ����� �������, �������� 'False' �� 'True'
        ret = CopyFile(strPathFrom, strPathTo, False)

        If ret <> 0 Then
            CopyFileTo = True
            ' ����� �������� ������ ��� ������, ���� ����
            ResetReadOnly4File strPathTo
        Else
            CopyFileTo = False
            MsgBox strMessages(42) & vbNewLine & "From: " & strPathFrom & vbNewLine & "To:" & strPathTo & vbNewLine & "Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError), vbExclamation, strProductName
            If mbDebugStandart Then DebugMode vbTab & "Copy file: False: " & strPathFrom & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        CopyFileTo = False
        If mbDebugStandart Then DebugMode vbTab & "Copy file: False : " & strPathFrom & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CopyFolderByShell
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sSource (String)
'                              sDestination (String)
'!--------------------------------------------------------------------------------
Public Function CopyFolderByShell(ByVal sSource As String, ByVal sDestination As String) As Long

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
'! Procedure   (�������)   :   Function CreateIfNotExistPath
'! Description (��������)  :   [�������� ��������, ���� �� ����������]
'! Parameters  (����������):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Function CreateIfNotExistPath(ByVal strFolderPath As String) As Boolean

    If LenB(strFolderPath) Then

        ' ���� ���, �� ������� �������
        If PathExists(strFolderPath) = False Then
            CreateNewDirectory strFolderPath
            CreateIfNotExistPath = PathIsAFolder(strFolderPath)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub CreateNewDirectory
'! Description (��������)  :   [�������� ������ ��������, ����������]
'! Parameters  (����������):   sNewDirectory (String)
'!--------------------------------------------------------------------------------
Public Sub CreateNewDirectory(ByVal sNewDirectory As String)

    Dim SecAttrib  As SECURITY_ATTRIBUTES
    Dim sPath      As String
    Dim iCounter   As Integer
    Dim sTempDir   As String
    Dim ret        As Long

    sPath = BackslashAdd2Path(sNewDirectory)
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
                If mbDebugStandart Then DebugMode str2VbTab & "CreateNewDirectory: False : " & sTempDir & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If
        End If
    Loop

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function DeleteFiles
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPathFile (String)
'!--------------------------------------------------------------------------------
Public Function DeleteFiles(ByVal strPathFile As String) As Boolean

    Dim ret             As Long
    Dim lngFilePathPtr  As Long
    
    If PathIsValidUNC(strPathFile) = False Then
        lngFilePathPtr = StrPtr("\\?\" & strPathFile)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strPathFile, Len(strPathFile) - 2))
    End If
    ret = DeleteFile(lngFilePathPtr)

    If ret = 0 Then
        ' ���� ��� �������, �� �������� ������� ������ ��� ������, �������� ����� � ����� ������� ����
        If Err.LastDllError = 5 Then
            ResetReadOnly4File strPathFile
            ret = DeleteFile(lngFilePathPtr)
            If ret = 0 Then
                If mbDebugStandart Then DebugMode vbTab & "DeleteFiles: False : " & strPathFile & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If
        Else
            If mbDebugStandart Then DebugMode vbTab & "DeleteFiles: False : " & strPathFile & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
        
    End If

    DeleteFiles = CBool(ret)
        
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function DeleteFolder
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Function DeleteFolder(ByVal strFolderPath As String) As Boolean

    Dim ret             As Long
    Dim lngFilePathPtr  As Long
    
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
        ' ����� �� �����
        If Err.LastDllError = 145 Then
            If mbDebugDetail Then DebugMode vbTab & "DeleteFiles: False : " & strFolderPath & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        Else
            If mbDebugStandart Then DebugMode vbTab & "DeleteFolder: False : " & strFolderPath & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    DeleteFolder = CBool(ret)
    
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DelFolderBackUp
'! Description (��������)  :   [�������� ���������� ��������, ���� �������� �����]
'! Parameters  (����������):   strFolderPath (String)
'!--------------------------------------------------------------------------------
Public Sub DelFolderBackUp(ByVal strFolderPath As String)

    Dim ret             As Long
    Dim lngFilePathPtr  As Long

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
            If mbDebugStandart Then DebugMode vbTab & "DelFolderBackUp: False : " & strFolderPath & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
    End If

    On Error GoTo 0

    If mbDebugDetail Then DebugMode "DelFolderBackUp-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DelRecursiveFolder
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sFolder (String)
'!--------------------------------------------------------------------------------
Public Sub DelRecursiveFolder(ByVal sFolder As String)

    Dim retDelete   As Long
    Dim retStrMsg   As String
    Dim sRoot       As String

    sRoot = BackslashDelFromPath(sFolder)
    If mbDebugStandart Then DebugMode vbTab & "DeleteFolder: " & sRoot

    If PathExists(sRoot) Then
        SearchFilesInRoot sRoot, ALL_FILES, True, False, True
        SearchFoldersInRoot sRoot, ALL_FOLDERS_EX, True, True

        ' �������� ������ ���������, ���� ��������
        If PathExists(sRoot) Then
            retDelete = DelTree(sRoot)

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
'! Procedure   (�������)   :   Sub DelTemp
'! Description (��������)  :   [�������� ���������� ��������, ���� �������� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub DelTemp()

    Dim lngTimeScriptRun    As Currency
    Dim lngTimeScriptFinish As Currency
    Dim ret                 As Long
    Dim lngFilePathPtr      As Long

    On Error Resume Next

    If mbDebugDetail Then DebugMode "DelTemp-Start"
    lngTimeScriptRun = GetTimeStart

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
            If mbDebugStandart Then DebugMode vbTab & "DelTemp: False : " & strWorkTemp & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If
        
    End If

    lngTimeScriptFinish = GetTimeStop(lngTimeScriptRun)
    If mbDebugStandart Then DebugMode "DelTemp-End: Time to Delete: " & CalculateTime(lngTimeScriptFinish, True)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function DelTree
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strDir (String)
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

                Do While LenB(strFile)

                    If AscW(strFile) <> vbDot Then
                        If StrComp(strFile, str2Dot) <> 0 Then
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
                    If mbDebugStandart Then DebugMode vbTab & "DelTree: False : " & strDir & " Error: �" & retLasrErr & " - " & ApiErrorText(retLasrErr)
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
'! Procedure   (�������)   :   Function ExpandFileNamebyEnvironment
'! Description (��������)  :   [��������� ��� ����� - ������������� ���������� %%]
'! Parameters  (����������):   strFileName (String)
'!--------------------------------------------------------------------------------
Public Function ExpandFileNameByEnvironment(ByVal strFileName As String) As String

    Dim r            As String
    Dim str_OSVer    As String
    Dim str_OSBit    As String
    Dim str_DATE     As String
    Dim str_PCMODEL  As String

    If InStr(strFileName, strPercent) Then
        ' ���������������� ������ �� %OSVer%
        str_OSVer = "wnt" & Left$(strOSCurrentVersion, 1)

        ' ���������������� �������� �� %OSBit%
        If mbIsWin64 Then
            str_OSBit = "x64"
        Else
            str_OSBit = "x32"
        End If

        ' ���������������� ���� %DATE%
        str_DATE = SafeDir(Replace$(CStr(Now()), strDot, strDash))
        ' ���������������� %PCMODEL%
        str_PCMODEL = SafeDir(Replace$(strCompModel, "_", strDash))
        
        ' ������ �������� ����������
        r = strFileName
        r = Replace$(r, "%PCNAME%", strCompModel, , , vbTextCompare)
        r = Replace$(r, "%PCMODEL%", str_PCMODEL, , , vbTextCompare)
        r = Replace$(r, "%OSVer%", str_OSVer, , , vbTextCompare)
        r = Replace$(r, "%OSBit%", str_OSBit, , , vbTextCompare)
        r = Replace$(r, "%DATE%", str_DATE, , , vbTextCompare)
        r = Trim$(r)
        ExpandFileNameByEnvironment = r
    Else
        ExpandFileNameByEnvironment = strFileName
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileisReadOnly
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPathFile (String)
'!--------------------------------------------------------------------------------
Public Function FileisReadOnly(ByVal strPathFile As String) As Boolean
    FileisReadOnly = GetAttr(strPathFile) And vbReadOnly
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileisSystemAttr
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPathFile (String)
'!--------------------------------------------------------------------------------
Public Function FileIsSystemAttr(ByVal strPathFile As String) As Boolean
    FileIsSystemAttr = GetAttr(strPathFile) And vbSystem
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileReadData
'! Description (��������)  :   [Read data from file with check for unicode yes/no]
'! Parameters  (����������):   sFileName (String)
'                              ByRef strResult (String)
'                              LocaleID (Long)
'!--------------------------------------------------------------------------------
Public Sub FileReadData(ByVal sFileName As String, ByRef strResult As String, Optional ByVal lngLocaleID As Long = 1033)

    Dim sText       As String
    Dim fNum        As Long
    Dim B1(0 To 1)  As Byte
    
    fNum = FreeFile

    Open sFileName For Binary Access Read Lock Write As fNum
    ' read first 2 byte, for check on Unicode
    Get #fNum, 1, B1()
    
    ' if Unicode &HFF and &HFE 255-254
    If B1(0) = &HFF And B1(1) = &HFE Then
        'sText = Space$(LOF(fNum) - 2)
        sText = MemAPIs.AllocStr(vbNullString, LOF(fNum) - 2)
        Seek #fNum, 3
        Get #fNum, , sText
        strResult = StrConv(sText, vbFromUnicode, lngLocaleID)
    ' If ANSI
    Else
        'sText = Space$(LOF(fNum))
        sText = MemAPIs.AllocStr(vbNullString, LOF(fNum))
        Seek #fNum, 1
        Get #fNum, , sText
        strResult = sText
    End If
    
    Close #fNum

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileWriteData
'! Description (��������)  :   [Write data to file with check]
'! Parameters  (����������):   sFileName (String)
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
'! Procedure   (�������)   :   Function FileWriteDataFromArray
'! Description (��������)  :   [Write data to file with check]
'! Parameters  (����������):   sFileName (String)
'                              sStringOut() (String)
'!--------------------------------------------------------------------------------
Public Sub FileWriteDataFromArray(ByVal sFileName As String, ByRef sStringOut() As String)

    Dim fNum As Integer
    Dim I As Long
    
    fNum = FreeFile

    Open sFileName For Binary Access Write Lock Write As fNum
    
    For I = LBound(sStringOut) To UBound(sStringOut)
        Put #fNum, , sStringOut(I) & vbNewLine
    Next I
    
    Close #fNum

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileWriteDataAPI
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sFilePath (String)
'                              strData (String)
'!--------------------------------------------------------------------------------
Public Sub FileWriteDataAPI(ByVal sFilePath As String, ByVal strData As String)
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
'! Procedure   (�������)   :   Function FileWriteDataAPIUni
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sFilePath (String)
'                              strData (String)
'!--------------------------------------------------------------------------------
Private Sub FileWriteDataAPIUni(ByVal sFilePath As String, ByVal strData As String)
    Dim fHandle         As Long
    Dim fSuccess        As Long
    Dim lBytesWritten   As Long
    Dim BytesToWrite    As Long
    Dim anArray()       As Byte
    Dim lngFilePathPtr  As Long
    Dim lngStringSize   As Long
    
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
'! Procedure   (�������)   :   Function FileWriteDataAppend
'! Description (��������)  :   [Read data from file with check for unicode yes/no]
'! Parameters  (����������):   sFileName (String)
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
'! Procedure   (�������)   :   Function GetEnviron
'! Description (��������)  :   [��������� ���������� ���������� ���������]
'! Parameters  (����������):   strEnv (String)
'                              mbCollectFull (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function GetEnviron(ByVal strEnv As String, Optional ByVal mbCollectFull As Boolean = False) As String

    Dim strTemp        As String
    Dim strTempEnv     As String
    Dim strNumPosition As Long

    strNumPosition = InStr(strEnv, strPercent)

    If strNumPosition Then
        strTemp = Mid$(strEnv, strNumPosition + 1, Len(strEnv) - strNumPosition)
        strNumPosition = InStr(strTemp, strPercent)

        If strNumPosition Then
            strTemp = Left$(strTemp, strNumPosition - 1)
        End If
    End If

    strTempEnv = Environ$(strTemp)

    If mbCollectFull Then
        GetEnviron = Replace$(strEnv, strPercent & strTemp & strPercent, strTempEnv, , , vbTextCompare)
    Else
        GetEnviron = strTempEnv
    End If

    If mbDebugStandart Then DebugMode str2VbTab & "GetEnviron: %" & strTemp & "%=" & strTempEnv & vbNewLine & _
              str2VbTab & "GetEnviron-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileName_woExt
'! Description (��������)  :   [�������� ��� ����� ��� ����������, ���� ��� �����]
'! Parameters  (����������):   strFileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileName_woExt(ByVal strFileName As String) As String

    Dim intLastSeparator As Long

    GetFileName_woExt = strFileName

    If LenB(strFileName) Then
        intLastSeparator = InStrRev(strFileName, strDot)

        If intLastSeparator Then
            GetFileName_woExt = Left$(strFileName, intLastSeparator - 1)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileNameOnly_woExt
'! Description (��������)  :   [�������� ������ ��� ����� ��� ����������, ���� ��� �����]
'! Parameters  (����������):   strFilePath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileNameOnly_woExt(ByVal strFilePath As String) As String

    Dim intLastSeparator As Long
    Dim strFileNameTemp  As String

    strFileNameTemp = strFilePath
    
    If LenB(strFileNameTemp) Then
    
        intLastSeparator = InStrRev(strFileNameTemp, vbBackslash)

        If intLastSeparator >= 0 Then
            strFileNameTemp = Right$(strFileNameTemp, Len(strFileNameTemp) - intLastSeparator)
        End If
        
        intLastSeparator = InStrRev(strFileNameTemp, strDot)
    
        If intLastSeparator Then
            strFileNameTemp = Right$(strFileNameTemp, Len(strFileNameTemp) - intLastSeparator)
        End If

    End If
    
    GetFileNameOnly_woExt = strFileNameTemp
    
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileNameExtension
'! Description (��������)  :   [�������� ���������� ����� �� ���� ��� ����� �����]
'! Parameters  (����������):   strFileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileNameExtension(ByVal strFileName As String) As String

    Dim intLastSeparator As Long

    intLastSeparator = InStrRev(strFileName, strDot)

    If intLastSeparator Then
        GetFileNameExtension = Right$(strFileName, Len(strFileName) - intLastSeparator)
    Else
        GetFileNameExtension = vbNullString
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileNameFromPath
'! Description (��������)  :   [�������� ��� ����� �� ������� ����]
'! Parameters  (����������):   strFilePath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileNameFromPath(ByVal strFilePath As String) As String

    Dim intLastSeparator As Long

    GetFileNameFromPath = strFilePath

    If LenB(strFilePath) Then
        intLastSeparator = InStrRev(strFilePath, vbBackslash)

        If intLastSeparator >= 0 Then
            GetFileNameFromPath = Right$(strFilePath, Len(strFilePath) - intLastSeparator)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileSizeByPath
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function GetFileSizeByPath(ByVal strPath As String) As Long

    Dim lHandle         As Long
    Dim lngFilePathPtr  As Long
    
    If PathIsValidUNC(strPath) = False Then
        lngFilePathPtr = StrPtr("\\?\" & strPath)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(strPath, Len(strPath) - 2))
    End If
    
    lHandle = CreateFile(lngFilePathPtr, GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)

    If lHandle <> INVALID_HANDLE_VALUE Then
        GetFileSizeByPath = GetFileSize(lHandle, 0&)
        CloseHandle lHandle
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetFileVersionOnly
'! Description (��������)  :   [Return file version information string.]
'! Parameters  (����������):   sFileName (String)
'!--------------------------------------------------------------------------------
Public Function GetFileVersionOnly(ByVal sFileName As String) As String
    Dim nUnused     As Long
    Dim sBuffer()   As Byte
    Dim nBufferSize As Long
    Dim lpBuffer    As Long
    Dim FFI         As VS_FIXEDFILEINFO
    Dim nVerSize    As Long
    Dim sResult     As String

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
'! Procedure   (�������)   :   Function GetPathNameFromPath
'! Description (��������)  :   [�������� ���� � ����� �� ������� ����]
'! Parameters  (����������):   strFilePath (String)
'!--------------------------------------------------------------------------------
Public Function GetPathNameFromPath(ByVal strFilePath As String) As String

    Dim intLastSeparator As Long

    GetPathNameFromPath = strFilePath
    
    intLastSeparator = InStrRev(strFilePath, vbBackslash)

    If intLastSeparator Then
        If intLastSeparator < Len(strFilePath) Then
            GetPathNameFromPath = Left$(strFilePath, intLastSeparator)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetUniqueTempFile
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
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
'! Procedure   (�������)   :   Function IsDriveCDRoom
'! Description (��������)  :   [�������� �� ����� ��������� � CD\DVD]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function IsDriveCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Long

    strDriveName = Left$(strAppPath, 3)

    ' ��������� �� ������ �� ����
    If InStr(strDriveName, vbBackslashDouble) = 0 Then
        '�������� ��� �����
        If PathIsRoot(strDriveName) Then
            xDrv = GetDriveType(strDriveName)
    
            If xDrv = DRIVE_CDROM Then
                IsDriveCDRoom = True
            End If
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function MoveFileTo
'! Description (��������)  :   [��������� ���� 'strPathFrom' � ���������� 'strPathTo', ���� ���� ����������, �� �� ����� ����������� ����� ������.]
'! Parameters  (����������):   strPathFrom (String)
'                              strPathTo (String)
'!--------------------------------------------------------------------------------
Public Function MoveFileTo(ByVal strPathFrom As String, ByVal strPathTo As String) As Boolean

    Dim ret As Long

    If StrComp(strPathFrom, strPathTo, vbTextCompare) <> 0 Then
        If FileExists(strPathFrom) Then
            ' ��� ���� ������, ����� �������� ������ ��� ������, � ��������� ���� ����
            ResetReadOnly4File strPathTo
            ' ���������� �����������
            '���� �� ������, ����� ����� ���� �� ����������� �� ����� �������, �������� 'False' �� 'True'
            ret = MoveFile(strPathFrom, strPathTo)

            If ret <> 0 Then
                MoveFileTo = True
                ' ����� �������� ������ ��� ������, ���� ����
                ResetReadOnly4File strPathTo
            Else
                MoveFileTo = False
                MsgBox strMessages(42) & vbNewLine & "From: " & strPathFrom & vbNewLine & "To:" & strPathTo & vbNewLine & "Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError), vbExclamation, strProductName
                If mbDebugStandart Then DebugMode vbTab & "Move file: False: " & strPathFrom & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
            End If

        Else
            MoveFileTo = False
            If mbDebugStandart Then DebugMode vbTab & "Move file: False : " & strPathFrom & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        End If

    Else
        If mbDebugStandart Then DebugMode vbTab & "Move file: Source and Destination are identicaly (" & strPathFrom & " ; " & strPathTo & ")"
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function ParserInf4Strings
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strInfFilePath (String)
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
    Dim strFileContent As String
    Dim Key            As String
    Dim Value          As String
    Dim I              As Long
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
    strFileContent = vbNullString
    lngFileDBSize = GetFileSizeByPath(strInfFilePath)

    If lngFileDBSize Then
        FileReadData strInfFilePath, strFileContent
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "DevParserByRegExp: File is zero = 0 bytes:" & strInfFilePath
    End If

    ' Find [strings] section
    Strings = vbNullString
    StringHash.CompareMode = TextCompare
    Set MatchesStrSect = RegExpStrSect.Execute(strFileContent)

    If MatchesStrSect.count Then
        Set objMatch = MatchesStrSect.item(0)
        Strings = objMatch.SubMatches(0) & objMatch.SubMatches(1)
        Set MatchesStrDefs = RegExpStrDefs.Execute(Strings)

        For I = 0 To MatchesStrDefs.count - 1
            Set objMatch1 = MatchesStrDefs.item(I)
            Key = objMatch1.SubMatches(0)
            Value = objMatch1.SubMatches(1)

            If LenB(Value) = 0 Then
                Value = objMatch1.SubMatches(2)
            End If

            If Not StringHash.Exists(Key) Then
                StringHash.Add Key, Value
                StringHash.Add strPercent & Key & strPercent, Value
            End If

        Next

    End If

    ' ���������� ���� ���� ����������
    Pos = InStr(strSearchString, strPercent)

    If Pos Then
        varname = Mid$(strSearchString, Pos, InStrRev(strSearchString, strPercent))
        valval = StringHash.item(varname)

        If LenB(valval) = 0 Then
            If mbDebugDetail Then DebugMode "ParserInf4Strings: Error in inf: Cannot find '" & strSearchString & "'"
        Else
            ParserInf4Strings = Replace$(strSearchString, varname, valval)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function PathCollect
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sPath (String)
'!--------------------------------------------------------------------------------
Public Function PathCollect(ByVal sPath As String) As String

    If InStr(sPath, strColon) = 2 Then
        PathCollect = sPath
    ElseIf Left$(sPath, 2) = vbBackslashDouble And PathIsValidUNC(sPath) Then
        PathCollect = sPath
    Else

        If Left$(sPath, 2) = ".\" Then
            PathCollect = PathCombine(strAppPath, sPath)
        Else

            If InStr(sPath, vbBackslash) = 1 Then
                PathCollect = strAppPath & sPath
            Else

                If Left$(sPath, 3) = "..\" Then
                    PathCollect = PathCombine(strAppPath, sPath)
                Else

                    If InStr(sPath, strPercent) Then
                        PathCollect = GetEnviron(sPath, True)
                    Else

                        If LenB(GetFileNameExtension(sPath)) Then
                            If GetFileNameFromPath(sPath) = sPath Then
                                PathCollect = sPath
                            Else
                                PathCollect = strAppPathBackSL & sPath
                            End If

                        Else
                            PathCollect = strAppPathBackSL & sPath
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect, vbBackslashDouble) Then
        If Left$(strAppPath, 2) <> vbBackslashDouble Then
            PathCollect = Replace$(PathCollect, vbBackslashDouble, vbBackslash)
        End If
    End If

    If PathIsAFolder(PathCollect) Then
        PathCollect = BackslashAdd2Path(PathCollect)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function PathCollect4Dest
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sPath (String)
'                              strDest (String)
'!--------------------------------------------------------------------------------
Public Function PathCollect4Dest(ByVal sPath As String, ByVal strDest As String) As String

    If InStr(sPath, strColon) = 2 Then
        PathCollect4Dest = sPath
    Else

        If Left$(sPath, 2) = ".\" Then
            PathCollect4Dest = strDest & Mid$(sPath, 2, Len(sPath) - 1)
        Else

            If InStr(sPath, vbBackslash) = 1 Then
                PathCollect4Dest = strDest & sPath
            Else

                If Left$(sPath, 3) = "..\" Then
                    PathCollect4Dest = GetPathNameFromPath(strDest) & Mid$(sPath, 4, Len(sPath) - 1)
                Else

                    If InStr(sPath, strPercent) Then
                        PathCollect4Dest = GetEnviron(sPath, True)
                    Else

                        If LenB(GetFileNameExtension(sPath)) Then
                            If GetFileNameFromPath(sPath) = sPath Then
                                PathCollect4Dest = sPath
                            Else
                                PathCollect4Dest = BackslashAdd2Path(strDest) & sPath
                            End If

                        Else
                            PathCollect4Dest = BackslashAdd2Path(strDest) & sPath
                        End If
                    End If
                End If
            End If
        End If
    End If

    If InStr(PathCollect4Dest, vbBackslashDouble) Then
        PathCollect4Dest = Replace$(PathCollect4Dest, vbBackslashDouble, vbBackslash)

        If Left$(strDest, 2) = vbBackslashDouble Then
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
'! Procedure   (�������)   :   Function PathCombine
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strDirectory (String)
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
'! Procedure   (�������)   :   Function PathExists
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strPath (String)
'!--------------------------------------------------------------------------------
Public Function PathExists(ByVal strPath As String) As Boolean
    PathExists = PathFileExists(StrPtr(strPath & vbNullChar))
    'PathExists = PathIsDirectory(StrPtr(strPath & vbNullChar))
    'PathExists = FolderExists(strPath)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function PathIsAFolder
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sPath (String)
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
'! Procedure   (�������)   :   Function PathIsValidUNC
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   sPath (String)
'!--------------------------------------------------------------------------------
Public Function PathIsValidUNC(ByVal sPath As String) As Boolean
    ' Returns True if the string is a valid UNC path.
    PathIsValidUNC = PathIsUNC(StrPtr(sPath))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ResetReadOnly4File
'! Description (��������)  :   [����� �������� "������ ��� ������"]
'! Parameters  (����������):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Public Sub ResetReadOnly4File(ByVal strPathFile As String)

    If FileExists(strPathFile) Then
        If (GetAttr(strPathFile) And vbReadOnly) Then
            SetAttr strPathFile, vbNormal
        End If

        If (GetAttr(strPathFile) And vbSystem) Then
            SetAttr strPathFile, vbNormal
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function SafeDir
'! Description (��������)  :   [function to replace special chars to create dirs correctly]
'! Parameters  (����������):   str (String)
'!--------------------------------------------------------------------------------
Public Function SafeDir(ByVal str As String) As String

    If InStr(str, vbBackslash) Then
        str = Replace$(str, vbBackslash, strDash)
    End If
    
    If InStr(str, "/") Then
        str = Replace$(str, "/", strDash)
    End If
    
    If InStr(str, "*") Then
        str = Replace$(str, "*", strDash)
    End If
    
    If InStr(str, strColon) Then
        str = Replace$(str, strColon, strDash)
    End If
    
    If InStr(str, strVopros) Then
        str = Replace$(str, strVopros, strDash)
    End If
    
    If InStr(str, ">") Then
       str = Replace$(str, ">", strDash)
    End If
    
    If InStr(str, "<") Then
        str = Replace$(str, "<", strDash)
    End If
    
    If InStr(str, "|") Then
        str = Replace$(str, "|", strDash)
    End If
    
    If InStr(str, "@") Then
        str = Replace$(str, "@", strDash)
    End If
    
    If InStr(str, "'") Then
        str = Replace$(str, "'", vbNullString)
    End If
        
    If InStr(str, strSpace) Then
        str = Replace$(str, strSpace, strDash)
    End If
    
    If InStr(str, "(R)") Then
        str = Replace$(str, "(R)", strDash)
    End If
        
    If InStr(str, "---") Then
        str = Replace$(str, "---", strDash)
    End If
    
    If InStr(str, "--") Then
        str = Replace$(str, "--", strDash)
    End If
        
    SafeDir = Trim$(str)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function SafeFileName
'! Description (��������)  :   [function to replace special chars to create files correctly]
'! Parameters  (����������):   strString (Variant)
'!--------------------------------------------------------------------------------
Public Function SafeFileName(ByVal strString As String) As String
    ' ����������� vbNullChar � ��� ��� �����
    If InStr(strString, vbNullChar) Then
        strString = TrimNull(strString)
    End If
    
    ' �������� VbTab
    If InStr(strString, vbTab) Then
        strString = Replace$(strString, vbTab, vbNullString)
    End If

    ' ����������� ��� ����� ","
    If InStr(strString, strComma) Then
        strString = Left$(strString, InStr(strString, strComma) - 1)
    End If

    ' ����������� ��� ����� ";"
    If InStr(strString, strSemiColon) Then
        strString = Left$(strString, InStr(strString, strSemiColon) - 1)
    End If

    SafeFileName = Trim$(strString)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function WhereIsDir
'! Description (��������)  :   [function to discover dirs with inf code]
'! Parameters  (����������):   str (String)
'                              strInfFilePath (String)
'!--------------------------------------------------------------------------------
Public Function WhereIsDir(ByVal str As String, ByVal strInfFilePath As String) As String

    Dim strSpecDir       As String
    Dim str_x()          As String
    Dim mbAdditionalPath As Boolean

    If InStr(str, strSemiColon) Then
        str_x = Split(str, strSemiColon)
        str = Trim$(str_x(0))
    End If

    If InStr(str, strComma) Then
        str_x = Split(str, strComma)
        mbAdditionalPath = True
        str = str_x(0)
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
            strSpecDir = strSysDrive

        Case "10"
            strSpecDir = strWinDir

            'system32 ���������� �� �����
        Case "11"
            strSpecDir = strSysDir86

        Case "12"
            strSpecDir = strSysDir86 & "Drivers"

        Case "17"
            strSpecDir = strInfDir

        Case "18"
            strSpecDir = strWinDir & "Help"

        Case "20"
            strSpecDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "21"
            strSpecDir = vbNullString

            'viewer dir
        Case "23"
            strSpecDir = strSysDir86 & "spool\drivers\color"

        Case "24"
            strSpecDir = strSysDrive

        Case "25"
            strSpecDir = vbNullString

            'shared dir
        Case "30"
            strSpecDir = strSysDrive

        Case "50"
            strSpecDir = strWinDir & "system"

        Case "51"
            strSpecDir = strSysDir86 & "Spool"

        Case "52"
            strSpecDir = strSysDir86 & "Spool\Drivers"

        Case "53"
            strSpecDir = vbNullString

            'user profile dir
        Case "54"
            strSpecDir = vbNullString

            ' ntldr.exe dir
        Case "55"
            strSpecDir = strSysDir86 & "spool\prtprocs"

        Case "16384"
            strSpecDir = GetSpecialFolderPath(CSIDL_DESKTOPDIRECTORY)

        Case "16386"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAMS)

        Case "16389"
            strSpecDir = GetSpecialFolderPath(CSIDL_MYDOCUMENTS)

        Case "16391"
            strSpecDir = GetSpecialFolderPath(CSIDL_STARTUP)

        Case "16392"
            strSpecDir = GetSpecialFolderPath(CSIDL_RECENT)

        Case "16393"
            strSpecDir = GetSpecialFolderPath(CSIDL_SENDTO)

        Case "16395"
            strSpecDir = GetSpecialFolderPath(CSIDL_STARTMENU)

        Case "16397"
            strSpecDir = GetSpecialFolderPath(CSIDL_MYMUSIC)

        Case "16397"
            strSpecDir = GetSpecialFolderPath(CSIDL_MYVIDEO)

        Case "16400"
            strSpecDir = GetSpecialFolderPath(CSIDL_DESKTOP)

        Case "16403"
            strSpecDir = GetSpecialFolderPath(CSIDL_NETHOOD)

        Case "16404"
            strSpecDir = GetSpecialFolderPath(CSIDL_FONTS)

        Case "16405"
            strSpecDir = GetSpecialFolderPath(CSIDL_TEMPLATES)

        Case "16406"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_STARTMENU)

        Case "16407"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_PROGRAMS)

        Case "16408"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_STARTUP)

        Case "16409"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)

        Case "16410"
            strSpecDir = GetSpecialFolderPath(CSIDL_APPDATA)

        Case "16411"
            strSpecDir = GetSpecialFolderPath(CSIDL_PRINTHOOD)

        Case "16412"
            strSpecDir = GetSpecialFolderPath(CSIDL_LOCAL_APPDATA)

        Case "16415"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_FAVORITES)

        Case "16416"
            strSpecDir = GetSpecialFolderPath(CSIDL_INTERNET_CACHE)

        Case "16417"
            strSpecDir = GetSpecialFolderPath(CSIDL_COOKIES)

        Case "16418"
            strSpecDir = GetSpecialFolderPath(CSIDL_HISTORY)

        Case "16419"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_APPDATA)

        Case "16420"
            strSpecDir = strWinDir

        Case "16421"
            strSpecDir = strSysDir86

        Case "16422"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES)

        Case "16423"
            strSpecDir = GetSpecialFolderPath(CSIDL_MYPICTURES)

        Case "16424"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROFILE)

        Case "16425"
            strSpecDir = strSysDir64

        Case "16426"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16427"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMON)

        Case "16428"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILES_COMMONX86)

        Case "16429"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_TEMPLATES)

        Case "16430"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_DOCUMENTS)

        Case "16432"
            strSpecDir = GetSpecialFolderPath(CSIDL_PROGRAM_FILESX86)

        Case "16437"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_MUSIC)

        Case "16438"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_PICTURES)

        Case "16439"
            strSpecDir = GetSpecialFolderPath(CSIDL_COMMON_VIDEO)

        Case "16440"
            strSpecDir = strWinDir & "resources"

        Case "16441"
            strSpecDir = strWinDir & "resources\0409"

        Case "-1"
            strSpecDir = vbNullString

            ' absolute path
            'http://msdn.microsoft.com/en-us/library/ff560821.aspx
        Case "66000"
            strSpecDir = Getpath_PrinterDriverDirectory

            If LenB(strSpecDir) = 0 Then
                strSpecDir = strSysDir86 & "spool\Drivers\w32x86"
            End If

        Case "66001"
            strSpecDir = Getpath_PrintProcessorDirectory

            If LenB(strSpecDir) = 0 Then
                strSpecDir = strSysDir86 & "spool\prtprocs\w32x86"
            End If

        Case "66002"
            strSpecDir = strSysDir86

        Case "66003"
            strSpecDir = Getpath_PrinterColorDirectory

            If LenB(strSpecDir) = 0 Then
                strSpecDir = strSysDir86 & "spool\drivers\color"
            End If

        Case "66004"
            strSpecDir = strSysDir86 & "spool\Drivers\w32x86"

        Case Else
            strSpecDir = vbNullString
    End Select

    If InStr(strSpecDir, vbNullChar) Then
        strSpecDir = TrimNull(strSpecDir)
    End If

    If mbAdditionalPath Then
        strSpecDir = BackslashAdd2Path(strSpecDir) & Trim$(str_x(1))

        If InStr(strSpecDir, strPercent) Then
            strSpecDir = ParserInf4Strings(strInfFilePath, strSpecDir)
        End If
    End If

    strSpecDir = Replace$(strSpecDir, vbTab, vbNullString)
    strSpecDir = Replace$(strSpecDir, strQuotes, vbNullString)
    strSpecDir = BackslashAdd2Path(strSpecDir)
    WhereIsDir = TrimNull(strSpecDir)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FileIs7zip
'! Description (��������)  :   [�������� ��������� ����� �� ����������� ������ 7-zip]
'! Parameters  (����������):   strPathFileName (String)
'!--------------------------------------------------------------------------------
Public Function FileIs7zip(ByVal strPathFileName As String) As Boolean
    If FileExists(strPathFileName) = True Then
        Dim hFile As Long, Length As Long
        Dim B1(0 To 3) As Byte
        hFile = CreateFile(StrPtr("\\?\" & IIf(Left$(strPathFileName, 2) = "\\", "UNC\" & Mid$(strPathFileName, 3), strPathFileName)), GENERIC_READ, FILE_SHARE_READ, 0, OPEN_EXISTING, FILE_FLAG_SEQUENTIAL_SCAN, 0)
        If hFile <> INVALID_HANDLE_VALUE Then
            Length = GetFileSize(hFile, 0) ' File size >= 2^31 not supported.
            If Length > 4 Then
                ReadFile hFile, VarPtr(B1(0)), 4, 0, 0
            End If
            CloseHandle hFile
        End If
        If B1(0) = &H37 And B1(1) = &H7A And B1(2) = &HBC And B1(3) = &HAF Then
            FileIs7zip = True
        End If
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function FolderExists
'! Description (��������)  :   [�������� ������������� ��������]
'! Parameters  (����������):   strPathName (String)
'!--------------------------------------------------------------------------------
Public Function FolderExists(ByVal strPathName As String) As Boolean
    On Error Resume Next
    Dim Attributes As VbFileAttribute, ErrVal As Long
    Attributes = GetAttr(strPathName)
    ErrVal = Err.Number
    On Error GoTo 0
    If (Attributes And (vbDirectory Or vbVolume)) > 0 And ErrVal = 0 Then FolderExists = True
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function ListingDirectory
'! Description (��������)  :   [������� �������� � ���� ������]
'! Parameters  (����������):   strPath (String), mbRecursion (Boolean)
'!--------------------------------------------------------------------------------
Public Function ListingDirectory(ByVal strPath As String, ByVal mbRecursion As Boolean) As String

    Dim strFileList_x() As FindListStruct
    Dim strFileList     As String
    Dim strFileListTemp As String
    Dim ii              As Long
    Dim lngLBound       As Long
    Dim lngUbound       As Long

    If mbDebugDetail Then DebugMode "***ListingDirectory-Start: source=" & strPath

    If LenB(strPath) > 0 Then
        strFileList_x = SearchFilesInRoot(strPath, ALL_FILES, mbRecursion, False, False)
        strFileList = vbNullString

        If UBound(strFileList_x) >= 0 Then
            If LenB(strFileList_x(0).FullPath) Then

                lngLBound = LBound(strFileList_x)
                lngUbound = UBound(strFileList_x)

                For ii = lngLBound To lngUbound
                    strFileListTemp = strFileList_x(ii).Name

                    If LenB(strFileListTemp) Then
                        AppendStr strFileList, strFileListTemp, ";"
                    End If

                Next
            End If
        End If

    Else
        If mbDebugDetail Then DebugMode "***ListingDirectory-Source Path not defined"
    End If

    ListingDirectory = strFileList
    If mbDebugDetail Then DebugMode "***ListingDirectory-Finish"
End Function

