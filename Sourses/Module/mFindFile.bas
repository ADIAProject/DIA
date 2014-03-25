Attribute VB_Name = "mFindFile"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Const vbDot            As Integer = 46

Private fp                       As FILE_PARAMS    'holds search parameters
Private fp2                      As FOLDER_PARAMS  'holds search parameters
Private sResultFileList()        As FindListStruct
Private sResultFolderList()      As FindListStruct
Private lngResultFileListCount   As Long
Private lngResultFolderListCount As Long

Public Type FindListStruct
    Path            As String
    Name            As String
    FullPath        As String
    RelativePath    As String
    NameLcase       As String
    NameWoExt       As String
    Size            As Long
    SizeInString    As String
End Type


'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FileSizeApi
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sSource (String)
'!--------------------------------------------------------------------------------
Public Function FileSizeApi(ByVal sSource As String) As String

    Dim wfd             As WIN32_FIND_DATA
    Dim hFile           As Long
    Dim sSize           As String
    Dim lngFilePathPtr  As Long

    If PathIsValidUNC(sSource) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sSource)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sSource, Len(sSource) - 2))
    End If
    hFile = FindFirstFile(lngFilePathPtr, wfd)

    If hFile <> INVALID_HANDLE_VALUE Then
        sSize = String$(30, vbNullChar)

        If InStr(1, sSource, TrimNull(wfd.cFileName), vbTextCompare) Then
            StrFormatByteSizeW wfd.nFileSizeLow, wfd.nFileSizeHigh, ByVal StrPtr(sSize), 30
            FileSizeApi = TrimNull(sSize)
        Else
            FileSizeApi = "0 byte"
        End If
        
        FindClose hFile
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function MatchSpec
'! Description (Описание)  :   [Проверка на соответствие условиям поиска]
'! Parameters  (Переменные):   sFile (String)
'                              sSpec (String)
'!--------------------------------------------------------------------------------
Public Function MatchSpec(ByVal sFile As String, ByVal sSpec As String) As Boolean

    If LenB(sFile) Then
        If LenB(sSpec) Then
            MatchSpec = PathMatchSpec(StrPtr(sFile & vbNullChar), StrPtr(sSpec & vbNullChar))
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function rgbCopyFiles
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sSourcePath (String)
'                              sDestination (String)
'                              sFiles (String)
'!--------------------------------------------------------------------------------
Public Function rgbCopyFiles(ByVal sSourcePath As String, ByVal sDestination As String, ByVal sFiles As String) As Long

    Dim wfd                   As WIN32_FIND_DATA
    Dim sA                    As SECURITY_ATTRIBUTES
    Dim hFile                 As Long
    Dim copied                As Long
    Dim currFile              As String
    Dim currSourcePath        As String
    Dim lngNumFilesFromFolder As Long
    Dim lngFilePathPtr        As Long

    sSourcePath = BackslashAdd2Path(sSourcePath)
    sDestination = BackslashAdd2Path(sDestination)

    'Create the target directory if it doesn't exist
    If PathExists(sDestination) = False Then
        Call CreateDirectory(sDestination, sA)
    End If

    'Start searching for files in the Target directory.
    If PathIsValidUNC(sSourcePath) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sSourcePath & sFiles)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sSourcePath, Len(sSourcePath) - 2) & sFiles)
    End If
    hFile = FindFirstFile(lngFilePathPtr, wfd)

    If hFile <> INVALID_HANDLE_VALUE Then

        'Copy each file to the new directory
        Do
            currFile = TrimNull(wfd.cFileName)

            If Asc(wfd.cFileName) <> vbDot Then
                currSourcePath = sSourcePath & currFile

                If Not PathIsAFolder(currSourcePath) Then
                    If MatchSpec(currFile, sFiles) Then
                        'copy the file to the destination directory & increment the count
                        Call CopyFileTo(currSourcePath, sDestination & currFile)
                        copied = copied + 1
                    End If

                Else
                    ' Копируем содержимое архива
                    If mbDebugStandart Then DebugMode str2VbTab & "CopyFiles from SubFolder: " & currFile
                    lngNumFilesFromFolder = rgbCopyFiles(currSourcePath, sDestination & currFile, ALL_FILES)
                    If mbDebugStandart Then DebugMode str2VbTab & "CopyFiles SubFolder - count files: " & lngNumFilesFromFolder
                    copied = copied + lngNumFilesFromFolder
                End If
            End If

            'just to check what's happening
            'find the next file matching the initial file spec
        Loop While FindNextFile(hFile, wfd)

        'Close the search handle
        FindClose hFile
        
        'and return the number of files copied
        rgbCopyFiles = copied
    Else
        If mbDebugStandart Then DebugMode str2VbTab & "CopyAllFilesFromFolder: " & sSourcePath & " No " & sFiles & " files found."
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SearchFilesInRoot
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strRootDir (String)
'                              strSearchMask (String)
'                              mbSearchRecursion (Boolean)
'                              mbOnlyFirstFile (Boolean)
'                              mbDelete (Boolean = False)
'                              mbSort (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function SearchFilesInRoot(ByVal strRootDir As String, ByVal strSearchMask As String, ByVal mbSearchRecursion As Boolean, ByVal mbOnlyFirstFile As Boolean, Optional mbDelete As Boolean = False, Optional mbSort As Boolean = False) As FindListStruct()

    With fp
        .sFileRoot = BackslashAdd2Path(strRootDir)
        .sFileNameExt = strSearchMask
        .bRecurse = mbSearchRecursion
    End With

    SearchForFiles fp.sFileRoot, True, 100, mbDelete

    If Not mbDelete Then
        If mbOnlyFirstFile Then
            ReDim Preserve sResultFileList(0)
            SearchFilesInRoot = sResultFileList
        Else

            'If mbSort Then
                'QuickSortMDArray sResultFileList, 1, 0
            'End If

            SearchFilesInRoot = sResultFileList
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SearchFoldersInRoot
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strRootDir (String)
'                              strSearchMask (String)
'                              mbSearchRecursion (Boolean)
'                              mbOnlyFirstFile (Boolean)
'!--------------------------------------------------------------------------------
Public Function SearchFoldersInRoot(ByVal strRootDir As String, ByVal strSearchMask As String, Optional ByVal mbSearchRecursion As Boolean = False, Optional ByVal mbDelete As Boolean = False) As FindListStruct()

    With fp2
        .sFileRoot = BackslashAdd2Path(strRootDir)
        .sFileNameExt = strSearchMask
        .bRecurse = mbSearchRecursion
    End With

    SearchForFolders fp2.sFileRoot, True, 100, mbDelete
    
    If Not mbDelete Then
        SearchFoldersInRoot = sResultFolderList
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SearchForFiles
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sRoot (String)
'                              mbInitial (Boolean)
'                              miMaxCountArr (Long)
'                              mbDelete (Boolean = False)
'!--------------------------------------------------------------------------------
Private Sub SearchForFiles(ByVal sRoot As String, ByVal mbInitial As Boolean, ByRef miMaxCountArr As Long, Optional ByRef mbDelete As Boolean = False, Optional ByVal sRootInit As String = vbNullString)

    Dim wfd         As WIN32_FIND_DATA
    Dim hFile       As Long
    Dim sSize       As String
    Dim strFileName As String
    Dim lngFilePathPtr As Long

    If PathIsValidUNC(sRoot) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sRoot & ALL_FILES)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sRoot, Len(sRoot) - 2) & ALL_FILES)
    End If
    hFile = FindFirstFile(lngFilePathPtr, wfd)
    
    If Not mbDelete Then
        If mbInitial Then
        
            lngResultFileListCount = 0
            ReDim sResultFileList(miMaxCountArr)
            sRootInit = sRoot

        Else

            ReDim Preserve sResultFileList(miMaxCountArr)

        End If
    End If

    If hFile <> INVALID_HANDLE_VALUE Then

        Do
            strFileName = TrimNull(wfd.cFileName)

            'if a folder, and recurse specified, call method again
            If (wfd.dwFileAttributes And vbDirectory) Then
                If Asc(strFileName) <> vbDot Then
                    If fp.bRecurse Then
                        SearchForFiles sRoot & strFileName & vbBackslash, False, miMaxCountArr, mbDelete, sRootInit
                    End If
                End If

            Else

                'must be a file..
                If MatchSpec(strFileName, fp.sFileNameExt) Then
                        ' Если есть флаг удаления, то запускаем его, иначе добавляем значение в массив
                    If mbDelete Then
                        DeleteFiles sRoot & strFileName
                    Else

                        ' Переопределение массива если превышаем заданную размерность
                        If lngResultFileListCount >= miMaxCountArr Then
                            miMaxCountArr = 2 * miMaxCountArr

                            ReDim Preserve sResultFileList(miMaxCountArr)

                        End If

                        ' Полный путь файла
                        sResultFileList(lngResultFileListCount).FullPath = sRoot & strFileName
                        ' размер файла числовой в байтах
                        sResultFileList(lngResultFileListCount).Size = wfd.nFileSizeLow
                        ' размер файла строковый форматированный учитывая региональные настройки в байт/кбайт/мбайт и т.д
                        sSize = String$(30, vbNullChar)
                        StrFormatByteSizeW wfd.nFileSizeLow, wfd.nFileSizeHigh, ByVal StrPtr(sSize), 30
                        sResultFileList(lngResultFileListCount).SizeInString = TrimNull(sSize)
                        ' Путь до файла
                        sResultFileList(lngResultFileListCount).Path = sRoot
                        If Not mbInitial Then
                            sResultFileList(lngResultFileListCount).RelativePath = Replace$(sRoot, sRootInit, vbNullString)
                        End If
                        ' Имя файла
                        sResultFileList(lngResultFileListCount).Name = strFileName
                        ' Имя файла smallcase
                        sResultFileList(lngResultFileListCount).NameLcase = LCase$(strFileName)
                        sResultFileList(lngResultFileListCount).NameWoExt = GetFileName_woExt(strFileName)
                        lngResultFileListCount = lngResultFileListCount + 1
                    End If
                End If
            End If

        Loop While FindNextFile(hFile, wfd)

        FindClose hFile
    End If


    ' Переопределение массива на реальное кол-во записей
    If Not mbDelete Then
        If mbInitial Then
            If lngResultFileListCount Then
                ReDim Preserve sResultFileList(lngResultFileListCount - 1)
            Else
                ReDim Preserve sResultFileList(lngResultFileListCount)
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SearchForFolders
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sRoot (String)
'                              mbInitial (Boolean)
'                              miMaxCountArr (Long)
'!--------------------------------------------------------------------------------
Private Sub SearchForFolders(ByVal sRoot As String, ByVal mbInitial As Boolean, ByRef miMaxCountArr As Long, Optional ByRef mbDelete As Boolean = False, Optional ByVal sRootInit As String = vbNullString)

    Dim wfd         As WIN32_FIND_DATA
    Dim hFile       As Long
    Dim strFindData As String
    Dim lngFilePathPtr As Long

    If PathIsValidUNC(sRoot) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sRoot & ALL_FILES)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sRoot, Len(sRoot) - 2) & ALL_FILES)
    End If
    hFile = FindFirstFile(lngFilePathPtr, wfd)
    
    If Not mbDelete Then
        If mbInitial Then
        
            lngResultFolderListCount = 0
            ReDim sResultFolderList(miMaxCountArr)
            sRootInit = sRoot
    
        Else
    
            ReDim Preserve sResultFolderList(miMaxCountArr)
    
        End If
    End If
    
    If hFile <> INVALID_HANDLE_VALUE Then

        Do
            strFindData = TrimNull(wfd.cFileName)

            If (wfd.dwFileAttributes And vbDirectory) Then
                If Asc(strFindData) <> vbDot Then
                                                
                    If MatchSpec(strFindData, fp2.sFileNameExt) Then

                        ' Если есть флаг удаления, то запускаем его, иначе добавляем значение в массив
                        If mbDelete Then
                            If fp2.bRecurse Then
                                'if a folder, and recurse specified, call  method again
                                SearchForFolders sRoot & strFindData & vbBackslash, False, miMaxCountArr, mbDelete, sRootInit
                            End If
                            DeleteFolder sRoot & strFindData & vbBackslash
                        Else
                    
                            ' Переопределение массива если превышаем заданную размерность
                            If lngResultFolderListCount = miMaxCountArr Then
                                miMaxCountArr = 2 * miMaxCountArr
    
                                ReDim Preserve sResultFolderList(miMaxCountArr)
    
                            End If
    
                            ' Полный путь файла
                            sResultFolderList(lngResultFolderListCount).FullPath = sRoot & strFindData
                            sResultFolderList(lngResultFolderListCount).Path = sRoot
                            sResultFolderList(lngResultFolderListCount).Name = strFindData
                            sResultFolderList(lngResultFolderListCount).NameLcase = LCase$(strFindData)
                            If Not mbInitial Then
                                sResultFolderList(lngResultFolderListCount).RelativePath = Replace$(sRoot, sRootInit, vbNullString)
                            End If
                            lngResultFolderListCount = lngResultFolderListCount + 1
                            
                            If fp2.bRecurse Then
                                'if a folder, and recurse specified, call  method again
                                SearchForFolders sRoot & strFindData & vbBackslash, False, miMaxCountArr, mbDelete
                            End If
                        
                        End If
                    End If
                End If
            End If

        Loop While FindNextFile(hFile, wfd)

        FindClose hFile
    End If

    ' Переопределение массива на реальное кол-во записей
    If Not mbDelete Then
        If mbInitial Then
            If lngResultFolderListCount Then
                ReDim Preserve sResultFolderList(lngResultFolderListCount - 1)
            Else
                ReDim Preserve sResultFolderList(lngResultFolderListCount)
            End If
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FolderContainsSubfolders
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sRoot (String)
'!--------------------------------------------------------------------------------
Public Function FolderContainsSubfolders(sRoot As String) As Boolean

    Dim wfd   As WIN32_FIND_DATA
    Dim hFile As Long
    Dim lngFilePathPtr As Long

    If LenB(sRoot) Then
        sRoot = BackslashAdd2Path(sRoot)

        If PathIsValidUNC(sRoot) = False Then
            lngFilePathPtr = StrPtr("\\?\" & sRoot & ALL_FILES)
        Else
            '\\?\UNC\
            lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sRoot, Len(sRoot) - 2) & ALL_FILES)
        End If
        hFile = FindFirstFile(lngFilePathPtr, wfd)

        If hFile <> INVALID_HANDLE_VALUE Then

            Do

                If (wfd.dwFileAttributes And vbDirectory) Then

                    'an item with the vbDirectory bit was found
                    'but is it a system folder?
                    If (Left$(wfd.cFileName, 1) <> ".") And (Left$(wfd.cFileName, 2) <> "..") Then
                        'nope, it's a user folder
                        FolderContainsSubfolders = True

                        Exit Do

                    End If
                End If

            Loop While FindNextFile(hFile, wfd)

            FindClose hFile
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FolderContainsFiles
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sRoot (String)
'!--------------------------------------------------------------------------------
Public Function FolderContainsFiles(ByVal sRoot As String) As Boolean

    Dim wfd             As WIN32_FIND_DATA
    Dim hFile           As Long
    Dim lngFilePathPtr  As Long

    If LenB(sRoot) Then
        sRoot = BackslashAdd2Path(sRoot)

        If PathIsValidUNC(sRoot) = False Then
            lngFilePathPtr = StrPtr("\\?\" & sRoot & ALL_FILES)
        Else
            '\\?\UNC\
            lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sRoot, Len(sRoot) - 2) & ALL_FILES)
        End If
        hFile = FindFirstFile(lngFilePathPtr, wfd)

        If hFile <> INVALID_HANDLE_VALUE Then

            Do

                'if the vbDirectory bit's not set, it's a
                'file so we're done!
                If (Not (wfd.dwFileAttributes And vbDirectory) = vbDirectory) Then
                    FolderContainsFiles = True

                    Exit Do

                End If

            Loop While FindNextFile(hFile, wfd)

            FindClose hFile
        End If

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetDirectorySize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sRoot (String)
'                              fp (FILE_PARAMS)
'!--------------------------------------------------------------------------------
Private Sub GetDirectorySize(ByVal sRoot As String, fp As FILE_PARAMS)

    Dim wfd             As WIN32_FIND_DATA
    Dim hFile           As Long
    Dim lngFilePathPtr  As Long

    If PathIsValidUNC(sRoot) = False Then
        lngFilePathPtr = StrPtr("\\?\" & sRoot & ALL_FILES)
    Else
        '\\?\UNC\
        lngFilePathPtr = StrPtr("\\?\UNC\" & Right$(sRoot, Len(sRoot) - 2) & ALL_FILES)
    End If
    hFile = FindFirstFile(lngFilePathPtr, wfd)
    
    If hFile <> INVALID_HANDLE_VALUE Then

        Do

            If Asc(wfd.cFileName) <> vbDot Then
                If (wfd.dwFileAttributes And vbDirectory) Then
                    If fp.bRecurse Then
                        GetDirectorySize sRoot & TrimNull(wfd.cFileName) & vbBackslash, fp
                    End If

                Else
                    fp.nFileCount = fp.nFileCount + 1
                    fp.nFileSize = fp.nFileSize + ((wfd.nFileSizeHigh * (MAXDWORD + 1)) + wfd.nFileSizeLow)
                End If
            End If

            fp.nSearched = fp.nSearched + 1
        Loop While FindNextFile(hFile, wfd)

        FindClose hFile
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FolderSizeApi
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   sSource (String)
'                              bRecursion (Boolean)
'!--------------------------------------------------------------------------------
Public Function FolderSizeApi(ByVal sSource As String, ByVal bRecursion As Boolean) As String

    Dim fp As FILE_PARAMS

    With fp
        .sFileRoot = BackslashAdd2Path(sSource)
        .sFileNameExt = ALL_FILES
        .bRecurse = bRecursion
    End With

    GetDirectorySize fp.sFileRoot, fp
    FolderSizeApi = FormatByteSize(CSng(fp.nFileSize))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function FormatByteSize
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   dwBytes (Single)
'!--------------------------------------------------------------------------------
Private Function FormatByteSize(ByVal dwBytes As Single) As String

    Dim sBuff  As String
    Dim dwBuff As Long

    sBuff = String$(32, vbNullChar)
    dwBuff = Len(sBuff)

    If StrFormatByteSize(dwBytes, sBuff, dwBuff) <> 0 Then
        FormatByteSize = TrimNull(sBuff)
    End If

End Function
