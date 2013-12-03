Attribute VB_Name = "mApiFiles"
Option Explicit

Public Const vbDot                      As Integer = 46
Public Const MAX_PATH                   As Long = 260
Public Const MAX_PATH_UNICODE           As Long = 2 * MAX_PATH - 1
Public Const MAX_PATH_B                 As Long = 4000 * 2 - 1
Public Const MAXDWORD                   As Long = &HFFFFFFFF
Public Const INVALID_HANDLE_VALUE       As Integer = -1
Public Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Public Const GENERIC_WRITE              As Long = &H40000000
Public Const GENERIC_READ               As Long = &H80000000
Public Const OPEN_EXISTING              As Long = 3
Public Const vbBackslash                As String = "\"
Public Const ALL_FILES                  As String = "*.*"
Public Const ForWriting                 As Long = 2
Public Const ForAppending               As Long = 8    'Файла открыт для ДОБАВЛЕНИЯ
Public Const ForReading                 As Long = 1    'Файла открыт для ЧТЕНИЯ

Public Type SECURITY_ATTRIBUTES
    nLength                             As Long
    lpSecurityDescriptor                As Long
    bInheritHandle                      As Long

End Type

Public Type FILETIME
    dwLowDateTime                           As Long
    dwHighDateTime                      As Long

End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes                    As Long
    ftCreationTime                      As FILETIME
    ftLastAccessTime                    As FILETIME
    ftLastWriteTime                     As FILETIME
    nFileSizeHigh                       As Long
    nFileSizeLow                        As Long
    dwReserved0                         As Long
    dwReserved1                         As Long
    cFileName(MAX_PATH_B)               As Byte
    cAlternate(14 * 2 - 1)              As Byte
End Type

Public Type FILE_PARAMS
    bRecurse                            As Boolean
    nFileCount                          As Long
    nFileSize                           As Currency    '64 bit value
    nSearched                           As Long
    sFileNameExt                        As String
    sFileRoot                           As String
End Type

Public Type FOLDER_PARAMS
    bRecurse                                As Boolean
    sFileNameExt                        As String
    sFileRoot                           As String
End Type

Public Declare Function PathFileExistsW Lib "shlwapi.dll" (ByVal pszPath As Long) As Long

Public Declare Function CopyFile _
                         Lib "kernel32.dll" _
                             Alias "CopyFileA" (ByVal lpExistingFileName As String, _
                                                ByVal lpNewFileName As String, _
                                                ByVal bFailIfExists As Long) As Long

Public Declare Function DeleteFileW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long

Public Declare Function CreateDirectory _
                         Lib "kernel32.dll" _
                             Alias "CreateDirectoryA" (ByVal lpPathName As String, _
                                                       lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long

Public Declare Function RemoveDirectory _
                         Lib "kernel32.dll" _
                             Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long

Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Long

Public Declare Function PathAddBackslash _
                         Lib "shlwapi.dll" _
                             Alias "PathAddBackslashA" (ByVal Path As String) As Long

Public Declare Function PathRemoveBackslash _
                         Lib "shlwapi.dll" _
                             Alias "PathRemoveBackslashA" (ByVal Path As String) As Long

Public Declare Function MoveFile _
                         Lib "kernel32.dll" _
                             Alias "MoveFileA" (ByVal lpExistingFileName As String, _
                                                ByVal lpNewFileName As String) As Long

Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long

Public Declare Function ReadFile _
                         Lib "kernel32.dll" (ByVal hFile As Long, _
                                             lpBuffer As Any, _
                                             ByVal nNumberOfBytesToRead As Long, _
                                             lpNumberOfBytesRead As Long, _
                                             ByVal lpOverlapped As Any) As Long

Public Declare Function GetFileSize _
                         Lib "kernel32.dll" (ByVal hFile As Long, _
                                             lpFileSizeHigh As Long) As Long

Public Declare Function StrFormatByteSize Lib "shlwapi.dll" _
                                          Alias "StrFormatByteSizeA" _
                                          (ByVal dw As Long, _
                                           ByVal pszBuf As String, _
                                           ByVal cchBuf As Long) As Long

Public Declare Function StrFormatByteSizeW _
                         Lib "shlwapi.dll" (ByVal qdwLow As Long, _
                                            ByVal qdwHigh As Long, _
                                            pwszBuf As Any, _
                                            ByVal cchBuf As Long) As Long

Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long

Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long

Public Declare Function PathMatchSpec _
                         Lib "shlwapi.dll" _
                             Alias "PathMatchSpecW" (ByVal pszFileParam As Long, _
                                                     ByVal pszSpec As Long) As Long

Public Declare Function GetTempFileName _
                         Lib "kernel32.dll" _
                             Alias "GetTempFileNameA" (ByVal lpszPath As String, _
                                                       ByVal lpPrefixString As String, _
                                                       ByVal wUnique As Long, _
                                                       ByVal lpTempFileName As String) As Long

Public Declare Function PathIsUNC _
                         Lib "shlwapi.dll" _
                             Alias "PathIsUNCA" (ByVal pszPath As String) As Long

' Копирование файлов посредством shell
Public Type SHFILEOPSTRUCT
    hWnd                                    As Long
    wFunc                               As Long
    pFrom                               As String
    pTo                                 As String
    fFlags                              As Integer
    fAborted                            As Boolean
    hNameMaps                           As Long
    sProgress                           As String
End Type

Public Const FO_MOVE                    As Long = &H1
Public Const FO_COPY                    As Long = &H2
Public Const FO_DELETE                  As Long = &H3
Public Const FO_RENAME                  As Long = &H4

Public Const FOF_SILENT                 As Long = &H4
Public Const FOF_RENAMEONCOLLISION      As Long = &H8
Public Const FOF_NOCONFIRMATION         As Long = &H10
Public Const FOF_SIMPLEPROGRESS         As Long = &H100
Public Const FOF_ALLOWUNDO              As Long = &H40

Public Declare Function SHFileOperation Lib "shell32" _
                                        Alias "SHFileOperationA" _
                                        (lpFileOp As SHFILEOPSTRUCT) As Long

Public Declare Function PathCombineW Lib "shlwapi.dll" (ByVal lpszDest As Long, ByVal lpszDir As Long, ByVal lpszFile As Long) As Boolean
Public Declare Function PathIsUNCServerShare Lib "shlwapi.dll" Alias "PathIsUNCServerShareA" (ByVal pszPath As String) As Long
'Public Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCA" (ByVal pszPath As String) As Long

