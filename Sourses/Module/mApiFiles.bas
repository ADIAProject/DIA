Attribute VB_Name = "mApiFiles"
Option Explicit

Public Const MAX_PATH              As Long = 260
Public Const MAX_PATH_UNICODE      As Long = 2 * MAX_PATH - 1
Public Const MAX_PATH_B            As Long = 4000 * 2 - 1
Public Const MAXDWORD              As Long = &HFFFFFFFF

'Константы для CreateFile
Public Const INVALID_HANDLE_VALUE  As Integer = -1
Public Const FILE_ATTRIBUTE_NORMAL As Long = &H80
Public Const GENERIC_WRITE         As Long = &H40000000
Public Const GENERIC_READ          As Long = &H80000000
Public Const OPEN_EXISTING         As Long = 3

'Строковые константы частоиспользуемых функций
Public Const vbBackslash           As String = "\"
Public Const vbBackslashUNC        As String = "\\"
Public Const vbBackslashDouble     As String = "\\"
Public Const ALL_FILES             As String = "*.*"

Public Type SECURITY_ATTRIBUTES
    nLength                             As Long
    lpSecurityDescriptor                As Long
    bInheritHandle                      As Long
End Type

Public Type FILETIME
    dwLowDateTime                       As Long
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
    bRecurse                            As Boolean
    sFileNameExt                        As String
    sFileRoot                           As String
End Type

'Api-Функции для работы с файлами
Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Boolean
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Boolean
Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
Public Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal Path As String) As Long
Public Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Any) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function StrFormatByteSize Lib "shlwapi.dll" Alias "StrFormatByteSizeA" (ByVal dw As Long, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
Public Declare Function StrFormatByteSizeW Lib "shlwapi.dll" (ByVal qdwLow As Long, ByVal qdwHigh As Long, pwszBuf As Any, ByVal cchBuf As Long) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecW" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCW" (ByVal pszPath As Long) As Boolean
Public Declare Function SHFileOperation Lib "shell32" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function PathCombineW Lib "shlwapi.dll" (ByVal lpszDest As Long, ByVal lpszDir As Long, ByVal lpszFile As Long) As Boolean

' Структура для копирование файлов посредством shell
Public Type SHFILEOPSTRUCT
    hWnd                                As Long
    wFunc                               As Long
    pFrom                               As String
    pTo                                 As String
    fFlags                              As Integer
    fAborted                            As Boolean
    hNameMaps                           As Long
    sProgress                           As String
End Type

' Константы для копирования файлов посредством shell
Public Const FO_MOVE               As Long = &H1
Public Const FO_COPY               As Long = &H2
Public Const FO_DELETE             As Long = &H3
Public Const FO_RENAME             As Long = &H4
Public Const FOF_SILENT            As Long = &H4
Public Const FOF_RENAMEONCOLLISION As Long = &H8
Public Const FOF_NOCONFIRMATION    As Long = &H10
Public Const FOF_SIMPLEPROGRESS    As Long = &H100
Public Const FOF_ALLOWUNDO         As Long = &H40
