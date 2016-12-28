Attribute VB_Name = "mApiFiles"
Option Explicit

Public Const MAX_PATH               As Long = 260
Public Const MAX_PATH_UNICODE       As Long = 2 * MAX_PATH - 1
Public Const MAX_PATH_B             As Long = 4000 * 2 - 1
Public Const MAXDWORD               As Long = &HFFFFFFFF

'Константы для CreateFile
Public Const INVALID_HANDLE_VALUE       As Long = (-1)
Public Const ERROR_CALL_NOT_IMPLEMENTED As Long = 120
Public Const FILE_ATTRIBUTE_NORMAL      As Long = &H80
Public Const FILE_FLAG_SEQUENTIAL_SCAN  As Long = &H8000000
Public Const GENERIC_WRITE              As Long = &H40000000
Public Const GENERIC_READ               As Long = &H80000000
Public Const OPEN_EXISTING              As Long = 3
Public Const CREATE_ALWAYS              As Long = 2
Public Const FILE_SHARE_READ            As Long = &H1
Public Const FILE_SHARE_WRITE           As Long = &H2
Public Const FILE_SHARE_DELETE          As Long = &H4

'Константы для FindFileEx
Public Const FIND_FIRST_EX_LARGE_FETCH   As Long = 2 'Win7??

'Add to declaration section
Public Enum FINDEX_INFO_LEVELS
  FindExInfoStandard = 0&
  FindExInfoBasic = 1&      'supported in W7 and newer
  FindExInfoMaxInfoLevel = 2&
End Enum

Public Enum FINDEX_SEARCH_OPS
    FindExSearchNameMatch = 0&
    FindExSearchLimitToDirectories = 1&
    FindExSearchLimitToDevices = 2&
    FindExSearchMaxSearchOp = 3&
End Enum


'Строковые константы частоиспользуемых функций
Public Const vbBackslash            As String = "\"
Public Const vbBackslashDouble      As String = "\\"
Public Const ALL_FILES              As String = "*.*"
Public Const ALL_FOLDERS            As String = "*."
Public Const ALL_FOLDERS_EX         As String = "*"

Public Type SECURITY_ATTRIBUTES
    nLength                         As Long
    lpSecurityDescriptor            As Long
    bInheritHandle                  As Long
End Type

Public Type FILETIME
    dwLowDateTime                   As Long
    dwHighDateTime                  As Long
End Type

Public Type WIN32_FIND_DATA
    dwFileAttributes                As Long
    ftCreationTime                  As FILETIME
    ftLastAccessTime                As FILETIME
    ftLastWriteTime                 As FILETIME
    nFileSizeHigh                   As Long
    nFileSizeLow                    As Long
    dwReserved0                     As Long
    dwReserved1                     As Long
    cFileName(MAX_PATH_B)           As Byte
    cAlternate(14 * 2 - 1)          As Byte
End Type

Public Declare Function PathFileExists Lib "shlwapi.dll" Alias "PathFileExistsW" (ByVal pszPath As Long) As Boolean
Public Declare Function CopyFile Lib "kernel32.dll" Alias "CopyFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal bFailIfExists As Long) As Long
Public Declare Function DeleteFile Lib "kernel32.dll" Alias "DeleteFileW" (ByVal lpFileName As Long) As Long
Public Declare Function CreateDirectory Lib "kernel32.dll" Alias "CreateDirectoryA" (ByVal lpPathName As String, lpSecurityAttributes As SECURITY_ATTRIBUTES) As Long
Public Declare Function RemoveDirectory Lib "kernel32.dll" Alias "RemoveDirectoryW" (ByVal lpPathName As Long) As Long
Public Declare Function PathIsDirectory Lib "shlwapi.dll" Alias "PathIsDirectoryW" (ByVal pszPath As Long) As Boolean
Public Declare Function PathAddBackslash Lib "shlwapi.dll" Alias "PathAddBackslashA" (ByVal Path As String) As Long
Public Declare Function PathRemoveBackslash Lib "shlwapi.dll" Alias "PathRemoveBackslashA" (ByVal Path As String) As Long
Public Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function CreateFile Lib "kernel32.dll" Alias "CreateFileW" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function WriteFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToWrite As Long, ByRef NumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByVal lpBuffer As Long, ByVal NumberOfBytesToRead As Long, ByRef NumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function FlushFileBuffers Lib "kernel32.dll" (ByVal hFile As Long) As Long
Public Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long
Public Declare Function StrFormatByteSize Lib "shlwapi.dll" Alias "StrFormatByteSizeA" (ByVal dw As Single, ByVal pszBuf As String, ByVal cchBuf As Long) As Long
Public Declare Function StrFormatByteSizeW Lib "shlwapi.dll" (ByVal qdwLow As Long, ByVal qdwHigh As Long, pwszBuf As Any, ByVal cchBuf As Long) As Long
Public Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
Public Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileW" (ByVal lpFileName As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function FindFirstFileEx Lib "kernel32.dll" Alias "FindFirstFileExW" (ByVal lpFileName As Long, ByVal fInfoLevelId As Long, lpFindFileData As WIN32_FIND_DATA, ByVal fSearchOp As Long, ByVal lpSearchFilter As Long, ByVal dwAdditionalFlags As Long) As Long
Public Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileW" (ByVal hFindFile As Long, lpFindFileData As WIN32_FIND_DATA) As Long
Public Declare Function PathMatchSpec Lib "shlwapi.dll" Alias "PathMatchSpecW" (ByVal pszFileParam As Long, ByVal pszSpec As Long) As Long
Public Declare Function GetTempFileName Lib "kernel32.dll" Alias "GetTempFileNameA" (ByVal lpszPath As String, ByVal lpPrefixString As String, ByVal wUnique As Long, ByVal lpTempFileName As String) As Long
Public Declare Function PathIsUNC Lib "shlwapi.dll" Alias "PathIsUNCW" (ByVal pszPath As Long) As Boolean
Public Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Public Declare Function PathCombineW Lib "shlwapi.dll" (ByVal lpszDest As Long, ByVal lpszDir As Long, ByVal lpszFile As Long) As Boolean
Public Declare Function PathIsRoot Lib "shlwapi.dll" Alias "PathIsRootA" (ByVal pszPath As String) As Long
Public Declare Function GetDriveType Lib "kernel32.dll" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
Public Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Public Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Public Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Any, puLen As Long) As Long

'File info
Public Type VS_FIXEDFILEINFO
    dwSignature As Long
    dwStrucVersionl As Integer     '  e.g. = &h0000 = 0
    dwStrucVersionh As Integer     '  e.g. = &h0042 = .42
    dwFileVersionMSl As Integer    '  e.g. = &h0003 = 3
    dwFileVersionMSh As Integer    '  e.g. = &h0075 = .75
    dwFileVersionLSl As Integer    '  e.g. = &h0000 = 0
    dwFileVersionLSh As Integer    '  e.g. = &h0031 = .31
    dwProductVersionMSl As Integer '  e.g. = &h0003 = 3
    dwProductVersionMSh As Integer '  e.g. = &h0010 = .1
    dwProductVersionLSl As Integer '  e.g. = &h0000 = 0
    dwProductVersionLSh As Integer '  e.g. = &h0031 = .31
    dwFileFlagsMask As Long        '  = &h3F for version "0.42"
    dwFileFlags As Long            '  e.g. VFF_DEBUG Or VFF_PRERELEASE
    dwFileOS As Long               '  e.g. VOS_DOS_WINDOWS16
    dwFileType As Long             '  e.g. VFT_DRIVER
    dwFileSubtype As Long          '  e.g. VFT2_DRV_KEYBOARD
    dwFileDateMS As Long           '  e.g. 0
    dwFileDateLS As Long           '  e.g. 0
End Type

' Структура для копирование файлов посредством shell
Public Type SHFILEOPSTRUCT
    hWnd                            As Long
    wFunc                           As Long
    pFrom                           As String
    pTo                             As String
    fFlags                          As Integer
    fAborted                        As Boolean
    hNameMaps                       As Long
    sProgress                       As String
End Type

' Константы для копирования файлов посредством shell
Public Const FO_MOVE                As Long = &H1
Public Const FO_COPY                As Long = &H2
Public Const FO_DELETE              As Long = &H3
Public Const FO_RENAME              As Long = &H4
Public Const FOF_SILENT             As Long = &H4
Public Const FOF_RENAMEONCOLLISION  As Long = &H8
Public Const FOF_NOCONFIRMATION     As Long = &H10
Public Const FOF_SIMPLEPROGRESS     As Long = &H100
Public Const FOF_ALLOWUNDO          As Long = &H40

'Константы типов дисков
Public Const DRIVE_UNKNOWN      As Long = 0 'The drive type cannot be determined.
Public Const DRIVE_NO_ROOT_DIR  As Long = 1 'The root path is invalid; for example, there is no volume mounted at the specified path.
Public Const DRIVE_REMOVABLE    As Long = 2 'The drive has removable media; for example, a floppy drive, thumb drive, or flash card reader.
Public Const DRIVE_FIXED        As Long = 3 'The drive has fixed media; for example, a hard disk drive or flash drive.
Public Const DRIVE_REMOTE       As Long = 4 'The drive is a remote (network) drive.
Public Const DRIVE_CDROM        As Long = 5 'The drive is a CD-ROM drive.
Public Const DRIVE_RAMDISK      As Long = 6 'The drive is a RAM disk.

'Константы для FSO (Scripting.FileSystemObject)
Public Const ForWriting    As Long = 2
Public Const ForAppending  As Long = 8
Public Const ForReading    As Long = 1
