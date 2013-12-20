Attribute VB_Name = "mSpecialFolder"
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
Public Enum CSIDL_VALUES
    CSIDL_DESKTOP = &H0
    CSIDL_INTERNET = &H1
    CSIDL_PROGRAMS = &H2
    CSIDL_CONTROLS = &H3
    CSIDL_PRINTERS = &H4
    CSIDL_PERSONAL = &H5
    CSIDL_FAVORITES = &H6
    CSIDL_STARTUP = &H7
    CSIDL_RECENT = &H8
    CSIDL_SENDTO = &H9
    CSIDL_BITBUCKET = &HA
    CSIDL_STARTMENU = &HB
    CSIDL_MYDOCUMENTS = &HC
    CSIDL_MYMUSIC = &HD
    CSIDL_MYVIDEO = &HE
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_DRIVES = &H11
    CSIDL_NETWORK = &H12
    CSIDL_NETHOOD = &H13
    CSIDL_FONTS = &H14
    CSIDL_TEMPLATES = &H15
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_PROGRAMS = &H17
    CSIDL_COMMON_STARTUP = &H18
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_APPDATA = &H1A
    CSIDL_PRINTHOOD = &H1B
    CSIDL_LOCAL_APPDATA = &H1C
    CSIDL_ALTSTARTUP = &H1D
    CSIDL_COMMON_ALTSTARTUP = &H1E
    CSIDL_COMMON_FAVORITES = &H1F
    CSIDL_INTERNET_CACHE = &H20
    CSIDL_COOKIES = &H21
    CSIDL_HISTORY = &H22
    CSIDL_COMMON_APPDATA = &H23
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_MYPICTURES = &H27
    CSIDL_PROFILE = &H28
    CSIDL_SYSTEMX86 = &H29
    CSIDL_PROGRAM_FILESX86 = &H2A
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
    CSIDL_COMMON_TEMPLATES = &H2D
    CSIDL_COMMON_DOCUMENTS = &H2E
    CSIDL_COMMON_ADMINTOOLS = &H2F
    CSIDL_ADMINTOOLS = &H30
    CSIDL_CONNECTIONS = &H31
    CSIDL_COMMON_MUSIC = &H35
    CSIDL_COMMON_PICTURES = &H36
    CSIDL_COMMON_VIDEO = &H37
    CSIDL_RESOURCES = &H38
    CSIDL_RESOURCES_LOCALIZED = &H39
    CSIDL_COMMON_OEM_LINKS = &H3A
    CSIDL_CDBURN_AREA = &H3B
    CSIDL_COMPUTERSNEARME = &H3D
    CSIDL_FLAG_PER_USER_INIT = &H800
    CSIDL_FLAG_NO_ALIAS = &H1000
    CSIDL_FLAG_DONT_VERIFY = &H4000
    CSIDL_FLAG_CREATE = &H8000
    CSIDL_FLAG_MASK = &HFF00
End Enum

#If False Then

    Private CSIDL_DESKTOP, CSIDL_INTERNET, CSIDL_PROGRAMS, CSIDL_CONTROLS, CSIDL_PRINTERS, CSIDL_PERSONAL, CSIDL_FAVORITES, CSIDL_STARTUP
    Private CSIDL_RECENT, CSIDL_SENDTO, CSIDL_BITBUCKET, CSIDL_STARTMENU, CSIDL_MYDOCUMENTS, CSIDL_MYMUSIC, CSIDL_MYVIDEO, CSIDL_DESKTOPDIRECTORY
    Private CSIDL_DRIVES, CSIDL_NETWORK, CSIDL_NETHOOD, CSIDL_FONTS, CSIDL_TEMPLATES, CSIDL_COMMON_STARTMENU, CSIDL_COMMON_PROGRAMS
    Private CSIDL_COMMON_STARTUP, CSIDL_COMMON_DESKTOPDIRECTORY, CSIDL_APPDATA, CSIDL_PRINTHOOD, CSIDL_LOCAL_APPDATA, CSIDL_ALTSTARTUP
    Private CSIDL_COMMON_ALTSTARTUP, CSIDL_COMMON_FAVORITES, CSIDL_INTERNET_CACHE, CSIDL_COOKIES, CSIDL_HISTORY, CSIDL_COMMON_APPDATA
    Private CSIDL_WINDOWS, CSIDL_SYSTEM, CSIDL_PROGRAM_FILES, CSIDL_MYPICTURES, CSIDL_PROFILE, CSIDL_SYSTEMX86, CSIDL_PROGRAM_FILESX86
    Private CSIDL_PROGRAM_FILES_COMMON, CSIDL_PROGRAM_FILES_COMMONX86, CSIDL_COMMON_TEMPLATES, CSIDL_COMMON_DOCUMENTS, CSIDL_COMMON_ADMINTOOLS
    Private CSIDL_ADMINTOOLS, CSIDL_CONNECTIONS, CSIDL_COMMON_MUSIC, CSIDL_COMMON_PICTURES, CSIDL_COMMON_VIDEO, CSIDL_RESOURCES
    Private CSIDL_RESOURCES_LOCALIZED, CSIDL_COMMON_OEM_LINKS, CSIDL_CDBURN_AREA, CSIDL_COMPUTERSNEARME, CSIDL_FLAG_PER_USER_INIT
    Private CSIDL_FLAG_NO_ALIAS, CSIDL_FLAG_DONT_VERIFY, CSIDL_FLAG_CREATE, CSIDL_FLAG_MASK
#End If

Private Const SHGFP_TYPE_CURRENT As Long = &H0    'current value for user, verify it exists

'Private Const SHGFP_TYPE_DEFAULT    As Long = &H1
Private Const MAX_LENGTH         As Long = 260
Private Const S_OK               As Long = 0

'Private Const S_FALSE               As Long = 1
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hWndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long
Private Declare Function GetPrinterDriverDirectory Lib "winspool.drv" Alias "GetPrinterDriverDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pDriverDirectory As String, ByVal cbBuff As Long, pcbNeeded As Long) As Long

Public Declare Function GetPrintProcessorDirectory Lib "winspool.drv" Alias "GetPrintProcessorDirectoryA" (ByVal pName As String, ByVal pEnvironment As String, ByVal Level As Long, ByVal pPrintProcessorInfo As String, ByVal cdBuf As Long, pcbNeeded As Long) As Long
Declare Function GetColorDirectory Lib "mscms" Alias "GetColorDirectoryA" (ByVal pcstr As String, ByVal pstr As String, ByRef pdword As Long) As Long

Private Declare Function GetWindowsDirectory Lib "kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Getpath_PrinterColorDirectory
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Getpath_PrinterColorDirectory() As String

    Dim WindirS As String * 255
    Dim Temp    As Long
    Dim Result  As String

    'declares a full lenght string for DIR name(for getting the path)
    'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
    'a variable for holding the the output of the function
    Temp = GetColorDirectory(vbNullString, WindirS, 255)
    Result = TrimNull(WindirS)
    Getpath_PrinterColorDirectory = Result
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Getpath_PrinterDriverDirectory
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Getpath_PrinterDriverDirectory() As String

    Dim Level            As Long
    Dim cbBuff           As Long
    Dim pcbNeeded        As Long
    Dim pName            As String
    Dim pEnvironment     As String
    Dim pDriverDirectory As String

    'initialization to determine size of buffer required
    Level = 1
    'must be 1
    cbBuff = 0
    'must be 0 initially
    pDriverDirectory = vbNullString
    'must be null string initially
    'string that specifies the name of the
    'server on which the printer driver resides.
    'If this parameter is vbNullString the
    'local driver-directory path is retrieved.
    pName = vbNullString
    'string that specifies the environment
    '(for example, "Windows NT x86", "Windows NT R4000",
    '"Windows NT Alpha_AXP", or "Windows 4.0"). If
    'this parameter is NULL, the current environment
    'of the calling application and client machine
    '(not of the destination application and print
    'server) is used.
    pEnvironment = vbNullString

    'find out how large the buffer
    'needs to be (pcbNeeded). Call will return 0.
    If GetPrinterDriverDirectory(pName, pEnvironment, Level, pDriverDirectory, cbBuff, pcbNeeded) = 0 Then
        'create a buffer large enough for the
        'string and a trailing null
        pDriverDirectory = String$(pcbNeeded, vbNullChar)
        cbBuff = Len(pDriverDirectory)

        'call again. Success = 1
        If GetPrinterDriverDirectory(pName, pEnvironment, Level, pDriverDirectory, cbBuff, pcbNeeded) = 1 Then
            Getpath_PrinterDriverDirectory = Left$(pDriverDirectory, pcbNeeded)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Getpath_PrintProcessorDirectory
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Getpath_PrintProcessorDirectory() As String

    Dim Level            As Long
    Dim cbBuff           As Long
    Dim pcbNeeded        As Long
    Dim pName            As String
    Dim pEnvironment     As String
    Dim pDriverDirectory As String

    'initialization to determine size of buffer required
    Level = 1
    'must be 1
    cbBuff = 0
    'must be 0 initially
    pDriverDirectory = vbNullString
    'must be null string initially
    'string that specifies the name of the
    'server on which the printer driver resides.
    'If this parameter is vbNullString the
    'local driver-directory path is retrieved.
    pName = vbNullString
    'string that specifies the environment
    '(for example, "Windows NT x86", "Windows NT R4000",
    '"Windows NT Alpha_AXP", or "Windows 4.0"). If
    'this parameter is NULL, the current environment
    'of the calling application and client machine
    '(not of the destination application and print
    'server) is used.
    pEnvironment = vbNullString

    'find out how large the buffer
    'needs to be (pcbNeeded). Call will return 0.
    If GetPrintProcessorDirectory(pName, pEnvironment, Level, pDriverDirectory, cbBuff, pcbNeeded) = 0 Then
        'create a buffer large enough for the
        'string and a trailing null
        pDriverDirectory = String$(pcbNeeded, vbNullChar)
        cbBuff = Len(pDriverDirectory)

        'call again. Success = 1
        If GetPrintProcessorDirectory(pName, pEnvironment, Level, pDriverDirectory, cbBuff, pcbNeeded) = 1 Then
            Getpath_PrintProcessorDirectory = Left$(pDriverDirectory, pcbNeeded)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Getpath_SYSTEM
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function Getpath_SYSTEM() As String

    Dim WindirS As String * 255
    Dim Temp    As Long
    Dim Result  As String

    'declares a full lenght string for DIR name(for getting the path)
    'a temporarry variable that holds LENGHT OF THE FINAL PATH STRING!
    'a variable for holding the the output of the function
    Temp = GetSystemDirectory(WindirS, 255)
    Result = Left$(WindirS, Temp)
    Getpath_SYSTEM = BackslashAdd2Path(Result)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetSpecialFolderPath
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   csidl (CSIDL_VALUES)
'                              SHGFP_TYPE (Long = SHGFP_TYPE_CURRENT)
'!--------------------------------------------------------------------------------
Public Function GetSpecialFolderPath(csidl As CSIDL_VALUES, Optional SHGFP_TYPE As Long = SHGFP_TYPE_CURRENT) As String

    Dim buff      As String
    Dim dwFlags   As Long
    Dim lngResult As Long

    'fill buffer with the specified folder item
    buff = String$(MAX_LENGTH, vbNullChar)
    lngResult = SHGetFolderPath(App.hInstance, csidl Or dwFlags, -1, SHGFP_TYPE, buff)

    If lngResult = S_OK Then
        GetSpecialFolderPath = TrimNull(buff)
    Else
        DebugMode "GetSpecialFolderPath: csidl=" & csidl & " SHGFP_TYPE=" & SHGFP_TYPE & " ResultCode=" & lngResult
    End If

End Function
