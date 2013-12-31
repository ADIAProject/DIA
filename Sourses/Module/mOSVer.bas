Attribute VB_Name = "mOsVer"
Option Explicit

' Программные переменные
Public strOSArchitecture   As String
Public strOsCurrentVersion As String
Public OsCurrVersionStruct As OSInfoStruct

'Получение расширенной информации о версии Windows
Public Type OSVERSIONINFO
    dwOSVersionInfoSize                 As Long
    dwMajorVersion                      As Long
    dwMinorVersion                      As Long
    dwBuildNumber                       As Long
    dwPlatformID                        As Long
    szCSDVersion                        As String * 128
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize                 As Long
    dwMajorVersion                      As Long
    dwMinorVersion                      As Long
    dwBuildNumber                       As Long
    dwPlatformID                        As Long
    szCSDVersion                        As String * 128
    wServicePackMajor                   As Integer
    wServicePackMinor                   As Integer
    wSuiteMask                          As Integer
    wProductType                        As Byte
    wReserved                           As Byte
End Type

' Проверка процесса на 64 bit разрядность
Public Type SYSTEM_INFO
    wProcessorArchitecture              As Integer
    wReserved                           As Integer
    dwPageSize                          As Long
    lpMinimumApplicationAddress         As Long
    lpMaximumApplicationAddress         As Long
    dwActiveProcessorMask               As Long
    dwNumberOfProcessors                As Long
    dwProcessorType                     As Long
    dwAllocationGranularity             As Long
    wProcessorLevel                     As Integer
    wProcessorRevision                  As Integer
End Type

Public Type OSInfoStruct
    Name As String
    BuildNumber As String
    ServicePack As String
    VerFullwBuild As String
    VerFull As String
    VerMajor As String
    VerMinor As String
    ClientOrServer As Boolean
    IsInitialize As Boolean
End Type

Public Const PROCESSOR_ARCHITECTURE_AMD64 As Long = &H9
Public Const PROCESSOR_ARCHITECTURE_IA64  As Long = &H6
Public Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7

'Windows NT - constants for unicode support
Public Const VER_PLATFORM_WIN32_NT          As Long = 2
Public Const VER_NT_WORKSTATION             As Long = 1

Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinXPOrLater
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinXPOrLater() As Boolean

    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If

    IsWinXPOrLater = OsCurrVersionStruct.VerFull > "5.0"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinVistaOrLater
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinVistaOrLater() As Boolean

    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If

    IsWinVistaOrLater = OsCurrVersionStruct.VerFull >= "6.0"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWow64
'! Description (Описание)  :   [Проверяет является ли запущенный процесс 64-битным]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWow64() As Boolean

    Dim SI As SYSTEM_INFO

    strOSArchitecture = "x86"

    If APIFunctionPresent("GetNativeSystemInfo", "kernel32.dll") Then
        GetNativeSystemInfo SI

        Select Case SI.wProcessorArchitecture

            Case PROCESSOR_ARCHITECTURE_IA64
                IsWow64 = True
                strOSArchitecture = "ia64"

            Case PROCESSOR_ARCHITECTURE_AMD64
                IsWow64 = True
                strOSArchitecture = "amd64"

            Case Else
                IsWow64 = False
        End Select

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function OSInfo
'! Description (Описание)  :   [Получение расширенной информации о версии Windows]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function OSInfo() As OSInfoStruct

    Dim OSVerInfo As OSVERSIONINFOEX
    Dim OSN       As String

    OSVerInfo.dwOSVersionInfoSize = Len(OSVerInfo)

    If GetVersionEx(OSVerInfo) <> 0 Then

        With OSVerInfo
            'Имя операционной системы
            OSN = "UnSupported\Unknown"

            If .dwMajorVersion = 5 Then
                If .dwMinorVersion = 0 Then
                    OSN = "2000"
                ElseIf .dwMinorVersion = 1 Then
                    OSN = "XP"
                ElseIf .dwMinorVersion = 2 Then
                    OSN = "Server 2003"
                End If

            ElseIf .dwMajorVersion = 6 Then

                If .dwMinorVersion = 0 Then
                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "Vista"
                    Else
                        OSN = "Server 2008"
                    End If

                ElseIf .dwMinorVersion = 1 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "7"
                    Else
                        OSN = "Server 2008 R2"
                    End If

                ElseIf .dwMinorVersion = 2 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "8"
                    Else
                        OSN = "Server 2012"
                    End If

                ElseIf .dwMinorVersion = 3 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "8.1"
                    Else
                        OSN = "Server 2012 R2"
                    End If

                Else
                    OSN = "9 ?"
                End If

            Else
                OSN = "9 ?"
            End If

            OSInfo.Name = "Windows " & OSN
            OSInfo.BuildNumber = .dwBuildNumber
            OSInfo.ServicePack = TrimNull(.szCSDVersion)
            OSInfo.VerFullwBuild = .dwMajorVersion & "." & .dwMinorVersion & "." & .dwBuildNumber
            OSInfo.VerFull = .dwMajorVersion & "." & .dwMinorVersion
            OSInfo.VerMajor = .dwMajorVersion
            OSInfo.VerMinor = .dwMinorVersion
            OSInfo.ClientOrServer = .wProductType = VER_NT_WORKSTATION
            OSInfo.IsInitialize = True
        End With

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function OSInfoWMI
'! Description (Описание)  :   [Получение расширенной информации о версии Windows, альтернативная функция, использует WMI]
'! Parameters  (Переменные):   Nfo (Long)
'!--------------------------------------------------------------------------------
Public Function OSInfoWMI(ByVal Nfo As Long) As String

    'Defining Variables
    Dim objWMI   As Object
    Dim objItem  As Object
    Dim colItems As Object

    Const wbemFlagReturnImmediately = &H10
    Const wbemFlagForwardOnly = &H20

    Dim strComputer    As String
    Dim strOSVersion   As String
    Dim strBuildNumber As String
    Dim strCaption     As String
    Dim strCSDVersion  As String

    strComputer = "."
    '   Get the WMI object and query results
    Set objWMI = CreateObject("winmgmts:\\" & strComputer & "\root\CIMV2")
    Set colItems = objWMI.ExecQuery("Select * from Win32_OperatingSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

    'Get the OS version number (first two) and OS product type (server or desktop)
    For Each objItem In colItems

        If Not IsNull(objItem.Version) Then
            strOSVersion = CStr(objItem.Version)
        End If

        If Not IsNull(objItem.BuildNumber) Then
            strBuildNumber = CStr(objItem.BuildNumber)
        End If

        If Not IsNull(objItem.Caption) Then
            strCaption = CStr(objItem.Caption)
        End If

        If Not IsNull(objItem.CSDVersion) Then
            strCSDVersion = CStr(objItem.CSDVersion)
        End If

    Next

    Select Case Nfo

        Case 0
            OSInfoWMI = strCaption

        Case 1
            OSInfoWMI = strBuildNumber

        Case 2
            OSInfoWMI = TrimNull(strCSDVersion)

        Case 3
            OSInfoWMI = strOSVersion

        Case 4
            OSInfoWMI = Left$(strOSVersion, 3)

        Case 5
            OSInfoWMI = Left$(strOSVersion, 1)

        Case 6
            OSInfoWMI = Mid$(strOSVersion, 3, 1)

        Case Else
            OSInfoWMI = "ERROR!"
    End Select

    'Clear the memory
    Set colItems = Nothing
    Set objWMI = Nothing
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_Manufacturer
'! Description (Описание)  :   [Получение производителя материнской платы, используется WMI]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_Manufacturer() As String

    Dim colItems           As Object
    Dim objItem            As Object
    Dim objWMIService      As Object
    Dim sAnsComputerSystem As String
    Dim sAnsBaseBoard      As String
    Dim objRegExp          As RegExp
    Dim strTemp            As String

    Const wbemFlagReturnImmediately = &H10
    Const wbemFlagForwardOnly = &H20

    ' получение данных из Win32_ComputerSystem - чаще всего есть если Ноутбук
    Set objWMIService = CreateObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

    For Each objItem In colItems

        sAnsComputerSystem = sAnsComputerSystem & objItem.Manufacturer
    Next

    ' получение данных из Win32_ComputerSystem
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BaseBoard", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

    For Each objItem In colItems

        sAnsBaseBoard = sAnsBaseBoard & objItem.Manufacturer
    Next

    ' итог
    If StrComp(sAnsComputerSystem, "System manufacturer", vbTextCompare) = 0 Then
        strTemp = Trim$(sAnsBaseBoard)
        mbIsNotebok = False
    Else
        strTemp = Trim$(sAnsComputerSystem)
        mbIsNotebok = True
    End If

    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
    End With

    'получаем date1
    GetMB_Manufacturer = Trim$(objRegExp.Replace(strTemp, " "))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_Model
'! Description (Описание)  :   [Получение модели материнской платы, используется WMI]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_Model() As String

    Dim colItems           As Object
    Dim objItem            As Object
    Dim objWMIService      As Object
    Dim sAnsComputerSystem As String
    Dim sAnsBaseBoard      As String
    Dim objRegExp          As RegExp
    Dim strTemp            As String

    Const wbemFlagReturnImmediately = &H10
    Const wbemFlagForwardOnly = &H20

    ' получение данных из Win32_ComputerSystem - чаще всего есть если Ноутбук
    Set objWMIService = CreateObject("winmgmts:\\.\root\CIMV2")
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_ComputerSystem", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

    For Each objItem In colItems

        sAnsComputerSystem = sAnsComputerSystem & objItem.Model
    Next

    ' получение данных из Win32_ComputerSystem
    Set colItems = objWMIService.ExecQuery("SELECT * FROM Win32_BaseBoard", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)

    For Each objItem In colItems

        sAnsBaseBoard = sAnsBaseBoard & objItem.Product
    Next

    ' итог
    If StrComp(sAnsComputerSystem, "System Product Name", vbTextCompare) = 0 Then
        strTemp = Trim$(sAnsBaseBoard)
        mbIsNotebok = False
    Else
        strTemp = Trim$(sAnsComputerSystem)
        mbIsNotebok = True
    End If

    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
    End With

    'получаем date1
    GetMB_Model = Trim$(objRegExp.Replace(strTemp, " "))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMBInfo
'! Description (Описание)  :   [Итоговая строка производитель/модель материнской платы/ноутбука]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMBInfo() As String

    Dim strMB_Manufacturer As String
    Dim strMB_Model        As String

    strMB_Manufacturer = GetMB_Manufacturer()
    strMB_Model = GetMB_Model()

    If LenB(strMB_Manufacturer) > 0 And LenB(strMB_Model) > 0 Then
        GetMBInfo = strMB_Manufacturer & "_" & strMB_Model
    ElseIf LenB(strMB_Manufacturer) = 0 And LenB(strMB_Model) > 0 Then
        GetMBInfo = strMB_Model
    ElseIf LenB(strMB_Manufacturer) > 0 And LenB(strMB_Model) = 0 Then
        GetMBInfo = strMB_Manufacturer
    Else
        GetMBInfo = "Unknown"
    End If

End Function

