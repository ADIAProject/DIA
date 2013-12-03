Attribute VB_Name = "mOsVer"
Option Explicit

' Программные переменные
Public strOSArchitecture                As String
Public strOsCurrentVersion              As String
Public OsCurrVersionStruct           As OSInfoStruct

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

' Проверка процесса на 64 bit
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
Public Const PROCESSOR_ARCHITECTURE_IA64 As Long = &H6
Public Const PROCESSOR_ARCHITECTURE_INTEL As Long = 0
Public Const PROCESSOR_ARCHITECTURE_ALPHA = 2
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7

'Windows NT - constants for unicode support
Public Const VER_PLATFORM_WIN32_NT      As Long = 2
Public Const VER_NT_WORKSTATION         As Long = 1

Public Declare Function GetVersionEx _
                         Lib "kernel32.dll" _
                             Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long

Public Declare Sub GetNativeSystemInfo _
                    Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)

Public Function IsWinXPOrLater() As Boolean
    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If
    
    IsWinXPOrLater = OsCurrVersionStruct.VerFull > "5.0"
End Function

Public Function IsWinVistaOrLater() As Boolean
    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If

    IsWinVistaOrLater = OsCurrVersionStruct.VerFull >= "6.0"
End Function

'! -----------------------------------------------------------
'!  Функция     :  IsWow64
'!  Переменные  :
'!  Возвр. знач.:  As Boolean
'!  Описание    :  Проверяет является ли запущенный процесс 64-битным
'! -----------------------------------------------------------
Public Function IsWow64() As Boolean

Dim SI                                  As SYSTEM_INFO

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

'! -----------------------------------------------------------
'!  Функция     :  OSInfo
'!  Переменные  :  Nfo As Integer
'!  Возвр. знач.:  As String
'!  Описание    :  Получение расширенной информации о версии Windows
'! -----------------------------------------------------------
Public Function OSInfo() As OSInfoStruct

Dim OSVerInfo                           As OSVERSIONINFOEX
Dim OSN                                 As String

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


'! -----------------------------------------------------------
'!  Функция     :  OSInfoWMI
'!  Переменные  :  Nfo As Integer
'!  Возвр. знач.:  As String
'!  Описание    :  Получение расширенной информации о версии Windows
'! -----------------------------------------------------------
Public Function OSInfoWMI(ByVal Nfo As Long) As String

'Defining Variables
Dim objWMI                              As Object
Dim objItem                             As Object
Dim colItems                            As Object
Const wbemFlagReturnImmediately = &H10
Const wbemFlagForwardOnly = &H20

Dim strComputer                         As String
Dim strOSVersion                        As String
Dim strBuildNumber                      As String
Dim strCaption                          As String
Dim strCSDVersion                       As String

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
