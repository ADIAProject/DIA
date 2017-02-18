Attribute VB_Name = "mSysInfo"
' Получения подробной информации о версии операционной системы,
' а также модели компьтера/ноутбука/материнской платы
Option Explicit

' Not add to project (DBS/DIA) - option for compiler
#Const mbIDE_DBSProject = False

' Программные переменные
Public strOSArchitecture        As String        ' Архитетуктура ОС
Public strOSCurrentVersion      As String
Public OSCurrVersionStruct      As OSInfoStruct
Public mbIsWin64                As Boolean
Public mbIsNotebook             As Boolean
    
Public Type OSInfoStruct
    Name                        As String
    BuildNumber                 As String
    ServicePack                 As String
    VerFullwBuild               As String
    VerFull                     As String
    VerMajor                    As String
    VerMinor                    As String
    ClientOrServer              As Boolean
    IsInitialize                As Boolean
End Type

' API-Declared
Public Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformID                As Long
    szCSDVersion                As String * 128
End Type

Public Type OSVERSIONINFOEX
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformID                As Long
    szCSDVersion                As String * 128
    wServicePackMajor           As Integer
    wServicePackMinor           As Integer
    wSuiteMask                  As Integer
    wProductType                As Byte
    wReserved                   As Byte
End Type

' Проверка процесса на 64 bit разрядность
Public Type SYSTEM_INFO
    wProcessorArchitecture      As Integer
    wReserved                   As Integer
    dwPageSize                  As Long
    lpMinimumApplicationAddress As Long
    lpMaximumApplicationAddress As Long
    dwActiveProcessorMask       As Long
    dwNumberOfProcessors        As Long
    dwProcessorType             As Long
    dwAllocationGranularity     As Long
    wProcessorLevel             As Integer
    wProcessorRevision          As Integer
End Type

Public Const PROCESSOR_ARCHITECTURE_AMD64   As Long = &H9
Public Const PROCESSOR_ARCHITECTURE_IA64    As Long = &H6
Public Const PROCESSOR_ARCHITECTURE_INTEL   As Long = 0
Public Const PROCESSOR_ARCHITECTURE_ALPHA   As Long = 2
Public Const PROCESSOR_ARCHITECTURE_ALPHA64 As Long = 7
Public Const VER_PLATFORM_WIN32_NT          As Long = 2
Public Const VER_NT_WORKSTATION             As Long = 1

' Battery status
Private Type SYSTEM_POWER_STATUS
        ACLineStatus As Byte
        BatteryFlag As Byte
        BatteryLifePercent As Byte
        Reserved1 As Byte
        BatteryLifeTime As Long
        BatteryFullLifeTime As Long
End Type

Private Declare Function GetSystemPowerStatus Lib "kernel32" (lpSystemPowerStatus As SYSTEM_POWER_STATUS) As Long

Public Declare Function GetVersionEx Lib "kernel32.dll" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFOEX) As Long
Public Declare Sub GetNativeSystemInfo Lib "kernel32.dll" (ByRef lpSystemInfo As SYSTEM_INFO)
Public Declare Function IsUserAnAdmin Lib "shell32.dll" () As Long

Private Declare Function IsWow64Process Lib "kernel32" (ByVal hProcess As Long, ByRef Wow64Process As Long) As Long
Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, lpFreeBytesAvailableToCaller As Currency, lpTotalNumberOfBytes As Currency, lpTotalNumberOfFreeBytes As Currency) As Long

' These functions are for getting the process token information, which IsUserAnAdministrator uses to
' handle detecting an administrator that’s running in a non-elevated process under UAC.
Private Const TOKEN_READ As Long = &H20008
Private Const TOKEN_ELEVATION_TYPE As Long = 18
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   GetFileName4Snap
'! Description (Описание)  :   [Получение имени файла для снимка системы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetFileName4Snap() As String
Attribute GetFileName4Snap.VB_UserMemId = 1610612751
    If mbIsNotebook Then
        If Not OSCurrVersionStruct.ClientOrServer Then
            GetFileName4Snap = ExpandFileNameByEnvironment("hwids_%PCMODEL%-Notebook_" & strOSCurrentVersion & "-Server_%OSBIT%" & "_%DATE%")
        Else
            GetFileName4Snap = ExpandFileNameByEnvironment("hwids_%PCMODEL%-Notebook_" & strOSCurrentVersion & "_%OSBIT%" & "_%DATE%")
        End If
    Else
        If Not OSCurrVersionStruct.ClientOrServer Then
            GetFileName4Snap = ExpandFileNameByEnvironment("hwids_%PCMODEL%_" & strOSCurrentVersion & "-Server_%OSBIT%" & "_%DATE%")
        Else
            GetFileName4Snap = ExpandFileNameByEnvironment("hwids_%PCMODEL%_" & strOSCurrentVersion & "_%OSBIT%" & "_%DATE%")
        End If
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMBInfo
'! Description (Описание)  :   [Итоговая строка производитель/модель материнской платы/ноутбука]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMBInfo() As String
Attribute GetMBInfo.VB_UserMemId = 1610612739

    Dim strMB_Manufacturer As String
    Dim strMB_Model        As String
    Dim mbMB_Model         As Boolean
    Dim mbMB_Manufacturer  As Boolean

    ' Получаем производителя PC, через реестр, если невозможно, то через WMI
    strMB_Manufacturer = GetMB_ManufacturerReg
    If LenB(strMB_Manufacturer) = 0 Then
        strMB_Manufacturer = GetMB_ManufacturerWMI
    End If
    
    ' Получаем модель PC
    strMB_Model = GetMB_ModelReg
    If LenB(strMB_Model) = 0 Then
        strMB_Model = GetMB_ModelWMI
    End If
    
    mbMB_Model = LenB(strMB_Model)
    mbMB_Manufacturer = LenB(strMB_Manufacturer)

    If mbMB_Manufacturer Then
        If mbMB_Model Then
            GetMBInfo = strMB_Manufacturer & strDash & strMB_Model
        Else
            GetMBInfo = strMB_Manufacturer
        End If
    Else
        If mbMB_Model Then
            GetMBInfo = strMB_Model
        Else
            GetMBInfo = "Unknown"
        End If
    End If
    
    If InStr(GetMBInfo, "_") Then
        GetMBInfo = Replace$(GetMBInfo, "_", strDash)
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_ManufacturerWMI
'! Description (Описание)  :   [Получение производителя материнской платы, используется WMI]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_ManufacturerWMI() As String

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
    If Err.Number = 429 Then Exit Function
    
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
    If StrComp(LCase$(sAnsComputerSystem), "system manufacturer") = 0 Then
        strTemp = Trim$(sAnsBaseBoard)
        mbIsNotebook = False
    Else
        strTemp = Trim$(sAnsComputerSystem)
        mbIsNotebook = True
    End If

    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(CO., LTD.)|(ELECTRONICS)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
        'Заменяем найденные значения " "
        GetMB_ManufacturerWMI = Trim$(.Replace(strTemp, strSpace))
    End With
    
    Set objRegExp = Nothing

End Function

'How define PC manufacturing from Reestr
'[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SystemInformation]
'"SystemManufacturer"="Hewlett-Packard"
'"SystemProductName"="HP Compaq Pro 6300 MT"
'[HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS]
'"BaseBoardManufacturer"="ASUSTeK COMPUTER INC."
'"BaseBoardProduct"="P8H67-M LX"
'"SystemManufacturer"="System manufacturer"
'"SystemProductName"="System Product Name"
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_ManufacturerReg
'! Description (Описание)  :   [Получение производителя материнской платы, используется реестр]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_ManufacturerReg() As String

    Dim regParam      As String
    Dim objRegExp     As RegExp
    Dim strTemp       As String

    regParam = GetRegString(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SystemInformation", "SystemManufacturer")
    
    ' итог
    If StrComp(LCase$(regParam), "system manufacturer") = 0 Then
        ' Если стандартное значение, то определяем по другой веткее реестра
        regParam = GetRegString(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardManufacturer")
        strTemp = Trim$(regParam)
        mbIsNotebook = False
    Else
        strTemp = Trim$(regParam)
        mbIsNotebook = True
    End If

    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(CO., LTD.)|(ELECTRONICS)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
        'Заменяем найденные значения " "
        GetMB_ManufacturerReg = Trim$(.Replace(strTemp, strSpace))
    End With
    
    Set objRegExp = Nothing

End Function

'!-------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_ModelWMI
'! Description (Описание)  :   [Получение модели материнской платы, используется WMI]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_ModelWMI() As String

    Dim colItems           As Object
    Dim objItem            As Object
    Dim objWMIService      As Object
    Dim sAnsComputerSystem As String
    Dim sAnsBaseBoard      As String
    Dim objRegExp          As RegExp
    Dim strTemp            As String
    Dim objChassisType     As Variant

    Const wbemFlagReturnImmediately = &H10
    Const wbemFlagForwardOnly = &H20

    ' получение данных из Win32_ComputerSystem - чаще всего есть если Ноутбук
    Set objWMIService = CreateObject("winmgmts:\\.\root\CIMV2")
    If Err.Number = 429 Then Exit Function
    
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
    If StrComp(LCase$(sAnsComputerSystem), "system product name") = 0 Then
        strTemp = Trim$(sAnsBaseBoard)
        mbIsNotebook = False
    Else
        strTemp = Trim$(sAnsComputerSystem)
        mbIsNotebook = True
    End If
    Set colItems = Nothing
    Set objItem = Nothing
    
    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(ELECTRONICS)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
        'Заменяем найденные значения " "
        GetMB_ModelWMI = Trim$(.Replace(strTemp, strSpace))
    End With
    
    Set objRegExp = Nothing

End Function

' How define PC manufacturing from Reestr
'[HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Control\SystemInformation]
'"SystemManufacturer"="Hewlett-Packard"
'"SystemProductName"="HP Compaq Pro 6300 MT"
'[HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\System\BIOS]
'"BaseBoardManufacturer"="ASUSTeK COMPUTER INC."
'"BaseBoardProduct"="P8H67-M LX"
'"SystemManufacturer"="System manufacturer"
'"SystemProductName"="System Product Name"
'!-------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetMB_ModelReg
'! Description (Описание)  :   [Получение модели материнской платы, используется реестр]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetMB_ModelReg() As String

    Dim regParam      As String
    Dim objRegExp     As RegExp
    Dim strTemp       As String

    regParam = GetRegString(HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\SystemInformation", "SystemProductName")
    
    ' итог
    If StrComp(LCase$(regParam), "system product name") = 0 Then
        ' Если стандартное значение, то определяем по другой веткее реестра
        regParam = GetRegString(HKEY_LOCAL_MACHINE, "HARDWARE\DESCRIPTION\System\BIOS", "BaseBoardProduct")
        strTemp = Trim$(regParam)
        mbIsNotebook = False
    Else
        strTemp = Trim$(regParam)
        mbIsNotebook = True
    End If

    ' удаляем лишние символы в наименовании
    Set objRegExp = New RegExp

    With objRegExp
        .Pattern = "/(, inc.)|(inc.)|(corporation)|(corp.)|(computer)|(co., ltd.)|(co., ltd)|(co.,ltd)|(co.)|(ltd)|(international)|(CO., LTD.)|(ELECTRONICS)|(Technology)/ig"
        .IgnoreCase = True
        .Global = True
        'Заменяем найденные значения " "
        GetMB_ModelReg = Trim$(.Replace(strTemp, strSpace))
    End With
    
    Set objRegExp = Nothing

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   GetSystemDiskFreeSpace
'! Description (Описание)  :   [Определения свободного места на системном диске]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function GetSystemDiskFreeSpace(ByVal strDrive As String) As Long
Attribute GetSystemDiskFreeSpace.VB_UserMemId = 1610612752
Dim BytesFreeToCalller  As Currency
Dim TotalBytes          As Currency
Dim TotalFreeBytes      As Currency
Dim TotalBytesUsed      As Currency

    If LenB(strDrive) Then
        GetDiskFreeSpaceEx strDrive, BytesFreeToCalller, TotalBytes, TotalFreeBytes
        GetSystemDiskFreeSpace = (TotalFreeBytes * 10000 / 1024 / 1024)
    End If
    
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   GetSystemDiskFreeSpace
'! Description (Описание)  :   [уточнение про "статус" компьютер-ноутбук по корпусу или батарее]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub IsPCisNotebook()
    
    If mbIsNotebook Then
        
        Dim colItems           As Object
        Dim objItem            As Object
        Dim objWMIService      As Object
        Dim objChassisType     As Variant
    
        Const wbemFlagReturnImmediately = &H10
        Const wbemFlagForwardOnly = &H20
    
        ' получение данных из Win32_SystemEnclosure - тип корпуса
        Set objWMIService = CreateObject("winmgmts:\\.\root\CIMV2")
        If Err.Number = 429 Then Exit Sub
        
        Set colItems = objWMIService.ExecQuery("Select * from Win32_SystemEnclosure", "WQL", wbemFlagReturnImmediately + wbemFlagForwardOnly)
        
        For Each objItem In colItems
            For Each objChassisType In objItem.ChassisTypes
                Select Case objChassisType
                    Case 3
                        mbIsNotebook = False
                        GoTo ExitSub
                    Case 10
                        mbIsNotebook = True
                        GoTo ExitSub
                End Select
            Next
        Next

        ' Not add to project (if not DBS) - option for compile
        #If Not mbIDE_DBSProject Then
            'Если не определили по типу корпусу, то определяем по батарее
            Dim BatteryStatus As SYSTEM_POWER_STATUS
            Dim ii            As Long
            Dim mbBatDev      As Boolean
            
            For ii = 0 To UBound(arrHwidsLocal)
                If InStr(arrHwidsLocal(ii).HWID, "ACPI0003") Then
                    mbBatDev = True
                    Exit For
                End If
            Next ii
            
            'Get status system battery
            GetSystemPowerStatus BatteryStatus
            
            'Not (No system battery) or (battery device is exist)
            If BatteryStatus.BatteryFlag < 128 Or mbBatDev = True Then
                mbIsNotebook = True
            Else
                mbIsNotebook = False
            End If
        #End If

ExitSub:
        Set colItems = Nothing
        Set objItem = Nothing
        
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   IsUserAnAdministrator
'! Description (Описание)  :   [Пользователь является администратором]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsUserAnAdministrator() As Boolean
Attribute IsUserAnAdministrator.VB_UserMemId = 1610612753
    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If
    
    If OSCurrVersionStruct.VerMajor = 5 Then
        If IsUserAnAdmin() Then
            IsUserAnAdministrator = True
        End If
    Else
        ' If we’re on Vista onwards, check for UAC elevation token
        ' as we may be an admin but we’re not elevated yet, so the
        ' IsUserAnAdmin() function will return false
        If OSCurrVersionStruct.VerMajor >= 6 Then
            Dim Result As Long
            Dim hProcessID As Long
            Dim hToken As Long
            Dim lReturnLength As Long
            Dim tokenElevationType As Long
            
            ' We need to get the token for the current process
            'hProcessID = GetCurrentProcess()
            hProcessID = App.hInstance
            If hProcessID <> 0 Then
                If OpenProcessToken(hProcessID, TOKEN_READ, hToken) = 1 Then
                    Result = GetTokenInformation(hToken, TOKEN_ELEVATION_TYPE, tokenElevationType, 4, lReturnLength)
                    If Result = 0 Then
                        CloseHandle hProcessID
                        ' Couldn’t get token information
                        Exit Function
                    End If
                    If tokenElevationType <> 1 Then
                        IsUserAnAdministrator = True
                    End If
                    CloseHandle hToken
                End If
                CloseHandle hProcessID
            End If
            Exit Function
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWin10
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWin10() As Boolean
Attribute IsWin10.VB_UserMemId = 1610612741

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    IsWin10 = OSCurrVersionStruct.VerFull = "10.0"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWin10OrLater
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWin10OrLater() As Boolean
Attribute IsWin10OrLater.VB_UserMemId = 1610612740

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    IsWin10OrLater = OSCurrVersionStruct.VerFull >= "10.0"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWin7
'! Description (Описание)  :   [type_description_here]0
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWin7() As Boolean
Attribute IsWin7.VB_UserMemId = 1610612743

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    IsWin7 = OSCurrVersionStruct.VerFull = "6.1"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWin7OrLater
'! Description (Описание)  :   [type_description_here]0
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWin7OrLater() As Boolean
Attribute IsWin7OrLater.VB_UserMemId = 1610612742

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    If Not IsWin10OrLater Then
        IsWin7OrLater = OSCurrVersionStruct.VerFull >= "6.1"
    Else
        IsWin7OrLater = True
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinVista
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinVista() As Boolean
Attribute IsWinVista.VB_UserMemId = 1610612745

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    IsWinVista = OSCurrVersionStruct.VerFull = "6.0"

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinVistaOrLater
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinVistaOrLater() As Boolean
Attribute IsWinVistaOrLater.VB_UserMemId = 1610612744

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    If Not IsWin10OrLater Then
        IsWinVistaOrLater = OSCurrVersionStruct.VerFull >= "6.0"
    Else
        IsWinVistaOrLater = True
    End If
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinXP
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinXP() As Boolean
Attribute IsWinXP.VB_UserMemId = 1610612747

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    IsWinXP = OSCurrVersionStruct.VerFull = "5.1"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsWinXPOrLater
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function IsWinXPOrLater() As Boolean
Attribute IsWinXPOrLater.VB_UserMemId = 1610612746

    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    If Not IsWin10OrLater Then
        IsWinXPOrLater = OSCurrVersionStruct.VerFull > "5.0"
    Else
        IsWinXPOrLater = True
    End If
End Function

'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Windows NT\CurrentVersion\WinPE
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function OSInfo
'! Description (Описание)  :   [Получение расширенной информации о версии Windows]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function OSInfo() As OSInfoStruct
Attribute OSInfo.VB_UserMemId = 1610612749

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
                ElseIf .dwMinorVersion = 4 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "10 (Beta)"
                    Else
                        OSN = "Server 2014 (Beta)"
                    End If
                End If

            ElseIf .dwMajorVersion = 10 Then
                
                If .dwMinorVersion = 0 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "10"
                    Else
                        OSN = "Server 2016"
                    End If

                ElseIf .dwMinorVersion = 1 Then

                    If .wProductType = VER_NT_WORKSTATION Then
                        OSN = "10.1 ?"
                    Else
                        OSN = "Server 2016 R2 ?"
                    End If

                Else
                    OSN = "11 ?"
                End If

            Else
                OSN = "11 ?"
            End If

            OSInfo.Name = "Windows " & OSN
            OSInfo.BuildNumber = .dwBuildNumber
            OSInfo.ServicePack = TrimNull(.szCSDVersion)
            OSInfo.VerFullwBuild = .dwMajorVersion & strDot & .dwMinorVersion & strDot & .dwBuildNumber
            OSInfo.VerFull = .dwMajorVersion & strDot & .dwMinorVersion
            OSInfo.VerMajor = .dwMajorVersion
            OSInfo.VerMinor = .dwMinorVersion
            OSInfo.ClientOrServer = .wProductType = VER_NT_WORKSTATION
            OSInfo.IsInitialize = True
            strOSCurrentVersion = OSInfo.VerFull
        End With

    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function OSInfoWMI
'! Description (Описание)  :   [Получение расширенной информации о версии Windows, альтернативная функция, использует WMI]
'! Parameters  (Переменные):   Nfo (Long)
'!--------------------------------------------------------------------------------
Public Function OSInfoWMI(ByVal Nfo As Long) As String
Attribute OSInfoWMI.VB_UserMemId = 1610612750

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

    strComputer = strDot
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
'! Procedure   (Функция)   :   Function IsWow64
'! Description (Описание)  :   [Проверяет является ли запущенный процесс 64-битным]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function OS_Is_x64() As Boolean
Attribute OS_Is_x64.VB_UserMemId = 1610612748

    Dim SI As SYSTEM_INFO
    Dim lngIts64 As Long

    strOSArchitecture = "x86"

    If APIFunctionPresent("GetNativeSystemInfo", "kernel32.dll") Then
        GetNativeSystemInfo SI

        Select Case SI.wProcessorArchitecture

            Case PROCESSOR_ARCHITECTURE_IA64
                OS_Is_x64 = True
                strOSArchitecture = "ia64"

            Case PROCESSOR_ARCHITECTURE_AMD64
                OS_Is_x64 = True
                strOSArchitecture = "amd64"

            Case Else
                OS_Is_x64 = False
        End Select
        
        If APIFunctionPresent("IsWow64Process", "kernel32.dll") Then
            ' IsWow64Process function exists
            ' Now use the function to determine if
            ' we are running under Wow64
            IsWow64Process GetCurrentProcess(), lngIts64
            If mbDebugStandart Then DebugMode "IsWow64: " & CBool(lngIts64)
        End If

    End If

End Function
