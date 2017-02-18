Attribute VB_Name = "mCollectHwid"
Option Explicit

' Flag to determine intel generation for correct install USB30 driver
Public mbIUSB_RootHubExist  As Boolean
Public mbIntel2thGeneration As Boolean
Public mbIntel4thGeneration As Boolean

' Intel USB3 Root Hub device id
Private Const strIUSB30     As String = "IUSB3\ROOT_HUB30"
Private Const strIUSB30_2th As String = "IUSB3\ROOT_HUB30&VID_8086&PID_1E31"
Private Const strIUSB30_4th As String = "IUSB3\ROOT_HUB30&VID_8086&PID_8C31|IUSB3\ROOT_HUB30&VID_8086&PID_9C31|IUSB3\ROOT_HUB30&VID_8086&PID_0F35|IUSB3\ROOT_HUB30&VID_8086&PID_8CB1|IUSB3\ROOT_HUB30&VID_8086&PID_9CB1"

                                      
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CollectHwidFromReestr
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CollectHwidFromReestr()

    Dim strDateDRV        As String
    Dim strVersionDRV     As String
    Dim strID             As String
    Dim strInfName        As String
    Dim strProviderName   As String
    Dim strCompatID       As String
    Dim strMatchesID      As String
    Dim strStrDescription As String
    Dim ii                As Long
    Dim iii               As Long
    Dim strIUSB30_4th_x() As String
    Dim regNameEnum       As String
    Dim regDriverClass    As String
    Dim regNameClass      As String
    Dim strDeviceDesc     As String
    Dim strMfg            As String
    Dim strCompatibleIDs  As String

    If mbDebugDetail Then DebugMode vbTab & "CollectHwidFromReestr-Start"

    strIUSB30_4th_x = Split(strIUSB30_4th, "|")
    ' максимальное кол-во элементов в массиве
    For ii = 0 To UBound(arrHwidsLocal)
        strID = arrHwidsLocal(ii).HWIDOrig
        ' Получаем данные об устройстве
        regNameEnum = "SYSTEM\CurrentControlSet\Enum\" & strID & vbBackslash
        ' список ID оборудования
        strCompatID = UCase$(GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "HardwareID", True))
        strCompatibleIDs = UCase$(GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "CompatibleIDs", True))

        If LenB(strCompatibleIDs) Then
            If LenB(strCompatID) Then
                strCompatID = strCompatID & (" | " & strCompatibleIDs)
            End If
        End If

        ' Убираем глюк с появлением &CTLR_ в HWID устройства
        If LenB(strCompatID) Then
            If InStr(strCompatID, "&CTLR_") Then
                If mbDebugDetail Then DebugMode vbTab & "CollectHwidFromReestr-!!! Replace for HWID: " & strID & " in CompatibleIDs '&CTLR_' ---> &_"
                strCompatID = Replace$(strCompatID, "&CTLR_", "&_")
            End If
            
            ' Check for USB30 support
            If InStr(strCompatID, strIUSB30) Then
                mbIUSB_RootHubExist = True
            End If
            ' Check version of intel generation
            If mbIUSB_RootHubExist Then
                If InStr(strCompatID, strIUSB30_2th) Then
                    mbIntel2thGeneration = True
                Else
                    
                    For iii = LBound(strIUSB30_4th_x) To UBound(strIUSB30_4th_x)
                        If InStr(strCompatID, strIUSB30_4th_x(iii)) Then
                            mbIntel4thGeneration = True
                            Exit For
                        End If
                    Next iii
                End If
            End If
        End If

        strDeviceDesc = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "DeviceDesc", True)
        strMfg = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "Mfg", True)
        regDriverClass = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "Driver", True)

        ' Получаем данные о драйвере
        If LenB(regDriverClass) Then
            regNameClass = "SYSTEM\CurrentControlSet\Control\Class\" & regDriverClass & vbBackslash
            'SYSTEM\CurrentControlSet\Control\Class\"+pos+"\\"
            ' Получаем данные о драйвере
            strProviderName = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "ProviderName", True)
            strDateDRV = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDate", True)
            strVersionDRV = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverVersion", True)
            strStrDescription = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "DriverDesc", True)
            strMatchesID = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "MatchingDeviceId", True)
            strInfName = GetKeyValue(HKEY_LOCAL_MACHINE, regNameClass, "InfPath", True)
        Else
            strProviderName = vbNullString
            strDateDRV = vbNullString
            strVersionDRV = vbNullString
            strStrDescription = vbNullString
            strInfName = vbNullString
        End If

        ' Если нет данных о драйвере, то подменяем их данными об устройстве
        If LenB(strProviderName) = 0 Then
            strProviderName = strMfg
        End If

        If LenB(strStrDescription) = 0 Then
            strStrDescription = strDeviceDesc
        End If

        'var tmp2 = RegRead(pos + "InfSection");
        'var tmp3 = RegRead(pos + "InfSectionExt");
        ' если необходимо конвертировать дату в формат dd/mm/yyyy, а также в формат русского
        If LenB(strDateDRV) Then
            ConvertDate2Rus strDateDRV
        End If

        If LenB(strDateDRV) Then
            If LenB(strVersionDRV) Then
                strVersionDRV = strDateDRV & strComma & strVersionDRV
            Else
                strVersionDRV = strUnknownLCase
            End If
        Else
            strVersionDRV = strUnknownLCase
        End If

        If LenB(strVersionDRV) Then
            arrHwidsLocal(ii).VerLocal = Trim$(strVersionDRV)
        Else
            arrHwidsLocal(ii).VerLocal = strUnknownLCase
        End If

        If LenB(strProviderName) Then
            arrHwidsLocal(ii).Provider = Trim$(strProviderName)
        Else
            arrHwidsLocal(ii).Provider = strUnknownLCase
        End If

        If LenB(strCompatID) Then
            arrHwidsLocal(ii).HWIDCompat = Trim$(strCompatID)
        Else
            arrHwidsLocal(ii).HWIDCompat = strUnknownUCase
        End If

        If LenB(strStrDescription) Then
            arrHwidsLocal(ii).Description = Trim$(strStrDescription)
        Else
            arrHwidsLocal(ii).Description = strUnknownLCase
        End If

        If LenB(strInfName) Then
            arrHwidsLocal(ii).HWIDMatches = UCase$(Trim$(strMatchesID))
        Else
            arrHwidsLocal(ii).HWIDMatches = strUnknownUCase
        End If

        If LenB(strInfName) Then
            arrHwidsLocal(ii).InfName = Trim$(strInfName)
        Else
            arrHwidsLocal(ii).InfName = strUnknownLCase
        End If

    Next

    '0 - strDevHwid
    '1 - strDevName
    '2 - strDevStatus
    '3 - strDevVerLocal
    '4 - strOrigHwid
    '5 - strProvider
    '6 - strCompatID
    '7 - strStrDescription
    '8 - strPriznakSravnenia
    '9 - strSection
    '10 - strIDCutting
    '11 - strMatchesID
    '12 - strInfName
    '13 - Есть драйвера или нет
    '14 - Список пакетов где обнуружены драйвера
    If mbDebugStandart Then DebugMode vbTab & "CollectHwidFromReestr: Found Devices: " & ii & vbNewLine & _
              vbTab & "CollectHwidFromReestr-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ReCollectHWID
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ReCollectHWID()
    'Поиск новых устройств
    RunDevconRescan
    ' Сбор сведений о PC
    ChangeStatusBarText strMessages(94)
    RunDevcon
    DevParserLocalHwids2
    ChangeStatusBarText strMessages(95)
    ' Обновляем данные из реестра
    CollectHwidFromReestr
    ChangeStatusBarText strMessages(114)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveHWIDs2File
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub SaveHWIDs2File()
    
    If SaveHwidsArray2File(strResultHwidsExtTxtPath, arrHwidsLocal) = False Then
        If Not mbSilentRun Then
            MsgBox strMessages(45) & vbNewLine & strResultHwidsExtTxtPath, vbCritical + vbInformation, strProductName
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveSnapReport
'! Description (Описание)  :   [Сохранение снимка системы для подсистемы эмулирования]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub SaveSnapReport(Optional ByVal strDirPathTo As String = vbNullString)
    
    Dim strFileReport       As String
    Dim strDirPathToExpand  As String
    
    If LenB(strDirPathTo) Then
        strDirPathToExpand = PathCollect(strDirPathTo)
    Else
        strDirPathToExpand = GetPathNameFromPath(strDebugLogFullPath)
    End If
    
    If PathExists(strDirPathToExpand) = False Then
        CreateNewDirectory strDirPathToExpand
    End If
    
    ' Раскрываем параметр strFilePathTo, может содержать пеерменные Environ
    strFileReport = PathCombine(strDirPathToExpand, GetFileName4Snap & ".txt")
    
    ' Если снимок есть, копируем файл снимка по назначению
    If FileExists(strResultHwidsExtTxtPath) Then
        CopyFileTo strResultHwidsExtTxtPath, strFileReport
    Else

        ' Создание временного файла снимка системы
        If SaveHwidsArray2File(strResultHwidsExtTxtPath, arrHwidsLocal) Then
            If FileExists(strResultHwidsExtTxtPath) Then
                CopyFileTo strResultHwidsExtTxtPath, strFileReport
            Else
                If mbDebugStandart Then DebugMode strMessages(45) & vbNewLine & strFileReport
            End If
        Else
            If mbDebugStandart Then DebugMode strMessages(45) & vbNewLine & strFileReport
        End If
    End If
End Sub

