Attribute VB_Name = "mCollectHwid"
Option Explicit

Private Const wbemFlagReturnImmediately As Long = &H10
Private Const wbemFlagForwardOnly       As Long = &H20

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CollectHwidFromReestr
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub CollectHwidFromReestr()

    Dim strDateDRV        As String
    Dim strVersionDRV     As String
    Dim strID             As String
    Dim RecCountArr       As Long
    Dim strInfName        As String
    Dim strProviderName   As String
    Dim strCompatID       As String
    Dim strMatchesID      As String
    Dim strStrDescription As String
    Dim i                 As Long
    Dim regNameEnum       As String
    Dim regDriverClass    As String
    Dim regNameClass      As String
    Dim strDeviceDesc     As String
    Dim strMfg            As String
    Dim strCompatibleIDs  As String

    DebugMode vbTab & "CollectHwidFromReestr-Start"

    ' максимальное кол-во элементов в массиве
    For i = LBound(arrHwidsLocal) To UBound(arrHwidsLocal)
        strID = arrHwidsLocal(i).HWIDOrig
        ' Получаем данные об устройстве
        regNameEnum = "SYSTEM\CurrentControlSet\Enum\" & strID & vbBackslash
        ' список ID оборудования
        strCompatID = UCase$(GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "HardwareID", True))
        strCompatibleIDs = UCase$(GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "CompatibleIDs", True))

        If LenB(strCompatibleIDs) > 0 Then
            If LenB(strCompatID) > 0 Then
                strCompatID = strCompatID & (" | " & strCompatibleIDs)
            End If
        End If

        ' Убираем глюк с появлением &CTLR_ в HWID устройства
        If LenB(strCompatID) > 0 Then
            If InStr(strCompatID, "&CTLR_") Then
                DebugMode vbTab & "CollectHwidFromReestr-Start - !!! Replace for HWID: " & strID & " in CompatibleIDs '&CTLR_' ---> &_", 1
                strCompatID = Replace$(strCompatID, "&CTLR_", "&_")
            End If
        End If

        strDeviceDesc = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "DeviceDesc", True)
        strMfg = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "Mfg", True)
        regDriverClass = GetKeyValue(HKEY_LOCAL_MACHINE, regNameEnum, "Driver", True)

        ' Получаем данные о драйвере
        If LenB(regDriverClass) > 0 Then
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
        If LenB(strDateDRV) > 0 Then
            ConvertDate2Rus strDateDRV
        End If

        If LenB(strDateDRV) > 0 And LenB(strVersionDRV) > 0 Then
            strVersionDRV = strDateDRV & "," & strVersionDRV
        Else
            strVersionDRV = "unknown"
        End If

        If LenB(strVersionDRV) > 0 Then
            arrHwidsLocal(i).VerLocal = Trim$(strVersionDRV)
        Else
            arrHwidsLocal(i).VerLocal = "unknown"
        End If

        If LenB(strProviderName) > 0 Then
            arrHwidsLocal(i).Provider = Trim$(strProviderName)
        Else
            arrHwidsLocal(i).Provider = "unknown"
        End If

        If LenB(strCompatID) > 0 Then
            arrHwidsLocal(i).HWIDCompat = Trim$(strCompatID)
        Else
            arrHwidsLocal(i).HWIDCompat = "UNKNOWN"
        End If

        If LenB(strStrDescription) > 0 Then
            arrHwidsLocal(i).Description = Trim$(strStrDescription)
        Else
            arrHwidsLocal(i).Description = "unknown"
        End If

        If LenB(strInfName) > 0 Then
            arrHwidsLocal(i).HWIDMatches = UCase$(Trim$(strMatchesID))
        Else
            arrHwidsLocal(i).HWIDMatches = "UNKNOWN"
        End If

        If LenB(strInfName) > 0 Then
            arrHwidsLocal(i).InfName = Trim$(strInfName)
        Else
            arrHwidsLocal(i).InfName = "unknown"
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
    DebugMode vbTab & "CollectHwidFromReestr: Found Devices: " & i & vbNewLine & _
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
    ChangeStatusTextAndDebug strMessages(94)
    RunDevcon
    DevParserLocalHwids2
    ChangeStatusTextAndDebug strMessages(95)
    ' Обновляем данные из реестра
    CollectHwidFromReestr
    ChangeStatusTextAndDebug strMessages(114)
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveHWIDs2File
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub SaveHWIDs2File()

    If SaveHwidsArray2File(strResultHwidsExtTxtPath, arrHwidsLocal) = False Then
        MsgBox strMessages(45) & vbNewLine & strResultHwidsExtTxtPath, vbCritical + vbInformation, strProductName
    End If

End Sub
