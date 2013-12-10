Attribute VB_Name = "mMain"
Option Explicit

' Основные параметры программы
Public Const strDateProgram             As String = "10/12/2013"

' Текущая версия базы данных
Public Const lngDevDBVersion            As Long = 5

' Версии лицензионного соглашения и файла Donate
Public Const strEULA_Version            As String = "02/02/2010"
Public Const strEULA_MD5RTF             As String = "68da44c8b1027547e4763472e0ecb727"
Public Const strEULA_MD5RTF_Eng         As String = "0cbd9d50eec41b26d24c5465c4be70bc"
Public Const strDONATE_MD5RTF           As String = "637e1aacdfcfa01fdc8827eb48796b1b"
Public Const strDONATE_MD5RTF_Eng       As String = "ca762ec290f0d9bedf2e09319661921a"

'Константы путей дополнительных утилит
Public Const strDevManView_Path         As String = "Tools\DevManView\DevManView.exe"
Public Const strDevManView_Path64       As String = "Tools\DevManView\DevManView-x64.exe"
Public Const strSIV_Path                As String = "Tools\SIV\SIV32X.exe"
Public Const strSIV_Path64              As String = "Tools\SIV\SIV64X.exe"
Public Const strUDI_Path                As String = "Tools\UDI\UnknownDeviceIdentifier.exe"
Public Const strDoubleDriver_Path       As String = "Tools\DoubleDriver\dd.exe"
Public Const strUnknownDevices_Path     As String = "Tools\UnknownDevices\UnknownDevices.exe"

' Переменная название программы
Public strProductName                   As String
Public strProductVersion                As String

' рабочий файл настроек
Public strSysIni                        As String
Public mbLoadIniTmpAfterRestart         As Boolean

' кэпшн основной формы
Public strFrmMainCaptionTemp            As String
Public strFrmMainCaptionTempDate        As String

' Шрифт основной формы и шрифта подсказок
Public strMainForm_FontName             As String
Public lngMainForm_FontSize             As Long

' Шрифт других форм
Public strOtherForm_FontName            As String
Public lngOtherForm_FontSize            As Long
Public mbCalcDriverScore                As Boolean

Public Type arrHwidsStruct
    HWID                                As String
    DevName                             As String
    Status                              As Long
    VerLocal                            As String
    HWIDOrig                            As String
    Provider                            As String
    HWIDCompat                          As String
    Description                         As String
    PriznakSravnenia                    As String
    InfSection                          As String
    HWIDCutting                         As String
    HWIDMatches                         As String
    InfName                             As String
    DRVExist                            As Long
    DPsList                             As String
    DRVScore                            As Long
End Type

Public Type arrOSStruct
    Ver                                 As String
    Name                                As String
    drpFolder                           As String
    drpFolderFull                       As String
    devIDFolder                         As String
    devIDFolderFull                     As String
    is64bit                             As Long
    DPFolderNotExist                    As Boolean
    PathPhysX                           As String
    PathLanguages                       As String
    CntBtn                              As Long
    ExcludeFileName                     As String
    PathRuntimes                        As String
End Type

' Массивы данных
Public arrHwidsLocal()                  As arrHwidsStruct
Public arrOSList()                      As arrOSStruct
Public arrTTipStatusIcon()              As String
Public arrCheckDP()                     As String
Public arrUtilsList()                   As String
Public arrTTip()                        As String
Public arrTTipSize()                    As String
Public arrDevIDs()                      As String
Public arrDriversList()                 As String

Public lngMaxDriversArrCount            As Long
Public lngDriversArrCount               As Long

' Массив служебных сообщений
Public lngOSCount                       As Long
Public lngOSCountPerRow                 As Long
Public lngUtilsCount                       As Long

'Пути до исполняемых файлов и других рабочих файлов
Public strDevconCmdPath                 As String
Public strDevConExePath                 As String
Public strDevConExePath64               As String
Public strDevConExePathW2k              As String
Public strDPInstExePath                 As String
Public strDPInstExePath64               As String
Public strDPInstExePath86               As String
Public strArh7zExePATH                  As String
Public strHwidsTxtPath                  As String
Public strHwidsTxtPathView              As String

Private strHwidsTxtPathVersion          As String
Private strHwidsTxtPathDRVFiles         As String

Public strResultHwidsTxtPath            As String
Public strResultHwidsExtTxtPath         As String
Public strWorkTemp                      As String
Public strWorkTempBackSL                As String
Public strWinTemp                       As String
Public strWinDir                        As String
Public strSysDir                        As String
Public strSysDir64                      As String
Public strSysDir86                      As String
Public strSysDirCatRoot                 As String
Public strSysDirDrivers                 As String
Public strSysDirDrivers64               As String
Public strSysDirDRVStore                As String
Public strSysDrive                      As String
Public strWinDirHelp                    As String
Public strInfDir                        As String

'Прочие параметры программы
Public mbIsWin64                        As Boolean
Public mbFirstStart                     As Boolean
Public mbStartMaximazed                 As Boolean
Public mbDelTmpAfterClose               As Boolean
Public mbUpdateCheck                    As Boolean
Public mbUpdateCheckBeta                As Boolean
Public mbUpdateToolTip                  As Boolean
Public miStartMode                      As Long
Public mbPatnAbs                        As Boolean
Public mbRecursion                      As Boolean
Public mbSaveSizeOnExit                 As Boolean

' стартовая вкладка для типов пакетов
Public lngStartModeTab2                 As Long

' Добавляем к описанию в главном окне в названии программы
Public strThisBuildBy                   As String

' Параметры каталога %Temp%
Public mbTempPath                       As Boolean
Public strAlternativeTempPath           As String

Private mbInitXPStyle                   As Boolean

' Пользователь администратор?
Private mbIsUserAnAdmin                 As Boolean

Public dtStartTimeProg                  As Long
Public dtEndTimeProg                    As Long
Public dtAllTimeProg                    As String
Public strExcludeHWID                   As String

' Переменные для парсинга
Public InfTempPathList(3000)            As String
Public InfTempPathListCount             As Long
Public IndexDevIDMass                   As Long
Public mbDevParserRun                   As Boolean

'номер последнего элемента в списке ОС
Public LastIdOS                         As Long

'номер последнего элемента в списке утилит
Public LastIdUtil                       As Long

'режим работы с элементом listview - либо изменние либо добавление
Public mbAddInList                      As Boolean
Public mbIsDriveCDRoom                  As Boolean
Public mbTabBlock                       As Boolean
Public mbTabHide                        As Boolean
Public mbOffSideButton                  As Boolean
Public miOffSideCount                   As Long
Public mbButtonTextUpCase               As Boolean
Public CurrentSelButtonIndex            As Long
Public mbLoadFinishFile                 As Boolean
Public mbReadClasses                    As Boolean
Public mbBreakUpdateDBAll               As Boolean
Public mbReadDPName                     As Boolean
Public mbConvertDPName                  As Boolean
Public strExcludeFileName               As String
Public strTTipTextHeaders               As String
Public strImageStatusButtonName         As String
Public strImageMainName                 As String

'Public strImageMenuName                         As String
Public mbTasks                          As Boolean
Public strPathDRPList                   As String
Public mbooSelectInstall                As Boolean
Public mbCheckDRVOk                     As Boolean
Public mbGroupTask                      As Boolean
Public mbIgnorStatusHwid                As Boolean
Public mbDRVNotInstall                  As Boolean

' Лицензионное соглашение
Private mbShowLicence                   As Boolean
Private strLicenceDate                  As String

Public mbEULAAgree                      As Boolean

Private mbAllFolderDRVNotExist          As Boolean

' Запуск с коммандной строкой
Public mbRunWithParam                   As Boolean

Private mbRunWithParamS                 As Boolean
Private strRunWithParam                 As String

' Работаем в тихом режиме
Public mbSilentRun                      As Boolean
Public miSilentRunTimer                 As Integer
Public mbSilentDLL                      As Boolean
Public strSilentSelectMode              As String

' Параметры DPinst
Public mbDpInstLegacyMode               As Boolean
Public mbDpInstPromptIfDriverIsNotBetter As Boolean
Public mbDpInstForceIfDriverIsNotBetter As Boolean
Public mbDpInstSuppressAddRemovePrograms As Boolean
Public mbDpInstSuppressWizard           As Boolean
Public mbDpInstQuietInstall             As Boolean
Public mbDpInstScanHardware             As Boolean
Public mbLogNotOnCDRoom                 As Boolean
Public mbHideOtherProcess               As Boolean
Public mbChangeResolution               As Boolean

' Сравнение версий драйверов по дате
Public mbCompareDrvVerByDate            As Boolean

' Грузить\негрузить драйвера для несовместимых ОС
Public mbLoadUnSupportedOS              As Boolean

'Маркер перезапуска программы
Public mbRestartProgram                 As Boolean

'Проверка соответствия даты в региональных настрйоках формату mm/dd/уууу
Private mbCorrectShortDateFormat        As Boolean

'Флаг сообщает о том что драйвер был удален
Public mbDeleteDriverByHwid             As Boolean

' Автообновление конфигурации при удалении драйвера
Public mbAutoInfoAfterDelDRV            As Boolean

' Автообновление конфигурации при удалении драйвера
Public mbDateFormatRus                  As Boolean

' Искать новые устройства при запуске программы
Public mbSearchOnStart                  As Boolean
Public lngPauseAfterSearch              As Long
Public mbCreateRestorePoint             As Boolean
Public mbMatchingHWID                   As Boolean
Public mbCompatiblesHWID                As Boolean
Public mbSearchCompatibleDriverOtherOS  As Boolean

' Переменная для оперделния режима только распаковки драйверов
Public mbOnlyUnpackDP                   As Boolean

' Переменная для определения выключения DEP
Private mbDisableDEP                    As Boolean

' Глубина поиска совместимых HWID
Public lngCompatiblesHWIDCount          As Long

' Глубина поиска совместимых HWID
Public mbMatchHWIDbyDPName              As Boolean

' Расширенное меню
'Public mbExMenu                              As Boolean
'-------------------- Переменные размеров Формы и кнопок ------------------'
Public MainFormWidth                    As Long
Public MainFormHeight                   As Long
Public miButtonWidth                    As Long
Public miButtonHeight                   As Long
Public miButtonLeft                     As Long
Public miButtonTop                      As Long
Public miBtn2BtnLeft                    As Long
Public miBtn2BtnTop                     As Long

' Минимальные значения размеров формы
Public Const MainFormWidthMin           As Long = 9350
Public Const MainFormHeightMin          As Long = 6500

' Дефолтные значения размеров формы
Private Const MainFormWidthDef          As Long = 11800
Private Const MainFormHeightDef         As Long = 8400

' Минимальные значения размеров кнопки
Public Const ButtonWidthMin             As Long = 1500
Public Const ButtonHeightMin            As Long = 350

' Дефолтные значения размеров кнопки
Private Const ButtonWidthDef            As Long = 2150
Private Const ButtonHeightDef           As Long = 550

' Дефолтные значения размеров кнопки
Private Const ButtonLeftDef             As Long = 100
Private Const ButtonTopDef              As Long = 480
Private Const Btn2BtnLeftDef            As Long = 100
Private Const Btn2BtnTopDef             As Long = 100

' Дефолтные значения размеров колонок в всплывающем сообщении
Public lngSizeRow1                      As Long
Public lngSizeRow2                      As Long
Public lngSizeRow3                      As Long
Public lngSizeRow4                      As Long
Public lngSizeRow5                      As Long
Public lngSizeRow6                      As Long
Public lngSizeRow9                      As Long
Public lngSizeRow13                     As Long
Public maxSizeRowAllLine                As Long
' Максимальные значения размеров колонок в всплывающем сообщении

Public lngSizeRowDPMax                  As Long
Public lngSizeRow1Max                   As Long
Public lngSizeRow2Max                   As Long
Public lngSizeRow3Max                   As Long
Public lngSizeRow4Max                   As Long
Public lngSizeRow5Max                   As Long
Public lngSizeRow6Max                   As Long
Public lngSizeRow9Max                   As Long
Public lngSizeRow13Max                  As Long
Public maxSizeRowAllLineMax             As Long


'strTableHwidHeader1    = "-HWID-"
'strTableHwidHeader2    = "-Путь-"
'strTableHwidHeader3    = "-Файл-"
'strTableHwidHeader4    = "-Версия(БД)-"
'strTableHwidHeader5    = "-Версия(PC)-"
'strTableHwidHeader6    = "-Статус-"
'strTableHwidHeader7    = "-Наименование устройства-"
'strTableHwidHeader8    = "-Пакет драйверов-"
'strTableHwidHeader9    = "!"
'strTableHwidHeader10   = "Производитель"
'strTableHwidHeader11   = "Совместимые HWID"
'strTableHwidHeader12   = "Код устройства"
Public strTableHwidHeader1              As String
Public strTableHwidHeader2              As String
Public strTableHwidHeader3              As String
Public strTableHwidHeader4              As String
Public strTableHwidHeader5              As String
Public strTableHwidHeader6              As String
Public strTableHwidHeader7              As String
Public strTableHwidHeader8              As String
Public strTableHwidHeader9              As String
Public strTableHwidHeader10             As String
Public strTableHwidHeader11             As String
Public strTableHwidHeader12             As String
Public strTableHwidHeader13             As String
Public strTableHwidHeader14             As String

Public lngTableHwidHeader1              As Long
Public lngTableHwidHeader2              As Long
Public lngTableHwidHeader3              As Long
Public lngTableHwidHeader4              As Long
Public lngTableHwidHeader5              As Long
Public lngTableHwidHeader6              As Long
Public lngTableHwidHeader7              As Long
Public lngTableHwidHeader8              As Long
Public lngTableHwidHeader9              As Long
Public lngTableHwidHeader10             As Long
Public lngTableHwidHeader11             As Long
Public lngTableHwidHeader12             As Long
Public lngTableHwidHeader13             As Long
Public lngTableHwidHeader14             As Long

' Переменные для определения модели компа
Public strCompModel                     As String
Public mbIsNotebok                      As Boolean

Public arrNotebookFilterList()          As String
Public arrNotebookFilterListDef()     As String
Public mbDP_Is_aFolder                  As Boolean
Public mbCheckUpdNotEnd                 As Boolean

'! -----------------------------------------------------------
'!  Функция     :  Main
'!  Переменные  :
'!  Описание    :  Основная функция запуска программы
'! -----------------------------------------------------------
Private Sub Main()

Dim mbShowFormLicence                   As Boolean
Dim strSysIniTMP                        As String

    On Error Resume Next
    dtStartTimeProg = GetTickCount
    
    Set objFSO = New Scripting.FileSystemObject
    Kavichki = ChrW$(34)
    
    ' Запоминаем app.path и прочее в переменные
    GetCurAppPath
    strProductVersion = App.Major & "." & App.Minor & "." & App.Revision
    strProductName = App.ProductName
    
    'считываем версию операционки
    If Not OsCurrVersionStruct.IsInitialize Then
        OsCurrVersionStruct = OSInfo
    End If
    strOsCurrentVersion = OsCurrVersionStruct.VerFull
    
    'Получаем временный каталог windows и каталог windows
    strWinDir = BackslashAdd2Path(Environ$("WINDIR"))
    strWinTemp = BackslashAdd2Path(Environ$("TMP"))

    If InStr(strWinTemp, " ") Then
        strWinTemp = strWinDir & "TEMP"
    End If

    ' Если временный каталог windows  (%windir%\temp)недоступен
    If PathFileExists(strWinTemp) = 0 Then
        MsgBox "Windows TempPath not Exist or Environ %TMP% undefined. Program is exit!!!", vbInformation, strProductName
        End
    End If
    
    ' Инициализация массива вендоров ноутбуков
    LoadNotebookList
    
    '******************************************
    ' Проверяем работает ли программа в режиме IDE
    ' Программа уже запущена???
    If App.PrevInstance And Not InIDE() Then
        MsgBoxEx "Found a running application 'Drivers Installer Assistant'. If you restart the program from the settings menu, then save the settings, the program waits until the previous session..." & str2vbNewLine & "This window will close automatically in 5 seconds. Please wait or click OK", 6, vbExclamation + vbSystemModal, strProductName
        ShowPrevInstance
    Else
        '******************************************
        ' - Инициализируем стиль WindowsXP
        Call ComCtlsInitIDEStopProtection
        Call InitVisualStyles
    End If

  
    ' Если каталог tools недоступен
    If PathFileExists(strAppPathBackSL & "Tools\") = 0 Then
        MsgBox "Not found the main program subfolder '.\Tools'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName
        End
    End If

    ' Рабочий временный каталог
    strWorkTemp = strWinTemp & "DriversInstaller"
    strWorkTempBackSL = BackslashAdd2Path(strWorkTemp)

    ' Создаем временный рабочий каталог
    If PathFileExists(strAppPathBackSL & "DriversInstaller.ini") = 0 Then
        strSysIni = strAppPathBackSL & "Tools\DriversInstaller.ini"
    Else
        strSysIni = strAppPathBackSL & "DriversInstaller.ini"
    End If
    
   ' Запущена ли программа с CD
    mbIsDriveCDRoom = IsDriveCDRoom
    ' Создаем файл настроек при необходимости
    CreateIni
    ' Считаваем язык операционки
    LoadLanguageOS

    'загружаем языковые файлы
    If PathFileExists(strAppPathBackSL & "Tools\Lang") = 1 Then
        mbMultiLanguage = LoadLanguageList
    End If

    'загружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    
    ' Получение настроек из ini-файла
    GetMainIniParam
 
    'Получение значений маркеров
    GetSummaryDPMarkers

    ' Если стоит настройка проверять временный путь на наличие ini, то перезагружаем файл параметров
    If mbLoadIniTmpAfterRestart Then
        If GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP", False) Then
            ' Reload Main ini
            strSysIniTMP = GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP_PATH", vbNullString)

            If LenB(strSysIniTMP) > 0 Then
                If PathFileExists(strSysIniTMP) = 1 Then
                    strSysIni = strSysIniTMP
                    ' Собственно перезагрузка настроек
                    GetMainIniParam
                End If
            End If
        End If
    End If

    If PathFileExists(strWorkTemp) = 0 Then
        CreateNewDirectory strWorkTemp
    End If

    'Перегружаем языковые файлы
    If PathFileExists(strAppPathBackSL & "Tools\Lang") = 1 Then
        mbMultiLanguage = LoadLanguageList
    End If

    'перегружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    
    strPathImageStatusButton = strAppPathBackSL & "Tools\Graphics\StatusButton\"
    strPathImageMain = strAppPathBackSL & "Tools\Graphics\Main\"
    'strPathImageMenu = strAppPathBackSL & "Tools\Graphics\Menu\"
    
    LoadIconImagePath
    ' Находится ли лог на CD
    mbLogNotOnCDRoom = LogNotOnCDRoom
    ' Очищаем лог-историю
    MakeCleanHistory
    ' Получаем размеры рабочей области программы
    GetWorkArea

    ' Проверяем на запуск с параметрами
    strRunWithParam = CStr(Command)
    If LenB(strRunWithParam) > 0 Then
        ' Парсинг строки запуска
        cmdLineParsing
        cmdLineAnalize
    End If
    
    If APIFunctionPresent("IsUserAnAdmin", "shell32.dll") Then
        mbIsUserAnAdmin = IsUserAnAdmin
    Else
        mbIsUserAnAdmin = True
    End If

    If Not mbDebugTime2File Then
        DebugMode "Current Date: " & Now()
    End If
    DebugMode "Version: " & strProductName & " v." & strProductVersion
    DebugMode "Build: " & strDateProgram
    DebugMode "ExeName: " & App.EXEName & ".exe"
    DebugMode "AppWork: " & strAppPath
    DebugMode "is User an Admin?: " & mbIsUserAnAdmin

    If mbIsUserAnAdmin Then
        ' записываем в реестр мой сертификат, для ЭЦП на exe-файлы
        DebugMode "SaveSert2Reestr"
        SaveSert2Reestr
    Else
        If Not mbRunWithParam Then
            If MsgBox(strMessages(138), vbYesNo + vbQuestion, strProductName) = vbNo Then
                End
            End If
        End If
    End If

    DebugMode "WinDir: " & strWinDir
    DebugMode "TmpDir: " & strWinTemp
    DebugMode "WorkTemp: " & strWorkTemp
    DebugMode "IsDriveCDRoom: " & mbIsDriveCDRoom

    If strOsCurrentVersion > "5.0" Then
        ' Определение windows x64
        mbIsWin64 = IsWow64
        DebugMode "IsWow64: " & mbIsWin64

        If mbIsWin64 Then
            Win64ReloadOptions
        End If
    ElseIf strOsCurrentVersion = "5.0" Then
        ' Для win2k надо старый devcon
        strDevConExePath = strDevConExePathW2k
    End If

    ' Disable DEP for current process
    If mbDisableDEP Then
        SetDEPDisable
    End If

    DebugMode "OsCurrentVersion: " & strOsCurrentVersion
    DebugMode "OS Language: ID=" & strPCLangID & " Name=" & strPCLangEngName & "(" & strPCLangLocaliseName & ")"
    
    ' Служебные файлы
    InitializePathHwidsTxt
    
    ' Регистрация внешних компонент
    RegisterAddComponent
    
    ' Если не существует каталогов с драйверами прописанных в настрйках, то выводим сообщение
    If mbAllFolderDRVNotExist Then
        MsgBox strMessages(6), vbCritical + vbApplicationModal, strProductName
        DebugMode strMessages(6)
        End
    End If
    
    DebugMode "InitXPStyle: " & mbInitXPStyle
    If APIFunctionPresent("IsAppThemed", "uxtheme.dll") Then
        mbAppThemed = IsAppThemed
        DebugMode "IsAppThemed: " & mbAppThemed
    End If
    mbAeroEnabled = IsAeroEnabled
    DebugMode "IsAeroEnabled : " & mbAeroEnabled

    mbCorrectShortDateFormat = IsCorrectShortDateFormat()
    
    ' изменяем разрешающую способность экрана монитора при необходимости
    SetVideoMode
    GetWorkArea
    
    ' Переменные для использовании при создании имени архива
    strCompModel = GetMBInfo()
    DebugMode "isNotebook: " & mbIsNotebok
    DebugMode "Notebook/Motherboard Model: " & strCompModel

    hc_Handle_Hand = LoadCursor(0, IDC_HAND)

    ' Показ лицензионного соглашения
    mbFirstStart = True
    mbShowLicence = GetSetting(App.ProductName, "Licence", "Show at Startup", True)
    strLicenceDate = GetSetting(App.ProductName, "Licence", "EULA_DATE", strEULA_Version)

    If InStr(1, strLicenceDate, strEULA_Version, vbTextCompare) Then
        If mbShowLicence Then
            If Not mbRunWithParam Then
                mbShowFormLicence = True
            End If

            If mbEULAAgree Then
                mbShowFormLicence = False
            End If

        Else
            mbShowFormLicence = False
        End If

    Else

        If Not mbRunWithParam Then
            mbShowFormLicence = True
        End If

        If mbEULAAgree Then
            mbShowFormLicence = False
        End If
    End If

    If Not mbRunWithParam Then
        If Not CheckBallonTip Then
            If MsgBox(strMessages(9), vbYesNo + vbQuestion, strMessages(10)) = vbNo Then
                End
            End If
        End If
    End If

    If mbShowFormLicence Then
        'Открываем форму лицензионного соглашения
        frmLicence.Show
    Else
        'Открываем основную форму
        frmMain.Show vbModeless
    End If

End Sub

'! -----------------------------------------------------------
'!  Функция     :  ChangeStatusTextAndDebug
'!  Переменные  :  Optional strSimpleText As String, Optional strDebugText As String
'!  Описание    :  Изменение текста статустной строки и отладочной информации
'! -----------------------------------------------------------
Public Sub ChangeStatusTextAndDebug(Optional strPanel2Text As String, _
                                    Optional strDebugText As String, _
                                    Optional ByVal mbEqual As Boolean = False, _
                                    Optional ByVal mbDoEvents As Boolean = True, _
                                    Optional strPanel1Text As String)

    If LenB(strPanel2Text) > 0 Then
        If mbDoEvents Then
            DoEvents
        End If

        If frmMain.ctlUcStatusBar1.PanelCount >= 2 Then
            frmMain.ctlUcStatusBar1.PanelText(2) = strPanel2Text
        Else
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel2Text
        End If

        'frmMain.pbStatusBar.Refresh
        If LenB(strPanel1Text) > 0 Then
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel1Text
        End If
    End If

    If LenB(strDebugText) > 0 Then
        If mbEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text

            End If

        Else
            DebugMode strDebugText
        End If

    Else

        If mbEqual Then
            If LenB(strPanel1Text) > 0 Then
                DebugMode strPanel1Text & ": " & strPanel2Text
            Else
                DebugMode strPanel2Text
            End If
        End If

    End If

End Sub

Private Function CheckBallonTip() As Boolean
    regParam = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "EnableBalloonTips")

    If LenB(regParam) = 0 Then
        CheckBallonTip = True
    Else
        CheckBallonTip = regParam = "1"
    End If

    DebugMode "EnableBalloonTips: " & regParam & "(" & CheckBallonTip & ")"

End Function

'! -----------------------------------------------------------
'!  Функция     :  CreateIni
'!  Переменные  :
'!  Описание    :  Сохранение настроек в ини файл если файла нет
'! -----------------------------------------------------------
Private Sub CreateIni()
Dim cnt                                 As Long

    If PathFileExists(strSysIni) = 0 Then
        If mbIsDriveCDRoom Then
            strSysIni = strWorkTempBackSL & "DriversInstaller.ini"
            MsgBox "File DriversInstaller.ini is not Exist!" & vbNewLine & "This program works from CD\DVD, so we create temporary DriversInstaller.ini-file" & vbNewLine & strSysIni, vbInformation + vbApplicationModal, strProductName
        End If

        'Секция Main
        IniWriteStrPrivate "Main", "DelTmpAfterClose", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheck", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheckBeta", "1", strSysIni
        IniWriteStrPrivate "Main", "StartMode", "1", strSysIni
        IniWriteStrPrivate "Main", "EULAAgree", "0", strSysIni
        IniWriteStrPrivate "Main", "HideOtherProcess", "1", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTemp", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTempPath", "%Temp%", strSysIni
        IniWriteStrPrivate "Main", "AutoLanguage", "1", strSysIni
        IniWriteStrPrivate "Main", "StartLanguageID", "0409", strSysIni
        IniWriteStrPrivate "Main", "IconMainSkin", "Standart", strSysIni
        IniWriteStrPrivate "Main", "SilentDLL", "0", strSysIni
        IniWriteStrPrivate "Main", "AutoInfoAfterDelDRV", "1", strSysIni
        IniWriteStrPrivate "Main", "SearchOnStart", "0", strSysIni
        IniWriteStrPrivate "Main", "PauseAfterSearch", "1", strSysIni
        IniWriteStrPrivate "Main", "CreateRestorePoint", "1", strSysIni
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", "0", strSysIni
        'Секция Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "%SYSTEMDRIVE%", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLog2AppPath", "0", strSysIni
        IniWriteStrPrivate "Debug", "Time2File", "0", strSysIni
        'Секция Devcon
        IniWriteStrPrivate "Devcon", "PathExe", "Tools\Devcon\devcon.exe", strSysIni
        IniWriteStrPrivate "Devcon", "PathExe64", "Tools\Devcon\devcon64.exe", strSysIni
        IniWriteStrPrivate "Devcon", "PathExeW2K", "Tools\Devcon\devconw2k.exe", strSysIni
        IniWriteStrPrivate "Devcon", "CollectHwidsCmd", "Tools\Devcon\devcon_c.cmd", strSysIni
        'Секция DPInst
        IniWriteStrPrivate "DPInst", "PathExe", "Tools\DPInst\DPInst.exe", strSysIni
        IniWriteStrPrivate "DPInst", "PathExe64", "Tools\DPInst\DPInst64.exe", strSysIni
        IniWriteStrPrivate "DPInst", "LegacyMode", 1, strSysIni
        IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", 1, strSysIni
        IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", 0, strSysIni
        IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", 0, strSysIni
        IniWriteStrPrivate "DPInst", "SuppressWizard", 0, strSysIni
        IniWriteStrPrivate "DPInst", "QuietInstall", 0, strSysIni
        IniWriteStrPrivate "DPInst", "ScanHardware", 1, strSysIni
        'Секция Arc
        IniWriteStrPrivate "Arc", "PathExe", "Tools\Arc\7za.exe", strSysIni
        'Секция OS
        IniWriteStrPrivate "OS", "OSCount", "4", strSysIni
        IniWriteStrPrivate "OS", "OSCountPerRow", "4", strSysIni
        IniWriteStrPrivate "OS", "Recursion", "1", strSysIni
        IniWriteStrPrivate "OS", "TabBlock", "1", strSysIni
        IniWriteStrPrivate "OS", "TabHide", 0, strSysIni
        IniWriteStrPrivate "OS", "LoadFinishFile", "1", strSysIni
        IniWriteStrPrivate "OS", "ReadClasses", "1", strSysIni
        IniWriteStrPrivate "OS", "ReadDPName", "1", strSysIni
        IniWriteStrPrivate "OS", "ConvertDPName", "1", strSysIni
        IniWriteStrPrivate "OS", "ExcludeHWID", "USB\ROOT_HUB*;ROOT\*;STORAGE\*;USBSTOR\*;PCIIDE\IDECHANNEL;PCI\CC_0604", strSysIni
        IniWriteStrPrivate "OS", "CompareDrvVerByDate", "1", strSysIni
        IniWriteStrPrivate "OS", "DateFormatRus", "0", strSysIni
        IniWriteStrPrivate "OS", "MatchingHWID", "1", strSysIni
        IniWriteStrPrivate "OS", "CompatiblesHWID", "1", strSysIni
        IniWriteStrPrivate "OS", "CompatiblesHWIDCount", "10", strSysIni
        IniWriteStrPrivate "OS", "LoadUnSupportedOS", "0", strSysIni
        IniWriteStrPrivate "OS", "CalcDriverScore", "1", strSysIni
        IniWriteStrPrivate "OS", "SearchCompatibleDriverOtherOS", "1", strSysIni
        'IniWriteStrPrivate "OS", "MatchHWIDbyMarkers", "1", strSysIni
        IniWriteStrPrivate "OS", "MatchHWIDbyDPName", "1", strSysIni
        'Секция OS_1
        IniWriteStrPrivate "OS_1", "Ver", "5.0;5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_1", "Name", "2000/XP/2003 Server", strSysIni
        IniWriteStrPrivate "OS_1", "drpFolder", "drivers\xp", strSysIni
        IniWriteStrPrivate "OS_1", "devIDFolder", "drivers\xp\dev_db", strSysIni
        IniWriteStrPrivate "OS_1", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_1", "PathPhysX", "drivers\XP\DP_Graphics_PhysX*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "PathLanguages", "drivers\XP\DP_Graphics_Languages*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "PathRuntimes", "drivers\XP\DP_Runtimes*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        'Секция OS_2
        IniWriteStrPrivate "OS_2", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_2", "Name", "Vista/7/8/8.1/Server 2008", strSysIni
        IniWriteStrPrivate "OS_2", "drpFolder", "drivers\vista", strSysIni
        IniWriteStrPrivate "OS_2", "devIDFolder", "drivers\vista\dev_db", strSysIni
        IniWriteStrPrivate "OS_2", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_2", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        'Секция OS_3
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_3", "Name", "Vista/7/8/8.1/Server 2008 x64", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista64", strSysIni
        IniWriteStrPrivate "OS_3", "devIDFolder", "drivers\vista64\dev_db", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "1", strSysIni
        IniWriteStrPrivate "OS_3", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        'Секция OS_4
        IniWriteStrPrivate "OS_4", "Ver", "5.0;5.1;5.2;6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_4", "Name", "Windows XP / 2000 / Server 2003 / Vista / Server 2008 / 7 / 8 / 8.1", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\All", strSysIni
        IniWriteStrPrivate "OS_4", "devIDFolder", "drivers\All\dev_db", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "2", strSysIni
        IniWriteStrPrivate "OS_4", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        'Секция Utils
        IniWriteStrPrivate "Utils", "UtilsCount", "3", strSysIni
        'Секция Utils_1
        IniWriteStrPrivate "Utils_1", "Name", "DirectX Diagnostics", strSysIni
        IniWriteStrPrivate "Utils_1", "Path", "%WINDIR%\system32\dxdiag.exe", strSysIni
        IniWriteStrPrivate "Utils_1", "Params", vbNullString, strSysIni
        'Секция Utils_2
        IniWriteStrPrivate "Utils_2", "Name", "Disk Managment", strSysIni
        IniWriteStrPrivate "Utils_2", "Path", "diskmgmt.msc", strSysIni
        IniWriteStrPrivate "Utils_2", "Params", vbNullString, strSysIni
        'Секция Utils_3
        IniWriteStrPrivate "Utils_3", "Name", "Remove BugFix with Installation of Video Drivers Nvidia", strSysIni
        IniWriteStrPrivate "Utils_3", "Path", "Tools\Nvidia\PatchPostInstall.cmd", strSysIni
        IniWriteStrPrivate "Utils_3", "Params", vbNullString, strSysIni
        'Секция MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(MainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(MainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Courier New", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "8", strSysIni
        IniWriteStrPrivate "MainForm", "HighlightColor", "32896", strSysIni
        'Секция Buttons
        IniWriteStrPrivate "Button", "Width", CStr(ButtonWidthDef), strSysIni
        IniWriteStrPrivate "Button", "Height", CStr(ButtonHeightDef), strSysIni
        IniWriteStrPrivate "Button", "Left", "100", strSysIni
        IniWriteStrPrivate "Button", "Top", "100", strSysIni
        IniWriteStrPrivate "Button", "Btn2BtnLeft", "100", strSysIni
        IniWriteStrPrivate "Button", "Btn2BtnTop", "100", strSysIni
        IniWriteStrPrivate "Button", "TextUpCase", "0", strSysIni
        IniWriteStrPrivate "Button", "FontName", "MS Sans Serif", strSysIni
        IniWriteStrPrivate "Button", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Button", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Button", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Button", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Button", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Button", "FontColor", "0", strSysIni
        IniWriteStrPrivate "Button", "IconStatusSkin", "Standart", strSysIni
        'Секция Tab
        IniWriteStrPrivate "Tab", "FontName", "MS Sans Serif", strSysIni
        IniWriteStrPrivate "Tab", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Tab", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontColor", "0", strSysIni
        'Секция Tab2
        IniWriteStrPrivate "Tab2", "StartMode", "1", strSysIni
        IniWriteStrPrivate "Tab2", "FontName", "MS Sans Serif", strSysIni
        IniWriteStrPrivate "Tab2", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Tab2", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontColor", "&H8000000D", strSysIni
        'Секция NotebookVendor
        IniWriteStrPrivate "NotebookVendor", "FilterCount", "22", strSysIni

        'Секция "NotebookVendor"
        IniWriteStrPrivate "NotebookVendor", "FilterCount", UBound(arrNotebookFilterListDef), strSysIni
        For cnt = 0 To UBound(arrNotebookFilterListDef) - 1
            IniWriteStrPrivate "NotebookVendor", "Filter_" & cnt + 1, arrNotebookFilterListDef(cnt), strSysIni
        Next

        ' Приводим Ini файл к читабельному виду
        NormIniFile strSysIni
        ' Активация отладки после создания ini-файла
        mbDebugEnable = True
        mbCleanHistory = True
        strDebugLogPathTemp = "%SYSTEMDRIVE%"
        strDebugLogNameTemp = "DIA-LOG_%DATE%.txt"

    End If

End Sub

Public Function DeleteDriverbyHwid(ByVal strHwid As String) As Boolean

Dim cmdString                           As String
Dim strDevConTemp                       As String

    If mbIsWin64 Then
        strDevConTemp = strDevConExePath64
    Else

        If strOsCurrentVersion = "5.0" Then
            strDevConTemp = strDevConExePathW2k
        Else
            strDevConTemp = strDevConExePath
        End If
    End If

    cmdString = Kavichki & strDevconCmdPath & Kavichki & " " & Kavichki & strDevConTemp & Kavichki & " " & Kavichki & strHwidsTxtPath & Kavichki & " 4 " & Kavichki & strHwid & Kavichki

    If RunAndWaitNew(cmdString, strWorkTemp, vbNormalFocus) = False Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        DeleteDriverbyHwid = False
    Else
        DeleteDriverbyHwid = True
    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  GetMainIniParam
'!  Переменные  :
'!  Описание    :  Получение настроек из ини файла
'! -----------------------------------------------------------
Private Sub GetMainIniParam()

Dim i                                   As Long
Dim mbAllFolderDRVNotExistCount         As Integer
Dim cntOsInIni                          As Integer
Dim cntUtilsInIni                       As Integer
Dim strDebugLogPathFolder               As String
Dim NotebookFilterCount                 As Long
Dim numFilter                           As Long

    'SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", True
    'SaveSetting App.ProductName, "Settings", "LOAD_INI_PATH", strSysIni
    '[Description]
    strThisBuildBy = GetIniValueString(strSysIni, "Description", "BuildBy", vbNullString)
    'strThisBuildBy = "www.SamLab.Ws"
    '[Debug]
    ' Активация отладки
    mbDebugEnable = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 0)
    ' Очистка истории
    mbCleanHistory = GetIniValueBoolean(strSysIni, "Debug", "CleenHistory", 0)
    ' Путь до лог файла
    strDebugLogPathTemp = PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%"))
    strDebugLogPath = PathCollect(PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%")))
    ' Имя лог-файла
    strDebugLogNameTemp = GetIniValueString(strSysIni, "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt")
    strDebugLogName = ExpandFileNamebyEnvironment(GetIniValueString(strSysIni, "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt"))
    ' Деталировка отладки - по умолчанию=1
    lngDetailMode = GetIniValueLong(strSysIni, "Debug", "DetailMode", 1)
    ' Записывать время в лог-файл
    mbDebugTime2File = GetIniValueBoolean(strSysIni, "Debug", "Time2File", 0)
    ' Создавать лог-файл в подпапке "logs" программы
    mbDebugLog2AppPath = GetIniValueBoolean(strSysIni, "Debug", "DebugLog2AppPath", 0)

    If Not mbDebugLog2AppPath Then
        strDebugLogFullPath = strDebugLogPath & strDebugLogName

        If mbDebugEnable Then
            strDebugLogPathFolder = strDebugLogPath

            If PathFileExists(strDebugLogPathFolder) = 0 Then
                CreateNewDirectory strDebugLogPathFolder
            End If
        End If

    Else
        strDebugLogPath2AppPath = strAppPathBackSL & "logs\" & strDebugLogName
        strDebugLogFullPath = strDebugLogPath2AppPath

        If Not LogNotOnCDRoom Then
            If mbDebugEnable Then
                If PathFileExists(strAppPathBackSL & "logs\") = 0 Then
                    CreateNewDirectory strAppPathBackSL & "logs\"

                End If
            End If
        Else
            strDebugLogFullPath = strDebugLogPath & strDebugLogName
        End If

    End If

    If lngDetailMode < 1 Then
        lngDetailMode = 1
    ElseIf lngDetailMode > 2 Then
        lngDetailMode = 2
    End If

    '[Main]
    ' удаление при выходе
    mbDelTmpAfterClose = GetIniValueBoolean(strSysIni, "Main", "DelTmpAfterClose", 1)
    ' проверка обновлений при старте (Секция MAIN)
    mbUpdateCheck = GetIniValueBoolean(strSysIni, "Main", "UpdateCheck", 1)
    ' проверка обновлений при старте (Секция MAIN)
    mbUpdateCheckBeta = GetIniValueBoolean(strSysIni, "Main", "UpdateCheckBeta", 1)
    ' погасить EULA
    mbEULAAgree = GetIniValueBoolean(strSysIni, "Main", "EULAAgree", 0)
    ' Автоопределение языка
    mbAutoLanguage = GetIniValueBoolean(strSysIni, "Main", "AutoLanguage", 1)

    If Not mbAutoLanguage Then
        strStartLanguageID = IniStringPrivate("Main", "StartLanguageID", strSysIni)
    End If

    ' Получение альтернативного пути Temp
    strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)

    If strAlternativeTempPath = "no_key" Then
        strAlternativeTempPath = strWinTemp
    End If

    ' при необходимости используем альтернативный temp
    mbTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mbTempPath Then
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathFileExists(strAlternativeTempPath) = 1 Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & "DriversInstaller"

            ' Если нет, то создаем временный рабочий каталог
            If PathFileExists(strWorkTemp) = 0 Then
                CreateNewDirectory strWorkTemp
            End If

        Else
            DebugMode "Alternative TempPath not Exist. Use Windows Temp"
        End If

    End If

    mbSearchOnStart = GetIniValueBoolean(strSysIni, "Main", "SearchOnStart", 0)
    lngPauseAfterSearch = GetIniValueLong(strSysIni, "Main", "PauseAfterSearch", 1)
    mbCreateRestorePoint = GetIniValueBoolean(strSysIni, "Main", "CreateRestorePoint", 1)
    mbLoadIniTmpAfterRestart = GetIniValueBoolean(strSysIni, "Main", "LoadIniTmpAfterRestart", 0)
    mbDisableDEP = GetIniValueBoolean(strSysIni, "Main", "DisableDEP", 1)
    '[OS]
    mbDP_Is_aFolder = GetIniValueBoolean(strSysIni, "OS", "DP_Is_aFolder", 0)
    ' обработка вложенных каталогов (Секция ОС)
    mbRecursion = GetIniValueBoolean(strSysIni, "OS", "Recursion", 1)
    ' Делать неактивными вкладки не моей ОС
    mbTabBlock = GetIniValueBoolean(strSysIni, "OS", "TabBlock", 1)
    ' Скрывать вкладки не моей ОС
    mbTabHide = GetIniValueBoolean(strSysIni, "OS", "TabHide", 0)
    ' Расчитывать баллы драйвера
    mbCalcDriverScore = GetIniValueBoolean(strSysIni, "OS", "CalcDriverScore", 1)
    ' получение Кол-ва систем (Секция OS) и построение массива ОС
    lngOSCount = IniLongPrivate("OS", "OSCount", strSysIni)

    If lngOSCount = 0 Or lngOSCount = 9999 Then
        MsgBox strMessages(5), vbExclamation, strMessages(4)
        DebugMode "The List supported operating systems is empty. Functioning the program impossible"
        End
    Else
        ReDim arrOSList(lngOSCount - 1)

        For i = 0 To UBound(arrOSList)
            cntOsInIni = i + 1
            arrOSList(i).Ver = IniStringPrivate("OS_" & cntOsInIni, "Ver", strSysIni)
            arrOSList(i).Name = IniStringPrivate("OS_" & cntOsInIni, "Name", strSysIni)
            arrOSList(i).drpFolder = IniStringPrivate("OS_" & cntOsInIni, "drpFolder", strSysIni)

            If arrOSList(i).drpFolder <> "no_key" Then
                arrOSList(i).drpFolderFull = PathCollect(arrOSList(i).drpFolder)
                If PathFileExists(arrOSList(i).drpFolderFull) = 0 Then
                    DebugMode "Not find folder with package driver" & vbNewLine & "for OS: " & arrOSList(i).Name & str2vbNewLine & "Folder is not Exist: " & vbNewLine & arrOSList(i).drpFolderFull
                    arrOSList(i).DPFolderNotExist = True
                    mbAllFolderDRVNotExistCount = mbAllFolderDRVNotExistCount + 1

                    If i <> UBound(arrOSList) Then
                        mbAllFolderDRVNotExist = True
                    Else
                        mbAllFolderDRVNotExist = mbAllFolderDRVNotExist And mbAllFolderDRVNotExistCount = UBound(arrOSList) + 1
                    End If

                Else
                    mbAllFolderDRVNotExist = False
                    arrOSList(i).DPFolderNotExist = False
                End If

            Else
                DebugMode "Folder with package driver" & vbNewLine & "for OS: " & arrOSList(i).Name & vbNewLine & "Is Not present in options. Correct and start the program again."
            End If

            arrOSList(i).devIDFolder = IniStringPrivate("OS_" & cntOsInIni, "devIDFolder", strSysIni)
            arrOSList(i).devIDFolderFull = PathCollect(arrOSList(i).devIDFolder)

            arrOSList(i).is64bit = IniLongPrivate("OS_" & cntOsInIni, "is64bit", strSysIni)

            If arrOSList(i).is64bit = 9999 Then
                arrOSList(i).is64bit = 0
            End If

            arrOSList(i).PathPhysX = IniStringPrivate("OS_" & cntOsInIni, "PathPhysX", strSysIni)
            If arrOSList(i).PathPhysX = "no_key" Then
                arrOSList(i).PathPhysX = vbNullString
            End If

            arrOSList(i).PathLanguages = IniStringPrivate("OS_" & cntOsInIni, "PathLanguages", strSysIni)
            If arrOSList(i).PathLanguages = "no_key" Then
                arrOSList(i).PathLanguages = vbNullString
            End If

            arrOSList(i).ExcludeFileName = IniStringPrivate("OS_" & cntOsInIni, "ExcludeFileName", strSysIni)
            If arrOSList(i).ExcludeFileName = "no_key" Then
                arrOSList(i).ExcludeFileName = vbNullString
            End If

            arrOSList(i).PathRuntimes = IniStringPrivate("OS_" & cntOsInIni, "PathRuntimes", strSysIni)
            If arrOSList(i).PathRuntimes = "no_key" Then
                arrOSList(i).PathRuntimes = vbNullString
            End If
        Next
    End If

    ' получение Кол-ва вкладок на одной строке (Секция Main)
    lngOSCountPerRow = IniLongPrivate("OS", "OSCountPerRow", strSysIni)

    If lngOSCountPerRow = 0 Or lngOSCountPerRow = 9999 Then
        lngOSCountPerRow = 4
    End If

    '[Utils]
    ' получение Кол-ва утилит
    lngUtilsCount = IniLongPrivate("Utils", "UtilsCount", strSysIni)

    If lngUtilsCount = 0 Or lngUtilsCount = 9999 Then
        'MsgBox "Список поддерживаемых операционых систем пуст. Работа программы немозможна", vbExclamation, "Работа программы невозможна"
        ReDim arrUtilsList(0, 3)
        arrUtilsList(0, 0) = "List_Empty"
        arrUtilsList(0, 1) = vbNullString
        arrUtilsList(0, 2) = vbNullString
        arrUtilsList(0, 3) = vbNullString
    Else
        ReDim arrUtilsList(lngUtilsCount - 1, 3)

        For i = 0 To UBound(arrUtilsList)
            cntUtilsInIni = i + 1
            arrUtilsList(i, 0) = IniStringPrivate("Utils_" & cntUtilsInIni, "Name", strSysIni)
            arrUtilsList(i, 1) = IniStringPrivate("Utils_" & cntUtilsInIni, "Path", strSysIni)
            arrUtilsList(i, 2) = IniStringPrivate("Utils_" & cntUtilsInIni, "Path64", strSysIni)
            arrUtilsList(i, 3) = IniStringPrivate("Utils_" & cntUtilsInIni, "Params", strSysIni)

            If arrUtilsList(i, 2) = "no_key" Then
                arrUtilsList(i, 2) = vbNullString
            End If

            If arrUtilsList(i, 3) = "no_key" Or arrUtilsList(i, 3) = "Дополнительные параметры запуска" Then
                arrUtilsList(i, 3) = vbNullString
            End If
        Next
    End If

    '--------------------- Получение путей до файлов ---------------------
    '[DevCon]
    ' DEVCON_CMD
    strDevconCmdPath = IniStringPrivate("DevCon", "CollectHwidsCmd", strSysIni)
    strDevconCmdPath = PathCollect(strDevconCmdPath)

    If PathFileExists(strDevconCmdPath) = 0 Then
        strDevconCmdPath = strAppPathBackSL & "Tools\Devcon\devcon_c.cmd"

        If PathFileExists(strDevconCmdPath) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDevconCmdPath, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE
    strDevConExePath = IniStringPrivate("DevCon", "PathExe", strSysIni)

    If InStr(strDevConExePath, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePath = PathCollect(strDevConExePath)

    If PathFileExists(strDevConExePath) = 0 Then
        strDevConExePath = strAppPathBackSL & "Tools\Devcon\devcon.exe"

        If PathFileExists(strDevConExePath) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDevConExePath, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE64
    strDevConExePath64 = IniStringPrivate("DevCon", "PathExe64", strSysIni)

    If InStr(strDevConExePath64, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePath64 = PathCollect(strDevConExePath64)

    If PathFileExists(strDevConExePath64) = 0 Then
        strDevConExePath64 = strAppPathBackSL & "Tools\Devcon\devcon64.exe"

        If PathFileExists(strDevConExePath64) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDevConExePath64, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE_W2k
    strDevConExePathW2k = IniStringPrivate("DevCon", "PathExeW2k", strSysIni)

    If InStr(strDevConExePathW2k, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePathW2k = PathCollect(strDevConExePathW2k)

    If PathFileExists(strDevConExePathW2k) = 0 Then
        strDevConExePathW2k = strAppPathBackSL & "Tools\Devcon\devconw2k.exe"

        If PathFileExists(strDevConExePathW2k) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDevConExePathW2k, vbInformation, strProductName
        End If
    End If

    '[DPInst]
    ' DPInst.exe
    strDPInstExePath86 = IniStringPrivate("DPInst", "PathExe", strSysIni)

    If InStr(strDPInstExePath86, ":") Then
        mbPatnAbs = True
    End If

    strDPInstExePath86 = PathCollect(strDPInstExePath86)

    If PathFileExists(strDPInstExePath86) = 0 Then
        strDPInstExePath86 = strAppPathBackSL & "Tools\DPInst\DPInst.exe"

        If PathFileExists(strDPInstExePath86) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath86, vbInformation, strProductName
        End If
    End If

    strDPInstExePath = strDPInstExePath86
    ' DPInst64.exe
    strDPInstExePath64 = IniStringPrivate("DPInst", "PathExe64", strSysIni)

    If InStr(strDPInstExePath64, ":") Then
        mbPatnAbs = True
    End If

    strDPInstExePath64 = PathCollect(strDPInstExePath64)

    If PathFileExists(strDPInstExePath64) = 0 Then
        strDPInstExePath64 = strAppPathBackSL & "Tools\DPInst\DPInst64.exe"

        If PathFileExists(strDPInstExePath64) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath64, vbInformation, strProductName
        End If
    End If

    ' Настройки DpInst
    mbDpInstLegacyMode = GetIniValueBoolean(strSysIni, "DPInst", "LegacyMode", 1)
    mbDpInstPromptIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "PromptIfDriverIsNotBetter", 1)
    mbDpInstForceIfDriverIsNotBetter = GetIniValueBoolean(strSysIni, "DPInst", "ForceIfDriverIsNotBetter", 0)
    mbDpInstSuppressAddRemovePrograms = GetIniValueBoolean(strSysIni, "DPInst", "SuppressAddRemovePrograms", 0)
    mbDpInstSuppressWizard = GetIniValueBoolean(strSysIni, "DPInst", "SuppressWizard", 0)
    mbDpInstQuietInstall = GetIniValueBoolean(strSysIni, "DPInst", "QuietInstall", 0)
    mbDpInstScanHardware = GetIniValueBoolean(strSysIni, "DPInst", "ScanHardware", 1)
    '[Arc]
    ' 7za.exe
    strArh7zExePATH = IniStringPrivate("Arc", "PathExe", strSysIni)

    If InStr(strArh7zExePATH, ":") Then
        mbPatnAbs = True
    End If

    strArh7zExePATH = PathCollect(strArh7zExePATH)

    If PathFileExists(strArh7zExePATH) = 0 Then
        strArh7zExePATH = strAppPathBackSL & "Tools\Arc\7za.exe"

        If PathFileExists(strArh7zExePATH) = 0 Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePATH, vbInformation, strProductName
        End If

    End If

    '[MainForm]
    ' Сохранять настройки при выходе
    mbSaveSizeOnExit = GetIniValueBoolean(strSysIni, "MainForm", "SaveSizeOnExit", 0)
    'Ширина основной формы
    MainFormWidth = GetIniValueLong(strSysIni, "MainForm", "Width", MainFormWidthDef)

    'Если полученное значение меньше минимального, то устанавливаем значение по умолчанию
    If MainFormWidth < MainFormWidthMin Then
        MainFormWidth = MainFormWidthDef
    End If

    'Высота основной формы
    MainFormHeight = GetIniValueLong(strSysIni, "MainForm", "Height", MainFormHeightDef)

    'Если полученное значение меньше минимального, то устанавливаем значение по умолчанию
    If MainFormHeight < MainFormHeightMin Then
        MainFormHeight = MainFormHeightDef
    End If

    ' получение вида запуска (Секция MainForm)
    mbStartMaximazed = GetIniValueBoolean(strSysIni, "MainForm", "StartMaximazed", 0)
    strMainForm_FontName = GetIniValueString(strSysIni, "MainForm", "FontName", "Courier New")
    lngMainForm_FontSize = GetIniValueLong(strSysIni, "MainForm", "FontSize", 8)
    ' Подсветка активного элемента
    glHighlightColor = GetIniValueLong(strSysIni, "MainForm", "HighlightColor", 32896)
    ' получение вида запуска (Секция OtherForm)
    strOtherForm_FontName = GetIniValueString(strSysIni, "OtherForm", "FontName", "Tahoma")
    lngOtherForm_FontSize = GetIniValueLong(strSysIni, "OtherForm", "FontSize", 8)
    '[Buttons]
    miButtonWidth = GetIniValueLong(strSysIni, "Button", "Width", ButtonWidthDef)
    miButtonHeight = GetIniValueLong(strSysIni, "Button", "Height", ButtonHeightDef)
    miButtonLeft = GetIniValueLong(strSysIni, "Button", "Left", ButtonLeftDef)
    miButtonTop = GetIniValueLong(strSysIni, "Button", "Top", ButtonTopDef)
    miBtn2BtnLeft = GetIniValueLong(strSysIni, "Button", "Btn2BtnLeft", Btn2BtnLeftDef)
    miBtn2BtnTop = GetIniValueLong(strSysIni, "Button", "Btn2BtnTop", Btn2BtnTopDef)
    ' текст кнопок в верхнем регистре (Секция Button)
    mbButtonTextUpCase = GetIniValueBoolean(strSysIni, "Button", "TextUpCase", 0)
    '[OS]
    ' Обрабатывать файлы Finish
    mbLoadFinishFile = GetIniValueBoolean(strSysIni, "OS", "LoadFinishFile", 1)
    ' Считывать класс пакета из файла Finish
    mbReadClasses = GetIniValueBoolean(strSysIni, "OS", "ReadClasses", 1)
    ' Считывать имя пакета
    mbReadDPName = GetIniValueBoolean(strSysIni, "OS", "ReadDPName", 1)
    ' Преобразовывать имена пакетов
    mbConvertDPName = GetIniValueBoolean(strSysIni, "OS", "ConvertDPName", 1)
    ' Исключаемые HWID из обработки
    strExcludeHWID = GetIniValueString(strSysIni, "OS", "ExcludeHWID", "USB\ROOT_HUB*;ROOT\*;STORAGE\*;USBSTOR\*;PCIIDE\IDECHANNEL;PCI\CC_0604")
    ' Сравнение версий драйверов
    mbCompareDrvVerByDate = GetIniValueBoolean(strSysIni, "OS", "CompareDrvVerByDate", 1)
    ' Отображать дату версии в формате dd/mm/yyyy
    mbDateFormatRus = GetIniValueBoolean(strSysIni, "OS", "DateFormatRus", 0)
    ' Обрабатывать совместимые HWID
    mbMatchingHWID = GetIniValueBoolean(strSysIni, "OS", "MatchingHWID", 1)
    mbCompatiblesHWID = GetIniValueBoolean(strSysIni, "OS", "CompatiblesHWID", 1)
    lngCompatiblesHWIDCount = GetIniValueLong(strSysIni, "OS", "CompatiblesHWIDCount", 5)
    'Проверять совместимость по имени или маркеру
    'mbMatchHWIDbyMarkers = GetIniValueBoolean(strSysIni, "OS", "MatchHWIDbyMarkers", 1)
    mbMatchHWIDbyDPName = GetIniValueBoolean(strSysIni, "OS", "MatchHWIDbyDPName", 1)
    ' Обрабатывать совместимые HWID
    mbLoadUnSupportedOS = GetIniValueBoolean(strSysIni, "OS", "LoadUnSupportedOS", 0)
    mbSearchCompatibleDriverOtherOS = GetIniValueBoolean(strSysIni, "OS", "SearchCompatibleDriverOtherOS", 1)
    '[Button]
    ' Шрифт Кнопок
    strDialog_FontName = GetIniValueString(strSysIni, "Button", "FontName", "MS Sans Serif")
    miDialog_FontSize = GetIniValueLong(strSysIni, "Button", "FontSize", 8)
    mbDialog_Bold = GetIniValueBoolean(strSysIni, "Button", "FontBold", 0)
    mbDialog_Italic = GetIniValueBoolean(strSysIni, "Button", "FontItalic", 0)
    mbDialog_Underline = GetIniValueBoolean(strSysIni, "Button", "FontUnderline", 0)
    mbDialog_Strikethru = GetIniValueBoolean(strSysIni, "Button", "FontStrikethru", 0)
    lngDialog_Color = GetIniValueLong(strSysIni, "Button", "FontColor", 0)
    strImageStatusButtonName = GetIniValueString(strSysIni, "Button", "IconStatusSkin", "Standart")
    '[Tab]
    ' Шрифт и настройки ЗАКЛАДОК
    strDialogTab_FontName = GetIniValueString(strSysIni, "Tab", "FontName", "MS Sans Serif")
    miDialogTab_FontSize = GetIniValueLong(strSysIni, "Tab", "FontSize", 8)
    mbDialogTab_Bold = GetIniValueBoolean(strSysIni, "Tab", "FontBold", 0)
    mbDialogTab_Italic = GetIniValueBoolean(strSysIni, "Tab", "FontItalic", 0)
    mbDialogTab_Underline = GetIniValueBoolean(strSysIni, "Tab", "FontUnderline", 0)
    mbDialogTab_Strikethru = GetIniValueBoolean(strSysIni, "Tab", "FontStrikethru", 0)
    lngDialogTab_Color = GetIniValueLong(strSysIni, "Tab", "FontColor", 0)
    '[Tab2]
    ' Шрифт и настройки ЗАКЛАДОК
    strDialogTab2_FontName = GetIniValueString(strSysIni, "Tab2", "FontName", "MS Sans Serif")
    miDialogTab2_FontSize = GetIniValueLong(strSysIni, "Tab2", "FontSize", 8)
    mbDialogTab2_Bold = GetIniValueBoolean(strSysIni, "Tab2", "FontBold", 0)
    mbDialogTab2_Italic = GetIniValueBoolean(strSysIni, "Tab2", "FontItalic", 0)
    mbDialogTab2_Underline = GetIniValueBoolean(strSysIni, "Tab2", "FontUnderline", 0)
    mbDialogTab2_Strikethru = GetIniValueBoolean(strSysIni, "Tab2", "FontStrikethru", 0)
    lngDialogTab2_Color = GetIniValueLong(strSysIni, "Tab2", "FontColor", &H8000000D)
    lngStartModeTab2 = GetIniValueLong(strSysIni, "Tab2", "StartMode", 1)
    '[Main]
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' расширенное меню
    'mbExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    ' Скрывать прочие процессы
    mbHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' Тихая регистрация DLL
    mbSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' Показывать напоминание об обновлении (всплывающее окно)
    mbUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)
    ' Автообновление информации после удаления драйвера
    mbAutoInfoAfterDelDRV = GetIniValueBoolean(strSysIni, "Main", "AutoInfoAfterDelDRV", 1)
    ' Стартовый режим
    miStartMode = GetIniValueLong(strSysIni, "Main", "StartMode", 1)
    '[NotebookVendor]
    NotebookFilterCount = IniLongPrivate("NotebookVendor", "FilterCount", strSysIni)
    If NotebookFilterCount = 0 Or NotebookFilterCount = 9999 Then
        arrNotebookFilterList() = arrNotebookFilterListDef()
    Else
        ReDim arrNotebookFilterList(NotebookFilterCount)

        For i = 0 To UBound(arrNotebookFilterList) - 1
            numFilter = i + 1
            arrNotebookFilterList(i) = IniStringPrivate("NotebookVendor", "Filter_" & numFilter, strSysIni)
            If arrNotebookFilterList(i) = "no_key" Then
                arrNotebookFilterList(i) = arrNotebookFilterListDef(i)
            End If
        Next
    End If

End Sub

Public Function GetMB_Manufacturer() As String

Dim colItems                            As Object
Dim objItem                             As Object
Dim objWMIService                       As Object
Dim sAnsComputerSystem                  As String
Dim sAnsBaseBoard                       As String
Dim objRegExp                           As RegExp
Dim strTemp                             As String

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

Public Function GetMB_Model() As String

Dim colItems                            As Object
Dim objItem                             As Object
Dim objWMIService                       As Object
Dim sAnsComputerSystem                  As String
Dim sAnsBaseBoard                       As String
Dim objRegExp                           As RegExp
Dim strTemp                             As String

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

Private Function GetMBInfo() As String
Dim strMB_Manufacturer As String
Dim strMB_Model As String

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

Private Sub LoadNotebookList()

ReDim arrNotebookFilterListDef(35) As String

    arrNotebookFilterListDef(0) = "3Q;*3q*"
    arrNotebookFilterListDef(1) = "Acer;*acer*;*emachines*;*packard*bell*;*gateway*"
    arrNotebookFilterListDef(2) = "Apple;*apple*"
    arrNotebookFilterListDef(3) = "Asus;*asus*"
    arrNotebookFilterListDef(4) = "BenQ;*benq*"
    arrNotebookFilterListDef(5) = "Clevo;*clevo*"
    arrNotebookFilterListDef(6) = "Dell;*dell*;*alienware*"
    arrNotebookFilterListDef(7) = "Eurocom;*eurocom*"
    arrNotebookFilterListDef(8) = "Fujitsu;*fujitsu*;*sieme*"
    arrNotebookFilterListDef(9) = "Getac;*getac*"
    arrNotebookFilterListDef(10) = "Gigabyte;*gigabyte*;*ecs*;*elitegroup*"
    arrNotebookFilterListDef(11) = "HP;*hp*;*hewle*;*compaq*"
    arrNotebookFilterListDef(12) = "Intel;*intel*"
    arrNotebookFilterListDef(13) = "Inventec;*inventec*"
    arrNotebookFilterListDef(14) = "iRU;*iru*"
    arrNotebookFilterListDef(15) = "Lenovo;*lenovo*;*compal*;*ibm*"
    arrNotebookFilterListDef(16) = "LG;*lg*"
    arrNotebookFilterListDef(17) = "Matsushita;*matsushita*"
    arrNotebookFilterListDef(18) = "Mitac;*mitac*;*MTC*"
    arrNotebookFilterListDef(19) = "MSI;*msi*;*micro-star*"
    arrNotebookFilterListDef(20) = "NEC;*nec*"
    arrNotebookFilterListDef(21) = "Panasonic;*panasonic*"
    arrNotebookFilterListDef(22) = "Pegatron;*pegatron*"
    arrNotebookFilterListDef(23) = "PROLiNK;*prolink*"
    arrNotebookFilterListDef(24) = "Quanta;*quanta*"
    arrNotebookFilterListDef(25) = "Roverbook;*roverbook*"
    arrNotebookFilterListDef(26) = "Sager;*sager*"
    arrNotebookFilterListDef(27) = "Samsung;*samsung*"
    arrNotebookFilterListDef(28) = "Shuttle;*shuttle*"
    arrNotebookFilterListDef(29) = "SiS;*sis*"
    arrNotebookFilterListDef(30) = "Sony;*sony*;*vaio*"
    arrNotebookFilterListDef(31) = "Toshiba;*toshiba*"
    arrNotebookFilterListDef(32) = "ViewSonic;*viewsonic*;*ViewBook*;*viewbook*"
    arrNotebookFilterListDef(33) = "VIZIO;*vizio*"
    arrNotebookFilterListDef(34) = "Wistron;*wistron*"

End Sub

' Считавыем язык операционки
Private Sub LoadLanguageOS()
Dim LCID                                As Long

    ' Считавыем язык операционки
    LCID = GetSystemDefaultLCID()
    'language id
    strPCLangID = GetUserLocaleInfo(LCID, LOCALE_ILANGUAGE)
    'localized name of language
    strPCLangLocaliseName = GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE)
    'English name of language
    strPCLangEngName = GetUserLocaleInfo(LCID, LOCALE_SENGLANGUAGE)
End Sub

'! -----------------------------------------------------------
'!  Функция     :  Win64ReloadOptions
'!  Переменные  :
'!  Описание    :  Переназначение переменных для Win x64
'! -----------------------------------------------------------
Private Sub Win64ReloadOptions()
    DebugMode "Win64ReloadOptions"
    strDPInstExePath = strDPInstExePath64
End Sub

Private Sub cmdLineParsing()
Dim argRetCMD                           As Collection
Dim i                                   As Integer
Dim intArgCount                         As Integer
Dim strArg                              As String
Dim strArg_x()                          As String
Dim iArgRavno                           As Integer
Dim iArgDvoetoch                        As Integer
Dim strArgParam                         As String

    With New cCMDArguments
        .CommandLine = "CMDLineParams " & Command$
        Set argRetCMD = .Arguments
        intArgCount = argRetCMD.Count
    End With

    For i = 2 To intArgCount
        strArg = argRetCMD(i)
        iArgRavno = InStr(strArg, "=")
        iArgDvoetoch = InStr(strArg, ":")

        If iArgRavno > 0 Then
            strArg_x = Split(strArg, "=")
            strArg = strArg_x(0)
            strArgParam = strArg_x(1)
        ElseIf iArgDvoetoch > 0 Then
            'strArg_x = Split(strArg, ":")
            strArg = Left$(argRetCMD(i), iArgDvoetoch - 1)
            strArgParam = Right$(argRetCMD(i), Len(argRetCMD(i)) - iArgDvoetoch)
        End If

        Select Case LCase$(strArg)
            Case "/?", "/h", "-help", "/help", "-h", "--h", "--help"
                ShowHelpMsg
                End

            Case "/extractdll", "-extractdll", "--extractdll"
                ExtractrResToFolder strArgParam
                End

            Case "/regdll", "-regdll", "--regdll"
                RegisterAddComponent
                End

            Case "/t", "-t", "--t"
                If IsNumeric(strArgParam) Then
                    miSilentRunTimer = CInt(strArgParam)
                Else
                    miSilentRunTimer = 10
                End If

                mbDebugEnable = True
                mbUpdateCheck = False

            Case "/s", "-s", "--s"
                mbRunWithParamS = True

                Select Case LCase$(strArgParam)
                    Case "n"
                        'новые
                        strSilentSelectMode = "n"

                    Case "q"
                        'неустановленные
                        strSilentSelectMode = "q"

                    Case "a"
                        'Все на вкладке
                        strSilentSelectMode = "a"

                    Case "n2"
                        'новые
                        strSilentSelectMode = "n2"

                    Case "q2"
                        'неустановленные
                        strSilentSelectMode = "q2"

                    Case "a2"
                        'Все на вкладке
                        strSilentSelectMode = "a2"

                    Case Else
                        'по умолчанию
                        strSilentSelectMode = "n"
                End Select

                ' на случай если не указано время ожидания запуска
                If miSilentRunTimer <= 0 Then
                    miSilentRunTimer = 10
                End If

                mbDebugEnable = True
                mbUpdateCheck = False

            Case Else
                ShowHelpMsg
                End
        End Select
    Next i
End Sub

' Показ окна с параметрами запуска
Private Sub ShowHelpMsg()
    MsgBox strMessages(137), vbInformation & vbOKOnly, strProductName & " " & strProductVersion
End Sub

' Извлечение ресурсов программы в каталог
Private Sub ExtractrResToFolder(strArg As String)
Dim strArg_x()                          As String
Dim strPathToTemp                       As String
Dim strPathTo                           As String

    ' Извлекаем путь из параметра
    strPathToTemp = strArg

    ' Проверяем существоание каталога
    If LenB(strPathToTemp) > 0 Then
        If PathFileExists(strPathToTemp) = 0 Then
            CreateNewDirectory strPathToTemp
        End If

        strPathTo = BackslashAdd2Path(strPathToTemp)
    Else
        strPathTo = strWorkTemp
    End If

    ' Запуск извлечения всех (dll-ocx) ресурсов программы
    If ExtractResourceAll(strPathTo) Then
        If MsgBox(strMessages(135), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If
    Else
        If MsgBox(strMessages(136), vbYesNo + vbInformation, strProductName) = vbYes Then
            ShellEx strPathTo, essSW_SHOWNORMAL
        End If
    End If

End Sub



'HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA
'"Blob"=hex:03,00,00,00,01,00,00,00,14,00,00,00,a3,1d,3e,0a,4d,99,33,5e,bd,9b,\
'  6f,18,e0,91,54,90,f1,35,25,ca,20,00,00,00,01,00,00,00,28,02,00,00,30,82,02,\
'  24,30,82,01,91,a0,03,02,01,02,02,10,82,58,85,44,28,61,9e,bc,48,c0,05,a4,40,\
'  6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,30,1f,31,1d,30,1b,06,03,55,04,03,\
'  13,14,77,77,77,2e,61,64,69,61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,1e,17,\
'  0d,31,33,30,33,31,31,30,39,35,37,34,30,5a,17,0d,33,39,31,32,33,31,32,33,35,\
'  39,35,39,5a,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,61,\
'  2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,81,9f,30,0d,06,09,2a,86,48,86,f7,0d,\
'  01,01,01,05,00,03,81,8d,00,30,81,89,02,81,81,00,c4,4e,f8,78,d3,eb,fc,45,49,\
'  13,31,a0,fc,f6,50,1d,3c,b3,4b,9e,d5,73,45,4c,06,93,70,e7,ee,c8,6b,25,82,16,\
'  4b,58,ea,22,40,ab,82,d7,c7,c9,90,0c,31,45,aa,7f,79,27,e6,b5,47,fe,7d,48,ad,\
'  70,e6,9a,46,25,64,0b,50,74,ce,ea,f1,8c,92,6c,82,2e,08,4b,aa,a8,10,05,d1,e8,\
'  9b,9b,fb,ce,79,3e,42,a4,49,88,03,c8,22,6f,b6,21,a2,3f,68,f2,84,5d,ac,29,a5,\
'  02,71,87,6d,81,ec,e3,d0,17,be,cf,48,58,a3,ab,ed,f5,9d,5f,02,03,01,00,01,a3,\
'  69,30,67,30,13,06,03,55,1d,25,04,0c,30,0a,06,08,2b,06,01,05,05,07,03,03,30,\
'  50,06,03,55,1d,01,04,49,30,47,80,10,01,60,4c,5b,6f,d2,c8,c6,60,6b,50,24,03,\
'  4b,9b,a7,a1,21,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,\
'  61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,82,10,82,58,85,44,28,61,9e,bc,48,c0,\
'  05,a4,40,6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,03,81,81,00,08,a6,57,6e,\
'  3c,a5,7c,ad,41,ab,61,f9,8f,41,0e,6e,e0,b2,6e,bd,35,16,cc,0c,05,d1,e2,d9,d4,\
'  b2,71,50,70,fd,28,a0,c7,7f,8f,23,63,4a,c4,e0,1b,0e,98,37,c1,24,1f,4f,ae,ae,\
'  db,8d,ce,b8,cb,9e,13,6e,b0,a8,b0,0f,90,1b,22,94,97,fa,47,b6,29,b1,eb,98,4a,\
'  26,28,23,a5,0a,ef,59,43,b1,be,25,49,2b,cf,8d,bc,82,37,20,cd,b7,db,90,0b,d7,\
'  3d,7b,e9,f5,87,7b,87,bb,ae,f2,53,de,5d,17,72,25,18,f9,61,bd,4e,cd,6c,c8

' процедура прописывания сертификата для проверки валидности цифровой подписи моего exe
Private Sub SaveSert2Reestr()
Dim strBuffer                           As String
Dim strBuffer_x()                       As String
Dim strByteArray()                      As Byte
Dim i                                   As Long

    On Error Resume Next

    strBuffer = "03,00,00,00,01,00,00,00,14,00,00,00,a3,1d,3e,0a,4d,99,33,5e,bd,9b," & _
                "6f,18,e0,91,54,90,f1,35,25,ca,20,00,00,00,01,00,00,00,28,02,00,00,30,82,02," & _
                "24,30,82,01,91,a0,03,02,01,02,02,10,82,58,85,44,28,61,9e,bc,48,c0,05,a4,40," & _
                "6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,30,1f,31,1d,30,1b,06,03,55,04,03," & _
                "13,14,77,77,77,2e,61,64,69,61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,1e,17," & _
                "0d,31,33,30,33,31,31,30,39,35,37,34,30,5a,17,0d,33,39,31,32,33,31,32,33,35," & _
                "39,35,39,5a,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,61," & _
                "2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,81,9f,30,0d,06,09,2a,86,48,86,f7,0d," & _
                "01,01,01,05,00,03,81,8d,00,30,81,89,02,81,81,00,c4,4e,f8,78,d3,eb,fc,45,49," & _
                "13,31,a0,fc,f6,50,1d,3c,b3,4b,9e,d5,73,45,4c,06,93,70,e7,ee,c8,6b,25,82,16," & _
                "4b,58,ea,22,40,ab,82,d7,c7,c9,90,0c,31,45,aa,7f,79,27,e6,b5,47,fe,7d,48,ad," & _
                "70,e6,9a,46,25,64,0b,50,74,ce,ea,f1,8c,92,6c,82,2e,08,4b,aa,a8,10,05,d1,e8," & _
                "9b,9b,fb,ce,79,3e,42,a4,49,88,03,c8,22,6f,b6,21,a2,3f,68,f2,84,5d,ac,29,a5," & _
                "02,71,87,6d,81,ec,e3,d0,17,be,cf,48,58,a3,ab,ed,f5,9d,5f,02,03,01,00,01,a3," & _
                "69,30,67,30,13,06,03,55,1d,25,04,0c,30,0a,06,08,2b,06,01,05,05,07,03,03,30," & _
                "50,06,03,55,1d,01,04,49,30,47,80,10,01,60,4c,5b,6f,d2,c8,c6,60,6b,50,24,03," & _
                "4b,9b,a7,a1,21,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69," & _
                "61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,82,10,82,58,85,44,28,61,9e,bc,48,c0," & _
                "05,a4,40,6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,03,81,81,00,08,a6,57,6e," & _
                "3c,a5,7c,ad,41,ab,61,f9,8f,41,0e,6e,e0,b2,6e,bd,35,16,cc,0c,05,d1,e2,d9,d4," & _
                "b2,71,50,70,fd,28,a0,c7,7f,8f,23,63,4a,c4,e0,1b,0e,98,37,c1,24,1f,4f,ae,ae," & _
                "db,8d,ce,b8,cb,9e,13,6e,b0,a8,b0,0f,90,1b,22,94,97,fa,47,b6,29,b1,eb,98,4a," & _
                "26,28,23,a5,0a,ef,59,43,b1,be,25,49,2b,cf,8d,bc,82,37,20,cd,b7,db,90,0b,d7," & _
                "3d,7b,e9,f5,87,7b,87,bb,ae,f2,53,de,5d,17,72,25,18,f9,61,bd,4e,cd,6c,c8"

    strBuffer_x = Split(strBuffer, ",")
    ReDim strByteArray(UBound(strBuffer_x))
    For i = LBound(strBuffer_x) To UBound(strBuffer_x)
        strByteArray(i) = CLng("&H" & strBuffer_x(i))
    Next

    SetRegBin HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA", "Blob", strByteArray

End Sub

Private Sub InitializePathHwidsTxt()
    ' Служебные файлы
    strHwidsTxtPath = strWorkTempBackSL & "HWIDS.txt"
    strHwidsTxtPathView = strWorkTempBackSL & "HWIDS_ForView.txt"
    strHwidsTxtPathVersion = strWorkTempBackSL & "HWIDS_Version.txt"
    strHwidsTxtPathDRVFiles = strWorkTempBackSL & "HWIDS_DRVFiles.txt"
    strResultHwidsTxtPath = strWorkTempBackSL & "HwidsTemp.txt"
    strResultHwidsExtTxtPath = strWorkTempBackSL & "HwidsTempExt.txt"
End Sub

' Функция анализа коммандной строки и присвоение переменных на основании передеваемых комманд
Private Sub cmdLineAnalize()
Dim miSilentRunTimerTemp                As String
Dim strRunWithParam_x()                 As String
Dim strRunWithParamTemp                 As String
Dim strSilentSelectModeTemp             As String
Dim i                                   As Long

    DebugMode "CmdString: " & Command
    strRunWithParam = Trim$(strRunWithParam)
    strRunWithParam_x = Split(strRunWithParam, " ")

    For i = LBound(strRunWithParam_x) To UBound(strRunWithParam_x)
        strRunWithParamTemp = strRunWithParam_x(i)

        If InStr(1, strRunWithParamTemp, "-t", vbTextCompare) = 1 Or InStr(1, strRunWithParamTemp, "t", vbTextCompare) = 1 Then
            mbRunWithParam = True
            miSilentRunTimerTemp = Replace$(strRunWithParamTemp, "-", vbNullString)
            miSilentRunTimerTemp = Replace$(miSilentRunTimerTemp, "t", vbNullString)

            If IsNumeric(miSilentRunTimerTemp) Then
                miSilentRunTimer = CInt(miSilentRunTimerTemp)
            Else
                miSilentRunTimer = 10
            End If

            mbDebugEnable = True
            mbUpdateCheck = False
        End If

        If InStr(1, strRunWithParamTemp, "-s", vbTextCompare) = 1 Or InStr(1, strRunWithParamTemp, "s", vbTextCompare) = 1 Then
            mbRunWithParamS = True
            strSilentSelectModeTemp = Replace$(strRunWithParamTemp, "-", vbNullString)
            strSilentSelectModeTemp = Replace$(strSilentSelectModeTemp, "s", vbNullString)

            Select Case LCase$(strSilentSelectModeTemp)

                Case "n"
                    ' новые
                    strSilentSelectMode = "n"

                Case "q"
                    ' неустановленные
                    strSilentSelectMode = "q"

                Case "a"
                    ' Все на вкладке
                    strSilentSelectMode = "a"

                Case "n2"
                    ' новые
                    strSilentSelectMode = "n2"

                Case "q2"
                    ' неустановленные
                    strSilentSelectMode = "q2"

                Case "a2"
                    ' Все на вкладке
                    strSilentSelectMode = "a2"

                Case Else
                    ' по умолчанию
                    strSilentSelectMode = "n"
            End Select

            mbDebugEnable = True
            mbUpdateCheck = False
        Else
            strSilentSelectMode = "n"
        End If

    Next

    ' Если стоит только параметр -s и нет -t, то делаем -t10
    If mbRunWithParamS Then
        If Not mbRunWithParam Then
            mbRunWithParam = True
            miSilentRunTimer = 10
        End If
    End If
End Sub
