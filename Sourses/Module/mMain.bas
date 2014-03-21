Attribute VB_Name = "mMain"
Option Explicit

'Основные параметры программы
Public Const strDateProgram         As String = "21/03/2014"

'Основные переменные проекта (название, версия и т.д)
Public strProductName               As String
Public strProductVersion            As String
Public Const strProjectName         As String = "DriversInstaller"
Public Const strUrl_MainWWWSite     As String = "http://adia-project.net/"                   ' Домашний сайт проекта
Public Const strUrl_MainWWWForum    As String = "http://adia-project.net/forum/index.php"    ' Домашний форум проекта
Public Const strUrlOsZoneNetThread  As String = "http://forum.oszone.net/thread-139908.html" ' Топик программы на сайте Oszone.net

'Константы путей основных каталогов и файла настроек (вынесены отдельно для универсальности кода под разные проекты)
Public Const strToolsLang_Path      As String = "Tools\DIA\Lang"            ' Каталог с языковыми файлами
Public Const strToolsDocs_Path      As String = "Tools\DIA\Docs"            ' Каталог с документацией на программу
Public Const strToolsGraphics_Path  As String = "Tools\DIA\Graphics"        ' Каталог с графическими ресурсами программы
Public Const strSettingIniFile      As String = "DIA.ini"  ' INI-Файл настроек программы

' Версии лицензионного соглашения и файла Donate
Public Const strEULA_Version        As String = "02/02/2010"
Public Const strEULA_MD5RTF         As String = "68da44c8b1027547e4763472e0ecb727"
Public Const strEULA_MD5RTF_Eng     As String = "0cbd9d50eec41b26d24c5465c4be70bc"
Public Const strDONATE_MD5RTF       As String = "637e1aacdfcfa01fdc8827eb48796b1b"
Public Const strDONATE_MD5RTF_Eng   As String = "ca762ec290f0d9bedf2e09319661921a"

'Константы путей дополнительных утилит
Public Const strDevManView_Path     As String = "Tools\DevManView\DevManView.exe"
Public Const strDevManView_Path64   As String = "Tools\DevManView\DevManView-x64.exe"
Public Const strSIV_Path            As String = "Tools\SIV\SIV32X.exe"
Public Const strSIV_Path64          As String = "Tools\SIV\SIV64X.exe"
Public Const strUDI_Path            As String = "Tools\UDI\UnknownDeviceIdentifier.exe"
Public Const strDoubleDriver_Path   As String = "Tools\DoubleDriver\dd.exe"
Public Const strUnknownDevices_Path As String = "Tools\UnknownDevices\UnknownDevices.exe"

'Описание структуры массива информации по HWID
Public Type arrHwidsStruct
    HWID                            As String           ' HWID устройства (обрезанный без "сессии" устройства)
    HWIDOrig                        As String           ' HWID устройства полный
    HWIDCutting                     As String           ' HWID устройства обрезанный до первой /
    HWIDCompat                      As String           ' HWID Совместимый (другой вариант написания)
    HWIDMatches                     As String           ' HWID Список подходящих (другие варианты написания в порядке совместимости)
    DevName                         As String           ' Имя устройства
    Provider                        As String           ' Производитель драйвера устройства
    Status                          As Long             ' Статус устройства
    VerLocal                        As String           ' Версия драйвера устройства
    Description                     As String           ' Описание
    PriznakSravnenia                As String           ' Результат сравнения драйверов по дате с базой индексов
    InfSection                      As String           ' Секция inf-файла в которой найден HWID (используется для анализа совместимости)
    InfName                         As String           ' Имя inf-файла драйвера
    DPsList                         As String           ' Список пакетов драйверов в котором есть подходящий драйвер
    DRVScore                        As Long             ' Балл найденного драйвера
End Type

'Описание структуры массива для поддерживаемой ОС
Public Type arrOSStruct
    Ver                             As String           ' Версия ОС
    Name                            As String           ' Имя ОС
    is64bit                         As Long             ' 64-битная ОС
    drpFolder                       As String           ' Каталог с пакетами драйверов (относительный путь)
    drpFolderFull                   As String           ' Каталог с пакетами драйверов (полный путь)
    devIDFolder                     As String           ' Каталог с базой индексов (относительный путь)
    devIDFolderFull                 As String           ' Каталог с базой индексов  (полный путь)
    DPFolderNotExist                As Boolean          ' Каталог не сущестует
    PathPhysX                       As String           ' Путь до файла Physx
    PathLanguages                   As String           ' Путь до файла Languages
    PathRuntimes                    As String           ' Путь до файла Runtimes
    CntBtn                          As Long             ' Количество пакетов в текущей ОС
    ExcludeFileName                 As String           ' Исключаемые имена пакетов драйверов
End Type

'Массивы данных
Public arrHwidsLocal()              As arrHwidsStruct   ' Массив информации о драйверах устройств
Public arrOSList()                  As arrOSStruct      ' Массив поддерживаемых ОС
Public arrTTipStatusIcon()          As String           ' Массив статусных сообщений - подсказки к картинкам
Public arrCheckDP()                 As String           ' Массив выделенных пакетов драйверов
Public arrUtilsList()               As String           ' Массив настраиваемых утилит
Public arrTTip()                    As String           ' Массив подсказок для пакетов драйверов
Public arrTTipSize()                As String           ' Массив размеров колонок в таблице для подсказки
Public arrDevIDs()                  As String           ' Меню для удаления драйверов устройств
Public arrDriversList()             As String           ' Меню для копирования HWID драйверов устройств
Public lngArrDriversListCountMax    As Long             ' Максимальная размерность массива HWID
Public lngArrDriversIndex           As Long             ' Текущий максимальный индекс массива HWID

'Пути до системных каталогов и других рабочих файлов
Public strHwidsTxtPath              As String
Public strHwidsTxtPathView          As String
Public strResultHwidsTxtPath        As String
Public strResultHwidsExtTxtPath     As String
Public strWorkTemp                  As String           ' Рабочий временный каталог
Public strWorkTempBackSL            As String           ' Рабочий временный каталог   + \
Public strWinTemp                   As String           ' Системный временный каталог + \
Public strWinDir                    As String           ' Системный каталог Windows   + \
Public strSysDir                    As String           ' Системный каталог System32  + \
Public strSysDir64                  As String           ' Системный каталог Windows\System32  + \
Public strSysDir86                  As String           ' Системный каталог Windows\Wow64  + \
Public strSysDirCatRoot             As String           ' c:\Windows\System32\catroot\
Public strSysDirDrivers             As String           ' Системный каталог Windows\System32\drivers  + \
Public strSysDirDrivers64           As String           ' Системный каталог Windows\Wow64\drivers  + \
Public strSysDirDRVStore            As String           ' Системный каталог System32\DriverStore\
Public strSysDrive                  As String           ' Системный диск
Public strWinDirHelp                As String           ' c:\Windows\Help\
Public strInfDir                    As String           ' c:\Windows\inf\

'Переменные и маркеры используемые в коде программы
Public mbFirstStart                 As Boolean          ' Флаг указывающий окочание запуска программы
Public mbIsDriveCDRoom              As Boolean          ' Флаг, указывающий что рабочий диск является CDRoom
Public mbAddInList                  As Boolean          ' режим работы с элементом listview - либо изменние либо добавление, для обмена между формами frmOptions,frmOSEdit,frmUtilsEdit
Public lngLastIdOS                  As Long             ' номер последнего элемента в списке ОС, для обмена между формами frmOptions и frmOSEdit
Public lngLastIdUtil                As Long             ' номер последнего элемента в списке утилит
Public lngCurrentBtnIndex           As Long             ' Текущая выделенная кнопка
Public strPathDRPList               As String           ' Список папок для распаковки
Public mbooSelectInstall            As Boolean          ' Флаг указывающий выборочную установку
Public mbCheckDRVOk                 As Boolean          ' Флаг, указывающий нажатие кнопки ОК на форме frmListHwid
Public mbGroupTask                  As Boolean          ' Флаг указывающий групповую задачу
Public mbRestartProgram             As Boolean          ' Маркер перезапуска программы
Public mbOnlyUnpackDP               As Boolean          ' Переменная для определения режима - только распаковка драйверов
Public mbDeleteDriverByHwid         As Boolean          ' Флаг сообщает о том что драйвер был удален на форме frmListHwidAll
Public strCompModel                 As String           ' Модель компьютера/материнской платы
Public strFrmMainCaptionTemp        As String           ' кэпшн основной формы
Public strFrmMainCaptionTempDate    As String           ' кэпшн основной формы - дата релиза программы

'Заголовки таблицы для подсказок
Public strTableHwidHeader1          As String           ' "-HWID-"
Public strTableHwidHeader2          As String           ' "-Путь-"
Public strTableHwidHeader3          As String           ' "-Файл-"
Public strTableHwidHeader4          As String           ' "-Версия(БД)-"
Public strTableHwidHeader5          As String           ' "-Версия(PC)-"
Public strTableHwidHeader6          As String           ' "-Статус-"
Public strTableHwidHeader7          As String           ' "-Наименование устройства-"
Public strTableHwidHeader8          As String           ' "-Пакет драйверов-"
Public strTableHwidHeader9          As String           ' "!"
Public strTableHwidHeader10         As String           ' "-Производитель-"
Public strTableHwidHeader11         As String           ' "-Совместимые HWID-"
Public strTableHwidHeader12         As String           ' "-Код устройства-"
Public strTableHwidHeader13         As String           ' "-Секция-"
Public strTableHwidHeader14         As String           ' "-Найден в пакете-"
'Размеры заголовков таблицы для подсказок, высчитываем как Len()
Public lngTableHwidHeader1          As Long
Public lngTableHwidHeader2          As Long
Public lngTableHwidHeader3          As Long
Public lngTableHwidHeader4          As Long
Public lngTableHwidHeader5          As Long
Public lngTableHwidHeader6          As Long
Public lngTableHwidHeader7          As Long
Public lngTableHwidHeader8          As Long
Public lngTableHwidHeader9          As Long
Public lngTableHwidHeader10         As Long
Public lngTableHwidHeader11         As Long
Public lngTableHwidHeader12         As Long
Public lngTableHwidHeader13         As Long
Public lngTableHwidHeader14         As Long
'Максимальные значения размеров колонок в всплывающем сообщении
Public lngSizeRowDPMax              As Long
Public lngSizeRow1Max               As Long
Public lngSizeRow2Max               As Long
Public lngSizeRow3Max               As Long
Public lngSizeRow4Max               As Long
Public lngSizeRow5Max               As Long
Public lngSizeRow6Max               As Long
Public lngSizeRow9Max               As Long
Public lngSizeRow13Max              As Long
Public maxSizeRowAllLineMax         As Long
'Дефолтные значения размеров колонок в всплывающем сообщении
'Расчитываются при старте исходя из длины наименования колонки
Public lngSizeRow1                  As Long
Public lngSizeRow2                  As Long
Public lngSizeRow3                  As Long
Public lngSizeRow4                  As Long
Public lngSizeRow5                  As Long
Public lngSizeRow6                  As Long
Public lngSizeRow9                  As Long
Public lngSizeRow13                 As Long
Public maxSizeRowAllLine            As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Main
'! Description (Описание)  :   [Основная функция запуска программы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Main()

    Dim mbShowFormLicence As Boolean
    Dim strSysIniTMP      As String
    Dim strLicenceDate    As String  ' дата лицензионного соглашения из реестра
    Dim mbIsUserAnAdmin   As Boolean ' Пользователь администратор?

    On Error Resume Next

    dtStartTimeProg = GetTickCount
    Set objFSO = New Scripting.FileSystemObject

    ' Запоминаем app.path и прочее в переменные
    GetMyAppProperties

    'считываем версию операционки
    If Not OSCurrVersionStruct.IsInitialize Then
        OSCurrVersionStruct = OSInfo
    End If

    'Получаем временный каталог windows и каталог windows
    strWinDir = BackslashAdd2Path(Environ$("WINDIR"))
    strWinTemp = BackslashAdd2Path(Environ$("TMP"))

    If InStr(strWinTemp, " ") Then
        strWinTemp = BackslashAdd2Path(PathCombine(strWinDir, "TEMP"))
    End If

    ' Если временный каталог windows  (%windir%\temp)недоступен
    If PathExists(strWinTemp) = False Then
        MsgBox "Windows TempPath not Exist or Environ %TMP% undefined. Program is exit!!!", vbInformation, strProductName

        'End
        GoTo ExitSub

    End If

    ' Инициализация массива вендоров ноутбуков
    LoadNotebookList
    'Получение значений маркеров
    GetSummaryDPMarkers

    '******************************************
    ' Проверяем работает ли программа в режиме IDE
    ' Программа уже запущена???
    If App.PrevInstance And Not InIDE() Then
        MsgBoxEx "Found a running application 'Drivers Installer Assistant'. If you restart the program from the settings menu, then save the settings, the program waits until the previous session..." & str2vbNewLine & _
                                    "This window will close automatically in 5 seconds. Please wait or click OK", vbExclamation + vbSystemModal, strProductName, 6
        ShowPrevInstance
    Else
        '******************************************
        ' - Инициализируем стиль WindowsXP
        Call ComCtlsInitIDEStopProtection
        Call InitVisualStyles
    End If

    ' Если каталог tools недоступен
    If PathExists(strAppPathBackSL & "Tools\") = False Then
        MsgBox "Not found the main program subfolder '.\Tools'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName

        'End
        GoTo ExitSub

    End If
    
    ' Если каталог tools недоступен
    If PathExists(strAppPathBackSL & "Tools\DIA\") = False Then
        MsgBox "Not found the main program subfolder '.\Tools\DIA'." & vbNewLine & "Program is exit!!!", vbInformation, strProductName

        'End
        GoTo ExitSub

    End If

    ' Рабочий временный каталог
    strWorkTemp = strWinTemp & strProjectName
    strWorkTempBackSL = BackslashAdd2Path(strWorkTemp)

    ' Создаем временный рабочий каталог
    If PathExists(strAppPathBackSL & strSettingIniFile) = False Then
        strSysIni = strAppPathBackSL & "Tools\" & strSettingIniFile
    Else
        strSysIni = strAppPathBackSL & strSettingIniFile
    End If

    ' Запущена ли программа с CD
    mbIsDriveCDRoom = IsDriveCDRoom
    ' Создаем файл настроек при необходимости
    CreateIni
    ' Считаваем язык операционки
    LoadLanguageOS

    'загружаем языковые файлы
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    'загружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    ' Получение настроек из ini-файла
    GetMainIniParam

    ' Если стоит настройка проверять временный путь на наличие ini, то перезагружаем файл параметров
    If mbLoadIniTmpAfterRestart Then
        If GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP", False) Then
            ' Reload Main ini
            strSysIniTMP = GetSetting(App.ProductName, "Settings", "LOAD_INI_TMP_PATH", vbNullString)

            If LenB(strSysIniTMP) Then
                If PathExists(strSysIniTMP) Then
                    strSysIni = strSysIniTMP
                    ' Собственно перезагрузка настроек
                    GetMainIniParam
                End If
            End If
        End If
    End If

    If PathExists(strWorkTemp) = False Then
        CreateNewDirectory strWorkTemp
    End If

    'Перегружаем языковые файлы
    If PathExists(strAppPathBackSL & strToolsLang_Path) Then
        mbMultiLanguage = LoadLanguageList
    End If

    'перегружаем программные сообщения
    LocaliseMessage strPCLangCurrentPath
    strPathImageStatusButton = strAppPathBackSL & strToolsGraphics_Path & "\StatusButton\"
    strPathImageMain = strAppPathBackSL & strToolsGraphics_Path & "\Main\"
    'strPathImageMenu = strAppPathBackSL & strToolsGraphics_Path & "\Menu\"
    LoadIconImagePath
    ' Очищаем лог-историю
    MakeCleanHistory
    ' Получаем размеры рабочей области программы
    GetWorkArea
    
    ' Проверяем на запуск с параметрами
    If LenB(Command) Then
        ' Запуск парсинга строки запуска
        CmdLineParsing
    End If

    If APIFunctionPresent("IsUserAnAdmin", "shell32.dll") Then
        mbIsUserAnAdmin = IsUserAnAdmin
    Else
        mbIsUserAnAdmin = True
    End If

    If Not mbDebugTime2File Then
        If mbDebugStandart Then DebugMode "Current Date: " & Now()
    End If

    If mbDebugStandart Then DebugMode "Version: " & strProductName & vbNewLine & _
              "Build: " & strDateProgram & vbNewLine & _
              "ExeName: " & strAppEXEName & ".exe" & vbNewLine & _
              "AppWork: " & strAppPath & vbNewLine & _
              "is User an Admin?: " & mbIsUserAnAdmin

    If mbIsUserAnAdmin Then
        ' записываем в реестр мой сертификат, для ЭЦП на exe-файлы
        If mbDebugStandart Then DebugMode "SaveSert2Reestr"
        SaveSert2Reestr
    Else

        If Not mbRunWithParam Then
            If MsgBox(strMessages(138), vbYesNo + vbQuestion, strProductName) = vbNo Then

                End

            End If
        End If
    End If

    If mbDebugStandart Then DebugMode "WinDir: " & strWinDir & vbNewLine & _
              "TmpDir: " & strWinTemp & vbNewLine & _
              "WorkTemp: " & strWorkTemp & vbNewLine & _
              "IsDriveCDRoom: " & mbIsDriveCDRoom

    If strOSCurrentVersion > "5.0" Then
        ' Определение windows x64
        mbIsWin64 = IsWow64
        If mbDebugStandart Then DebugMode "IsWow64: " & mbIsWin64

        If mbIsWin64 Then
            Win64ReloadOptions
        End If

    ElseIf strOSCurrentVersion = "5.0" Then
        ' Для win2k надо старый devcon
        strDevConExePath = strDevConExePathW2k
    End If

    ' Disable DEP for current process
    If mbDisableDEP Then
        SetDEPDisable
    End If

    ' Регистрация внешних компонент
    RegisterAddComponent

    If mbDebugStandart Then DebugMode "OsCurrentVersion: " & strOSCurrentVersion & vbNewLine & _
              "Architecture: " & strOSArchitecture & vbNewLine & _
              "OS Language: ID=" & strPCLangID & " Name=" & strPCLangEngName & "(" & strPCLangLocaliseName & ")"

    ' Служебные файлы
    InitializePathHwidsTxt

    ' Если не существует каталогов с драйверами прописанных в настройках, то выводим сообщение
    If mbAllFolderDRVNotExist Then
        MsgBox strMessages(6), vbCritical + vbApplicationModal, strProductName
        If mbDebugStandart Then DebugMode strMessages(6)

        'End
        GoTo ExitSub

    End If

    If APIFunctionPresent("IsAppThemed", "uxtheme.dll") Then
        mbAppThemed = IsAppThemed
        If mbDebugStandart Then DebugMode "IsAppThemed: " & mbAppThemed
    End If

    mbAeroEnabled = IsAeroEnabled
    If mbDebugStandart Then DebugMode "IsAeroEnabled : " & mbAeroEnabled
    ' изменяем разрешающую способность экрана монитора при необходимости
    SetVideoMode
    GetWorkArea
    
    ' Получаем имя производителя материнской платы/ноутбука
    strCompModel = GetMBInfo()
    If mbDebugStandart Then DebugMode "isNotebook: " & mbIsNotebok & vbNewLine & _
              "Notebook/Motherboard Model: " & strCompModel
              
    ' Маркер указывающий что это "первый" запуск программы, нужен для события активации формы и других процедур
    mbFirstStart = True
    
    ' Если запуск программы произошел не с параметрами, то....
    If Not mbRunWithParam Then
        ' Показ лицензионного соглашения
        strLicenceDate = GetSetting(App.ProductName, "Licence", "EULA_DATE", strEULA_Version)
        mbShowFormLicence = GetSetting(App.ProductName, "Licence", "Show at Startup", True)
        If mbShowFormLicence Then
            If Not mbEULAAgree Then
                mbShowFormLicence = StrComp(strLicenceDate, strEULA_Version) <> 0
            End If
        End If
        
        ' Если не поддерживается всыпылваюющие подсказки
        If Not CheckBallonTip Then
            If MsgBox(strMessages(9), vbYesNo + vbQuestion, strMessages(10)) = vbNo Then
                GoTo ExitSub
            End If
        End If
    End If
    
    'Because Ambient.UserMode does not report IDE behavior properly, we use our own UserMode tracker.  Many thanks to
    ' Kroc of camendesign.com for suggesting this fix.
    g_UserModeFix = True

    If mbShowFormLicence Then
        'Открываем форму лицензионного соглашения
        frmLicence.Show
    Else
        'Открываем основную форму
        frmMain.Show vbModeless
    End If

ExitSub:

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ChangeStatusTextAndDebug
'! Description (Описание)  :   [Изменение текста статустной строки и отладочной информации]
'! Parameters  (Переменные):   strPanel2Text (String)
'                              strDebugText (String)
'                              mbEqual (Boolean = False)
'                              mbDoEvents (Boolean = True)
'                              strPanel1Text (String)
'!--------------------------------------------------------------------------------
Public Sub ChangeStatusTextAndDebug(ByVal strPanel2Text As String, Optional ByVal strPanel1Text As String = vbNullString, Optional ByVal mbDoEvents As Boolean = True)

    If LenB(strPanel2Text) Then

        If frmMain.ctlUcStatusBar1.PanelCount >= 2 Then
            frmMain.ctlUcStatusBar1.PanelText(2) = strPanel2Text
        Else
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel2Text
        End If

        If LenB(strPanel1Text) Then
            frmMain.ctlUcStatusBar1.PanelText(1) = strPanel1Text
        End If
        
        If mbDoEvents Then
            DoEvents
        End If
    End If
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SaveSert2Reestr
'! Description (Описание)  :   [процедура прописывания сертификата для проверки валидности цифровой подписи моего exe]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub SaveSert2Reestr()

    Dim strBuffer      As String
    Dim strBuffer_x()  As String
    Dim strByteArray() As Byte
    Dim i              As Long

    On Error Resume Next
    
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
'

    strBuffer = "03,00,00,00,01,00,00,00,14,00,00,00,a3,1d,3e,0a,4d,99,33,5e,bd,9b," & "6f,18,e0,91,54,90,f1,35,25,ca,20,00,00,00,01,00,00,00,28,02,00,00,30,82,02," & "24,30,82,01,91,a0,03,02,01,02,02,10,82,58,85,44,28,61,9e,bc,48,c0,05,a4,40," & _
                                "6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,30,1f,31,1d,30,1b,06,03,55,04,03," & "13,14,77,77,77,2e,61,64,69,61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,1e,17," & _
                                "0d,31,33,30,33,31,31,30,39,35,37,34,30,5a,17,0d,33,39,31,32,33,31,32,33,35," & "39,35,39,5a,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69,61," & _
                                "2d,70,72,6f,6a,65,63,74,2e,6e,65,74,30,81,9f,30,0d,06,09,2a,86,48,86,f7,0d," & "01,01,01,05,00,03,81,8d,00,30,81,89,02,81,81,00,c4,4e,f8,78,d3,eb,fc,45,49," & _
                                "13,31,a0,fc,f6,50,1d,3c,b3,4b,9e,d5,73,45,4c,06,93,70,e7,ee,c8,6b,25,82,16," & "4b,58,ea,22,40,ab,82,d7,c7,c9,90,0c,31,45,aa,7f,79,27,e6,b5,47,fe,7d,48,ad," & _
                                "70,e6,9a,46,25,64,0b,50,74,ce,ea,f1,8c,92,6c,82,2e,08,4b,aa,a8,10,05,d1,e8," & "9b,9b,fb,ce,79,3e,42,a4,49,88,03,c8,22,6f,b6,21,a2,3f,68,f2,84,5d,ac,29,a5," & _
                                "02,71,87,6d,81,ec,e3,d0,17,be,cf,48,58,a3,ab,ed,f5,9d,5f,02,03,01,00,01,a3," & "69,30,67,30,13,06,03,55,1d,25,04,0c,30,0a,06,08,2b,06,01,05,05,07,03,03,30," & _
                                "50,06,03,55,1d,01,04,49,30,47,80,10,01,60,4c,5b,6f,d2,c8,c6,60,6b,50,24,03," & "4b,9b,a7,a1,21,30,1f,31,1d,30,1b,06,03,55,04,03,13,14,77,77,77,2e,61,64,69," & _
                                "61,2d,70,72,6f,6a,65,63,74,2e,6e,65,74,82,10,82,58,85,44,28,61,9e,bc,48,c0," & "05,a4,40,6f,ce,eb,30,09,06,05,2b,0e,03,02,1d,05,00,03,81,81,00,08,a6,57,6e," & _
                                "3c,a5,7c,ad,41,ab,61,f9,8f,41,0e,6e,e0,b2,6e,bd,35,16,cc,0c,05,d1,e2,d9,d4," & "b2,71,50,70,fd,28,a0,c7,7f,8f,23,63,4a,c4,e0,1b,0e,98,37,c1,24,1f,4f,ae,ae," & _
                                "db,8d,ce,b8,cb,9e,13,6e,b0,a8,b0,0f,90,1b,22,94,97,fa,47,b6,29,b1,eb,98,4a," & "26,28,23,a5,0a,ef,59,43,b1,be,25,49,2b,cf,8d,bc,82,37,20,cd,b7,db,90,0b,d7," & _
                                "3d,7b,e9,f5,87,7b,87,bb,ae,f2,53,de,5d,17,72,25,18,f9,61,bd,4e,cd,6c,c8"
    strBuffer_x = Split(strBuffer, ",")

    ReDim strByteArray(UBound(strBuffer_x))

    For i = LBound(strBuffer_x) To UBound(strBuffer_x)
        strByteArray(i) = CLng("&H" & strBuffer_x(i))
    Next

    SetRegBin HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\SystemCertificates\ROOT\Certificates\A31D3E0A4D99335EBD9B6F18E0915490F13525CA", "Blob", strByteArray
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub Win64ReloadOptions
'! Description (Описание)  :   [Переназначение переменных для Win x64]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub Win64ReloadOptions()
    If mbDebugStandart Then DebugMode "Win64ReloadOptions"
    strDPInstExePath = strDPInstExePath64
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub InitializePathHwidsTxt
'! Description (Описание)  :   [Служебные файлы]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Sub InitializePathHwidsTxt()
    strHwidsTxtPath = strWorkTempBackSL & "HWIDS.txt"
    strHwidsTxtPathView = strWorkTempBackSL & "HWIDS_ForView.txt"
    strResultHwidsTxtPath = strWorkTempBackSL & "HwidsTemp.txt"
    strResultHwidsExtTxtPath = strWorkTempBackSL & "HwidsTempExt.txt"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckBallonTip
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function CheckBallonTip() As Boolean
    regParam = GetKeyValue(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "EnableBalloonTips")

    If LenB(regParam) = 0 Then
        CheckBallonTip = True
    Else
        CheckBallonTip = regParam = "1"
    End If

    If mbDebugStandart Then DebugMode "EnableBalloonTips: " & regParam & "(" & CheckBallonTip & ")"
End Function
