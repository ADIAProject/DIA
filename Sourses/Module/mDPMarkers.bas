Attribute VB_Name = "mDPMarkers"
'============================================================================
'Новая структура драйвер-паков
'
'Старый формат драйвер-паков подразумевал имена файлов DP_NAME_wnt6-x64_DATE.7z, типа, DP_Bluetooth_wnt6-x64_1210.7z. Приставка wnt5_x86-32 означала Windows XP x86, wnt6-x64 - Windows Vista/7/8 x64, wnt6-x86 - Windows Vista/7/8 x86 и располагались они соответственно в 3х разных папках. Но с увеличением числа систем на базе технологии NT и ядре 6.x, в каждом драйвер-паке может находиться минимум по 3 драйвера для каждой системы Vista/7/8 не учитывая серверные платформы, плюс некоторые могут быть универсальными и они в каждом драйвер-паке дублируются, а т.к. DPS не умеет сравнивать драйвера в разных драйвер-паках, то подбор нужного драйвера происходит не оптимально и могут предлагаться к установке в лучшем случае дублирующиеся драйвера, а в версиях DPS до билда 269 предлагались и вообще драйвера от другой системы (например для Win8 от Win7).
'
'Понятно одно, что структуру старых драйвер-паков пришло время реорганизовывать в соответствие с текущими реалиями и новая структура будущих драйвер-паков будет выглядеть следующим образом:
'
'1. Название
'Общий формат имени драйвер-паков будет следующим: DP_NAME_DATE.7z - подробнее:
'DP - сокращение от слова DriverPack (драйвер-пак - набор драйверов в одном общем архиве)
'NAME - имя драйвер-пака по типу драйверов в них и по их бренду: “Bluetooth” или “Sound_Realtek”
'- первый случай, когда драйверов для какого-то типа устройства много и они небольшие по размеру: модемов несколько десятков устройств и драйвера для них по размеру не более пары мегабайт
'- второй случай, когда драйвера для какого-то одного бренда большие по размеру: например драйверы для видеокарт фирмы nVidia для ПК и ноутбуков занимают сотни мегабайт
'DATE - дата создания в формате ГодМесяцНеделя: “12113” - 2012 год, 11 месяц ноябрь, 3 неделя
'7z - тип архива драйвер-пака - все драйвер-паки будут упаковываться архиватором 7-Zip версии 9.x
'
'2. Расположение
'Новые драйвер-паки будут располагаться все в одном общем каталоге, а не как старые версии были распределены по трем различным подкаталогам, что затрудняло навигацию и обновление файлов.
'
'3. Содержание
'а) Драйвер-паки с именами по типу драйвера: DP_NAME_DATE\имя_бренда\маркер_системы\драйвер:
'DP_Bluetooth_12113\Broadcom\5x86\папка_драйвера\ или DP_Modem_12112\Acorp\NTx64\папка_драйвера\
'
'б) Драйвер-паки с именами по типу драйвера и бренду: DP_NAME_DATE\маркер_системы\драйвер:
'DP_Sound_Realtek_12114\8x64\папка_драйвера\ или DP_Video_nVIDIA_12112\NTx64\папка_драйвера\
'
'* Маркеры системы:
'5x64 - Windows XP x64
'5x86 - Windows XP x86
'6x64 - Windows Vista x64
'6x86 - Windows Vista x86
'7x64 - Windows 7 x64
'7x86 - Windows 7 x86
'8x64 - Windows 8 x64
'8x86 - Windows 8 x86
'NTx64 - Windows Vista/7/8 x64
'NTx86 - Windows Vista/7/8 x86
'Allx64 - Все Windows x64
'Allx86 - Все Windows x86
'AllXP - Windows XP x86/x64
'All6 - Windows Vista x86/x64
'All7 - Windows 7 x86/x64
'All8 - Windows 8 x86/x64
'WinAll - Все Windows
'var ver_51x64="5x64";
'var ver_51x86="5x86";
'var ver_60x64="6x64|NTx64|AllNT";
'var ver_60x86="6x86|NTx86|AllNT";
'var ver_61x64="7x64|NTx64|AllNT";
'var ver_61x86="7x86|NTx86|AllNT";
'var ver_62x64="8x64|NTx64|AllNT|All8x64|All8x64";
'var ver_62x86="8x86|NTx86|AllNT|All8x86";
'var ver_63x64="81x64|NTx64|AllNT|All8x64";
'var ver_63x86="81x86|NTx86|AllNT|All8x86";
'
'STRICT - Если маркер следует после другого маркера, то следует что драйвер предназначен только для той ОС
'Все будущие драйвер-паки будут иметь именно такую структуру
'============================================================================
Option Explicit

' поддерживаемые программой маркеры операционных систем
Public Const strVer_51x64   As String = "5x64"
Public Const strVer_51x86   As String = "5x86"
Public Const strVer_60x64   As String = "6x64|NTx64|AllNT"
Public Const strVer_60x86   As String = "6x86|NTx86|AllNT"
Public Const strVer_61x64   As String = "7x64|781x64|NTx64|AllNT"
Public Const strVer_61x86   As String = "7x86|781x86|NTx86|AllNT"
Public Const strVer_62x64   As String = "8x64|All8x64|NTx64|AllNT"
Public Const strVer_62x86   As String = "8x86|All8x86|NTx86|AllNT"
Public Const strVer_63x64   As String = "81x64|781x64|All8x64|NTx64|AllNT"
Public Const strVer_63x86   As String = "81x86|781x86|All8x86|NTx86|AllNT"
Public Const strVer_XXx64   As String = "Allx64"
Public Const strVer_XXx86   As String = "Allx86"
Public Const strVer_51xXX   As String = "AllXP"
Public Const strVer_60xXX   As String = "All6"
Public Const strVer_61xXX   As String = "All7"
Public Const strVer_62xXX   As String = "All8"
Public Const strVer_63xXX   As String = "All81"
Public Const strVer_XXxXX   As String = "WinAll"
Public Const strVerSTRICT   As String = "STRICT"
Public Const strVerFORCED   As String = "FORCED"

' Служебные переменные для всех версий маркеров
Public strVer_Known_Ver     As String
Public strVer_All_Known_Ver As String
Public strVer_Any86         As String
Public strVer_Any64         As String

' Массив вендлров ноутбуков
Public arrNotebookFilterList()           As String ' Массив производителей ноутбуков, загружается шз файла настроек (для корректного определения модели тачпада)
Public arrNotebookFilterListDef()        As String ' Массив производителей ноутбуков, по умолчанию, если не прописано в файле настроек

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub GetSummaryDPMarkers
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub GetSummaryDPMarkers()
    strVer_Known_Ver = strVer_51x64 + "|" + strVer_51x86 + "|" + strVer_60x64 + "|" + strVer_60x86 + "|" + strVer_61x64 + "|" + strVer_61x86 + "|" + strVer_62x64 + "|" + strVer_62x86 + "|" + strVer_63x64 + "|" + strVer_63x86
    strVer_Any86 = strVer_51x86 + "|" + strVer_60x86 + "|" + strVer_61x86 + "|" + strVer_62x86 + "|" + strVer_63x86 + "|" + strVer_XXx86
    strVer_Any64 = strVer_51x64 + "|" + strVer_60x64 + "|" + strVer_61x64 + "|" + strVer_62x64 + "|" + strVer_63x64 + "|" + strVer_XXx64
    strVer_All_Known_Ver = strVer_Any86 + "|" + strVer_Any64 + "|" + strVer_51xXX + "|" + strVer_60xXX + "|" + strVer_61xXX + "|" + strVer_62xXX + "|" + strVer_63xXX + "|" + strVer_XXxXX
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadNotebookList
'! Description (Описание)  :   [Загрузка массива производителей ноутбуков, для корректного оперделения модели тачпада]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub LoadNotebookList()

    ReDim arrNotebookFilterListDef(35)

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

