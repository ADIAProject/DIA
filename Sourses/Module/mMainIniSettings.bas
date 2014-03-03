Attribute VB_Name = "mMainIniSettings"
Option Explicit

Public mbPatnAbs                         As Boolean     ' ���� �� �������� �������� �����������, ������������ � ���������� frmOptions
Public mbAllFolderDRVNotExist            As Boolean     ' ��� �������� � �������� ���������, ��������� � ���������� �� ����������

' ��������� ��������� ����������� �� ini-�����
Public strSysIni                         As String      ' ������� ���� ��������
Public mbLoadIniTmpAfterRestart          As Boolean     ' ��������� ini �� ��������� �����
Public lngOSCount                        As Long        ' ���������� �� �������������� ����������
Public lngOSCountPerRow                  As Long        ' ���������� ��, ������������ �� ����� ������
Public lngUtilsCount                     As Long        ' ���������� ������, ����������� � ����������
Public strDevconCmdPath                  As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DevCon\devcon_c.cmd
Public strArh7zExePATH                   As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\Arc\7z.exe
Public strDevConExePath                  As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DevCon\devcon.exe
Public strDevConExePath64                As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DevCon\devcon64.exe
Public strDevConExePathW2k               As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DevCon\devconw2k.exe
Public strDPInstExePath64                As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DPInst\DPInst64.exe
Public strDPInstExePath86                As String      ' ���� �� ����������� ������ � ������ ������� ������ - .\Tools\DPInst\DPInst.exe
Public strDPInstExePath                  As String      ' ���� �� ����������� ������ � ������ ������� ������ - ����������, � ����������� �� �����������, �� ���������� ����
Public mbDelTmpAfterClose                As Boolean
Public mbUpdateCheck                     As Boolean
Public mbUpdateCheckBeta                 As Boolean
Public mbUpdateToolTip                   As Boolean
Public miStartMode                       As Long
Public mbRecursion                       As Boolean
Public mbSaveSizeOnExit                  As Boolean
Public strExcludeHWID                    As String
Public lngStartModeTab2                  As Long        ' ��������� ������� ��� ����� �������
Public strThisBuildBy                    As String      ' ��������� � �������� � ������� ���� � �������� ���������
Public mbTabBlock                        As Boolean
Public mbTabHide                         As Boolean
Public mbButtonTextUpCase                As Boolean
Public mbLoadFinishFile                  As Boolean
Public mbReadClasses                     As Boolean
Public mbReadDPName                      As Boolean
Public mbConvertDPName                   As Boolean
Public strExcludeFileName                As String
Public strImageStatusButtonName          As String
Public strImageMainName                  As String
Public mbEULAAgree                       As Boolean
Public mbCompareDrvVerByDate             As Boolean     ' ��������� ������ ��������� �� ����
Public mbLoadUnSupportedOS               As Boolean     ' �������\��������� �������� ��� ������������� ��
Public mbAutoInfoAfterDelDRV             As Boolean     ' �������������� ������������ ��� �������� ��������
Public mbDateFormatRus                   As Boolean     ' �������������� ������������ ��� �������� ��������
Public mbCreateRestorePoint              As Boolean     ' ���������� ��� ������ �������� ����� ��������������
Public mbDisableDEP                      As Boolean     ' ���������� ��� ����������� ���������� DEP
Public mbHideOtherProcess                As Boolean     ' �������� ��������� �������� ��� �������
Public mbDP_Is_aFolder                   As Boolean     ' ������ ��������� � ���� ����� - �.� ������������� ����� ���������
Public mbStartMaximazed                  As Boolean     ' ��������� ��������� ����������� �� ���� �����
Public mbTempPath                        As Boolean     ' ������������ �������������� ������� Temp - �.� �������� �������
Public strAlternativeTempPath            As String      ' ���� ��� ��������������� �������� Temp
Public mbDpInstLegacyMode                As Boolean     ' ��������� DPinst
Public mbDpInstPromptIfDriverIsNotBetter As Boolean     ' ��������� DPinst
Public mbDpInstForceIfDriverIsNotBetter  As Boolean     ' ��������� DPinst
Public mbDpInstSuppressAddRemovePrograms As Boolean     ' ��������� DPinst
Public mbDpInstSuppressWizard            As Boolean     ' ��������� DPinst
Public mbDpInstQuietInstall              As Boolean     ' ��������� DPinst
Public mbDpInstScanHardware              As Boolean     ' ��������� DPinst
Public mbSearchOnStart                   As Boolean     ' ������ ����� ���������� ��� ������� ���������
Public lngPauseAfterSearch               As Long        ' ������ ����� ������ ����� ���������� ��� ������� ���������
Public mbCalcDriverScore                 As Boolean     ' ������������ ��� ������� ��������� ���� ���������� ��������, �� ��������� ��������� �������
Public mbCompatiblesHWID                 As Boolean     ' ������������ ��� ������ ���������� ��������� ������ CompatiblesHWID, ������� �� �������
Public mbSearchCompatibleDriverOtherOS   As Boolean     ' ������ ���������� �������� �� ���� ��������, � �� ������ �� �����������
Public lngCompatiblesHWIDCount           As Long        ' ������� ������ ����������� HWID
Public mbMatchHWIDbyDPName               As Boolean     ' ������ ����� ����� ��� ���������� ������������� ��������
Public lngMainFormWidth                  As Long        ' ������ �������� �����
Public lngMainFormHeight                 As Long        ' ������ �������� �����
Public lngButtonWidth                    As Long        ' ������ ������
Public lngButtonHeight                   As Long        ' ������ ������
Public lngButtonLeft                     As Long        ' ������ ����� ��� ������
Public lngButtonTop                      As Long        ' ������ ������ ��� ������
Public lngBtn2BtnLeft                    As Long        ' �������� ����� �������� �� �����������
Public lngBtn2BtnTop                     As Long        ' �������� ����� �������� �� ���������
'Public strImageMenuName                  As String
'Public mbExMenu                           As Boolean ' ����������� ����

'-------------------- ��������� �������� ���� � ������  ------------------'
Public Const lngMainFormWidthMin         As Long = 9350     ' ����������� �������� �������� �����
Public Const lngMainFormHeightMin        As Long = 6500     ' ����������� �������� �������� �����
Public Const lngButtonWidthMin           As Long = 1500     ' ����������� �������� �������� ������ - ������
Public Const lngButtonHeightMin          As Long = 350      ' ����������� �������� �������� ������ - ������
Private Const lngMainFormWidthDef        As Long = 11800    ' ��������� �������� �������� �����
Private Const lngMainFormHeightDef       As Long = 8400     ' ��������� �������� �������� �����
Private Const lngButtonWidthDef          As Long = 2150     ' ��������� �������� �������� ������ - ������
Private Const lngButtonHeightDef         As Long = 550      ' ��������� �������� �������� ������ - ������
Private Const lngButtonLeftDef           As Long = 100      ' ��������� �������� �������� ������ - ������ ����� ��� ������
Private Const lngButtonTopDef            As Long = 480      ' ��������� �������� �������� ������ - ������ ������ ��� ������
Private Const lngBtn2BtnLeftDef          As Long = 100      ' ��������� �������� �������� ������ - �������� ����� �������� �� �����������
Private Const lngBtn2BtnTopDef           As Long = 100      ' ��������� �������� �������� ������ - �������� ����� �������� �� ���������



'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub CreateIni
'! Description (��������)  :   [���������� �������� � ��� ���� ���� ����� ���]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub CreateIni()

    Dim cnt As Long

    If PathExists(strSysIni) = False Then
        If mbIsDriveCDRoom Then
            strSysIni = strWorkTempBackSL & strSettingIniFile
            MsgBox "File " & strSettingIniFile & " is not Exist!" & vbNewLine & "This program works from CD\DVD, so we create temporary " & strSettingIniFile & "-file" & vbNewLine & strSysIni, vbInformation + vbApplicationModal, strProductName
        End If

        '������ Main
        IniWriteStrPrivate "Main", "DelTmpAfterClose", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheck", "1", strSysIni
        IniWriteStrPrivate "Main", "UpdateCheckBeta", "0", strSysIni
        IniWriteStrPrivate "Main", "StartMode", "1", strSysIni
        IniWriteStrPrivate "Main", "EULAAgree", "0", strSysIni
        IniWriteStrPrivate "Main", "HideOtherProcess", "1", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTemp", "0", strSysIni
        IniWriteStrPrivate "Main", "AlternativeTempPath", "%Temp%", strSysIni
        IniWriteStrPrivate "Main", "AutoLanguage", "1", strSysIni
        IniWriteStrPrivate "Main", "StartLanguageID", "0409", strSysIni
        IniWriteStrPrivate "Main", "IconMainSkin", "Standart", strSysIni
        IniWriteStrPrivate "Main", "SilentDLL", "0", strSysIni
        IniWriteStrPrivate "Main", "LoadIniTmpAfterRestart", "0", strSysIni
        IniWriteStrPrivate "Main", "AutoInfoAfterDelDRV", "1", strSysIni
        IniWriteStrPrivate "Main", "SearchOnStart", "0", strSysIni
        IniWriteStrPrivate "Main", "PauseAfterSearch", "1", strSysIni
        IniWriteStrPrivate "Main", "CreateRestorePoint", "1", strSysIni

        '������ Debug
        IniWriteStrPrivate "Debug", "DebugEnable", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogPath", "%SYSTEMDRIVE%", strSysIni
        IniWriteStrPrivate "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt", strSysIni
        IniWriteStrPrivate "Debug", "CleenHistory", "1", strSysIni
        IniWriteStrPrivate "Debug", "DetailMode", "1", strSysIni
        IniWriteStrPrivate "Debug", "DebugLog2AppPath", "0", strSysIni
        IniWriteStrPrivate "Debug", "Time2File", "0", strSysIni
        '������ DPInst
        IniWriteStrPrivate "DPInst", "PathExe", "Tools\DPInst\DPInst.exe", strSysIni
        IniWriteStrPrivate "DPInst", "PathExe64", "Tools\DPInst\DPInst64.exe", strSysIni
        IniWriteStrPrivate "DPInst", "LegacyMode", 1, strSysIni
        IniWriteStrPrivate "DPInst", "PromptIfDriverIsNotBetter", 1, strSysIni
        IniWriteStrPrivate "DPInst", "ForceIfDriverIsNotBetter", 0, strSysIni
        IniWriteStrPrivate "DPInst", "SuppressAddRemovePrograms", 0, strSysIni
        IniWriteStrPrivate "DPInst", "SuppressWizard", 0, strSysIni
        IniWriteStrPrivate "DPInst", "QuietInstall", 0, strSysIni
        IniWriteStrPrivate "DPInst", "ScanHardware", 1, strSysIni
        '������ Arc
        IniWriteStrPrivate "Arc", "PathExe", "Tools\Arc\7za.exe", strSysIni
        '������ Devcon
        IniWriteStrPrivate "Devcon", "PathExe", "Tools\Devcon\devcon.exe", strSysIni
        IniWriteStrPrivate "Devcon", "PathExe64", "Tools\Devcon\devcon64.exe", strSysIni
        IniWriteStrPrivate "Devcon", "PathExeW2K", "Tools\Devcon\devconw2k.exe", strSysIni
        IniWriteStrPrivate "Devcon", "CollectHwidsCmd", "Tools\Devcon\devcon_c.cmd", strSysIni
        '������ OS
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
        IniWriteStrPrivate "OS", "CompatiblesHWID", "1", strSysIni
        IniWriteStrPrivate "OS", "CompatiblesHWIDCount", "10", strSysIni
        IniWriteStrPrivate "OS", "LoadUnSupportedOS", "0", strSysIni
        IniWriteStrPrivate "OS", "CalcDriverScore", "1", strSysIni
        IniWriteStrPrivate "OS", "SearchCompatibleDriverOtherOS", "1", strSysIni
        IniWriteStrPrivate "OS", "MatchHWIDbyDPName", "1", strSysIni
        IniWriteStrPrivate "OS", "DP_is_aFolder", "0", strSysIni
        '������ OS_1
        IniWriteStrPrivate "OS_1", "Ver", "5.0;5.1;5.2", strSysIni
        IniWriteStrPrivate "OS_1", "Name", "2000/XP/2003 Server", strSysIni
        IniWriteStrPrivate "OS_1", "drpFolder", "drivers\xp", strSysIni
        IniWriteStrPrivate "OS_1", "devIDFolder", "drivers\xp\dev_db", strSysIni
        IniWriteStrPrivate "OS_1", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_1", "PathPhysX", "drivers\XP\DP_Graphics_PhysX*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "PathLanguages", "drivers\XP\DP_Graphics_Languages*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "PathRuntimes", "drivers\XP\DP_Runtimes*.7z", strSysIni
        IniWriteStrPrivate "OS_1", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        '������ OS_2
        IniWriteStrPrivate "OS_2", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_2", "Name", "Vista/7/8/8.1/Server 2008", strSysIni
        IniWriteStrPrivate "OS_2", "drpFolder", "drivers\vista", strSysIni
        IniWriteStrPrivate "OS_2", "devIDFolder", "drivers\vista\dev_db", strSysIni
        IniWriteStrPrivate "OS_2", "is64bit", "0", strSysIni
        IniWriteStrPrivate "OS_2", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_2", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        '������ OS_3
        IniWriteStrPrivate "OS_3", "Ver", "6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_3", "Name", "Vista/7/8/8.1/Server 2008 x64", strSysIni
        IniWriteStrPrivate "OS_3", "drpFolder", "drivers\vista64", strSysIni
        IniWriteStrPrivate "OS_3", "devIDFolder", "drivers\vista64\dev_db", strSysIni
        IniWriteStrPrivate "OS_3", "is64bit", "1", strSysIni
        IniWriteStrPrivate "OS_3", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_3", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        '������ OS_4
        IniWriteStrPrivate "OS_4", "Ver", "5.0;5.1;5.2;6.0;6.1;6.2;6.3", strSysIni
        IniWriteStrPrivate "OS_4", "Name", "Windows XP / 2000 / Server 2003 / Vista / Server 2008 / 7 / 8 / 8.1", strSysIni
        IniWriteStrPrivate "OS_4", "drpFolder", "drivers\All", strSysIni
        IniWriteStrPrivate "OS_4", "devIDFolder", "drivers\All\dev_db", strSysIni
        IniWriteStrPrivate "OS_4", "is64bit", "2", strSysIni
        IniWriteStrPrivate "OS_4", "PathPhysX", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "PathLanguages", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "PathRuntimes", vbNullString, strSysIni
        IniWriteStrPrivate "OS_4", "ExcludeFileName", "DPsFnshr*.7z", strSysIni
        
        '������ Utils
        IniWriteStrPrivate "Utils", "UtilsCount", "3", strSysIni
        '������ Utils_1
        IniWriteStrPrivate "Utils_1", "Name", "DirectX Diagnostics", strSysIni
        IniWriteStrPrivate "Utils_1", "Path", "%WINDIR%\system32\dxdiag.exe", strSysIni
        IniWriteStrPrivate "Utils_1", "Params", vbNullString, strSysIni
        '������ Utils_2
        IniWriteStrPrivate "Utils_2", "Name", "Disk Managment", strSysIni
        IniWriteStrPrivate "Utils_2", "Path", "diskmgmt.msc", strSysIni
        IniWriteStrPrivate "Utils_2", "Params", vbNullString, strSysIni
        '������ Utils_3
        IniWriteStrPrivate "Utils_3", "Name", "Remove BugFix with Installation of Video Drivers Nvidia", strSysIni
        IniWriteStrPrivate "Utils_3", "Path", "Tools\Nvidia\PatchPostInstall.cmd", strSysIni
        IniWriteStrPrivate "Utils_3", "Params", vbNullString, strSysIni
        '������ MainForm
        IniWriteStrPrivate "MainForm", "Width", CStr(lngMainFormWidthDef), strSysIni
        IniWriteStrPrivate "MainForm", "Height", CStr(lngMainFormHeightDef), strSysIni
        IniWriteStrPrivate "MainForm", "StartMaximazed", "0", strSysIni
        IniWriteStrPrivate "MainForm", "SaveSizeOnExit", "0", strSysIni
        IniWriteStrPrivate "MainForm", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "MainForm", "FontSize", "8", strSysIni
        IniWriteStrPrivate "MainForm", "HighlightColor", "32896", strSysIni
        '������ Buttons
        IniWriteStrPrivate "Button", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "Button", "FontSize", "9", strSysIni
        IniWriteStrPrivate "Button", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Button", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Button", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Button", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Button", "FontColor", "0", strSysIni
        IniWriteStrPrivate "Button", "Width", lngButtonWidthDef, strSysIni
        IniWriteStrPrivate "Button", "Height", lngButtonHeightDef, strSysIni
        IniWriteStrPrivate "Button", "Left", lngButtonLeftDef, strSysIni
        IniWriteStrPrivate "Button", "Top", lngButtonTopDef, strSysIni
        IniWriteStrPrivate "Button", "Btn2BtnLeft", lngBtn2BtnLeftDef, strSysIni
        IniWriteStrPrivate "Button", "Btn2BtnTop", lngBtn2BtnTopDef, strSysIni
        IniWriteStrPrivate "Button", "TextUpCase", "0", strSysIni
        IniWriteStrPrivate "Button", "IconStatusSkin", "Standart", strSysIni
        '������ Tab
        IniWriteStrPrivate "Tab", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "Tab", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Tab", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Tab", "FontColor", "0", strSysIni
        '������ Tab2
        IniWriteStrPrivate "Tab2", "StartMode", "1", strSysIni
        IniWriteStrPrivate "Tab2", "FontName", "Tahoma", strSysIni
        IniWriteStrPrivate "Tab2", "FontSize", "8", strSysIni
        IniWriteStrPrivate "Tab2", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontBold", "0", strSysIni
        IniWriteStrPrivate "Tab2", "FontColor", "&H8000000D", strSysIni
        '������ ToolTip
        'IniWriteStrPrivate "ToolTip", "FontName", "Courier New", strSysIni
        IniWriteStrPrivate "ToolTip", "FontName", "Lucida Console", strSysIni
        IniWriteStrPrivate "ToolTip", "FontSize", "8", strSysIni
        IniWriteStrPrivate "ToolTip", "FontUnderline", "0", strSysIni
        IniWriteStrPrivate "ToolTip", "FontStrikethru", "0", strSysIni
        IniWriteStrPrivate "ToolTip", "FontItalic", "0", strSysIni
        IniWriteStrPrivate "ToolTip", "FontBold", "0", strSysIni
        IniWriteStrPrivate "ToolTip", "FontColor", "0", strSysIni
        '������ NotebookVendor
        IniWriteStrPrivate "NotebookVendor", "FilterCount", "22", strSysIni
        '������ "NotebookVendor"
        IniWriteStrPrivate "NotebookVendor", "FilterCount", UBound(arrNotebookFilterListDef), strSysIni

        For cnt = 0 To UBound(arrNotebookFilterListDef) - 1
            IniWriteStrPrivate "NotebookVendor", "Filter_" & cnt + 1, arrNotebookFilterListDef(cnt), strSysIni
        Next

        ' �������� Ini ���� � ������������ ����
        NormIniFile strSysIni
        ' ��������� ������� ����� �������� ini-�����
        mbDebugEnable = True
        mbCleanHistory = True
        strDebugLogPathTemp = "%SYSTEMDRIVE%"
        strDebugLogNameTemp = "DIA-LOG_%DATE%.txt"
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetMainIniParam
'! Description (��������)  :   [��������� �������� �� ��� �����]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub GetMainIniParam()

    Dim i                           As Long
    Dim mbAllFolderDRVNotExistCount As Integer
    Dim cntOsInIni                  As Integer
    Dim cntUtilsInIni               As Integer
    Dim strDebugLogPathFolder       As String
    Dim NotebookFilterCount         As Long
    Dim numFilter                   As Long

    'SaveSetting App.ProductName, "Settings", "LOAD_INI_TMP", True
    'SaveSetting App.ProductName, "Settings", "LOAD_INI_PATH", strSysIni
    '[Description]
    strThisBuildBy = GetIniValueString(strSysIni, "Description", "BuildBy", vbNullString)
    'strThisBuildBy = "www.SamLab.Ws"
    '[Debug]
    ' ��������� �������
    mbDebugEnable = GetIniValueBoolean(strSysIni, "Debug", "DebugEnable", 0)
    ' ������� �������
    mbCleanHistory = GetIniValueBoolean(strSysIni, "Debug", "CleenHistory", 1)
    ' ���� �� ��� �����
    strDebugLogPathTemp = PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%"))
    strDebugLogPath = PathCollect(PathNameFromPath(GetIniValueString(strSysIni, "Debug", "DebugLogPath", "%SYSTEMDRIVE%")))
    ' ��� ���-�����
    strDebugLogNameTemp = GetIniValueString(strSysIni, "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt")
    strDebugLogName = ExpandFileNamebyEnvironment(GetIniValueString(strSysIni, "Debug", "DebugLogName", "DIA-LOG_%DATE%.txt"))
    ' ����������� ������� - �� ���������=1
    lngDetailMode = GetIniValueLong(strSysIni, "Debug", "DetailMode", 1)
    ' ���������� ����� � ���-����
    mbDebugTime2File = GetIniValueBoolean(strSysIni, "Debug", "Time2File", 0)
    ' ��������� ���-���� � �������� "logs" ���������
    mbDebugLog2AppPath = GetIniValueBoolean(strSysIni, "Debug", "DebugLog2AppPath", 0)

    If Not mbDebugLog2AppPath Then
        strDebugLogFullPath = strDebugLogPath & strDebugLogName

        If mbDebugEnable Then
            strDebugLogPathFolder = strDebugLogPath

            If PathExists(strDebugLogPathFolder) = False Then
                CreateNewDirectory strDebugLogPathFolder
            End If
        End If

    Else
        strDebugLogFullPath = strAppPathBackSL & "logs\" & strDebugLogName

        If Not LogNotOnCDRoom Then
            If mbDebugEnable Then
                If PathExists(strAppPathBackSL & "logs\") = False Then
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
    ' �������� ��� ������
    mbDelTmpAfterClose = GetIniValueBoolean(strSysIni, "Main", "DelTmpAfterClose", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mbUpdateCheck = GetIniValueBoolean(strSysIni, "Main", "UpdateCheck", 1)
    ' �������� ���������� ��� ������ (������ MAIN)
    mbUpdateCheckBeta = GetIniValueBoolean(strSysIni, "Main", "UpdateCheckBeta", 1)
    ' �������� EULA
    mbEULAAgree = GetIniValueBoolean(strSysIni, "Main", "EULAAgree", 0)
    ' ��������������� �����
    mbAutoLanguage = GetIniValueBoolean(strSysIni, "Main", "AutoLanguage", 1)

    If Not mbAutoLanguage Then
        strStartLanguageID = IniStringPrivate("Main", "StartLanguageID", strSysIni)
    End If

    ' ��������� ��������������� ���� Temp
    strAlternativeTempPath = IniStringPrivate("Main", "AlternativeTempPath", strSysIni)

    If strAlternativeTempPath = "no_key" Then
        strAlternativeTempPath = strWinTemp
    End If

    ' ��� ������������� ���������� �������������� temp
    mbTempPath = GetIniValueBoolean(strSysIni, "Main", "AlternativeTemp", 0)

    If mbTempPath Then
        strAlternativeTempPath = PathCollect(strAlternativeTempPath)
        DebugMode "AlternativeTempPath: " & strAlternativeTempPath

        If PathExists(strAlternativeTempPath) Then
            strWinTemp = strAlternativeTempPath
            strWorkTemp = strWinTemp & strProjectName

            ' ���� ���, �� ������� ��������� ������� �������
            If PathExists(strWorkTemp) = False Then
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
    ' ��������� ��������� ��������� (������ ��)
    mbRecursion = GetIniValueBoolean(strSysIni, "OS", "Recursion", 1)
    ' ������ ����������� ������� �� ���� ��
    mbTabBlock = GetIniValueBoolean(strSysIni, "OS", "TabBlock", 1)
    ' �������� ������� �� ���� ��
    mbTabHide = GetIniValueBoolean(strSysIni, "OS", "TabHide", 0)
    ' ����������� ����� ��������
    mbCalcDriverScore = GetIniValueBoolean(strSysIni, "OS", "CalcDriverScore", 1)
    ' ��������� ���-�� ������ (������ OS) � ���������� ������� ��
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

                If PathExists(arrOSList(i).drpFolderFull) = False Then
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

    ' ��������� ���-�� ������� �� ����� ������ (������ Main)
    lngOSCountPerRow = IniLongPrivate("OS", "OSCountPerRow", strSysIni)

    If lngOSCountPerRow = 0 Or lngOSCountPerRow = 9999 Then
        lngOSCountPerRow = 4
    End If

    '[Utils]
    ' ��������� ���-�� ������
    lngUtilsCount = IniLongPrivate("Utils", "UtilsCount", strSysIni)

    If lngUtilsCount = 0 Or lngUtilsCount = 9999 Then

        'MsgBox "������ �������������� ����������� ������ ����. ������ ��������� ����������", vbExclamation, "������ ��������� ����������"
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

            If arrUtilsList(i, 3) = "no_key" Or arrUtilsList(i, 3) = "�������������� ��������� �������" Then
                arrUtilsList(i, 3) = vbNullString
            End If

        Next

    End If

    '--------------------- ��������� ����� �� ������ ---------------------
    '[DevCon]
    ' DEVCON_CMD
    strDevconCmdPath = IniStringPrivate("DevCon", "CollectHwidsCmd", strSysIni)
    strDevconCmdPath = PathCollect(strDevconCmdPath)

    If PathExists(strDevconCmdPath) = False Then
        strDevconCmdPath = strAppPathBackSL & "Tools\Devcon\devcon_c.cmd"

        If PathExists(strDevconCmdPath) = False Then
            MsgBox strMessages(7) & vbNewLine & strDevconCmdPath, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE
    strDevConExePath = IniStringPrivate("DevCon", "PathExe", strSysIni)

    If InStr(strDevConExePath, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePath = PathCollect(strDevConExePath)

    If PathExists(strDevConExePath) = False Then
        strDevConExePath = strAppPathBackSL & "Tools\Devcon\devcon.exe"

        If PathExists(strDevConExePath) = False Then
            MsgBox strMessages(7) & vbNewLine & strDevConExePath, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE64
    strDevConExePath64 = IniStringPrivate("DevCon", "PathExe64", strSysIni)

    If InStr(strDevConExePath64, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePath64 = PathCollect(strDevConExePath64)

    If PathExists(strDevConExePath64) = False Then
        strDevConExePath64 = strAppPathBackSL & "Tools\Devcon\devcon64.exe"

        If PathExists(strDevConExePath64) = False Then
            MsgBox strMessages(7) & vbNewLine & strDevConExePath64, vbInformation, strProductName
        End If
    End If

    ' DEVCON_EXE_W2k
    strDevConExePathW2k = IniStringPrivate("DevCon", "PathExeW2k", strSysIni)

    If InStr(strDevConExePathW2k, ":") Then
        mbPatnAbs = True
    End If

    strDevConExePathW2k = PathCollect(strDevConExePathW2k)

    If PathExists(strDevConExePathW2k) = False Then
        strDevConExePathW2k = strAppPathBackSL & "Tools\Devcon\devconw2k.exe"

        If PathExists(strDevConExePathW2k) = False Then
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

    If PathExists(strDPInstExePath86) = False Then
        strDPInstExePath86 = strAppPathBackSL & "Tools\DPInst\DPInst.exe"

        If PathExists(strDPInstExePath86) = False Then
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

    If PathExists(strDPInstExePath64) = False Then
        strDPInstExePath64 = strAppPathBackSL & "Tools\DPInst\DPInst64.exe"

        If PathExists(strDPInstExePath64) = False Then
            MsgBox strMessages(7) & vbNewLine & strDPInstExePath64, vbInformation, strProductName
        End If
    End If

    ' ��������� DpInst
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

    If PathExists(strArh7zExePATH) = False Then
        strArh7zExePATH = strAppPathBackSL & "Tools\Arc\7za.exe"

        If PathExists(strArh7zExePATH) = False Then
            MsgBox strMessages(7) & vbNewLine & strArh7zExePATH, vbInformation, strProductName
        End If
    End If

    '[MainForm]
    ' ��������� ��������� ��� ������
    mbSaveSizeOnExit = GetIniValueBoolean(strSysIni, "MainForm", "SaveSizeOnExit", 0)
    '������ �������� �����
    lngMainFormWidth = GetIniValueLong(strSysIni, "MainForm", "Width", lngMainFormWidthDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If lngMainFormWidth < lngMainFormWidthMin Then
        lngMainFormWidth = lngMainFormWidthDef
    End If

    '������ �������� �����
    lngMainFormHeight = GetIniValueLong(strSysIni, "MainForm", "Height", lngMainFormHeightDef)

    '���� ���������� �������� ������ ������������, �� ������������� �������� �� ���������
    If lngMainFormHeight < lngMainFormHeightMin Then
        lngMainFormHeight = lngMainFormHeightDef
    End If

    ' ��������� ���� ������� (������ MainForm)
    mbStartMaximazed = GetIniValueBoolean(strSysIni, "MainForm", "StartMaximazed", 0)
    strFontMainForm_Name = GetIniValueString(strSysIni, "MainForm", "FontName", "Tahoma")
    lngFontMainForm_Size = GetIniValueLong(strSysIni, "MainForm", "FontSize", 8)
    ' ��������� ��������� ��������
    glHighlightColor = GetIniValueLong(strSysIni, "MainForm", "HighlightColor", 32896)
    ' ��������� ���� ������� (������ OtherForm)
    strFontOtherForm_Name = GetIniValueString(strSysIni, "OtherForm", "FontName", "Tahoma")
    lngFontOtherForm_Size = GetIniValueLong(strSysIni, "OtherForm", "FontSize", 8)
    '[Buttons]
    lngButtonWidth = GetIniValueLong(strSysIni, "Button", "Width", lngButtonWidthDef)
    lngButtonHeight = GetIniValueLong(strSysIni, "Button", "Height", lngButtonHeightDef)
    lngButtonLeft = GetIniValueLong(strSysIni, "Button", "Left", lngButtonLeftDef)
    lngButtonTop = GetIniValueLong(strSysIni, "Button", "Top", lngButtonTopDef)
    lngBtn2BtnLeft = GetIniValueLong(strSysIni, "Button", "Btn2BtnLeft", lngBtn2BtnLeftDef)
    lngBtn2BtnTop = GetIniValueLong(strSysIni, "Button", "Btn2BtnTop", lngBtn2BtnTopDef)
    ' ����� ������ � ������� �������� (������ Button)
    mbButtonTextUpCase = GetIniValueBoolean(strSysIni, "Button", "TextUpCase", 0)
    '[OS]
    ' ������������ ����� Finish
    mbLoadFinishFile = GetIniValueBoolean(strSysIni, "OS", "LoadFinishFile", 1)
    ' ��������� ����� ������ �� ����� Finish
    mbReadClasses = GetIniValueBoolean(strSysIni, "OS", "ReadClasses", 1)
    ' ��������� ��� ������
    mbReadDPName = GetIniValueBoolean(strSysIni, "OS", "ReadDPName", 1)
    ' ��������������� ����� �������
    mbConvertDPName = GetIniValueBoolean(strSysIni, "OS", "ConvertDPName", 1)
    ' ����������� HWID �� ���������
    strExcludeHWID = GetIniValueString(strSysIni, "OS", "ExcludeHWID", "USB\ROOT_HUB*;ROOT\*;STORAGE\*;USBSTOR\*;PCIIDE\IDECHANNEL;PCI\CC_0604")
    ' ��������� ������ ���������
    mbCompareDrvVerByDate = GetIniValueBoolean(strSysIni, "OS", "CompareDrvVerByDate", 1)
    ' ���������� ���� ������ � ������� dd/mm/yyyy
    mbDateFormatRus = GetIniValueBoolean(strSysIni, "OS", "DateFormatRus", 0)
    ' ������������ ����������� HWID
    mbCompatiblesHWID = GetIniValueBoolean(strSysIni, "OS", "CompatiblesHWID", 1)
    lngCompatiblesHWIDCount = GetIniValueLong(strSysIni, "OS", "CompatiblesHWIDCount", 5)
    '��������� ������������� �� ����� ��� �������
    'mbMatchHWIDbyMarkers = GetIniValueBoolean(strSysIni, "OS", "MatchHWIDbyMarkers", 1)
    mbMatchHWIDbyDPName = GetIniValueBoolean(strSysIni, "OS", "MatchHWIDbyDPName", 1)
    ' ������������ ����������� HWID
    mbLoadUnSupportedOS = GetIniValueBoolean(strSysIni, "OS", "LoadUnSupportedOS", 0)
    mbSearchCompatibleDriverOtherOS = GetIniValueBoolean(strSysIni, "OS", "SearchCompatibleDriverOtherOS", 1)
    '[Button]
    ' ����� ������
    strFontBtn_Name = GetIniValueString(strSysIni, "Button", "FontName", "Tahoma")
    miFontBtn_Size = GetIniValueLong(strSysIni, "Button", "FontSize", 8)
    mbFontBtn_Bold = GetIniValueBoolean(strSysIni, "Button", "FontBold", 0)
    mbFontBtn_Italic = GetIniValueBoolean(strSysIni, "Button", "FontItalic", 0)
    mbFontBtn_Underline = GetIniValueBoolean(strSysIni, "Button", "FontUnderline", 0)
    mbFontBtn_Strikethru = GetIniValueBoolean(strSysIni, "Button", "FontStrikethru", 0)
    lngFontBtn_Color = GetIniValueLong(strSysIni, "Button", "FontColor", 0)
    strImageStatusButtonName = GetIniValueString(strSysIni, "Button", "IconStatusSkin", "Standart")
    '[Tab]
    ' ����� � ��������� ��������
    strFontTab_Name = GetIniValueString(strSysIni, "Tab", "FontName", "Tahoma")
    miFontTab_Size = GetIniValueLong(strSysIni, "Tab", "FontSize", 8)
    mbFontTab_Bold = GetIniValueBoolean(strSysIni, "Tab", "FontBold", 0)
    mbFontTab_Italic = GetIniValueBoolean(strSysIni, "Tab", "FontItalic", 0)
    mbFontTab_Underline = GetIniValueBoolean(strSysIni, "Tab", "FontUnderline", 0)
    mbFontTab_Strikethru = GetIniValueBoolean(strSysIni, "Tab", "FontStrikethru", 0)
    lngFontTab_Color = GetIniValueLong(strSysIni, "Tab", "FontColor", 0)
    '[Tab2]
    ' ����� � ��������� ��������
    strFontTab2_Name = GetIniValueString(strSysIni, "Tab2", "FontName", "Tahoma")
    miFontTab2_Size = GetIniValueLong(strSysIni, "Tab2", "FontSize", 8)
    mbFontTab2_Bold = GetIniValueBoolean(strSysIni, "Tab2", "FontBold", 0)
    mbFontTab2_Italic = GetIniValueBoolean(strSysIni, "Tab2", "FontItalic", 0)
    mbFontTab2_Underline = GetIniValueBoolean(strSysIni, "Tab2", "FontUnderline", 0)
    mbFontTab2_Strikethru = GetIniValueBoolean(strSysIni, "Tab2", "FontStrikethru", 0)
    lngFontTab2_Color = GetIniValueLong(strSysIni, "Tab2", "FontColor", &H8000000D)
    lngStartModeTab2 = GetIniValueLong(strSysIni, "Tab2", "StartMode", 1)
    '[ToolTip]
    ' ����� � ��������� ToolTip
    'strFontTT_Name = GetIniValueString(strSysIni, "ToolTip", "FontName", "Courier New")
    strFontTT_Name = GetIniValueString(strSysIni, "ToolTip", "FontName", "Lucida Console")
    miFontTT_Size = GetIniValueLong(strSysIni, "ToolTip", "FontSize", 8)
    mbFontTT_Bold = GetIniValueBoolean(strSysIni, "ToolTip", "FontBold", 0)
    mbFontTT_Italic = GetIniValueBoolean(strSysIni, "ToolTip", "FontItalic", 0)
    mbFontTT_Underline = GetIniValueBoolean(strSysIni, "ToolTip", "FontUnderline", 0)
    mbFontTT_Strikethru = GetIniValueBoolean(strSysIni, "ToolTip", "FontStrikethru", 0)
    lngFontTT_Color = GetIniValueLong(strSysIni, "ToolTip", "FontColor", 0)
    '[Main]
    strImageMainName = GetIniValueString(strSysIni, "Main", "IconMainSkin", "Standart")
    ' ����������� ����
    'mbExMenu = GetIniValueBoolean(strSysIni, "Main", "ExMenu", 1)
    'strImageMenuName = GetIniValueString(strSysIni, "Main", "IconMenuSkin", "Standart")
    ' �������� ������ ��������
    mbHideOtherProcess = GetIniValueBoolean(strSysIni, "Main", "HideOtherProcess", 1)
    ' ����� ����������� DLL
    mbSilentDLL = GetIniValueBoolean(strSysIni, "Main", "SilentDll", 0)
    ' ���������� ����������� �� ���������� (����������� ����)
    mbUpdateToolTip = GetIniValueBoolean(strSysIni, "Main", "UpdateToolTip", 1)
    ' �������������� ���������� ����� �������� ��������
    mbAutoInfoAfterDelDRV = GetIniValueBoolean(strSysIni, "Main", "AutoInfoAfterDelDRV", 1)
    ' ��������� �����
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

