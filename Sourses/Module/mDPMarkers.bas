Attribute VB_Name = "mDPMarkers"
'============================================================================
'����� ��������� �������-�����
'
'������ ������ �������-����� ������������ ����� ������ DP_NAME_wnt6-x64_DATE.7z, ����, DP_Bluetooth_wnt6-x64_1210.7z. ��������� wnt5_x86-32 �������� Windows XP x86, wnt6-x64 - Windows Vista/7/8 x64, wnt6-x86 - Windows Vista/7/8 x86 � ������������� ��� �������������� � 3� ������ ������. �� � ����������� ����� ������ �� ���� ���������� NT � ���� 6.x, � ������ �������-���� ����� ���������� ������� �� 3 �������� ��� ������ ������� Vista/7/8 �� �������� ��������� ���������, ���� ��������� ����� ���� �������������� � ��� � ������ �������-���� �����������, � �.�. DPS �� ����� ���������� �������� � ������ �������-�����, �� ������ ������� �������� ���������� �� ���������� � ����� ������������ � ��������� � ������ ������ ������������� ��������, � � ������� DPS �� ����� 269 ������������ � ������ �������� �� ������ ������� (�������� ��� Win8 �� Win7).
'
'������� ����, ��� ��������� ������ �������-����� ������ ����� ���������������� � ������������ � �������� �������� � ����� ��������� ������� �������-����� ����� ��������� ��������� �������:
'
'1. ��������
'����� ������ ����� �������-����� ����� ���������: DP_NAME_DATE.7z - ���������:
'DP - ���������� �� ����� DriverPack (�������-��� - ����� ��������� � ����� ����� ������)
'NAME - ��� �������-���� �� ���� ��������� � ��� � �� �� ������: �Bluetooth� ��� �Sound_Realtek�
'- ������ ������, ����� ��������� ��� ������-�� ���� ���������� ����� � ��� ��������� �� �������: ������� ��������� �������� ��������� � �������� ��� ��� �� ������� �� ����� ���� ��������
'- ������ ������, ����� �������� ��� ������-�� ������ ������ ������� �� �������: �������� �������� ��� ��������� ����� nVidia ��� �� � ��������� �������� ����� ��������
'DATE - ���� �������� � ������� ��������������: �12113� - 2012 ���, 11 ����� ������, 3 ������
'7z - ��� ������ �������-���� - ��� �������-���� ����� ������������� ����������� 7-Zip ������ 9.x
'
'2. ������������
'����� �������-���� ����� ������������� ��� � ����� ����� ��������, � �� ��� ������ ������ ���� ������������ �� ���� ��������� ������������, ��� ���������� ��������� � ���������� ������.
'
'3. ����������
'�) �������-���� � ������� �� ���� ��������: DP_NAME_DATE\���_������\������_�������\�������:
'DP_Bluetooth_12113\Broadcom\5x86\�����_��������\ ��� DP_Modem_12112\Acorp\NTx64\�����_��������\
'
'�) �������-���� � ������� �� ���� �������� � ������: DP_NAME_DATE\������_�������\�������:
'DP_Sound_Realtek_12114\8x64\�����_��������\ ��� DP_Video_nVIDIA_12112\NTx64\�����_��������\
'
'* ������� �������:
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
'Allx64 - ��� Windows x64
'Allx86 - ��� Windows x86
'AllXP - Windows XP x86/x64
'All6 - Windows Vista x86/x64
'All7 - Windows 7 x86/x64
'All8 - Windows 8 x86/x64
'WinAll - ��� Windows
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
'STRICT - ���� ������ ������� ����� ������� �������, �� ������� ��� ������� ������������ ������ ��� ��� ��
'��� ������� �������-���� ����� ����� ������ ����� ���������
'============================================================================
Option Explicit

' �������������� ���������� ������� ������������ ������
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

' ��������� ���������� ��� ���� ������ ��������
Public strVer_Known_Ver     As String
Public strVer_All_Known_Ver As String
Public strVer_Any86         As String
Public strVer_Any64         As String

' ������ �������� ���������
Public arrNotebookFilterList()           As String ' ������ �������������� ���������, ����������� �� ����� �������� (��� ����������� ����������� ������ �������)
Public arrNotebookFilterListDef()        As String ' ������ �������������� ���������, �� ���������, ���� �� ��������� � ����� ��������

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub GetSummaryDPMarkers
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub GetSummaryDPMarkers()
    strVer_Known_Ver = strVer_51x64 + "|" + strVer_51x86 + "|" + strVer_60x64 + "|" + strVer_60x86 + "|" + strVer_61x64 + "|" + strVer_61x86 + "|" + strVer_62x64 + "|" + strVer_62x86 + "|" + strVer_63x64 + "|" + strVer_63x86
    strVer_Any86 = strVer_51x86 + "|" + strVer_60x86 + "|" + strVer_61x86 + "|" + strVer_62x86 + "|" + strVer_63x86 + "|" + strVer_XXx86
    strVer_Any64 = strVer_51x64 + "|" + strVer_60x64 + "|" + strVer_61x64 + "|" + strVer_62x64 + "|" + strVer_63x64 + "|" + strVer_XXx64
    strVer_All_Known_Ver = strVer_Any86 + "|" + strVer_Any64 + "|" + strVer_51xXX + "|" + strVer_60xXX + "|" + strVer_61xXX + "|" + strVer_62xXX + "|" + strVer_63xXX + "|" + strVer_XXxXX
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadNotebookList
'! Description (��������)  :   [�������� ������� �������������� ���������, ��� ����������� ����������� ������ �������]
'! Parameters  (����������):
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

