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
'81x64 - Windows 8.1 x64
'81x86 - Windows 8.1 x86
'NTx64 - Windows Vista/7/8/8.1 x64
'NTx86 - Windows Vista/7/8/8.1 x86
'Allx64 - ��� Windows x64
'Allx86 - ��� Windows x86
'AllXP - Windows XP x86/x64
'All6 - Windows Vista x86/x64
'All7 - Windows 7 x86/x64
'All8 - Windows 8 x86/x64
'All81 - Windows 8.1 x86/x64
'All9 - Windows 9 x86/x64
'All10 - Windows 10 x86/x64
'WinAll - ��� Windows
'// Markers
'var ver_51x64="5x64";
'var ver_51x86="5x86";
'var ver_60x64="6x64|NTx64|AllNT|67x64|6Xx64";
'var ver_60x86="6x86|NTx86|AllNT|67x86|6Xx86";
'var ver_61x64="7x64|NTx64|AllNT|67x64|78x64|781x64|78110x64|6Xx64";
'var ver_61x86="7x86|NTx86|AllNT|67x86|78x86|781x86|78110x86|6Xx86";
'var ver_62x64="8x64|NTx64|AllNT|78x64|All8x64|6Xx64|AllMx64";
'var ver_62x86="8x86|NTx86|AllNT|78x86|All8x86|6Xx86|AllMx86";
'var ver_63x64="81x64|NTx64|AllNT|781x64|All8x64|78110x64|8110x64|6Xx64|AllMx64";
'var ver_63x86="81x86|NTx86|AllNT|781x86|All8x86|78110x86|8110x86|6Xx86|AllMx86";
'var ver_64x64="9x64|NTx64|AllNT|All8x64|81x64|6Xx64|AllMx64";
'var ver_64x86="9x86|NTx86|AllNT|All8x86|81x86|6Xx86|AllMx86";
'var ver_100x64="10x64|NTx64|AllNT|78110x64|8110x64|All8x64|AllMx64";
'var ver_100x86="10x86|NTx86|AllNT|78110x86|8110x86|All8x86|AllMx86";
'
'STRICT - ���� ������ ������� ����� ������� �������, �� ������� ��� ������� ������������ ������ ��� ��� ��
'��� ������� �������-���� ����� ����� ������ ����� ���������
'============================================================================
Option Explicit

' �������������� ���������� ������� ������������ ������
Public Const strVer_51x64   As String = "5x64"
Public Const strVer_51x86   As String = "5x86"
Public Const strVer_60x64   As String = "6x64|NTx64|AllNT|67x64|6Xx64"
Public Const strVer_60x86   As String = "6x86|NTx86|AllNT|67x86|6Xx86"
Public Const strVer_61x64   As String = "7x64|NTx64|AllNT|67x64|78x64|781x64|78110x64|6Xx64"
Public Const strVer_61x86   As String = "7x86|NTx86|AllNT|67x86|78x86|781x86|78110x86|6Xx86"
Public Const strVer_62x64   As String = "8x64|NTx64|AllNT|78x64|All8x64|6Xx64|AllMx64"
Public Const strVer_62x86   As String = "8x86|NTx86|AllNT|78x86|All8x86|6Xx86|AllMx86"
Public Const strVer_63x64   As String = "81x64|NTx64|AllNT|781x64|All8x64|78110x64|8110x64|6Xx64|AllMx64"
Public Const strVer_63x86   As String = "81x86|NTx86|AllNT|781x86|All8x86|78110x86|8110x86|6Xx86|AllMx86"
Public Const strVer_64x64   As String = "9x64|NTx64|AllNT|All8x64|81x64|6Xx64|AllMx64"
Public Const strVer_64x86   As String = "9x86|NTx86|AllNT|All8x86|81x86|6Xx86|AllMx86"
Public Const strVer_100x64  As String = "10x64|NTx64|AllNT|78110x64|8110x64|All8x64|AllMx64"
Public Const strVer_100x86  As String = "10x86|NTx86|AllNT|78110x86|8110x86|All8x86|AllMx86"
Public Const strVer_XXx64   As String = "Allx64"
Public Const strVer_XXx86   As String = "Allx86"
Public Const strVer_51xXX   As String = "AllXP"
Public Const strVer_60xXX   As String = "All6"
Public Const strVer_61xXX   As String = "All7"
Public Const strVer_62xXX   As String = "All8"
Public Const strVer_63xXX   As String = "All81"
Public Const strVer_64xXX   As String = "All9"
Public Const strVer_100xXX  As String = "All10"
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
    strVer_Known_Ver = strVer_51x64 + "|" + strVer_51x86 + "|" + strVer_60x64 + "|" + strVer_60x86 + "|" + strVer_61x64 + "|" + strVer_61x86 + "|" + strVer_62x64 + "|" + strVer_62x86 + "|" + strVer_63x64 + "|" + strVer_63x86 + "|" + strVer_64x64 + "|" + strVer_64x86 + "|" + strVer_100x64 + "|" + strVer_100x86
    strVer_Any86 = strVer_51x86 + "|" + strVer_60x86 + "|" + strVer_61x86 + "|" + strVer_62x86 + "|" + strVer_63x86 + "|" + strVer_64x86 + "|" + strVer_100x86 + "|" + strVer_XXx86
    strVer_Any64 = strVer_51x64 + "|" + strVer_60x64 + "|" + strVer_61x64 + "|" + strVer_62x64 + "|" + strVer_63x64 + "|" + strVer_64x64 + "|" + strVer_100x64 + "|" + strVer_XXx64
    strVer_All_Known_Ver = strVer_Any86 + "|" + strVer_Any64 + "|" + strVer_51xXX + "|" + strVer_60xXX + "|" + strVer_61xXX + "|" + strVer_62xXX + "|" + strVer_63xXX + "|" + strVer_64xXX + "|" + strVer_100xXX + "|" + strVer_XXxXX
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadNotebookList
'! Description (��������)  :   [�������� ������� �������������� ���������, ��� ����������� ����������� ������ �������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub LoadNotebookList()

    ReDim arrNotebookFilterListDef(14)

    arrNotebookFilterListDef(0) = "Acer;*acer*;*emachines*;*packard*bell*;*gateway*;*aspire*"
    arrNotebookFilterListDef(1) = "Apple;*apple*"
    arrNotebookFilterListDef(2) = "Asus;*asus*"
    arrNotebookFilterListDef(3) = "OEM;*benq*;*clevo*;*depo*;*durabook*;*ecs*;*elitegroup*;*eurocom*;*getac*;*gigabyte*;*intel*;*iru*;*k-systems*;*medion*;*mitac*;*mtc*;*nec*;*pegatron*;*prolink*;*quanta*;*sager*;*shuttle*;*twinhead*;*rover*;*roverbook*;*viewbook*;*viewsonic*;*vizio*;*wistron*"
    arrNotebookFilterListDef(4) = "Dell;*dell*;*alienware*;*arima*;*jetway*;*gericom*"
    arrNotebookFilterListDef(5) = "Fujitsu;*fujitsu*;*sieme*"
    arrNotebookFilterListDef(6) = "HP;*hp*;*hewle*;*compaq*"
    arrNotebookFilterListDef(7) = "Lenovo;*lenovo*;*compal*;*ibm*;"
    arrNotebookFilterListDef(8) = "LG;*lg*"
    arrNotebookFilterListDef(9) = "MSI;*msi*;*micro-star*"
    arrNotebookFilterListDef(10) = "Panasonic;*panasonic*;*matsushita*"
    arrNotebookFilterListDef(11) = "Samsung;*samsung*"
    arrNotebookFilterListDef(12) = "Sony;*sony*;*vaio*"
    arrNotebookFilterListDef(13) = "Toshiba;*toshiba*"

End Sub

'Design by SamLab
'function ManufactorerAliases(str){
'        str = str.toLowerCase();
'        if ((str.indexOf('acer')==0) || (str.indexOf('emachine')==0) || (str.indexOf('gateway')!=-1) || (str.indexOf('bell')!=-1)  || (str.indexOf('aspire')!=-1)) { return 'Acer'; }
'        if (str.indexOf('apple')!=-1) { return 'Apple'; }
'        if (str.indexOf('asus')!=-1) { return 'Asus'; }
'        if ((str.indexOf('dell')==0) || (str.indexOf('alienware')!=-1) || (str.indexOf('arima')!=-1) || (str.indexOf('gericom')!=-1) || (str.indexOf('jetway')!=-1)) { return 'Dell'; }
'        if ((str.indexOf('fujitsu')!=-1) || (str.indexOf('sieme')!=-1)) { return 'Fujitsu'; }
'        if ((str.indexOf('hp')==0) || (str.indexOf('hewle')!=-1) || (str.indexOf('compaq')!=-1) || (str.indexOf('to be filled by hpd')!=-1)) { return 'HP'; }
'        if ((str.indexOf('lenovo')!=-1) || (str.indexOf('ibm')==0) || (str.indexOf('compal')!=-1)) { return 'Lenovo'; }
'        if (str.indexOf('lg')==0) { return 'LG'; }
'        if ((str.indexOf('micro-star')!=-1) || (str.indexOf('msi')==0)) { return 'MSI'; }
'        if ((str.indexOf('panasonic')!=-1) || (str.indexOf('matsushita')!=-1)) { return 'Panasonic'; }
'        if (str.indexOf('samsung')!=-1) { return 'Samsung'; }
'        if ((str.indexOf('sony')==0) || (str.indexOf('vaio')!=-1)) { return 'Sony'; }
'        if (str.indexOf('toshiba')!=-1) { return 'Toshiba'; }
'        if ((str.indexOf('benq')==0) || (str.indexOf('clevo')==0) || (str.indexOf('depo')==0) || (str.indexOf('durabook')!=-1) || (str.indexOf('ecs')==0) || (str.indexOf('elitegroup')!=-1) || (str.indexOf('eurocom')==0) || (str.indexOf('getac')==0) || (str.indexOf('gigabyte')!=-1) || (str.indexOf('intel')==0) || (str.indexOf('iru')==0) || (str.indexOf('k-systems')==0) || (str.indexOf('medion')!=-1) || (str.indexOf('mitac')==0) || (str.indexOf('mtc')==0) || (str.indexOf('nec')==0) || (str.indexOf('pegatron')!=-1) || (str.indexOf('prolink')!=-1) || (str.indexOf('quanta')!=-1) || (str.indexOf('sager')==0) || (str.indexOf('shuttle')!=-1) || (str.indexOf('twinhead')!=-1) || (str.indexOf('rover')!=-1) || (str.indexOf('roverbook')==0) || (str.indexOf('viewbook')==0) || (str.indexOf('viewsonic')==0) || (str.indexOf('vizio')==0) || (str.indexOf('wistron')!=-1)) { return 'OEM'; }

