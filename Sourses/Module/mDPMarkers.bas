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
'WinAll - ��� Windows
'var ver_51x64="5x64";
'var ver_51x86="5x86";
'var ver_60x64="6x64|NTx64|AllNT";
'var ver_60x86="6x86|NTx86|AllNT";
'var ver_61x64="7x64|NTx64|AllNT|78x64|781x64";
'var ver_61x86="7x86|NTx86|AllNT|78x86|781x86";
'var ver_62x64="8x64|NTx64|AllNT|78x64|All8x64";
'var ver_62x86="8x86|NTx86|AllNT|78x86|All8x86";
'var ver_63x64="81x64|NTx64|AllNT|781x64|All8x64";
'var ver_63x86="81x86|NTx86|AllNT|781x86|All8x86";
'
'STRICT - ���� ������ ������� ����� ������� �������, �� ������� ��� ������� ������������ ������ ��� ��� ��
'��� ������� �������-���� ����� ����� ������ ����� ���������
'============================================================================
Option Explicit

' �������������� ���������� ������� ������������ ������
Public Const strVer_51x64   As String = "5x64"
Public Const strVer_51x86   As String = "5x86"
Public Const strVer_60x64   As String = "6x64|67x64|NTx64|AllNT"
Public Const strVer_60x86   As String = "6x86|67x86|NTx86|AllNT"
Public Const strVer_61x64   As String = "7x64|67x64|78x64|781x64|NTx64|AllNT"
Public Const strVer_61x86   As String = "7x86|67x86|78x86|781x86|NTx86|AllNT"
Public Const strVer_62x64   As String = "8x64|78x64|All8x64|NTx64|AllNT"
Public Const strVer_62x86   As String = "8x86|78x86|All8x86|NTx86|AllNT"
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

    ReDim arrNotebookFilterListDef(21)

    arrNotebookFilterListDef(0) = "Acer;*acer*;*emachines*;*packard*bell*;*gateway*;*aspire*"
    arrNotebookFilterListDef(1) = "Apple;*apple*"
    arrNotebookFilterListDef(2) = "Asus;*asus*"
    arrNotebookFilterListDef(3) = "Clevo;*clevo*;*eurocom*;*sager*;*iru*;*viewsonic*;*viewbook*"
    arrNotebookFilterListDef(4) = "Dell;*dell*;*alienware*;*arima*;*jetway*;*gericom*"
    arrNotebookFilterListDef(5) = "Fujitsu;*fujitsu*;*sieme*"
    arrNotebookFilterListDef(6) = "Gigabyte;*gigabyte*;*ecs*;*elitegroup*;*roverbook*;*rover*"
    arrNotebookFilterListDef(7) = "HP;*hp*;*hewle*;*compaq*"
    arrNotebookFilterListDef(8) = "Intel;*intel*"
    arrNotebookFilterListDef(9) = "Lenovo;*lenovo*;*compal*;*ibm*;*wistron*"
    arrNotebookFilterListDef(10) = "LG;*lg*"
    arrNotebookFilterListDef(11) = "MTC;*mitac*;*mtc*;*depo*;*getac*"
    arrNotebookFilterListDef(12) = "MSI;*msi*;*micro-star*"
    arrNotebookFilterListDef(13) = "Panasonic;*panasonic*;*matsushita*"
    arrNotebookFilterListDef(14) = "Quanta;*quanta*;*prolink*;*nec*;*k-systems*;*benq*;*vizio*"
    arrNotebookFilterListDef(15) = "Pegatron;*pegatron*;*medion*"
    arrNotebookFilterListDef(16) = "Samsung;*samsung*"
    arrNotebookFilterListDef(17) = "Shuttle;*shuttle*"
    arrNotebookFilterListDef(18) = "Sony;*sony*;*vaio*"
    arrNotebookFilterListDef(19) = "Toshiba;*toshiba*"
    arrNotebookFilterListDef(20) = "Twinhead;*twinhead*;*durabook*"

End Sub

'Design by SamLab
'[NotebookVendor]
'FilterCount = 21
'Filter_1=Acer;*acer*;*emachines*;*packard*bell*;*gateway*;*aspire*
'Filter_2=Apple;*apple*
'Filter_3=Asus;*asus*
'Filter_4=Clevo;*clevo*;*eurocom*;*sager*;*iru*;*viewsonic*;*viewbook*
'Filter_5=Dell;*dell*;*alienware*;*arima*;*jetway*;*gericom*
'Filter_6=Fujitsu;*fujitsu*;*sieme*
'Filter_7=Gigabyte;*gigabyte*;*ecs*;*elitegroup*;*roverbook*;*rover*
'Filter_8=HP;*hp*;*hewle*;*compaq*
'Filter_9=Intel;*intel*
'Filter_10=Lenovo;*lenovo*;*compal*;*ibm*;*wistron*
'Filter_11=LG;*lg*
'Filter_12=MTC;*mitac*;*mtc*;*depo*;*getac*
'Filter_13=MSI;*msi*;*micro-star*
'Filter_14=Panasonic;*panasonic*;*matsushita*
'Filter_15=Quanta;*quanta*;*prolink*;*nec*;*k-systems*;*benq*;*vizio*
'Filter_16=Pegatron;*pegatron*;*medion*
'Filter_17=Samsung;*samsung*
'Filter_18=Shuttle;*shuttle*
'Filter_19=Sony;*sony*;*vaio*
'Filter_20=Toshiba;*toshiba*
'Filter_21=Twinhead;*twinhead*;*durabook*
