Attribute VB_Name = "mUpdate"
Option Explicit

Public mbCheckUpdNotEnd         As Boolean ' ����, ������������ ��������� �������� ���������� (��� ��� ����������� �����)

Public strLink()                As String
Public strLinkFull()            As String
Public strLinkHistory           As String
Public strLinkHistory_en        As String
Public strVersion               As String
Public strDateProg              As String
Public strDescription           As String
Public strDescription_en        As String
Public strRelease               As String
Public strUpdVersions()         As String
Public strUpdDescription()      As String

Private XMLHTTP                 As MSXML2.XMLHTTP30

Private Const iTimeOutInSecs       As Integer = 5                           ' ������� �������� ������ �� ������� � ��������
Private Const strXMLMainSection    As String = "//driversinstaller"         ' ������ ����� Xml-��������� ����������
Private Const strUrl_ProjectFolder As String = "Project/"                   ' ������� ������� �� �������, � ��� ���� ��� ����� xml
Private Const strUrl_UpdFile       As String = "dia_update2.xml"            ' ���� ������� ���� ���������� ���������
Private Const strUrl_TestFile      As String = "test.txt"                   ' ���� ��� �������� ����������� ����� ���������
Private Const strUrl_TestWWW       As String = "http://ya.ru/"              ' ���� ��� �������� ������� ���������� ��������

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long
Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CheckConnection2Server
'! Description (��������)  :   [�������� ������������� ����� (��������� Url_Test_Site) �� ����� ���������]
'! Parameters  (����������):   URL (String)
'!--------------------------------------------------------------------------------
Function CheckConnection2Server(ByVal URL As String) As String

    ' ������� ��������� ���� �� ������ URL$ � ��������� ��� ��� ������ LocalPath$
    Dim strResultText As String
    Dim strResultCode As String
    Dim errNum        As Long
    Dim tmstart, tmcurr, iTimeTaken

    On Error GoTo ErrCode

    ' ���� ���� ��������-����������, ��
    If CheckInternetConnection Then
        Set XMLHTTP = New MSXML2.XMLHTTP30
        tmstart = Now

        With XMLHTTP
            .Open "GET", Replace$(URL, vbBackslash, "/"), "True"
            .sEnd ""

            Do
                tmcurr = Now
                iTimeTaken = CInt(DateDiff("s", tmstart, tmcurr))

                ' ���� �������, �� �������
                If iTimeTaken > iTimeOutInSecs Then
                    .abort

                    Exit Do

                End If

                Sleep 50
                DoEvents
            Loop While .readyState <> 4

            strResultText = .statusText
            strResultCode = .Status
        End With

        If StrComp(strResultText, "OK", vbTextCompare) = 0 Then
            CheckConnection2Server = "OK"
        Else
            CheckConnection2Server = "Error:" & strResultCode & " - " & strResultText & " - " & XMLHTTP.responseText
        End If
    End If

    Exit Function

ErrCode:
    errNum = Err.Number
    Debug.Print Err.Number & strSpace & Err.Description & strSpace & Err.LastDllError

    If errNum <> 0 Then
        If mbDebugStandart Then DebugMode str5VbTab & "CheckConnection2Server: " & " Error: �" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError) & vbNewLine & _
                  str5VbTab & "CheckConnection2Server: Err.Number: " & Err.Number & " Err.Description: " & Err.Description
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CheckInternetConnection
'! Description (��������)  :   [�������� ������� �������� ����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function CheckInternetConnection() As Boolean

    Dim aux As String * 255
    Dim r   As Long

    r = InternetGetConnectedStateEx(r, aux, 254, 0)

    If r = 1 Then
        CheckInternetConnection = True
    Else
        CheckInternetConnection = False
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub CheckUpd
'! Description (��������)  :   [�������� ����� ������ ��������� � �������������� MSXML]
'! Parameters  (����������):   Start (Boolean = True)
'!--------------------------------------------------------------------------------
Public Sub CheckUpd(Optional ByVal Start As Boolean = True)

    Dim strTextNodeName         As String
    Dim miNodeIndex             As Integer
    Dim strVerTemp              As String
    Dim lngResultCompare        As eVerCompareResult
    Dim strUrl_TestWWW_Result   As String
    Dim strUrl_Test_Site        As String
    Dim strUrl_Test_Site_Result As String
    Dim strUrl_Request          As String
    

    If mbDebugStandart Then DebugMode "CheckUpd-Start" & vbNewLine & _
               vbTab & "CheckUpd-Options: " & Start
    
    ' ������ ��������� �������� �������� ����������
    mbCheckUpdNotEnd = True

    On Error Resume Next

    '������ ������ ��������� (�������������)
    strVerTemp = strProductVersion
    
    ' �������� ������� ������� �� google/yandex
    strUrl_TestWWW_Result = CheckConnection2Server(strUrl_TestWWW)
    ' ���� ������ ����, ����� ��������� ������
    If StrComp(strUrl_TestWWW_Result, "OK", vbTextCompare) = 0 Then
        
        ' ��������� ������ ��� ���������� ��������� ����� � ����� ���������
        strUrl_Test_Site = strUrl_MainWWWSite & strUrl_TestFile
        
        ' �������� ������� ������� �� ����� adia-project - ���� test.txt
        strUrl_Test_Site_Result = CheckConnection2Server(strUrl_Test_Site)

        If StrComp(strUrl_Test_Site_Result, "OK", vbTextCompare) = 0 Then

            Dim xmlDoc       As DOMDocument30
            Dim nodeList     As IXMLDOMNodeList
            Dim xmlNode      As IXMLDOMNode
            Dim propertyNode As IXMLDOMElement
            
            ' ��������� ������ ��� ���������� ����� ������� ����������
            'strUrl_Request = strAppPathBackSL & "dia_update2.xml"
            strUrl_Request = strUrl_MainWWWSite & strUrl_ProjectFolder & strUrl_UpdFile
            
            Set xmlDoc = New DOMDocument
            xmlDoc.async = False
            
            ' ��������� ���� ������� ����������
            If Not xmlDoc.Load(strUrl_Request) Then
                ChangeStatusBarText strMessages(126)

                If Not Start Then
                    MsgBox strMessages(126), vbInformation, strMessages(54)
                End If

            Else
                Set nodeList = xmlDoc.documentElement.selectNodes(strXMLMainSection)
                Set xmlNode = nodeList(0)
                miNodeIndex = 0

                For Each propertyNode In xmlNode.childNodes

                    strTextNodeName = vbNullString
                    strTextNodeName = LCase$(xmlNode.childNodes(miNodeIndex).nodeName)

                    ' ������ �� ����� *_update2.xml
                    Select Case strTextNodeName
                        
                        ' ������ ���������
                        Case "version"
                            strVersion = xmlNode.childNodes(miNodeIndex).Text

                        ' ���� ���������
                        Case "date"
                            strDateProg = xmlNode.childNodes(miNodeIndex).Text

                        ' ��� ��������� - beta/release
                        Case "release"
                            strRelease = xmlNode.childNodes(miNodeIndex).Text

                        ' ������ �� ������ ������� ��������� - RUS
                        Case "linkhistory"
                            strLinkHistory = xmlNode.childNodes(miNodeIndex).Text

                        ' ������ �� ������ ������� ��������� - ENG
                        Case "linkhistory_en"
                            strLinkHistory_en = xmlNode.childNodes(miNodeIndex).Text
                    End Select

                    miNodeIndex = miNodeIndex + 1
                Next

                '**** ��������� ������ ��������
                lngResultCompare = CompareByVersion(strVersion, strVerTemp)

                ' ������ ����� ��������� � ����� ����
                Select Case lngResultCompare

                    Case crGreaterVer

                        If StrComp(strRelease, "beta", vbTextCompare) = 0 Then
                            If Not mbUpdateCheckBeta Then
                                If mbDebugStandart Then DebugMode vbTab & "The version on the site is Beta. In options check for beta are disable. Break function!!!"
                                ChangeStatusBarText strMessages(56)

                                If Not Start Then
                                    If MsgBox(strMessages(56) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                        frmCheckUpdate.Show vbModal, frmMain
                                    Else

                                        Exit Sub

                                    End If

                                Else
                                    ChangeStatusBarText strMessages(56)
                                End If

                            Else
                                frmCheckUpdate.Show vbModal, frmMain
                            End If

                        Else
                            frmCheckUpdate.Show vbModal, frmMain
                        End If

                    Case crEqualVer
                        ChangeStatusBarText strMessages(56)

                        If Not Start Then
                            If MsgBox(strMessages(56) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                frmCheckUpdate.Show vbModal, frmMain
                            End If
                        End If

                    Case crLessVer
                        ChangeStatusBarText strMessages(55)

                        If Not Start Then
                            If MsgBox(strMessages(55) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                frmCheckUpdate.Show vbModal, frmMain
                            End If
                        End If

                    Case Else
                        ChangeStatusBarText strMessages(102)

                        If Not Start Then
                            MsgBox strMessages(102), vbInformation, strProductName
                        End If

                End Select

                Set xmlNode = Nothing
                Set nodeList = Nothing
            End If

        Else
            If mbDebugStandart Then DebugMode vbTab & "CheckUPD-Site: " & strMessages(53) & vbNewLine & "Error: " & strUrl_Test_Site_Result
            ChangeStatusBarText strMessages(143)

            If Not Start Then
                MsgBox strMessages(143) & vbNewLine & "Error: " & strUrl_Test_Site_Result, vbInformation, strMessages(54)
            End If
        End If

    ' �� 99% �������� �����������
    Else
        If mbDebugStandart Then DebugMode vbTab & "CheckUPD-Inet: " & strMessages(53) & vbNewLine & "Error: " & strUrl_TestWWW_Result
        ChangeStatusBarText strMessages(53)

        If Not Start Then
            MsgBox strMessages(53) & vbNewLine & "Error: " & strUrl_TestWWW_Result, vbInformation, strMessages(54)
        End If
    End If

    Set xmlDoc = Nothing
    
    ' ������ ��������� �������� �������� ����������
    mbCheckUpdNotEnd = False

    On Error GoTo 0

    If mbDebugStandart Then DebugMode "CheckUpd-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function GetDeltaDay
'! Description (��������)  :   [������������ ������� � ���� ����� ����� ������]
'! Parameters  (����������):   dtFirstDate (Date)
'                              dtSecondDate (Date)
'!--------------------------------------------------------------------------------
Private Function GetDeltaDay(ByVal dtFirstDate As Date, ByVal dtSecondDate As Date) As Integer
    GetDeltaDay = CInt(DateDiff("d", dtFirstDate, dtSecondDate))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadUpdateData
'! Description (��������)  :   [�������� ������ ������ ������ ��������� � �������������� MSXML]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub LoadUpdateData()

    Dim xmlDoc          As DOMDocument
    Dim nodeList        As IXMLDOMNodeList
    Dim xmlNode         As IXMLDOMNode
    Dim propertyNode    As IXMLDOMElement
    Dim strTextNodeName As String
    Dim miNodeIndex     As Integer
    Dim strVersionsTemp As String
    Dim I               As Long
    Dim strUrl_Request  As String
    Dim lngUbound       As Long

    On Error Resume Next
   
    ' ��������� ������ ��� ���������� ����� ������� ����������
    'strUrl_Request = strAppPathBackSL & "d*_update2.xml"
    strUrl_Request = strUrl_MainWWWSite & strUrl_ProjectFolder & strUrl_UpdFile
    
    Set xmlDoc = New DOMDocument
    xmlDoc.async = False
            
    If Not xmlDoc.Load(strUrl_Request) Then
        ChangeStatusBarText strMessages(53)
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeList = xmlDoc.documentElement.selectNodes(strXMLMainSection)
        Set xmlNode = nodeList(0)
        miNodeIndex = 0

        ' ������ �� ����� d*_update2.xml - ������ ������
        For Each propertyNode In xmlNode.childNodes

            strTextNodeName = vbNullString
            strTextNodeName = LCase$(xmlNode.childNodes(miNodeIndex).nodeName)
            
            ' ���� ������ � ����� xml �� ������� ������ ���������
            If StrComp(strTextNodeName, "versions") = 0 Then
            
                strVersionsTemp = xmlNode.childNodes(miNodeIndex).Text
                strUpdVersions = Split(strVersionsTemp, strSemiColon)
                lngUbound = UBound(strUpdVersions)

                ReDim strUpdDescription(lngUbound, 2)
                ReDim strLink(lngUbound, 6)
                ReDim strLinkFull(lngUbound, 6)

                ' ������ �� ����� %ver%.xml - �������� �������� ���������
                For I = 0 To lngUbound
                    LoadUpdDescription strUpdVersions(I), I
                Next I

            End If

            miNodeIndex = miNodeIndex + 1
        Next

        Set xmlNode = Nothing
        Set nodeList = Nothing
    End If

    Set xmlDoc = Nothing

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub LoadUpdDescription
'! Description (��������)  :   [�������� ������ ����� ��������� ���������� ����������
'!                              �� ����� ��������: ������ �� ������������ ��� ����������, �������� ���������� � rtf-������� � rus/eng]
'! Parameters  (����������):   strVer (String)
'                              lngIndexVer (Long)
'!--------------------------------------------------------------------------------
Public Sub LoadUpdDescription(ByVal strVer As String, ByVal lngIndexVer As Long)

    Dim xmlDocVers       As DOMDocument
    Dim nodeListVers     As IXMLDOMNodeList
    Dim xmlNodeVers      As IXMLDOMNode
    Dim propertyNodeVers As IXMLDOMElement
    Dim strTextNodeName  As String
    Dim miNodeIndex      As Integer
    Dim strUrl_Request   As String

    ' ��������� ������ ��� ���������� ����� ��������� ����������
    'strUrl_Request = strAppPathBackSL & strVer & ".xml"
    strUrl_Request = strUrl_MainWWWSite & strUrl_ProjectFolder & strVer & ".xml"

    Set xmlDocVers = New DOMDocument
    xmlDocVers.async = False
    
    If Not xmlDocVers.Load(strUrl_Request) Then
        ChangeStatusBarText strMessages(53)
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeListVers = xmlDocVers.documentElement.selectNodes(strXMLMainSection)
        Set xmlNodeVers = nodeListVers(0)
        miNodeIndex = 0

        For Each propertyNodeVers In xmlNodeVers.childNodes

            strTextNodeName = vbNullString
            strTextNodeName = LCase$(xmlNodeVers.childNodes(miNodeIndex).nodeName)

            Select Case strTextNodeName

                ' �������� ���������
                Case "description"
                    strUpdDescription(lngIndexVer, 0) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "description_en"
                    strUpdDescription(lngIndexVer, 1) = xmlNodeVers.childNodes(miNodeIndex).Text

                ' ������/������� �� ���� ���������� - ������/���������
                Case "link"
                    strLink(lngIndexVer, 0) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "link_header"
                    strLink(lngIndexVer, 1) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "link_mirror1"
                    strLink(lngIndexVer, 2) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "link_header1"
                    strLink(lngIndexVer, 3) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "link_mirror2"
                    strLink(lngIndexVer, 4) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "link_header2"
                    strLink(lngIndexVer, 5) = xmlNodeVers.childNodes(miNodeIndex).Text

                ' ������/������� �� ���� ������� ������������ - ������/���������
                Case "linkfull"
                    strLinkFull(lngIndexVer, 0) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "linkfull_header"
                    strLinkFull(lngIndexVer, 1) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "linkfull_mirror1"
                    strLinkFull(lngIndexVer, 2) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "linkfull_header1"
                    strLinkFull(lngIndexVer, 3) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "linkfull_mirror2"
                    strLinkFull(lngIndexVer, 4) = xmlNodeVers.childNodes(miNodeIndex).Text

                Case "linkfull_header2"
                    strLinkFull(lngIndexVer, 5) = xmlNodeVers.childNodes(miNodeIndex).Text
            End Select

            miNodeIndex = miNodeIndex + 1
        Next

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub ShowUpdateToolTip
'! Description (��������)  :   [����� ��� ������������� ������������ ��������� � ��������� ������� ����������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub ShowUpdateToolTip()

    Dim mbShowToolTip As Boolean
    Dim intDeltaDay   As Integer
    Dim dtToolTipDate As Date
    Dim strTTipDate   As String
    
    If GetDeltaDay(Date, CDate(strDateProgram)) > 180 Then
        If mbUpdateToolTip Then
            ' ��������� ���� ����� ������������ ����������� ��������� ��������� ���
            strTTipDate = GetSetting(App.ProductName, "UpdateToolTip", "Show at Date", vbNullString)

            '���� �� ������������ (�.� �������� ������), �� ����������
            If LenB(strTTipDate) = 0 Then
                mbShowToolTip = True
            Else
                '���� ������������, �� ��������� ��� �����
                dtToolTipDate = CDate(strTTipDate)
                intDeltaDay = GetDeltaDay(Date, dtToolTipDate)

                ' ���� ������������ ����� ���� ���� �����, �� ���������� �����
                If intDeltaDay >= 5 Then
                    mbShowToolTip = True
                End If
            End If

        Else
            mbShowToolTip = False
        End If

    Else
        mbShowToolTip = False
    End If

    ' ���� ��� ������� ���������, �� ���������� ���������
    ' "��������, ������������ ����, ������ ��������� 'DIA/DBS' ��� ��������! "
    If mbShowToolTip Then
        ShowNotifyMessage strMessages(107)
    End If

End Sub

