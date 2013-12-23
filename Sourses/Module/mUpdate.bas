Attribute VB_Name = "mUpdate"
Option Explicit

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

Private Const iTimeOutInSecs    As Integer = 5
Private Const strXMLMainSection As String = "//driversinstaller"
Private Const Url_Request       As String = "http://www.adia-project.net/Project/dia_update2.xml"
Private Const Url_Test_WWW      As String = "http://ya.ru/"
Private Const Url_Test_Site     As String = "http://adia-project.net/test.txt"

Private Declare Function InternetGetConnectedStateEx Lib "wininet.dll" (ByRef lpdwFlags As Long, ByVal lpszConnectionName As String, ByVal dwNameLen As Integer, ByVal dwReserved As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckConnection2Server
'! Description (Описание)  :   [Проверка существования файла на сервере]
'! Parameters  (Переменные):   URL (String)
'!--------------------------------------------------------------------------------
Function CheckConnection2Server(ByVal URL As String) As String

    ' Функция скачивает файл по ссылке URL$
    ' и сохраняет его под именем LocalPath$
    Dim strResultText As String
    Dim strResultCode As String
    Dim errNum        As Long
    Dim tmstart, tmcurr, iTimeTaken

    On Error GoTo ErrCode

    If CheckInternetConnection Then
        Set XMLHTTP = New MSXML2.XMLHTTP30
        tmstart = Now

        With XMLHTTP
            '.Open "GET", Replace$(URL, vbBackslash, "/"), "False"
            .Open "GET", Replace$(URL, vbBackslash, "/"), "True"
            .sEnd ""

            Do
                tmcurr = Now
                iTimeTaken = CInt(DateDiff("s", tmstart, tmcurr))

                ' Если таймаут, то выходим
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
    Debug.Print Err.Number & " " & Err.Description & " " & Err.LastDllError

    If errNum <> 0 Then
        DebugMode str5VbTab & "CheckConnection2Server: " & " Error: №" & Err.LastDllError & " - " & ApiErrorText(Err.LastDllError)
        DebugMode str5VbTab & "CheckConnection2Server: Err.Number: " & Err.Number & " Err.Description: " & Err.Description
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckInternetConnection
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function CheckInternetConnection() As Boolean

    Dim aux As String * 255
    Dim R   As Long

    R = InternetGetConnectedStateEx(R, aux, 254, 0)

    If R = 1 Then
        CheckInternetConnection = True
    Else
        CheckInternetConnection = False
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub CheckUpd
'! Description (Описание)  :   [Проверка новых версий программы с использованием MSXML]
'! Parameters  (Переменные):   Start (Boolean = True)
'!--------------------------------------------------------------------------------
Public Sub CheckUpd(Optional ByVal Start As Boolean = True)

    Dim TextNodeName         As String
    Dim NodeIndex            As Integer
    Dim strVerTemp           As String
    Dim strResultCompare     As String
    Dim Url_Test_Result_WWW  As String
    Dim Url_Test_Result_Site As String

    DebugMode "CheckUpd-Start"
    mbCheckUpdNotEnd = True
    DebugMode vbTab & "CheckUpd-Options: " & Start

    On Error Resume Next

    'Узнаем версию программы (установленной)
    strVerTemp = strProductVersion
    'Url_Request = strAppPathBackSL & "dia_update2.xml"
    ' проверка наличия доступа до google
    Url_Test_Result_WWW = CheckConnection2Server(Url_Test_WWW)

    If StrComp(Url_Test_Result_WWW, "OK", vbTextCompare) = 0 Then
        ' проверка наличия доступа до сайта adia-project
        Url_Test_Result_Site = CheckConnection2Server(Url_Test_Site)

        If StrComp(Url_Test_Result_Site, "OK", vbTextCompare) = 0 Then

            Dim xmlDoc       As DOMDocument30
            Dim nodeList     As IXMLDOMNodeList
            Dim xmlNode      As IXMLDOMNode
            Dim propertyNode As IXMLDOMElement

            Set xmlDoc = New DOMDocument
            xmlDoc.async = False

            ' загружаем файл
            If Not xmlDoc.Load(Url_Request) Then
                ChangeStatusTextAndDebug strMessages(126)

                If Not Start Then
                    MsgBox strMessages(126), vbInformation, strMessages(54)
                End If

            Else
                Set nodeList = xmlDoc.documentElement.selectNodes(strXMLMainSection)
                Set xmlNode = nodeList(0)
                NodeIndex = 0

                For Each propertyNode In xmlNode.childNodes

                    TextNodeName = vbNullString
                    TextNodeName = xmlNode.childNodes(NodeIndex).nodeName

                    Select Case TextNodeName

                            ' Данные из файла dia_update2.xml
                            ' Версия проги
                        Case "version"
                            strVersion = xmlNode.childNodes(NodeIndex).Text

                            ' Дата проги
                        Case "date"
                            strDateProg = xmlNode.childNodes(NodeIndex).Text

                        Case "release"
                            strRelease = xmlNode.childNodes(NodeIndex).Text

                            ' Ссылка на Полную историю изменений
                        Case "linkHistory"
                            strLinkHistory = xmlNode.childNodes(NodeIndex).Text

                        Case "linkHistory_en"
                            strLinkHistory_en = xmlNode.childNodes(NodeIndex).Text
                    End Select

                    NodeIndex = NodeIndex + 1
                Next

                '**** Сравнение версий программ
                strResultCompare = CompareByVersion(strVersion, strVerTemp)

                ' Анализ итога сравнения и показ окна
                Select Case strResultCompare

                    Case ">"

                        If StrComp(strRelease, "beta", vbTextCompare) = 0 Then
                            If Not mbUpdateCheckBeta Then
                                DebugMode vbTab & "The version on the site is Beta. In options check for beta are disable. Break function!!!"
                                ChangeStatusTextAndDebug strMessages(56)

                                If Not Start Then
                                    If MsgBox(strMessages(56) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                        frmCheckUpdate.Show vbModal, frmMain
                                    Else

                                        Exit Sub

                                    End If

                                Else
                                    ChangeStatusTextAndDebug strMessages(56)
                                End If

                            Else
                                frmCheckUpdate.Show vbModal, frmMain
                            End If

                        Else
                            frmCheckUpdate.Show vbModal, frmMain
                        End If

                    Case "="
                        ChangeStatusTextAndDebug strMessages(56)

                        If Not Start Then
                            If MsgBox(strMessages(56) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                frmCheckUpdate.Show vbModal, frmMain
                            End If
                        End If

                    Case "<"
                        ChangeStatusTextAndDebug strMessages(55)

                        If Not Start Then
                            If MsgBox(strMessages(55) & strMessages(144), vbQuestion + vbYesNo, strProductName) = vbYes Then
                                frmCheckUpdate.Show vbModal, frmMain
                            End If
                        End If

                    Case Else
                        ChangeStatusTextAndDebug strMessages(102)

                        If Not Start Then
                            MsgBox strMessages(102), vbInformation, strProductName
                        End If

                End Select

                Set xmlNode = Nothing
                Set nodeList = Nothing
            End If

        Else
            DebugMode vbTab & "CheckUPD-Site: " & strMessages(53) & vbNewLine & "Error: " & Url_Test_Result_Site
            ChangeStatusTextAndDebug strMessages(143)

            If Not Start Then
                MsgBox strMessages(143) & vbNewLine & "Error: " & Url_Test_Result_Site, vbInformation, strMessages(54)
            End If
        End If

    Else
        DebugMode vbTab & "CheckUPD-Inet: " & strMessages(53) & vbNewLine & "Error: " & Url_Test_Result_WWW
        ChangeStatusTextAndDebug strMessages(53)

        If Not Start Then
            MsgBox strMessages(53) & vbNewLine & "Error: " & Url_Test_Result_WWW, vbInformation, strMessages(54)
        End If
    End If

    Set xmlDoc = Nothing
    mbCheckUpdNotEnd = False

    On Error GoTo 0

    DebugMode "CheckUpd-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DeltaDay
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Private Function DeltaDay() As Integer

    Dim CurrentDate As Date
    Dim BuildDate   As Date
    Dim DeltaTemp   As Integer

    CurrentDate = Date
    BuildDate = CDate(strDateProgram)
    DeltaTemp = CInt(CurrentDate - BuildDate)
    DeltaDay = DeltaTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function DeltaDayNew
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   dtFirstDate (Date)
'                              dtSecondDate (Date)
'!--------------------------------------------------------------------------------
Private Function DeltaDayNew(ByVal dtFirstDate As Date, ByVal dtSecondDate As Date) As Integer

    Dim DeltaTemp As Integer

    DeltaTemp = CInt(dtFirstDate - dtSecondDate)
    DeltaDayNew = DeltaTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadUpdateData
'! Description (Описание)  :   [Проверка новых версий программы с использованием MSXML]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub LoadUpdateData()

    Dim xmlDoc          As DOMDocument
    Dim nodeList        As IXMLDOMNodeList
    Dim xmlNode         As IXMLDOMNode
    Dim propertyNode    As IXMLDOMElement
    Dim Url_Request     As String
    Dim TextNodeName    As String
    Dim NodeIndex       As Integer
    Dim strVersionsTemp As String
    Dim i               As Long

    On Error Resume Next

    Set xmlDoc = New DOMDocument
    xmlDoc.async = False
    Url_Request = "http://www.adia-project.net/Project/dia_update2.xml"

    'Url_Request = strAppPathBackSL & "dia_update2.xml"
    If Not xmlDoc.Load(Url_Request) Then
        ChangeStatusTextAndDebug strMessages(53)
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeList = xmlDoc.documentElement.selectNodes(strXMLMainSection)
        Set xmlNode = nodeList(0)
        NodeIndex = 0

        For Each propertyNode In xmlNode.childNodes

            TextNodeName = vbNullString
            TextNodeName = xmlNode.childNodes(NodeIndex).nodeName

            Select Case TextNodeName

                    ' Данные из файла dia_update2.xml
                    ' массив версий
                Case "versions"
                    strVersionsTemp = xmlNode.childNodes(NodeIndex).Text
                    strUpdVersions = Split(strVersionsTemp, ";")

                    ReDim strUpdDescription(UBound(strUpdVersions), 2) As String
                    ReDim strLink(UBound(strUpdVersions), 6) As String
                    ReDim strLinkFull(UBound(strUpdVersions), 6) As String

                    ' Данные из файла %ver%.xml
                    'Загрузка описаний изменений
                    For i = LBound(strUpdVersions) To UBound(strUpdVersions)
                        LoadUpdDescription strUpdVersions(i), i
                    Next

            End Select

            NodeIndex = NodeIndex + 1
        Next

        Set xmlNode = Nothing
        Set nodeList = Nothing
    End If

    Set xmlDoc = Nothing

    On Error GoTo 0

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadUpdDescription
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strVer (String)
'                              lngIndexVer (Long)
'!--------------------------------------------------------------------------------
Public Sub LoadUpdDescription(ByVal strVer As String, ByVal lngIndexVer As Long)

    Dim xmlDocVers       As DOMDocument
    Dim nodeListVers     As IXMLDOMNodeList
    Dim xmlNodeVers      As IXMLDOMNode
    Dim propertyNodeVers As IXMLDOMElement
    Dim Url_Request      As String
    Dim TextNodeName     As String
    Dim NodeIndex        As Integer

    Set xmlDocVers = New DOMDocument
    xmlDocVers.async = False
    Url_Request = "http://www.adia-project.net/Project/" & strVer & ".xml"

    'Url_Request = strAppPath & vbBackslash & strVer & ".xml"
    If Not xmlDocVers.Load(Url_Request) Then
        ChangeStatusTextAndDebug strMessages(53)
        MsgBox strMessages(53), vbInformation, strMessages(54)
    Else
        Set nodeListVers = xmlDocVers.documentElement.selectNodes(strXMLMainSection)
        Set xmlNodeVers = nodeListVers(0)
        NodeIndex = 0

        For Each propertyNodeVers In xmlNodeVers.childNodes

            TextNodeName = vbNullString
            TextNodeName = xmlNodeVers.childNodes(NodeIndex).nodeName

            Select Case TextNodeName

                    ' Описание изменений
                Case "description"
                    strUpdDescription(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "description_en"
                    strUpdDescription(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                    ' Ссылка на обновление
                Case "link"
                    strLink(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header"
                    strLink(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_Mirror1"
                    strLink(lngIndexVer, 2) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header1"
                    strLink(lngIndexVer, 3) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_Mirror2"
                    strLink(lngIndexVer, 4) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "link_header2"
                    strLink(lngIndexVer, 5) = xmlNodeVers.childNodes(NodeIndex).Text

                    ' Ссылка на дистрибутив
                Case "linkFull"
                    strLinkFull(lngIndexVer, 0) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header"
                    strLinkFull(lngIndexVer, 1) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_Mirror1"
                    strLinkFull(lngIndexVer, 2) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header1"
                    strLinkFull(lngIndexVer, 3) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_Mirror2"
                    strLinkFull(lngIndexVer, 4) = xmlNodeVers.childNodes(NodeIndex).Text

                Case "linkFull_header2"
                    strLinkFull(lngIndexVer, 5) = xmlNodeVers.childNodes(NodeIndex).Text
            End Select

            NodeIndex = NodeIndex + 1
        Next

    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShowUpdateToolTip
'! Description (Описание)  :   [Показ напоминания об обновлении]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub ShowUpdateToolTip()

    Dim mbShowToolTip As Boolean
    Dim intDeltaDay   As Integer
    Dim dtToolTipDate As Date
    Dim strTTipDate   As String

    If DeltaDay > 180 Then
        If mbUpdateToolTip Then
            strTTipDate = GetSetting(App.ProductName, "UpdateToolTip", "Show at Date", vbNullString)

            If LenB(strTTipDate) = 0 Then
                mbShowToolTip = True
            Else
                dtToolTipDate = CDate(strTTipDate)
                intDeltaDay = DeltaDayNew(Date, dtToolTipDate)

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

    ' Если все условия выполнены, то показываем сообщение
    ' "Возможно, используемая вами, версия программы 'Помощник установки драйверов' уже устарела! "
    If mbShowToolTip Then
        ShowNotifyMessage strMessages(107)
    End If

End Sub
