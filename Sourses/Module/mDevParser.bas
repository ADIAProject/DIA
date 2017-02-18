Attribute VB_Name = "mDevParser"
Option Explicit

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CollectCmdString
'! Description (��������)  :   [�������� ���������� ������ ������� ��������� DPInst]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function CollectCmdString() As String

    Dim strCmdStringDPInstTemp As String

    If mbDpInstLegacyMode Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/LM "
    End If

    If mbDpInstPromptIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/P "
    End If

    If mbDpInstForceIfDriverIsNotBetter Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/F "
    End If

    If mbDpInstSuppressAddRemovePrograms Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SA "
    End If

    If mbDpInstSuppressWizard Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SW "
    End If

    If mbDpInstQuietInstall Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/Q "
    End If

    If mbDpInstScanHardware Then
        strCmdStringDPInstTemp = strCmdStringDPInstTemp & "/SH "
    End If

    ' �������������� ������
    CollectCmdString = strCmdStringDPInstTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CompareDevDBVersion
'! Description (��������)  :   [��������� ������ ��� ���������, � ���������� � ���������]
'! Parameters  (����������):   strDevDBFullFileName (String)
'                              strPathDRP (String)
'!--------------------------------------------------------------------------------
Public Function CompareDevDBVersion(strDevDBFullFileName As String) As Boolean

    Dim lngResult         As Long
    Dim strFilePath_woExt As String

    strFilePath_woExt = GetFileName_woExt(strDevDBFullFileName)
    lngResult = IniLongPrivate(GetFileNameFromPath(strFilePath_woExt), "Version", BackslashAdd2Path(GetPathNameFromPath(strFilePath_woExt)) & "DevDBVersions.ini")

    If lngResult = 9999 Then
        CompareDevDBVersion = False
    Else
        CompareDevDBVersion = Not (lngResult <> lngDevDBVersion)
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function DeleteDriverbyHwid
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strHwid (String)
'!--------------------------------------------------------------------------------
Public Function DeleteDriverbyHwid(ByVal strHwid As String) As Boolean

    Dim cmdString     As String
    Dim strDevConTemp As String

    cmdString = strQuotes & strDevConExePath & strQuotes & strSpace & strQuotes & strDevConTemp & strQuotes & strSpace & strQuotes & strHwidsTxtPath & strQuotes & " 4 " & strQuotes & strHwid & strQuotes

    If RunAndWaitNew(cmdString, strWorkTemp, vbNormalFocus) = False Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        DeleteDriverbyHwid = False
    Else
        DeleteDriverbyHwid = True
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DevParserLocalHwids2
'! Description (��������)  :   [������� ��������� ����� devcon ��� ��������� ���������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub DevParserLocalHwids2()

    Dim strContent          As String
    Dim ii                  As Long
    Dim lngStrCnt           As Long
    Dim miStatus            As Long
    Dim strID               As String
    Dim strIDOrig           As String
    Dim strIDCutting        As String
    Dim strName             As String
    Dim strName_x()         As String
    Dim miMaxCountArr       As Long
    Dim lngRecCountArr      As Long
    Dim strID_x()           As String
    Dim objRegExpDevcon     As RegExp
    Dim objMatchesDevcon    As MatchCollection
    Dim objMatch            As Match
    Dim strStatus           As String

    Set objRegExpDevcon = New RegExp

    With objRegExpDevcon
        .Pattern = "(^[^\n\r\s][^\n\r]+)\r\n(\s+[^\n\r]+\r\n)*[^\n\r]*((?:DEVICE IS|DEVICE HAS|DRIVER IS|DRIVER HAS)[^\r]+)"
        .MultiLine = True
        '.IgnoreCase = True
        .Global = True
    End With

    If mbDebugDetail Then DebugMode "DevParserLocalHwids2-Start"

    If FileExists(strHwidsTxtPath) Then
        FileReadData strHwidsTxtPath, strContent
        strContent = UCase$(strContent)
        Set objMatchesDevcon = objRegExpDevcon.Execute(strContent)
        miMaxCountArr = 100

        ' ������������ ���-�� ��������� � �������
        ReDim arrHwidsLocal(miMaxCountArr)

        lngStrCnt = objMatchesDevcon.count
        lngRecCountArr = 0

        For ii = 0 To lngStrCnt - 1
            Set objMatch = objMatchesDevcon.item(ii)

            ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
            If lngRecCountArr = miMaxCountArr Then
                miMaxCountArr = miMaxCountArr + miMaxCountArr

                ReDim Preserve arrHwidsLocal(miMaxCountArr)

            End If

            ' �������� ������
            With objMatch
                strID = .SubMatches(0)
                strName = .SubMatches(1)
                strStatus = .SubMatches(2)
            End With

            'objMatch
            strID = Trim$(Replace$(strID, vbNewLine, vbNullString))
            ' ��������� �� "\"
            strIDOrig = strID

            If InStr(strID, vbBackslash) Then
                strID_x = Split(strID, vbBackslash)
                strID = strID_x(0) & vbBackslash & strID_x(1)
            End If

            strIDCutting = ParseDoubleHwid(strID)

            '���� �� ������ � ������ ����������, �� ����������
            If strExcludeHWID <> "*" Then
                If Not MatchSpec(strID, strExcludeHWID) Then
                    '"���: " & strID & " present in " & strExcludeHWID
                    miStatus = 0

                    ' ���������� �������
                    If InStr(strStatus, "RUNNING") Then
                        miStatus = 1
                    End If

                    If LenB(strName) Then
                        strName_x = Split(strName, vbNewLine)
                        strName = strName_x(0)
                    End If

                    strName = Replace$(strName, vbNewLine, vbNullString)
                    strName = Replace$(strName, "NAME:", vbNullString)
                    strName = Trim$(strName)

                    If Len(strID) > 3 Then
                        arrHwidsLocal(lngRecCountArr).HWID = strID
                        arrHwidsLocal(lngRecCountArr).DevName = strName
                        arrHwidsLocal(lngRecCountArr).Status = miStatus
                        arrHwidsLocal(lngRecCountArr).HWIDOrig = strIDOrig
                        arrHwidsLocal(lngRecCountArr).HWIDCutting = strIDCutting
                        lngRecCountArr = lngRecCountArr + 1
                    End If
                End If

            Else
                miStatus = 0

                ' ���������� �������
                If InStr(strStatus, "RUNNING") Then
                    miStatus = 1
                End If

                If LenB(strName) Then
                    strName_x = Split(strName, vbNewLine)
                    strName = strName_x(0)
                End If

                strName = Replace$(strName, vbNewLine, vbNullString)
                strName = Replace$(strName, "NAME:", vbNullString)
                strName = Trim$(strName)

                If Len(strID) > 3 Then
                    ' ID ������������
                    arrHwidsLocal(lngRecCountArr).HWID = strID
                    ' ������������ ������������
                    arrHwidsLocal(lngRecCountArr).DevName = strName
                    ' ������ ������������
                    arrHwidsLocal(lngRecCountArr).Status = miStatus
                    arrHwidsLocal(lngRecCountArr).HWIDOrig = strIDOrig
                    arrHwidsLocal(lngRecCountArr).HWIDCutting = strIDCutting
                    lngRecCountArr = lngRecCountArr + 1
                End If
            End If

        Next

        ' ������������� ������ �� �������� ���-�� �������
        If lngRecCountArr Then

            ReDim Preserve arrHwidsLocal(lngRecCountArr - 1)

        Else

            ReDim Preserve arrHwidsLocal(0)

        End If

        If SaveHwidsArray2File(strResultHwidsTxtPath, arrHwidsLocal) = False Then
            MsgBox strMessages(45) & vbNewLine & strResultHwidsTxtPath, vbCritical + vbInformation, strProductName
        End If

    Else
        If mbDebugDetail Then DebugMode "DevParserLocalHwids2-False: " & strHwidsTxtPath & vbTab & strMessages(46)
        MsgBox strHwidsTxtPath & vbNewLine & strMessages(46), vbInformation, strProductName
        Unload frmMain
    End If

    If mbDebugDetail Then DebugMode "DevParserLocalHwids2-End"
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function ParseDoubleHwid
'! Description (��������)  :   [type_description_here]
'! Parameters  (����������):   strValuer (String)
'!--------------------------------------------------------------------------------
Public Function ParseDoubleHwid(ByVal strValuer As String) As String

    Dim strValuer_x() As String

    If LenB(strValuer) Then

        ' ��������� �� "\" - ��������� ������ xxx\yyy
        If InStr(strValuer, vbBackslash) Then
            strValuer_x = Split(strValuer, vbBackslash)

            If UBound(strValuer_x) Then
                strValuer = strValuer_x(0) & vbBackslash & strValuer_x(1)
            Else
                strValuer = strValuer_x(0)
            End If
        End If
    End If

    ParseDoubleHwid = strValuer
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function RunDevcon
'! Description (��������)  :   [������ ��������� Devcon ��� ������ ��������� ��� ����� ���������� �� HWID]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function RunDevcon() As Boolean

    Dim cmdString As String

    If FileExists(strHwidsTxtPath) Then
        DeleteFiles strHwidsTxtPath
    End If

    cmdString = "cmd.exe /c " & strQuotes & strQuotes & strDevConExePath & strQuotes & " status * > " & strQuotes & strHwidsTxtPath & strQuotes
    
    CreateIfNotExistPath strWorkTemp

    RunDevcon = RunAndWaitNew(cmdString, strWorkTemp, vbHide)
    If Not RunDevcon Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
    End If

    If GetFileSizeByPath(strHwidsTxtPath) Then
        PrintFileInDebugLog strHwidsTxtPath
    Else
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
        RunDevcon = False
    End If

    If mbDebugStandart Then DebugMode vbTab & "Run Devcon: " & RunDevcon
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function RunDevconRescan
'! Description (��������)  :   [����� ����� ��������� + ������ ��������� Devcon]
'! Parameters  (����������):   lngPause (Long = 1)
'!--------------------------------------------------------------------------------
Public Function RunDevconRescan(Optional ByVal lngPause As Long = 1) As Boolean

    Dim cmdString As String

    cmdString = strQuotes & strDevConExePath & strQuotes & " rescan"
    ChangeStatusBarText strMessages(96) & strSpace & cmdString
    CreateIfNotExistPath strWorkTemp

    RunDevconRescan = RunAndWaitNew(cmdString, strWorkTemp, vbHide)
    If Not RunDevconRescan Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
    End If

    If mbDebugDetail Then DebugMode vbTab & "Run RunDevconRescan: " & RunDevconRescan
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function RunDevconView
'! Description (��������)  :   [������ ��������� Devcon]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function RunDevconView() As Boolean

    Dim cmdString As String

    cmdString = strQuotes & strDevconCmdPath & strQuotes & strSpace & strQuotes & strDevConExePath & strQuotes & strSpace & strQuotes & strHwidsTxtPathView & strQuotes & " 3"

    RunDevconView = RunAndWaitNew(cmdString, strWorkTemp, vbHide)
    If Not RunDevconView Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
    End If

    If mbDebugDetail Then DebugMode vbTab & "Run DevconView: " & RunDevconView
End Function
