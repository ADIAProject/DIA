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
Public Function CompareDevDBVersion(strDevDBFullFileName As String, Optional ByVal strPathDRP As String) As Boolean

    Dim LngValue          As Long
    Dim strFilePath_woExt As String

    strFilePath_woExt = GetFileName_woExt(strDevDBFullFileName)
    LngValue = IniLongPrivate(GetFileNameFromPath(strFilePath_woExt), "Version", BackslashAdd2Path(GetPathNameFromPath(strFilePath_woExt)) & "DevDBVersions.ini")

    If LngValue = 9999 Then
        CompareDevDBVersion = False
    Else
        CompareDevDBVersion = Not (LngValue <> lngDevDBVersion)
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

    If mbIsWin64 Then
        strDevConTemp = strDevConExePath64
    Else

        If StrComp(strOSCurrentVersion, "5.0") = 0 Then
            strDevConTemp = strDevConExePathW2k
        Else
            strDevConTemp = strDevConExePath
        End If
    End If

    cmdString = strKavichki & strDevconCmdPath & strKavichki & strSpace & strKavichki & strDevConTemp & strKavichki & strSpace & strKavichki & strHwidsTxtPath & strKavichki & " 4 " & strKavichki & strHwid & strKavichki

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

    Dim str           As String
    Dim i             As Long
    Dim strCnt        As Long
    Dim miStatus      As Long
    Dim strID         As String
    Dim strIDOrig     As String
    Dim strIDCutting  As String
    Dim strName       As String
    Dim strName_x()   As String
    Dim miMaxCountArr As Long
    Dim RecCountArr   As Long
    Dim strID_x()     As String
    Dim RegExpDevcon  As RegExp
    Dim MatchesDevcon As MatchCollection
    Dim objMatch      As Match
    Dim strStatus     As String

    Set RegExpDevcon = New RegExp

    With RegExpDevcon
        .Pattern = "(^[^\n\r\s][^\n\r]+)\r\n(\s+[^\n\r]+\r\n)*[^\n\r]*((?:DEVICE IS|DEVICE HAS|DRIVER IS|DRIVER HAS)[^\r]+)"
        .MultiLine = True
        .IgnoreCase = True
        .Global = True
    End With

    If mbDebugDetail Then DebugMode "DevParserLocalHwids2-Start"

    If PathExists(strHwidsTxtPath) Then
        If Not PathIsAFolder(strHwidsTxtPath) Then
            str = FileReadData(strHwidsTxtPath)
            Set MatchesDevcon = RegExpDevcon.Execute(str)
            miMaxCountArr = 100

            ' ������������ ���-�� ��������� � �������
            ReDim arrHwidsLocal(miMaxCountArr)

            strCnt = MatchesDevcon.Count
            RecCountArr = 0

            'For i = 0 To MatchesDevcon.Count - 1
            For i = 0 To strCnt - 1
                Set objMatch = MatchesDevcon.item(i)

                ' ���� ������� � ������� ���������� ������ ��� ���������, �� ����������� ����������� �������
                If RecCountArr = miMaxCountArr Then
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
                strID = UCase$(Trim$(Replace$(strID, vbNewLine, vbNullString)))
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
                        If InStr(1, strStatus, "running", vbTextCompare) Then
                            miStatus = 1
                        End If

                        If LenB(strName) Then
                            strName_x = Split(strName, vbNewLine)
                            strName = strName_x(0)
                        End If

                        strName = Replace$(strName, vbNewLine, vbNullString)
                        strName = Replace$(strName, "Name:", vbNullString, , , vbTextCompare)
                        strName = Trim$(strName)

                        If Len(strID) > 3 Then
                            arrHwidsLocal(RecCountArr).HWID = strID
                            arrHwidsLocal(RecCountArr).DevName = strName
                            arrHwidsLocal(RecCountArr).Status = miStatus
                            arrHwidsLocal(RecCountArr).HWIDOrig = strIDOrig
                            arrHwidsLocal(RecCountArr).HWIDCutting = strIDCutting
                            RecCountArr = RecCountArr + 1
                        End If
                    End If

                Else
                    miStatus = 0

                    ' ���������� �������
                    If InStr(1, strStatus, "running", vbTextCompare) Then
                        miStatus = 1
                    End If

                    If LenB(strName) Then
                        strName_x = Split(strName, vbNewLine)
                        strName = strName_x(0)
                    End If

                    strName = Replace$(strName, vbNewLine, vbNullString)
                    strName = Replace$(strName, "Name:", vbNullString, , , vbTextCompare)
                    strName = Trim$(strName)

                    If Len(strID) > 3 Then
                        ' ID ������������
                        arrHwidsLocal(RecCountArr).HWID = strID
                        ' ������������ ������������
                        arrHwidsLocal(RecCountArr).DevName = strName
                        ' ������ ������������
                        arrHwidsLocal(RecCountArr).Status = miStatus
                        arrHwidsLocal(RecCountArr).HWIDOrig = UCase$(strIDOrig)
                        arrHwidsLocal(RecCountArr).HWIDCutting = UCase$(strIDCutting)
                        RecCountArr = RecCountArr + 1
                    End If
                End If

            Next

            ' ������������� ������ �� �������� ���-�� �������
            If RecCountArr Then

                ReDim Preserve arrHwidsLocal(RecCountArr - 1)

            Else

                ReDim Preserve arrHwidsLocal(0)

            End If

            If SaveHwidsArray2File(strResultHwidsTxtPath, arrHwidsLocal) = False Then
                MsgBox strMessages(45) & vbNewLine & strResultHwidsTxtPath, vbCritical + vbInformation, strProductName
            End If
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

    If PathExists(strHwidsTxtPath) Then
        DeleteFiles strHwidsTxtPath
    End If

    cmdString = "cmd.exe /c " & strKavichki & strKavichki & strDevConExePath & strKavichki & " status * > " & strKavichki & strHwidsTxtPath & strKavichki
    
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

    cmdString = strKavichki & strDevConExePath & strKavichki & " rescan"
    ChangeStatusTextAndDebug strMessages(96) & strSpace & cmdString
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

    cmdString = strKavichki & strDevconCmdPath & strKavichki & strSpace & strKavichki & strDevConExePath & strKavichki & strSpace & strKavichki & strHwidsTxtPathView & strKavichki & " 3"

    RunDevconView = RunAndWaitNew(cmdString, strWorkTemp, vbHide)
    If Not RunDevconView Then
        MsgBox strMessages(33) & str2vbNewLine & cmdString, vbInformation, strProductName
    End If

    If mbDebugDetail Then DebugMode vbTab & "Run DevconView: " & RunDevconView
End Function

