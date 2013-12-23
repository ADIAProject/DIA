Attribute VB_Name = "mDebug"
Option Explicit

'==========================================================================
'------------------ ��������� ����������� ������ --------------------------'
'==========================================================================
Public mbDebugEnable           As Boolean
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogPathTemp     As String
Public strDebugLogName         As String
Public strDebugLogNameTemp     As String
Public strDebugLogPath2AppPath As String
Public mbCleanHistory          As Boolean     '������� ������� ����������� ������
Public lngDetailMode           As Long
Public mbDebugLog2AppPath      As Boolean
Public mbDebugTime2File        As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DebugMode
'! Description (��������)  :   [������� ���������� ���������]
'! Parameters  (����������):   Msg (String)
'                              lngDetailModeTemp (Long = 1)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal Msg As String, Optional ByVal lngDetailModeTemp As Long = 1)

    Dim tsLogFile As TextStream

    ' ��������� �� ����� ���� ��� ����������� ��� ��������
    If mbDebugEnable Then
        If Not mbLogNotOnCDRoom Then
            If lngDetailModeTemp <= lngDetailMode Then
                If LenB(Msg) > 0 Then
                    If objFSO.FileExists(strDebugLogFullPath) Then
                        Set tsLogFile = objFSO.OpenTextFile(strDebugLogFullPath, ForAppending, False, TristateTrue)
                    Else
                        Set tsLogFile = objFSO.OpenTextFile(strDebugLogFullPath, ForWriting, True, TristateTrue)
                    End If

                    If mbDebugTime2File Then
                        tsLogFile.WriteLine CStr(Now()) & vbTab & Msg
                    Else
                        tsLogFile.WriteLine Msg
                    End If

                    tsLogFile.Close
                End If
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function LogNotOnCDRoom
'! Description (��������)  :   [�������� �� �������� ���-����� �� CD]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    LogNotOnCDRoom = False
    strDriveName = Left$(strDebugLogPath, 2)

    ' ��������� �� ������ �� ����
    If InStr(strDriveName, vbBackslash) = 0 Then
        '�������� ��� �����
        Set xDrv = objFSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            LogNotOnCDRoom = True
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub MakeCleanHistory
'! Description (��������)  :   [�������� ������� ����������� ������]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Sub MakeCleanHistory()

    Dim FileDel As File

    If mbCleanHistory Then
        If objFSO.FileExists(strDebugLogFullPath) Then
            If Not mbLogNotOnCDRoom Then
                Set FileDel = objFSO.GetFile(strDebugLogFullPath)
                FileDel.Delete
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub PrintFileInDebugLog
'! Description (��������)  :   [������ � DebugLog ����������� �����]
'! Parameters  (����������):   strFilePath (String)
'!--------------------------------------------------------------------------------
Public Sub PrintFileInDebugLog(ByVal strFilePath As String)

    Dim objTxtFile    As TextStream
    Dim strTxtFileAll As String

    If PathExists(strFilePath) Then
        If Not PathIsAFolder(strFilePath) Then
            If GetFileSizeByPath(strFilePath) > 0 Then
                Set objTxtFile = objFSO.OpenTextFile(strFilePath, ForReading, False, TristateUseDefault)
                strTxtFileAll = objTxtFile.ReadAll
                objTxtFile.Close
                DebugMode vbTab & "Content of file: " & strFilePath & vbNewLine & "*********************BEGIN FILE**************************" & vbNewLine & strTxtFileAll & vbNewLine & "**********************END FILE***************************"
            Else
                DebugMode vbTab & "Content of file: " & strFilePath & " Error - 0 bytes"
            End If
        End If
    End If

End Sub
