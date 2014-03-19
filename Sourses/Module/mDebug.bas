Attribute VB_Name = "mDebug"
Option Explicit

' ������ ��� ����������� ������ ���������� ��������� � ���-����
' �������� ����������� ������� ������ ����������� ���������� ���������

'==========================================================================
'------------------ ��������� ����������� ������ --------------------------'
'==========================================================================
' ��������� ����������� �� ini-����� ��� ������� ���������
Public mbDebugStandart           As Boolean   '����������� �������
Public mbDebugDetail           As Boolean   '��������� �������, ������ ���������� ���������
Public mbCleanHistory          As Boolean   '������� ������� ����������� ������
Public mbDebugTime2File        As Boolean   '���������� ����� ������� � ���-����
Public mbDebugLog2AppPath      As Boolean   '������� Logs ��������� � ����� � ����������
Public lngDetailMode           As Long      '����� ����������� ���-�����
' ��������� �������������� � ���� ������ ���������
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogName         As String
Public strDebugLogPathTemp     As String
Public strDebugLogNameTemp     As String

' ����� ��� ������ ����������� �����
Private tsDebugLogFile As TextStream

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DebugMode
'! Description (��������)  :   [������� ���������� ���������]
'! Parameters  (����������):   Msg (String)
'                              lngDetailModeTemp (Long = 1)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal Msg As String)

    ' ��������� �� ����� ���� ��� ����������� ��� ��������
    If PathExists(strDebugLogFullPath) Then
        Set tsDebugLogFile = objFSO.OpenTextFile(strDebugLogFullPath, ForAppending, False, TristateUseDefault)
    Else
        Set tsDebugLogFile = objFSO.OpenTextFile(strDebugLogFullPath, ForWriting, True, TristateUseDefault)
    End If

    If mbDebugTime2File Then
        tsDebugLogFile.WriteLine CStr(Now()) & vbTab & Msg
    Else
        tsDebugLogFile.WriteLine Msg
    End If

    tsDebugLogFile.Close

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function LogNotOnCDRoom
'! Description (��������)  :   [�������� �� �������� ���-����� �� CD]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom(Optional ByVal strLogFolder As String) As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    If LenB(strLogFolder) = 0 Then
        strDriveName = Left$(strDebugLogPath, 2)
    Else
        strDriveName = Left$(strLogFolder, 2)
    End If
    
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

    If mbCleanHistory Then
        If PathExists(strDebugLogFullPath) Then
            If Not LogNotOnCDRoom Then
                DeleteFiles (strDebugLogFullPath)
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

    If PathExists(strFilePath) Then
        If Not PathIsAFolder(strFilePath) Then
            If GetFileSizeByPath(strFilePath) > 0 Then
            
                Dim objTxtFile    As TextStream
                Dim strTxtFileAll As String
                
                Set objTxtFile = objFSO.OpenTextFile(strFilePath, ForReading, False, TristateUseDefault)
                strTxtFileAll = objTxtFile.ReadAll
                objTxtFile.Close
                If mbDebugStandart Then DebugMode vbTab & "Content of file: " & strFilePath & vbNewLine & "*********************BEGIN FILE**************************" & vbNewLine & strTxtFileAll & vbNewLine & "**********************END FILE***************************"
            Else
                If mbDebugStandart Then DebugMode vbTab & "Content of file: " & strFilePath & " Error - 0 bytes"
            End If
        End If
    End If

End Sub
