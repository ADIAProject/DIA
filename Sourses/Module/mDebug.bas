Attribute VB_Name = "mDebug"
Option Explicit

' ������ ��� ����������� ������ ���������� ��������� � ���-����
' �������� ����������� ������� ������ ����������� ���������� ���������

'==========================================================================
'------------------ ��������� ����������� ������ --------------------------'
'==========================================================================
' ��������� ����������� �� ini-����� ��� ������� ���������
Public mbDebugStandart         As Boolean   '����������� �������
Public mbDebugDetail           As Boolean   '��������� �������, ������ ���������� ���������
Public mbCleanHistory          As Boolean   '������� ������� ����������� ������
Public mbDebugTime2File        As Boolean   '���������� ����� ������� � ���-����
Public mbDebugLog2AppPath      As Boolean   '������� Logs ��������� � ����� � ����������
Public lngDetailMode           As Long      '����� ����������� ���-�����
Public strDebugLogPathTemp     As String    '���������� �������� ���-����� (���� ����� ���� ������������� � � environment-�����������)
Public strDebugLogNameTemp     As String    '��� ���-����� (�������������� ����������)
' ��������� �������������� � ���� ������ ���������
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogName         As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DebugMode
'! Description (��������)  :   [������� ���������� ���������]
'! Parameters  (����������):   Msg (String)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal strMsg As String)
    
    Dim mbFileExist As Boolean
    Dim fNum        As Integer
    
    mbFileExist = FileExists(strDebugLogFullPath)
    
    fNum = FreeFile
    Open strDebugLogFullPath For Binary Access Write As fNum
    
    If Not mbDebugTime2File Then
        ' ��������� �� ����� ���� ��� ����������� ��� ��������
        If mbFileExist Then
            Put #fNum, LOF(fNum), strMsg & vbNewLine
        Else
            Put #fNum, , strMsg & vbNewLine
        End If
    Else
        ' ��������� �� ����� ���� ��� ����������� ��� ��������
        If mbFileExist Then
            Put #fNum, LOF(fNum), (vbNewLine & CStr(Now()) & vbTab) & strMsg
        Else
            Put #fNum, , (vbNewLine & CStr(Now()) & vbTab) & strMsg
        End If

    End If
    
    Close #fNum

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function LogNotOnCDRoom
'! Description (��������)  :   [�������� �� �������� ���-����� �� CD]
'! Parameters  (����������):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom(Optional ByVal strLogFolder As String) As Boolean

    Dim strDriveName As String
    Dim xDrv         As Long

    If LenB(strLogFolder) = 0 Then
        strDriveName = Left$(strDebugLogPath, 3)
    Else
        strDriveName = Left$(strLogFolder, 3)
    End If
    
    ' ��������� �� ������ �� ����
    If InStr(strDriveName, vbBackslashDouble) = 0 Then
        '�������� ��� �����
        If PathIsRoot(strDriveName) Then
            xDrv = GetDriveType(strDriveName)
    
            If xDrv = DRIVE_CDROM Then
                LogNotOnCDRoom = True
            End If
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
        If FileExists(strDebugLogFullPath) Then
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
    Dim strTxtFileAll As String
    
    If FileExists(strFilePath) Then
        If GetFileSizeByPath(strFilePath) Then
                    
            If mbDebugStandart Then
                FileReadData strFilePath, strTxtFileAll
                DebugMode vbTab & "Content of file: " & strFilePath & vbNewLine & "*********************BEGIN FILE**************************" & vbNewLine & strTxtFileAll & vbNewLine & "**********************END FILE***************************"
            End If
        Else
            If mbDebugStandart Then DebugMode vbTab & "Content of file: " & strFilePath & " Error - 0 bytes"
        End If
    End If

End Sub
