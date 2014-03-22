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
' ��������� �������������� � ���� ������ ���������
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogName         As String
Public strDebugLogPathTemp     As String
Public strDebugLogNameTemp     As String

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Sub DebugMode
'! Description (��������)  :   [������� ���������� ���������]
'! Parameters  (����������):   Msg (String)
'                              lngDetailModeTemp (Long = 1)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal Msg As String)
    
    Dim mbFileExist As Boolean
    Dim fNum As Integer
    
    mbFileExist = PathExists(strDebugLogFullPath)
    
    fNum = FreeFile
    Open strDebugLogFullPath For Binary Access Write As fNum
    
    If Not mbDebugTime2File Then
        ' ��������� �� ����� ���� ��� ����������� ��� ��������
        If mbFileExist Then
            Put #fNum, LOF(fNum), Msg & vbNewLine
        Else
            Put #fNum, , Msg & vbNewLine
        End If
    Else
        ' ��������� �� ����� ���� ��� ����������� ��� ��������
        If mbFileExist Then
            Put #fNum, LOF(fNum), (vbNewLine & CStr(Now()) & vbTab) & Msg
        Else
            Put #fNum, , (vbNewLine & CStr(Now()) & vbTab) & Msg
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
    Dim strTxtFileAll As String
    
    If PathExists(strFilePath) Then
        If Not PathIsAFolder(strFilePath) Then
            If GetFileSizeByPath(strFilePath) Then
                        
                If mbDebugStandart Then
                    strTxtFileAll = FileReadData(strFilePath)
                    DebugMode vbTab & "Content of file: " & strFilePath & vbNewLine & "*********************BEGIN FILE**************************" & vbNewLine & strTxtFileAll & vbNewLine & "**********************END FILE***************************"
                End If
            Else
                If mbDebugStandart Then DebugMode vbTab & "Content of file: " & strFilePath & " Error - 0 bytes"
            End If
        End If
    End If

End Sub
