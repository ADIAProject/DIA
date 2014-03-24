Attribute VB_Name = "mTimeFunction"
Option Explicit

Public dtStartTimeProg                   As Long
Public dtEndTimeProg                     As Long
Public dtAllTimeProg                     As String

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'!--------------------------------------------------------------------------------
'! Procedure   (�������)   :   Function CalculateTime
'! Description (��������)  :   [������� ������� �������, ������ �� ���������� �������� � ������������� ������� GetTickCount]
'! Parameters  (����������):   lngStartTime (Long)
'                              lngEndTime (Long)
'                              mbmSec (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function CalculateTime(ByVal lngStartTime As Long, ByVal lngEndTime As Long, Optional ByVal mbmSec As Boolean = False) As String

    Dim lngWorkTimeTemp         As Double
    Dim lngWorkTimeSecound      As Long
    Dim lngWorkTimeMinutes      As Long
    Dim lngWorkTimeHours        As Long
    Dim lngWorkTimeMilliSecound As Long
    Dim strWorkTimeSecound      As String
    Dim strWorkTimeMinutes      As String
    Dim strWorkTimeHours        As String
    Dim strWorkTimeMilliSecound As String

    If lngEndTime > lngStartTime Then
        ' ��������� �������� � �������
        lngWorkTimeTemp = (lngEndTime - lngStartTime) / 1000
        ' ����������� ��������
        lngWorkTimeHours = (lngWorkTimeTemp \ 3600)
        lngWorkTimeMinutes = (lngWorkTimeTemp \ 60) Mod 60
        lngWorkTimeSecound = lngWorkTimeTemp Mod 60
        lngWorkTimeMilliSecound = (lngWorkTimeTemp - Fix(lngWorkTimeTemp)) * 1000

        ' ��������� ���������� ���� ��� �������������
        strWorkTimeHours = Format$(lngWorkTimeHours, "00")
        strWorkTimeMinutes = Format$(lngWorkTimeMinutes, "00")
        strWorkTimeSecound = Format$(lngWorkTimeSecound, "00")
        strWorkTimeMilliSecound = Format$(lngWorkTimeMilliSecound, "000")
    
        ' ���� ��������� ����� � �������������
        If mbmSec Then
            ' �������� �����
            If lngWorkTimeHours = 0 Then
                CalculateTime = strWorkTimeMinutes & ":" & strWorkTimeSecound & "." & strWorkTimeMilliSecound & " (mm:ss.ms)"
            Else
                CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & "." & strWorkTimeMilliSecound & " (hh:mm:ss.ms)"
            End If
        Else
            ' �������� �����
            If lngWorkTimeHours = 0 Then
                CalculateTime = strWorkTimeMinutes & ":" & strWorkTimeSecound & " (mm:ss)"
            Else
                CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & " (hh:mm:ss)"
            End If
        End If
    Else
        ' �������� �����
        CalculateTime = "00:00:00.000 (hh:mm:ss.ms)"
    End If

End Function
