Attribute VB_Name = "mTimeFunction"
Option Explicit

Public dtStartTimeProg                   As Long
Public dtEndTimeProg                     As Long
Public dtAllTimeProg                     As String

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CalculateTime
'! Description (Описание)  :   [Функция расчета времени, исходя из полученных значений в миллисекундах функции GetTickCount]
'! Parameters  (Переменные):   lngStartTime (Long)
'                              lngEndTime (Long)
'                              mbmSec (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function CalculateTime(ByVal lngStartTime As Long, ByVal lngEndTime As Long, Optional ByVal mbmSec As Boolean = False) As String

    Dim lngWorkTimeTemp         As Single
    Dim lngWorkTimeSecound      As Long
    Dim lngWorkTimeMinutes      As Long
    Dim lngWorkTimeHours        As Long
    Dim lngWorkTimeMilliSecound As Long
    Dim strWorkTimeSecound      As String
    Dim strWorkTimeMinutes      As String
    Dim strWorkTimeHours        As String
    Dim strWorkTimeMilliSecound As String

    If lngEndTime > lngStartTime Then
        lngWorkTimeTemp = (lngEndTime - lngStartTime) / 1000

        'время в секундах
        'Если надо то в миллисекундах
        If mbmSec Then
            lngWorkTimeMilliSecound = (lngWorkTimeTemp - Fix(lngWorkTimeTemp)) * 1000
        Else
            lngWorkTimeTemp = Fix(lngWorkTimeTemp)
        End If

        Select Case lngWorkTimeTemp

            Case 0 To 3599
                lngWorkTimeMinutes = lngWorkTimeTemp \ 60
                lngWorkTimeSecound = lngWorkTimeTemp Mod 60

            Case 3600
                lngWorkTimeHours = 1

            Case Is > 3600
                lngWorkTimeHours = lngWorkTimeTemp \ 3600
                lngWorkTimeTemp = lngWorkTimeTemp Mod 3600
                lngWorkTimeMinutes = lngWorkTimeTemp \ 60
                lngWorkTimeSecound = lngWorkTimeTemp Mod 60

            Case Else
                lngWorkTimeHours = 0
                lngWorkTimeMinutes = 0
                lngWorkTimeSecound = 0
        End Select

    End If

    ' Добавляем лидирующие нули при необходимости
    ' Часы
    If Len(CStr(lngWorkTimeHours)) = 1 Then
        strWorkTimeHours = "0" & CStr(lngWorkTimeHours)
    ElseIf Len(CStr(lngWorkTimeHours)) = 2 Then
        strWorkTimeHours = CStr(lngWorkTimeHours)
    Else
        strWorkTimeHours = "00"
    End If

    ' Минуты
    If Len(CStr(lngWorkTimeMinutes)) = 1 Then
        strWorkTimeMinutes = "0" & CStr(lngWorkTimeMinutes)
    ElseIf Len(CStr(lngWorkTimeMinutes)) = 2 Then
        strWorkTimeMinutes = CStr(lngWorkTimeMinutes)
    Else
        strWorkTimeMinutes = "00"
    End If

    ' Секунды
    If Len(CStr(lngWorkTimeSecound)) = 1 Then
        strWorkTimeSecound = "0" & CStr(lngWorkTimeSecound)
    ElseIf Len(CStr(lngWorkTimeSecound)) = 2 Then
        strWorkTimeSecound = CStr(lngWorkTimeSecound)
    Else
        strWorkTimeSecound = "00"
    End If

    ' МилиСекунды
    If mbmSec Then
        If Len(CStr(lngWorkTimeMilliSecound)) = 1 Then
            strWorkTimeMilliSecound = "00" & CStr(lngWorkTimeMilliSecound)
        ElseIf Len(CStr(lngWorkTimeMilliSecound)) = 2 Then
            strWorkTimeMilliSecound = "0" & CStr(lngWorkTimeMilliSecound)
        ElseIf Len(CStr(lngWorkTimeMilliSecound)) = 3 Then
            strWorkTimeMilliSecound = CStr(lngWorkTimeMilliSecound)
        Else
            strWorkTimeMilliSecound = "000"
        End If

        ' Итоговое время
        CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & "." & strWorkTimeMilliSecound & " (hh:mm:ss.ms)"
    Else
        ' Итоговое время
        CalculateTime = strWorkTimeHours & ":" & strWorkTimeMinutes & ":" & strWorkTimeSecound & " (hh:mm:ss)"
    End If

End Function
