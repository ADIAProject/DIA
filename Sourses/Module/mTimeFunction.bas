Attribute VB_Name = "mTimeFunction"
Option Explicit

Public dtStartTimeProg                   As Long
Public dtEndTimeProg                     As Long
Public dtAllTimeProg                     As String

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long
Public Declare Function QueryPerformanceFrequency Lib "kernel32" (lpFrequency As Any) As Long
Public Declare Function QueryPerformanceCounter Lib "kernel32" (lpPerformanceCount As Any) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CalculateTime
'! Description (Описание)  :   [Функция расчета времени, исходя из полученных значений в миллисекундах функции GetTickCount]
'! Parameters  (Переменные):   lngStartTime (Long)
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
        ' Переводим значение в секунды
        lngWorkTimeTemp = (lngEndTime - lngStartTime) / 1000
        ' Высчитываем значения
        lngWorkTimeHours = (lngWorkTimeTemp \ 3600)
        lngWorkTimeMinutes = (lngWorkTimeTemp \ 60) Mod 60
        lngWorkTimeSecound = lngWorkTimeTemp Mod 60
        lngWorkTimeMilliSecound = (lngWorkTimeTemp - Fix(lngWorkTimeTemp)) * 1000

        ' Добавляем лидирующие нули при необходимости
        strWorkTimeHours = Format$(lngWorkTimeHours, "00")
        strWorkTimeMinutes = Format$(lngWorkTimeMinutes, "00")
        strWorkTimeSecound = Format$(lngWorkTimeSecound, "00")
        strWorkTimeMilliSecound = Format$(lngWorkTimeMilliSecound, "000")
    
        ' Если результат нужен с милиСекундами
        If mbmSec Then
            ' Итоговое время
            If lngWorkTimeHours = 0 Then
                CalculateTime = strWorkTimeMinutes & strColon & strWorkTimeSecound & strDot & strWorkTimeMilliSecound & " (mm:ss.ms)"
            Else
                CalculateTime = strWorkTimeHours & strColon & strWorkTimeMinutes & strColon & strWorkTimeSecound & strDot & strWorkTimeMilliSecound & " (hh:mm:ss.ms)"
            End If
        Else
            ' Итоговое время
            If lngWorkTimeHours = 0 Then
                CalculateTime = strWorkTimeMinutes & strColon & strWorkTimeSecound & " (mm:ss)"
            Else
                CalculateTime = strWorkTimeHours & strColon & strWorkTimeMinutes & strColon & strWorkTimeSecound & " (hh:mm:ss)"
            End If
        End If
    Else
        ' Итоговое время
        CalculateTime = "00:00:00.000 (hh:mm:ss.ms)"
    End If

End Function

 '**************************************
' Name: Calculate timing23-May-2012
' Description:Basically, this is just one more way of calculating and displaying how much time a process took before finishing.
' This will check for a hi-performance timer and if none is found then uses the API GetTickCount.
' By: Kenaso
' http://www.Planet-Source-Code.com/vb/scripts/ShowCode.asp?txtCodeId=74366&lngWId=1
 
' ' Retrieves the frequency of the high-resolution performance counter,
' ' if one exists. The frequency cannot change while the system is running.
' ' If the function fails, the return value is zero.
'Private Declare Function QueryPerformanceFrequency Lib "kernel32" (curFrequency As Currency) As Long
' ' The QueryPerformanceCounter function retrieves the current value of the  ' high-resolution performance counter.
'Private Declare Function QueryPerformanceCounter Lib "kernel32" (curCounter As Currency) As Boolean

' ' This is a rough translation of the GetTickCount API. The
' ' tick count of a PC is only valid for the first 49.7 days
' ' since the last reboot. When you capture the tick count,
' ' you are capturing the total number of milliseconds elapsed
' ' since the last reboot. The elapsed time is stored as a
' ' DWORD value. Therefore, the time will wrap around to zero
' ' if the system is run continuously for 49.7 days.
'
'Private mcurFrequency As Currency ' High performance Frequency
'
'Public Function GetTimeStart() As Long
'    If mcurFrequency > 0 Then
'        QueryPerformanceCounter curEnd ' Hi-performance timer available
'    Else
'        curEnd = CDbl(GetTickCount) ' No hi-performance timer
'    End If
'
'    GetTimeStart = curEnd
'End Function
