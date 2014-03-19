Attribute VB_Name = "mDebug"
Option Explicit

' Модуль для организации записи отладочных сообщений в лог-файл
' Имееется возможность задания режима детализации отладочных сообщений

'==========================================================================
'------------------ Параметры отладочного режима --------------------------'
'==========================================================================
' Параметры считываются из ini-файла при запуске программы
Public mbDebugStandart           As Boolean   'Стандартная отладка
Public mbDebugDetail           As Boolean   'Детальная отладка, больше отладочных сообщений
Public mbCleanHistory          As Boolean   'Очистка истории отладочного режима
Public mbDebugTime2File        As Boolean   'Записывать время события в лог-файл
Public mbDebugLog2AppPath      As Boolean   'Каталог Logs находится в папке с программой
Public lngDetailMode           As Long      'Режим детализации лог-файла
' Параметры рассчитываемые в ходе работы программы
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogName         As String
Public strDebugLogPathTemp     As String
Public strDebugLogNameTemp     As String

' Поток для вывода отладочного файла
Private tsDebugLogFile As TextStream

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugMode
'! Description (Описание)  :   [Функция отладочных сообщений]
'! Parameters  (Переменные):   Msg (String)
'                              lngDetailModeTemp (Long = 1)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal Msg As String)

    ' создается ли новый файл или открывается для дозаписи
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
'! Procedure   (Функция)   :   Function LogNotOnCDRoom
'! Description (Описание)  :   [Проверка на хранение лог-файла на CD]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom(Optional ByVal strLogFolder As String) As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    If LenB(strLogFolder) = 0 Then
        strDriveName = Left$(strDebugLogPath, 2)
    Else
        strDriveName = Left$(strLogFolder, 2)
    End If
    
    ' Проверяем на запуск из сети
    If InStr(strDriveName, vbBackslash) = 0 Then
        'получаем тип диска
        Set xDrv = objFSO.GetDrive(strDriveName)

        If xDrv.DriveType = CDRom Then
            LogNotOnCDRoom = True
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub MakeCleanHistory
'! Description (Описание)  :   [Удаление истории отладочного режима]
'! Parameters  (Переменные):
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
'! Procedure   (Функция)   :   Sub PrintFileInDebugLog
'! Description (Описание)  :   [Запись в DebugLog сожержимого файла]
'! Parameters  (Переменные):   strFilePath (String)
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
