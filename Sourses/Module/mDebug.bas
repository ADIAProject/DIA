Attribute VB_Name = "mDebug"
Option Explicit

' Модуль для организации записи отладочных сообщений в лог-файл
' Имееется возможность задания режима детализации отладочных сообщений

'==========================================================================
'------------------ Параметры отладочного режима --------------------------'
'==========================================================================
' Параметры считываются из ini-файла при запуске программы
Public mbDebugStandart         As Boolean   'Стандартная отладка
Public mbDebugDetail           As Boolean   'Детальная отладка, больше отладочных сообщений
Public mbCleanHistory          As Boolean   'Очистка истории отладочного режима
Public mbDebugTime2File        As Boolean   'Записывать время события в лог-файл
Public mbDebugLog2AppPath      As Boolean   'Каталог Logs находится в папке с программой
Public lngDetailMode           As Long      'Режим детализации лог-файла
Public strDebugLogPathTemp     As String    'Директория создания лог-файла (путь может быть относительный и с environment-переменными)
Public strDebugLogNameTemp     As String    'Имя лог-файла (поддерживаются переменные)
' Параметры рассчитываемые в ходе работы программы
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogName         As String

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugMode
'! Description (Описание)  :   [Функция отладочных сообщений]
'! Parameters  (Переменные):   Msg (String)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal strMsg As String)
    
    Dim mbFileExist As Boolean
    Dim fNum        As Integer
    
    mbFileExist = FileExists(strDebugLogFullPath)
    
    fNum = FreeFile
    Open strDebugLogFullPath For Binary Access Write As fNum
    
    If Not mbDebugTime2File Then
        ' создается ли новый файл или открывается для дозаписи
        If mbFileExist Then
            Put #fNum, LOF(fNum), strMsg & vbNewLine
        Else
            Put #fNum, , strMsg & vbNewLine
        End If
    Else
        ' создается ли новый файл или открывается для дозаписи
        If mbFileExist Then
            Put #fNum, LOF(fNum), (vbNewLine & CStr(Now()) & vbTab) & strMsg
        Else
            Put #fNum, , (vbNewLine & CStr(Now()) & vbTab) & strMsg
        End If

    End If
    
    Close #fNum

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LogNotOnCDRoom
'! Description (Описание)  :   [Проверка на хранение лог-файла на CD]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom(Optional ByVal strLogFolder As String) As Boolean

    Dim strDriveName As String
    Dim xDrv         As Long

    If LenB(strLogFolder) = 0 Then
        strDriveName = Left$(strDebugLogPath, 3)
    Else
        strDriveName = Left$(strLogFolder, 3)
    End If
    
    ' Проверяем на запуск из сети
    If InStr(strDriveName, vbBackslashDouble) = 0 Then
        'получаем тип диска
        If PathIsRoot(strDriveName) Then
            xDrv = GetDriveType(strDriveName)
    
            If xDrv = DRIVE_CDROM Then
                LogNotOnCDRoom = True
            End If
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
        If FileExists(strDebugLogFullPath) Then
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
