Attribute VB_Name = "mDebug"
Option Explicit

'==========================================================================
'------------------ Параметры отладочного режима --------------------------'
'==========================================================================
Public mbDebugEnable           As Boolean
Public strDebugLogFullPath     As String
Public strDebugLogPath         As String
Public strDebugLogPathTemp     As String
Public strDebugLogName         As String
Public strDebugLogNameTemp     As String
Public strDebugLogPath2AppPath As String
Public mbCleanHistory          As Boolean     'Очистка истории отладочного режима
Public lngDetailMode           As Long
Public mbDebugLog2AppPath      As Boolean
Public mbDebugTime2File        As Boolean

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DebugMode
'! Description (Описание)  :   [Функция отладочных сообщений]
'! Parameters  (Переменные):   Msg (String)
'                              lngDetailModeTemp (Long = 1)
'!--------------------------------------------------------------------------------
Public Sub DebugMode(ByVal Msg As String, Optional ByVal lngDetailModeTemp As Long = 1)

    Dim tsLogFile As TextStream

    ' создается ли новый файл или открывается для дозаписи
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
'! Procedure   (Функция)   :   Function LogNotOnCDRoom
'! Description (Описание)  :   [Проверка на хранение лог-файла на CD]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function LogNotOnCDRoom() As Boolean

    Dim strDriveName As String
    Dim xDrv         As Drive

    LogNotOnCDRoom = False
    strDriveName = Left$(strDebugLogPath, 2)

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
'! Procedure   (Функция)   :   Sub PrintFileInDebugLog
'! Description (Описание)  :   [Запись в DebugLog сожержимого файла]
'! Parameters  (Переменные):   strFilePath (String)
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
