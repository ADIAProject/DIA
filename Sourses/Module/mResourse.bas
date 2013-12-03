Attribute VB_Name = "mResource"
Option Explicit

'! -----------------------------------------------------------
'!  Функция     :  GetBinaryFileFromResource
'!  Переменные  :  File_Path As String, ByVal ID As String, Resource As String
'!  Описание    :  Извлечение бинарного файла из ресурса
'! -----------------------------------------------------------
Public Function GetBinaryFileFromResource(ByVal File_Path As String, _
                                          ByVal ID As String, _
                                          ByVal Resource As String) As Boolean

Dim iFile                               As Long
Dim BinaryData()                        As Byte

    iFile = FreeFile
    GetBinaryFileFromResource = False

    'загрузка из ресурсов
    On Error GoTo HandErr

    BinaryData = LoadResData(ID, Resource)

    If LenB(BinaryData(1)) > 0 Then
        'Если что - то есть, то все гуд
        Open File_Path For Binary Access Write As #iFile
        'запись в файл
        Put #iFile, 1, BinaryData
        Close #iFile
        'операция успешна
        GetBinaryFileFromResource = True

    End If

ExitFromSub:
    Exit Function
HandErr:

    If err.Number = 326 Then
        If MsgBox("Error №: " & err.Number & vbNewLine & "Description: " & err.Description & str2vbNewLine & "This Error in Function 'GetBinaryFileFromResource'." & vbNewLine & "Executable file is corrupted, or required library removed from the resources of program." & str2vbNewLine & "Download the latest re-distribution program!!!" & vbNewLine & "If the error persists, please report it to the developer." & str2vbNewLine & "Normal work of program is not guaranteed, you want to continue?", vbCritical + vbYesNo, strProductName) = vbNo Then
            End

        End If

    ElseIf err.Number <> 0 Then
        GoTo ExitFromSub

    End If

End Function

Public Function ExtractResource(ByVal strOCXFileName As String, _
                                ByVal strPathOcx As String) As Boolean

Dim strCopyOcxFileTo                    As String

    strCopyOcxFileTo = BackslashAdd2Path(strPathOcx) & strOCXFileName

    ' Извлекаем ресурс в файл
    If GetBinaryFileFromResource(strCopyOcxFileTo, "OCX_" & FileName_woExt(strOCXFileName), "CUSTOM") Then
        DebugMode str2VbTab & strOCXFileName & ": BinaryFileFromResourse: True"
        ExtractResource = True
    Else
        DebugMode str2VbTab & strOCXFileName & ": BinaryFileFromResourse: False"
    End If

End Function

Public Function ExtractResourceAll(ByVal strPathOcxTo As String) As Boolean
    DebugMode "ExtractResourceAll - Start"
    ExtractResourceAll = True

    If ExtractResource("MSFLXGRD.OCX", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'MSFLXGRD.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If

        ExtractResourceAll = False
    End If

    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"

    If ExtractResource("RICHTX32.OCX", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'RICHTX32.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If

        ExtractResourceAll = False
    End If

    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"

    If ExtractResource("TABCTL32.OCX", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'TABCTL32.OCX' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If

        ExtractResourceAll = False
    End If

    DebugMode vbTab & "ExtractResourceAll - *****************Check Next File********************"

    If ExtractResource("vbscript.dll", strPathOcxTo) = False Then
        If MsgBox("Extract OCX or DLL: 'vbscript.dll' - False" & str2vbNewLine & strMessages(134), vbYesNo + vbQuestion, strProductName) = vbNo Then
            End
        End If

        ExtractResourceAll = False
    End If

    '    DebugMode VbTab & "ExtractResourceAll - *****************Check Next File********************"
    '
    '    If ExtractResource("capicom.dll", strPathOcxTo) = False Then
    '        If MsgBox("Extract OCX or DLL: capicom.dll' - False" & str2vbNewLine & strMessages(20), vbYesNo + vbQuestion, strProductName) = vbNo Then
    '            End
    '
    '        End If
    '
    '        ExtractResourceAll = False
    '
    '    End If

    DebugMode "ExtractResourceAll - End"

End Function
