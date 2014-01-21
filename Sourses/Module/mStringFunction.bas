Attribute VB_Name = "mStringFunction"
Option Explicit

Public Const str2vbNullChar = vbNullChar & vbNullChar
Public Const str2vbNewLine = vbNewLine & vbNewLine
Public Const str2VbTab = vbTab & vbTab
Public Const str3VbTab = vbTab & vbTab & vbTab
Public Const str4VbTab = vbTab & vbTab & vbTab & vbTab
Public Const str5VbTab = vbTab & vbTab & vbTab & vbTab & vbTab
Public Const str6VbTab = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
Public Const str7VbTab = vbTab & vbTab & vbTab & vbTab & vbTab & vbTab & vbTab
Public Const Percentage = "%"
Public Const Kavichki = """" 'ChrW$(34)

' Not use in project
' Сравнение строк с учетм регистра и без
'Private Declare Function StrCmpI Lib "shlwapi.dll" Alias "StrCmpIW" (ByVal ptr1 As Long, ByVal ptr2 As Long) As Long
'Private Declare Function StrCmp Lib "shlwapi.dll" Alias "StrCmpW" (ByVal ptr1 As Long, ByVal ptr2 As Long) As Long
' конвертация строк с учетом регистра
'Private Declare Function CharLower Lib "user32.dll" Alias "CharLowerA" (ByVal lpsz As String) As String
'Private Declare Function CharUpper Lib "user32.dll" Alias "CharUpperA" (ByVal lpsz As String) As String
Public Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
Public Declare Function lstrcat Lib "kernel32.dll" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Public Declare Function ArrPtr Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function AppendStr
'! Description (Описание)  :   [Добавляет строку к строке с нужным разделителем]
'! Parameters  (Переменные):   strHead (String)
'                              strAdd (String)
'                              strSep (String = " ")
'!--------------------------------------------------------------------------------
Public Function AppendStr(ByVal strHead As String, ByVal strAdd As String, Optional ByVal strSep As String = " ") As String

    If LenB(strAdd) > 0 Then
        If LenB(strHead) > 0 Then
            AppendStr = strHead & (strSep & strAdd)
        Else
            AppendStr = strAdd
        End If

    Else
        AppendStr = strHead & strSep
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompareByVersion
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strVersionBD (String)
'                              strVersionLocal (String)
'!--------------------------------------------------------------------------------
Public Function CompareByVersion(ByVal strVersionBD As String, ByVal strVersionLocal As String) As String

    Dim strDevVer_x()       As String
    Dim strDevVerLocal_x()  As String
    Dim strDevVer_xx        As String
    Dim strDevVerLocal_xx   As String
    Dim miDimension         As Integer
    Dim miDimensionLocal    As Integer
    Dim strVersionBD_x()    As String
    Dim strVersionLocal_x() As String
    Dim i                   As Integer
    Dim ResultTemp          As String

    DebugMode str6VbTab & "CompareByVersion-Start" & vbNewLine & _
              str7VbTab & "CompareByVersion-Start: " & strVersionBD & " compare with " & strVersionLocal
    ResultTemp = "?"
    strDevVer_x = Split(Trim$(strVersionBD), ",")
    miDimension = UBound(strDevVer_x)

    If InStr(strVersionBD, "unknown") = 0 Then
        If InStr(strVersionLocal, "unknown") = 0 Then
            If miDimension > 0 Then
                strDevVer_xx = Trim$(strDevVer_x(1))
            Else
                ResultTemp = "<"
                strDevVer_xx = strVersionBD
            End If

            strDevVerLocal_x = Split(Trim$(strVersionLocal), ",")
            miDimensionLocal = UBound(strDevVerLocal_x)

            If miDimensionLocal > 0 Then
                strDevVerLocal_xx = Trim$(strDevVerLocal_x(1))
            Else
                ResultTemp = ">"
                strDevVerLocal_xx = strVersionLocal
            End If

            If Right$(strDevVer_xx, 1) = "." Then
                strDevVer_xx = Left$(strDevVer_xx, Len(strDevVer_xx) - 1)
            End If

            If Right$(strDevVerLocal_xx, 1) = "." Then
                strDevVer_xx = Left$(strDevVerLocal_xx, Len(strDevVerLocal_xx) - 1)
            End If

            strVersionBD_x = Split(strDevVer_xx, ".")
            strVersionLocal_x = Split(strDevVerLocal_xx, ".")

            If LenB(Trim$(strDevVerLocal_xx)) > 0 Then
                If UBound(strVersionBD_x) > UBound(strVersionLocal_x) Then

                    For i = LBound(strVersionLocal_x) To UBound(strVersionLocal_x)

                        If IsNumeric(strVersionBD_x(i)) Then
                            If IsNumeric(strVersionLocal_x(i)) Then
                                If CLng(strVersionBD_x(i)) < CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = "<"

                                    Exit For

                                ElseIf CLng(strVersionBD_x(i)) > CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = ">"

                                    Exit For

                                Else

                                    If i = UBound(strVersionBD_x) Then
                                        ResultTemp = "="
                                    End If
                                End If
                            End If

                        Else
                            ResultTemp = "?"
                        End If

                    Next

                Else

                    For i = LBound(strVersionBD_x) To UBound(strVersionBD_x)

                        If IsNumeric(strVersionBD_x(i)) Then
                            If IsNumeric(strVersionLocal_x(i)) Then
                                If CLng(strVersionBD_x(i)) < CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = "<"

                                    Exit For

                                ElseIf CLng(strVersionBD_x(i)) > CLng(strVersionLocal_x(i)) Then
                                    ResultTemp = ">"

                                    Exit For

                                Else

                                    If i = UBound(strVersionBD_x) Then
                                        ResultTemp = "="
                                    End If
                                End If
                            End If

                        Else
                            ResultTemp = "?"
                        End If

                    Next

                End If

            Else
                ResultTemp = "?"
            End If

        Else
            ResultTemp = ">"
        End If

    Else
        ResultTemp = "?"
    End If

CompareFinish:
    CompareByVersion = ResultTemp
    DebugMode str7VbTab & "CompareByVersion-Result: " & strVersionBD & " " & ResultTemp & " " & strVersionLocal & vbNewLine & _
              str6VbTab & "CompareByVersion-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ConvertDate2Rus
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   dtDate (String)
'!--------------------------------------------------------------------------------
Public Sub ConvertDate2Rus(ByRef dtDate As String)

    Dim DD         As String
    Dim MM         As String
    Dim YYYY       As String
    Dim objRegExp  As RegExp
    Dim objMatch   As Match
    Dim objMatches As MatchCollection

    If LenB(dtDate) > 0 Then
        If InStr(dtDate, "unknown") = 0 Then
            Set objRegExp = New RegExp

            With objRegExp
                .Pattern = "(\d+).(\d+).(\d+)"
                .IgnoreCase = True
                .Global = True
            End With

            'получаем date1
            Set objMatches = objRegExp.Execute(dtDate)

            With objMatches

                If .Count > 0 Then
                    Set objMatch = .Item(0)
                    MM = Format$(objMatch.SubMatches(0), "00")
                    DD = Format$(objMatch.SubMatches(1), "00")
                    YYYY = DateTime.Year(dtDate)
                End If

            End With

            ' если необходимо конвертировать дату в формат dd/mm/yyyy
            If mbDateFormatRus Then
                dtDate = DD & "/" & MM & "/" & YYYY
            Else
                dtDate = MM & "/" & DD & "/" & YYYY
            End If
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompareByDate
'! Description (Описание)  :   [Check if date1 newer than date2]
'! Parameters  (Переменные):   Date1 (String)
'                              Date2 (String)
'!--------------------------------------------------------------------------------
Public Function CompareByDate(ByVal Date1 As String, ByVal Date2 As String) As String

    Dim objRegExp    As RegExp
    Dim objMatch     As Match
    Dim objMatches   As MatchCollection
    Dim m1           As Integer
    Dim M2           As Integer
    Dim d1           As Integer
    Dim d2           As Integer
    Dim Y1           As Integer
    Dim Y2           As Integer
    Dim strDate1     As String
    Dim strDate2     As String
    Dim strDate1_x() As String
    Dim strDate2_x() As String
    Dim strResult    As String

    DebugMode str5VbTab & "CompareByDate-Start" & vbNewLine & _
              str6VbTab & "CompareByDate: " & Date1 & " compare with " & Date2

    If InStr(Date1, "unknown") = 0 Then
        If InStr(Date1, ",") Then
            strDate1_x = Split(Trim$(Date1), ",")
            Date1 = strDate1_x(0)
        End If

        If InStr(Date2, ",") Then
            strDate2_x = Split(Trim$(Date2), ",")
            Date2 = strDate2_x(0)
        End If

        Set objRegExp = New RegExp

        With objRegExp
            .Pattern = "(\d+).(\d+).(\d+)"
            .IgnoreCase = True
            .Global = True
        End With

        'получаем date1
        Set objMatches = objRegExp.Execute(Date1)

        With objMatches

            If .Count > 0 Then
                Set objMatch = .Item(0)
                m1 = objMatch.SubMatches(0)
                d1 = objMatch.SubMatches(1)
                Y1 = objMatch.SubMatches(2)
            End If

        End With

        'получаем date2
        Set objMatches = objRegExp.Execute(Date2)

        With objMatches

            If .Count > 0 Then
                Set objMatch = .Item(0)
                M2 = objMatch.SubMatches(0)
                d2 = objMatch.SubMatches(1)
                Y2 = objMatch.SubMatches(2)
            End If

        End With

        If mbDateFormatRus Then
            strDate1 = Y1 & "." & d1 & "." & m1
            strDate2 = Y2 & "." & d2 & "." & M2
        Else
            strDate1 = Y1 & "." & m1 & "." & d1
            strDate2 = Y2 & "." & M2 & "." & d2
        End If

        strResult = CompareByVersion(strDate1, strDate2)
        CompareByDate = strResult
    Else
        CompareByDate = "?"
    End If

    DebugMode str6VbTab & "CompareByDate-Result: " & Date1 & " " & strResult & " " & Date1 & vbNewLine & _
              str5VbTab & "CompareByDate-End"
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ConvertString
'! Description (Описание)  :   [Заменяем в строке некоторые символы RegExp на константы VB]
'! Parameters  (Переменные):   strStringText (String)
'!--------------------------------------------------------------------------------
Public Function ConvertString(strStringText As String) As String

    If InStr(strStringText, "\t") Then
        strStringText = Replace$(strStringText, "\t", vbTab)
    End If

    If InStr(strStringText, "\r\n") Then
        strStringText = Replace$(strStringText, "\r\n", vbNewLine)
    End If

    If InStr(strStringText, "\r") Then
        strStringText = Replace$(strStringText, "\r", vbCr)
    End If

    If InStr(strStringText, "\n") Then
        strStringText = Replace$(strStringText, "\n", vbLf)
    End If

    ConvertString = strStringText
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function TrimNull
'! Description (Описание)  :   [получаем значение из буфера данных]
'! Parameters  (Переменные):   startstr (String)
'!--------------------------------------------------------------------------------
Public Function TrimNull(ByVal startstr As String) As String
    TrimNull = Left$(startstr, lstrlenW(StrPtr(startstr)))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub DelSpaceAfterZPT
'! Description (Описание)  :   [Удаление пробелов после запятой в строке с версией драйвера]
'! Parameters  (Переменные):   strVersion (String)
'!--------------------------------------------------------------------------------
Public Sub DelSpaceAfterZPT(ByRef strVersion As String)

    If InStr(strVersion, ",   ") Then
        strVersion = Replace$(strVersion, ",   ", ",", ",")
    End If

    If InStr(strVersion, ",  ") Then
        strVersion = Replace$(strVersion, ",  ", ",")
    End If

    If InStr(strVersion, ", ") Then
        strVersion = Replace$(strVersion, ", ", ",")
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ConvertVerByDate
'! Description (Описание)  :   [Check if date1 newer than date2]
'! Parameters  (Переменные):   strVersion1 (String)
'!--------------------------------------------------------------------------------
Public Sub ConvertVerByDate(ByRef strVersion1 As String)

    Dim strVer     As String
    Dim strVer_x() As String

    If LenB(strVersion1) > 0 Then
        If InStr(strVersion1, "unknown") = 0 Then
            If InStr(strVersion1, ",") Then
                strVer_x = Split(strVersion1, ",")
                strVersion1 = strVer_x(0)
                strVer = strVer_x(1)
            End If

            ConvertDate2Rus strVersion1
            strVersion1 = strVersion1 & "," & strVer
        End If
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ReplaceBadSymbol
'! Description (Описание)  :   [Удаляем лишние символы из строки]
'! Parameters  (Переменные):   strString (String)
'!--------------------------------------------------------------------------------
Public Function ReplaceBadSymbol(ByVal strString As String) As String

    ' Убираем символ ","
    If InStr(strString, ",") Then
        strString = Replace$(strString, ",", vbNullString)
    End If

    ' Убираем символ "*"
    If InStr(strString, "*") Then
        strString = Replace$(strString, "*", vbNullString)
    End If

    ' Убираем символ "!"
    If InStr(strString, "!") Then
        strString = Replace$(strString, "!", vbNullString)
    End If

    ' Убираем символ "@"
    If InStr(strString, "@") Then
        strString = Replace$(strString, "@", vbNullString)
    End If

    ' Убираем символ "#"
    If InStr(strString, "#") Then
        strString = Replace$(strString, "#", vbNullString)
    End If

    ' Убираем символ "™"
    If InStr(strString, "™") Then
        strString = Replace$(strString, "™", vbNullString)
    End If

    ' Убираем символ "®"
    If InStr(strString, "®") Then
        strString = Replace$(strString, "®", vbNullString)
    End If

    ' Убираем символ "?"
    If InStr(strString, "?") Then
        strString = Replace$(strString, "?", vbNullString)
    End If

    ' Убираем символ ";"
    If InStr(strString, ";") Then
        strString = Replace$(strString, ";", vbNullString)
    End If

    ' Убираем символ ":"
    If InStr(strString, ":") Then
        strString = Replace$(strString, ":", vbNullString)
    End If

    ' Убираем символ "   "
    If InStr(strString, "   ") Then
        strString = Replace$(strString, "   ", " ")
    End If

    ' Убираем символ "  "
    If InStr(strString, "  ") Then
        strString = Replace$(strString, "  ", " ")
    End If

    ReplaceBadSymbol = Trim$(strString)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ReplaceBadSymbolInDate
'! Description (Описание)  :   [Удаляем лишние символы из строки]
'! Parameters  (Переменные):   strString (String)
'!--------------------------------------------------------------------------------
Public Function ReplaceBadSymbolInDate(strString As String) As String

    ' Убираем символ "Fri"
    If InStr(1, strString, "fri", vbTextCompare) Then
        strString = Replace$(strString, "fri", vbNullString, , , vbTextCompare)
    End If

    ' Убираем символ "Thu"
    If InStr(1, strString, "thu", vbTextCompare) Then
        strString = Replace$(strString, "thu", vbNullString, , , vbTextCompare)
    End If

    ' Убираем символ " "
    If InStr(strString, " ") Then
        strString = Replace$(strString, " ", vbNullString)
    End If

    ReplaceBadSymbolInDate = strString
End Function

