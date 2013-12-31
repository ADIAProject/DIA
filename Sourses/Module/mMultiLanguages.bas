Attribute VB_Name = "mMultiLanguages"
Option Explicit

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Copyright ©1996-2009 VBnet, Randy Birch, All Rights Reserved.
' Some pages may also contain other copyrights by the author.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Distribution: You can freely use this code in your own
'               applications, but you may not reproduce
'               or publish this code on any web site,
'               online service, or distribute as source
'               on any media without express permission.
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Поддержка многоязычности в программе
Public mbMultiLanguage                As Boolean
Public arrLanguage()                  As String     ' Массив служебных сообщений
Public strPCLangID                    As String
Public strPCLangLocaliseName          As String
Public strPCLangEngName               As String
Public strPCLangCurrentPath           As String
Public strPCLangCurrentID             As String

'Язык программы при старте
Public mbAutoLanguage                 As Boolean
Public strStartLanguageID             As String

' Массив служебных сообщений
Public strMessages(150)               As String

' Api - переменные для работы с языками
Public Const LOCALE_ILANGUAGE         As Long = &H1    'language id
Public Const LOCALE_SLANGUAGE         As Long = &H2    'localized name of language
Public Const LOCALE_SENGLANGUAGE      As Long = &H1001    'English name of language

Private Const LOCALE_SABBREVLANGNAME  As Long = &H3    'abbreviated language name
Private Const LOCALE_SNATIVELANGNAME  As Long = &H4    'native name of language
Private Const LOCALE_IDEFAULTLANGUAGE As Long = &H9    'default language id

Public Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long

Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

' Получение Font.charset на основании кодовой страницы
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetCharsetFromLng
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngCodePage (Long)
'!--------------------------------------------------------------------------------
Public Function GetCharsetFromLng(lngCodePage As Long) As Long

    Dim lngCharset As Long

    Select Case lngCodePage

        Case 1251
            lngCharset = RUSSIAN_CHARSET

        Case 1250
            'EASTEUROPE_CHARSET = 238
            lngCharset = EASTEUROPE_CHARSET

        Case 1252
            'ANSI_CHARSET = 0
            lngCharset = ANSI_CHARSET

        Case 1253
            'GREEK_CHARSET = 161
            lngCharset = GREEK_CHARSET

        Case 1254
            'TURKISH_CHARSET = 162
            lngCharset = TURKISH_CHARSET

        Case 1255
            'HEBREW_CHARSET = 177
            lngCharset = HEBREW_CHARSET

        Case 1256
            'ARABIC_CHARSET = 178
            lngCharset = ARABIC_CHARSET

        Case 1257
            'BALTIC_CHARSET = 186
            lngCharset = BALTIC_CHARSET

        Case 1258
            'VIETNAMESE_CHARSET = 163
            lngCharset = VIETNAMESE_CHARSET

        Case 874
            lngCharset = THAI_CHARSET

        Case 932
            'SHIFTJIS_CHARSET = 128
            lngCharset = SHIFTJIS_CHARSET

        Case 949
            'HANGUL_CHARSET = 129
            lngCharset = HANGUL_CHARSET

        Case 936
            'GB2312_CHARSET = 134
            lngCharset = GB2312_CHARSET

        Case 950
            'CHINESEBIG5_CHARSET = 136
            lngCharset = CHINESEBIG5_CHARSET

        Case Else
            'DEFAULT_CHARSET = 1
            lngCharset = DEFAULT_CHARSET
    End Select

    GetCharsetFromLng = lngCharset
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetUserLocaleInfo
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   dwLocaleID (Long)
'                              dwLCType (Long)
'!--------------------------------------------------------------------------------
Public Function GetUserLocaleInfo(ByVal dwLocaleID As Long, ByVal dwLCType As Long) As String

    Dim sReturn As String
    Dim R       As Long

    'call the function passing the Locale type
    'variable to retrieve the required size of
    'the string buffer needed
    R = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, 0)

    'if successful..
    If R Then
        'pad the buffer with spaces
        sReturn = String$(R, vbNullChar)
        'and call again passing the buffer
        R = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

        'if successful (r > 0)
        If R Then
            'r holds the size of the string
            'including the terminating null
            GetUserLocaleInfo = TrimNull(sReturn)
            ', r - 1)
        End If
    End If

End Function

'! -----------------------------------------------------------
'!  Функция     :  LoadLanguageList
'!  Переменные  :
'!  Описание    :  Загрузка списка языков
'! -----------------------------------------------------------
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LoadLanguageList
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function LoadLanguageList() As Boolean

    Dim strFileList_x() As String
    Dim ii              As Integer
    Dim strTemp         As String
    Dim LngValue        As Long

    strFileList_x = SearchFilesInRoot(strAppPathBackSL & strToolsLang_Path, "*.lng", False, False)

    If UBound(strFileList_x, 2) = 0 Then
        If LenB(strFileList_x(0, 0)) = 0 Then
            LoadLanguageList = False

            Exit Function

        End If
    End If

    ReDim arrLanguage(6, UBound(strFileList_x, 2) + 1)

    For ii = LBound(strFileList_x, 2) To UBound(strFileList_x, 2)
        ' Путь до языкового файла
        strTemp = strFileList_x(0, ii)

        If strTemp <> "no_key" Then
            arrLanguage(1, ii + 1) = strTemp
        End If

        ' Имя языка
        strTemp = IniStringPrivate("Lang", "Name", strFileList_x(0, ii))

        If strTemp <> "no_key" Then
            arrLanguage(2, ii + 1) = strTemp
        End If

        ' Имя переводчика
        strTemp = IniStringPrivate("Lang", "TranslatorName", strFileList_x(0, ii))

        If strTemp <> "no_key" Then
            arrLanguage(4, ii + 1) = strTemp
        End If

        ' Адрес переводчика
        strTemp = IniStringPrivate("Lang", "TranslatorURL", strFileList_x(0, ii))

        If strTemp <> "no_key" Then
            arrLanguage(5, ii + 1) = strTemp
        End If

        ' Charset языка
        LngValue = GetIniValueLong(strFileList_x(0, ii), "Lang", "Charset", 1)
        arrLanguage(6, ii + 1) = LngValue
        ' ID языка
        strTemp = IniStringPrivate("Lang", "ID", strFileList_x(0, ii))

        If strTemp <> "no_key" Then
            arrLanguage(3, ii + 1) = strTemp

            If mbAutoLanguage Then
                If InStr(1, strTemp, strPCLangID, vbTextCompare) Then
                    strPCLangCurrentPath = arrLanguage(1, ii + 1)
                    strPCLangCurrentID = strPCLangID
                    lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(6, ii + 1)))
                End If

            Else

                If LenB(strStartLanguageID) > 0 Then
                    If InStr(1, strTemp, strStartLanguageID, vbTextCompare) Then
                        strPCLangCurrentPath = arrLanguage(1, ii + 1)
                        strPCLangCurrentID = strStartLanguageID
                        lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(6, ii + 1)))
                    End If
                End If
            End If
        End If

        LoadLanguageList = True
    Next

    If LenB(strPCLangCurrentPath) = 0 Then
        strPCLangCurrentPath = PathCombine(strAppPathBackSL & strToolsLang_Path, "English.lng")
        strPCLangCurrentID = "0409"
        lngFont_Charset = 1
    End If

End Function

'Локализация сообщений программы
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMessage
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Public Sub LocaliseMessage(StrPathFile As String)

    Dim i As Integer

    For i = 1 To UBound(strMessages)
        strMessages(i) = LocaliseString(StrPathFile, "Messages", "strMessages" & i, "strMessages" & i)
    Next i

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LocaliseString
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'                              strSection (String)
'                              strParam (String)
'                              strDefValue (String)
'!--------------------------------------------------------------------------------
Public Function LocaliseString(ByVal StrPathFile As String, ByVal strSection As String, ByVal strParam As String, ByVal strDefValue As String) As String

    Dim strTemp As String

    strTemp = IniStringPrivate(strSection, strParam, StrPathFile)

    If strTemp <> "no_key" Then
        LocaliseString = ConvertString(Trim$(strTemp))
    Else
        LocaliseString = strDefValue
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LoadLanguageOS
'! Description (Описание)  :   [Считываем язык операционной системы, и записываем в переменные Public]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Sub LoadLanguageOS()

    Dim LCID As Long

    ' Считываем язык операционой системы
    LCID = GetSystemDefaultLCID()
    'language id
    strPCLangID = GetUserLocaleInfo(LCID, LOCALE_ILANGUAGE)
    'localized name of language
    strPCLangLocaliseName = GetUserLocaleInfo(LCID, LOCALE_SLANGUAGE)
    'English name of language
    strPCLangEngName = GetUserLocaleInfo(LCID, LOCALE_SENGLANGUAGE)
End Sub

