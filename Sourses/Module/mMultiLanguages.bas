Attribute VB_Name = "mMultiLanguages"
Option Explicit

'Note: this file has been writen for use within Drivers Installer Assistant.

' Модуль для организации поддержки многоязычности в программе
' В процессе обработки создается массив данных о языковых файлах arrLanguage() , данные беруться из секции [Lang]
' Пример секции
' [Lang]
'    Name = English
'    TranslatorName = Romeo91 & stasys44
'    TranslatorURL = www.adia-project.net
'    ID=0409;0009;0c09;2809;1009;2409;4009;1809;2009;4409;1409;3409;4809;1c09;2c09;0809;3009;3c09;3809
'    Version=6.0.0 - справочное значение
'    Date=17/2/2014 - справочное значение
'    Charset=1252

' данные вносятся в языковой файл в подсекции с именами form
' параметрами являются имена контролов

Public mbMultiLanguage                As Boolean
Public arrLanguage()                  As String     ' Массив данных о языковых файлах
Public strPCLangID                    As String
Public strPCLangLocaliseName          As String
Public strPCLangEngName               As String
Public strPCLangCurrentPath           As String
Public strPCLangCurrentID             As String
Public strPCLangCurrentLangName       As String

'Язык программы при старте
Public mbAutoLanguage                 As Boolean
Public strStartLanguageID             As String

' Массив служебных сообщений
Public strMessages(163)               As String

' Api - переменные для работы с языками
Private Const LOCALE_ILANGUAGE         As Long = &H1     'language id
Private Const LOCALE_SLANGUAGE         As Long = &H2     'localized name of language
Private Const LOCALE_SENGLANGUAGE      As Long = &H1001  'English name of language
' Character sets
Private Const ANSI_CHARSET = 0
Private Const DEFAULT_CHARSET = 1
Private Const SYMBOL_CHARSET = 2
Private Const SHIFTJIS_CHARSET = 128
Private Const HANGEUL_CHARSET = 129
Private Const HANGUL_CHARSET = 129
Private Const GB2312_CHARSET = 134
Private Const CHINESEBIG5_CHARSET = 136
Private Const OEM_CHARSET = 255
Private Const JOHAB_CHARSET = 130
Private Const HEBREW_CHARSET = 177
Private Const ARABIC_CHARSET = 178
Private Const GREEK_CHARSET = 161
Private Const TURKISH_CHARSET = 162
Private Const VIETNAMESE_CHARSET = 163
Private Const THAI_CHARSET = 222
Private Const EASTEUROPE_CHARSET = 238
Private Const RUSSIAN_CHARSET = 204
Private Const MAC_CHARSET = 77
Private Const BALTIC_CHARSET = 186

'Private Const LOCALE_SABBREVLANGNAME  As Long = &H3     'abbreviated language name
'Private Const LOCALE_SNATIVELANGNAME  As Long = &H4     'native name of language
'Private Const LOCALE_IDEFAULTLANGUAGE As Long = &H9     'default language id

Private Declare Function GetSystemDefaultLCID Lib "kernel32.dll" () As Long
Private Declare Function GetLocaleInfo Lib "kernel32.dll" Alias "GetLocaleInfoA" (ByVal Locale As Long, ByVal LCType As Long, ByVal lpLCData As String, ByVal cchData As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetCharsetFromLng
'! Description (Описание)  :   [Получение Font.charset на основании кодовой страницы]
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
    Dim r       As Long

    'call the function passing the Locale type
    'variable to retrieve the required size of
    'the string buffer needed
    r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, 0)

    'if successful..
    If r Then
        'pad the buffer with spaces
        sReturn = FillNullChar(r)
        'and call again passing the buffer
        r = GetLocaleInfo(dwLocaleID, dwLCType, sReturn, Len(sReturn))

        'if successful (r > 0)
        If r Then
            'r holds the size of the string
            'including the terminating null
            GetUserLocaleInfo = TrimNull(sReturn)
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LoadLanguageList
'! Description (Описание)  :   [Загрузка списка языков]
'! Parameters  (Переменные):
'!--------------------------------------------------------------------------------
Public Function LoadLanguageList() As Boolean

    Dim strFileList_x() As FindListStruct
    Dim ii              As Integer
    Dim strTemp         As String
    Dim strLangFilePath As String
    Dim lngUbound       As Long

    strFileList_x = SearchFilesInRoot(strAppPathBackSL & strToolsLang_Path, "*.lng", False, False)

    lngUbound = UBound(strFileList_x)
    If lngUbound Then
        If LenB(strFileList_x(0).FullPath) Then
    
            ReDim arrLanguage(5, lngUbound)
        
            For ii = 0 To lngUbound
                
                ' Путь до языкового файла
                strLangFilePath = strFileList_x(ii).FullPath
                arrLanguage(0, ii) = strLangFilePath
                ' Имя языка
                arrLanguage(1, ii) = GetIniValueString(strLangFilePath, "Lang", "Name", vbNullString)
                ' Имя переводчика
                arrLanguage(3, ii) = GetIniValueString(strLangFilePath, "Lang", "TranslatorName", vbNullString)
                ' Адрес переводчика
                arrLanguage(4, ii) = GetIniValueString(strLangFilePath, "Lang", "TranslatorURL", vbNullString)
                ' Charset языка
                arrLanguage(5, ii) = GetIniValueLong(strLangFilePath, "Lang", "Charset", 1)
                ' ID языка
                strTemp = GetIniValueString(strLangFilePath, "Lang", "ID", vbNullString)
        
                If LenB(strTemp) Then
                    arrLanguage(2, ii) = strTemp
        
                    If mbAutoLanguage Then
                        If InStr(1, strTemp, strPCLangID, vbTextCompare) Then
                            strPCLangCurrentPath = arrLanguage(0, ii)
                            strPCLangCurrentLangName = arrLanguage(1, ii)
                            lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(5, ii)))
                            strPCLangCurrentID = strPCLangID
                        End If
        
                    Else
        
                        If LenB(strStartLanguageID) Then
                            If InStr(1, strTemp, strStartLanguageID, vbTextCompare) Then
                                strPCLangCurrentPath = arrLanguage(0, ii)
                                strPCLangCurrentLangName = arrLanguage(1, ii)
                                lngFont_Charset = GetCharsetFromLng(CLng(arrLanguage(5, ii)))
                                strPCLangCurrentID = strStartLanguageID
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
        End If
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

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub LocaliseMessage
'! Description (Описание)  :   [Локализация сообщений программы]
'! Parameters  (Переменные):   StrPathFile (String)
'!--------------------------------------------------------------------------------
Public Sub LocaliseMessage(strPathFile As String)

    Dim ii As Integer

    For ii = 1 To UBound(strMessages)
        strMessages(ii) = LocaliseString(strPathFile, "Messages", "strMessages" & ii, "strMessages" & ii)
    Next ii

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LocaliseString
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   StrPathFile (String)
'                              strSection (String)
'                              strParam (String)
'                              strDefValue (String)
'!--------------------------------------------------------------------------------
Public Function LocaliseString(ByVal strPathFile As String, ByVal strSection As String, ByVal strParam As String, ByVal strDefValue As String) As String

    Dim strTemp As String

    strTemp = Trim$(IniStringPrivate(strSection, strParam, strPathFile))

    If StrComp(strTemp, "no_key") <> 0 Then
        LocaliseString = ConvertString(strTemp)
    Else
        LocaliseString = strDefValue
    End If

End Function

