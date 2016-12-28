Attribute VB_Name = "mIniFile"
Option Explicit

Private IndexDevIDMass As Long

Private Declare Function GetPrivateProfileSection Lib "kernel32.dll" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileInt Lib "kernel32.dll" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32.dll" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStringW Lib "kernel32.dll" (ByVal lpApplicationName As Long, ByVal lpKeyName As Long, ByVal lpDefault As Long, ByVal lpReturnedString As Long, ByVal nSize As Long, ByVal lpFileName As Long) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32.dll" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As Any, ByVal lpFileName As String) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CheckIniSectionExists
'! Description (Описание)  :   [sub to load all keys from an ini section into a listbox]
'! Parameters  (Переменные):   strSection (String)
'                              strfullpath (String)
'!--------------------------------------------------------------------------------
Public Function CheckIniSectionExists(ByVal strSection As String, ByVal strIniPath As String) As Boolean

    Dim strBuffer As String
    Dim nTemp     As Long

    strBuffer = FillNullChar(5120)
    nTemp = GetPrivateProfileSection(strSection, strBuffer, Len(strBuffer), strIniPath)

    CheckIniSectionExists = nTemp

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetIniEmptySectionFromList
'! Description (Описание)  :   [Получение списка пустых секций из полученного ранее списка секций]
'! Parameters  (Переменные):   strSectionList (String)
'                              strIniPath (String)
'!--------------------------------------------------------------------------------
Public Function GetIniEmptySectionFromList(ByVal strSectionList As String, ByVal strIniPath As String) As String

    Dim strTmp             As String
    Dim strSectionList_x() As String
    Dim ii                 As Long
    Dim strManufSection    As String
    Dim sTemp              As String * 2048

    strSectionList_x = Split(strSectionList, "|")

    For ii = 0 To UBound(strSectionList_x)
        strManufSection = strSectionList_x(ii)
        sTemp = vbNullString
    
        If GetPrivateProfileSection(strManufSection, sTemp, 2048, strIniPath) = 0 Then
        
            If LenB(strTmp) Then
                strTmp = strTmp & strComma & strManufSection
            Else
                strTmp = strManufSection
            End If
        End If

    Next

    If LenB(strTmp) Then
        GetIniEmptySectionFromList = strTmp
    Else
        GetIniEmptySectionFromList = strDash
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetIniValueBoolean
'! Description (Описание)  :   [Получение Boolean значения переменной ini-файла с дефолтовым значением]
'! Parameters  (Переменные):   strIniPath (String)
'                              strIniSection (String)
'                              strIniValue (String)
'                              lngValueDefault (Long)
'!--------------------------------------------------------------------------------
Public Function GetIniValueBoolean(ByVal strIniPath As String, ByVal strIniSection As String, ByVal strIniValue As String, ByVal lngValueDefault As Long) As Boolean

    Dim lngResult As Long

    lngResult = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If lngResult = 9999 Then
        lngResult = lngValueDefault
    End If

    GetIniValueBoolean = CBool(lngResult)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetIniValueLong
'! Description (Описание)  :   [Получение Long значения переменной ini-файла с дефолтовым значением]
'! Parameters  (Переменные):   strIniPath (String)
'                              strIniSection (String)
'                              strIniValue (String)
'                              lngValueDefault (Long)
'!--------------------------------------------------------------------------------
Public Function GetIniValueLong(ByVal strIniPath As String, ByVal strIniSection As String, ByVal strIniValue As String, ByVal lngValueDefault As Long) As Long

    Dim lngResult As Long

    lngResult = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If lngResult = 9999 Then
        lngResult = lngValueDefault
    End If

    GetIniValueLong = lngResult
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetIniValueString
'! Description (Описание)  :   [Получение String значения переменной ini-файла с дефолтовым значением]
'! Parameters  (Переменные):   strIniPath (String)
'                              strIniSection (String)
'                              strIniValue (String)
'                              strValueDefault (String)
'!--------------------------------------------------------------------------------
Public Function GetIniValueString(ByVal strIniPath As String, ByVal strIniSection As String, ByVal strIniValue As String, ByVal strValueDefault As String) As String

    Dim strValue As String

    strValue = IniStringPrivate(strIniSection, strIniValue, strIniPath)

    If StrComp(strValue, "no_key") = 0 Then
        strValue = strValueDefault
    End If

    GetIniValueString = strValue
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetSectionMass
'! Description (Описание)  :   [Читает имена значений и переменных в массив в указанной секции .INI]
'! Parameters  (Переменные):   SekName (String) - имя секции (регистр не учитывается)
'                              IniFileName (String) - имя файла .ini (если путь к файлу не указан,файл ищется в папке Windows)
'                              mbFirstValue (Boolean)   - если требуется прочитать только первую строку в секции
'! Return Value:Возвр. знач.:  Малый буфер или Нет секции если есть ошибки в работе функции. Иначе возвращает массив переменная=значение
'!--------------------------------------------------------------------------------
Public Function GetSectionMass(ByVal SekName As String, ByVal IniFileName As String, Optional ByVal mbFirstValue As Boolean)

    Dim strBuffer        As String * 32767
    Dim strTemp          As String
    Dim intTemp          As Long
    Dim intTempSmallBuff As Long
    Dim intSize          As Long
    Dim lngIndex         As Long
    Dim arrSection()     As String
    Dim arrSectionTemp() As String
    Dim strKey           As String
    Dim strValue         As String
    Dim strString        As String
    Dim lpKeyValue()     As String
    Dim miRavnoPosition  As Long

    On Error GoTo PROC_ERR

    lngIndex = 1
    intSize = GetPrivateProfileSection(SekName, strBuffer, 32767, IniFileName)
    strTemp = Left$(strBuffer, intSize)

    If mbFirstValue Then

        ReDim arrSection(1, 2)

        arrSectionTemp = Split(strTemp, vbNullChar)
        intTempSmallBuff = InStrRev(strTemp, vbNullChar)

        If intTempSmallBuff Then
            strString = arrSectionTemp(0)
            miRavnoPosition = InStr(strString, strRavno)

            If miRavnoPosition Then
                strKey = Left$(strString, miRavnoPosition - 1)
                strValue = Mid$(strString, miRavnoPosition + 1)
            Else
                strKey = strString
                strValue = strString
            End If

            arrSection(lngIndex, 1) = strKey
            arrSection(lngIndex, 2) = strValue
            IndexDevIDMass = 1
            GoTo IF_EXIT
        Else

            ReDim arrSection(1, 2)

            arrSection(1, 1) = "small_buffer"
            arrSection(1, 2) = "small_buffer"
            IndexDevIDMass = 1
            GoTo IF_EXIT
        End If
    End If

    If LenB(strTemp) Then
        lpKeyValue = Split(strTemp, vbNullChar)

        ReDim arrSection(UBound(lpKeyValue), 2)

        Do Until LenB(strTemp) = 0
            intTempSmallBuff = InStrRev(strTemp, vbNullChar)

            If intTempSmallBuff Then
                intTemp = InStr(strTemp, vbNullChar)
                strString = Left$(strTemp, intTemp)

                If InStr(strString, "---") Then
                    strKey = "Строка без ID"
                    strValue = "Строка без ID"
                    GoTo Save_StrKey
                End If

                miRavnoPosition = InStr(strString, strRavno)

                If miRavnoPosition Then
                    strKey = Left$(strString, miRavnoPosition - 1)
                    strValue = Mid$(strString, miRavnoPosition + 1)
                Else
                    strKey = TrimNull(strString)
                    strValue = TrimNull(strString)
                End If

Save_StrKey:
                arrSection(lngIndex, 1) = strKey
                arrSection(lngIndex, 2) = strValue
                lngIndex = lngIndex + 1
                strTemp = Mid$(strTemp, intTemp + 1, Len(strTemp))
            Else

                ReDim arrSection(1, 2)

                arrSection(1, 1) = "small_buffer"
                arrSection(1, 2) = "small_buffer"
                IndexDevIDMass = 1
                GoTo IF_EXIT
            End If

        Loop

    Else

        ReDim arrSection(lngIndex, 2)

        arrSection(lngIndex, 1) = "no_section"
        arrSection(lngIndex, 2) = "no_section"
    End If

    IndexDevIDMass = lngIndex
IF_EXIT:
    GetSectionMass = arrSection
PROC_EXIT:

    Exit Function

PROC_ERR:

    If Not mbSilentRun Then
        MsgBox "Error:  Err.Number: " & Err.Number & " Err.Description: " & Err.Description, vbExclamation + vbOKOnly, "GetValueString"
    End If

    Resume PROC_EXIT

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IniDelAllKeyPrivate
'! Description (Описание)  :   [Удаляет все ключи в заданной секции в приватном файле .INI - заодно удаляет и саму секцию!?]
'! Parameters  (Переменные):   SekName (String) - имя секции (регистр не учитывается)
'                              IniFileName (String) - имя файла .ini (если путь к файлу не указан,файл ищется в папке Windows)
'!--------------------------------------------------------------------------------
Public Function IniDelAllKeyPrivate(SekName As String, IniFileName As String)
    WritePrivateProfileString SekName, vbNullString, vbNullString, IniFileName
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IniLongPrivate
'! Description (Описание)  :   [Читает целый параметр из любого файла .INI, возвращяет 9999, если ключ не найден]
'! Parameters  (Переменные):   SekName (String)
'                              KeyName (String)
'                              IniFileName (String)
'!--------------------------------------------------------------------------------
Public Function IniLongPrivate(ByVal SekName As String, ByVal KeyName As String, ByVal IniFileName As String) As Long
    IniLongPrivate = GetPrivateProfileInt(SekName, KeyName, 9999, IniFileName)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IniSectionIsEmpty
'! Description (Описание)  :   [Проверка на то что секция пустая]
'! Parameters  (Переменные):   strSection (String)
'                              strIni (String)
'!--------------------------------------------------------------------------------
Public Function IniSectionIsEmpty(ByVal strSection As String, ByVal strIni As String) As Boolean

    Dim sTemp As String * 2048

    'в неё запишется количество символов в строке ключа
    IniSectionIsEmpty = GetPrivateProfileSection(strSection, sTemp, 2048, strIni) = 0
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IniStringPrivate
'! Description (Описание)  :   [Читает строковый параметр из любого файла .INI,"no_key" - возвращаемое функцией значение, если ключ не найден]
'! Parameters  (Переменные):   SekName (String)
'                              KeyName (String)
'                              IniFileName (String)
'!--------------------------------------------------------------------------------
Public Function IniStringPrivate(ByVal SekName As String, ByVal KeyName As String, ByVal IniFileName As String) As String

    'строковый буфер(под значение ключа)
    Dim sTemp()     As Byte
    Dim nTemp       As Long

    ReDim sTemp(4096)
    'в неё запишется количество символов в строке ключа
    'ограничение - параметр не может быть больше 4096 символов
    nTemp = GetPrivateProfileStringW(StrPtr(SekName), StrPtr(KeyName), StrPtr("no_key"), VarPtr(sTemp(0)), -1, StrPtr(IniFileName))
    IniStringPrivate = Left$(sTemp(), nTemp * 2)
    IniStringPrivate = TrimNull(IniStringPrivate)
    Erase sTemp
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub IniWriteStrPrivate
'! Description (Описание)  :   [Записывает строковый параметр в любой файл .INI]
'! Parameters  (Переменные):   SekName (String)
'                              KeyName (String)
'                              Param (String)
'                              IniFileName (String)
'!--------------------------------------------------------------------------------
Public Sub IniWriteStrPrivate(ByVal SekName As String, ByVal KeyName As String, ByVal Param As String, ByVal IniFileName As String)
    WritePrivateProfileString SekName, KeyName, Param, IniFileName
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function LoadIniSectionKeys
'! Description (Описание)  :   [load all keys from an ini section]
'! Parameters  (Переменные):   strSection (String)
'                              strfullpath (String)
'                              mbKeys (Boolean = True) As String()
'!--------------------------------------------------------------------------------
Public Function LoadIniSectionKeys(ByVal strSection As String, ByVal strFullPath As String, Optional ByVal mbKeys As Boolean = True) As String()

    Dim strKeyAndVal() As String
    Dim strKey_Val()   As String
    Dim strBuffer      As String
    Dim intx           As Long
    Dim Z()            As String
    Dim n              As Long

    n = -1
    strBuffer = FillNullChar(5120)
    GetPrivateProfileSection strSection, strBuffer, Len(strBuffer), strFullPath
    strKeyAndVal = Split(strBuffer, vbNullChar)

    For intx = LBound(strKeyAndVal) To UBound(strKeyAndVal)

        If LenB(strKeyAndVal(intx)) = 0 Then

            Exit For

        End If

        strKey_Val = Split(strKeyAndVal(intx), strRavno)

        If UBound(strKey_Val) = -1 Then

            Exit For

        End If

        n = n + 1

        ReDim Preserve Z(n)

        If mbKeys Then
            ' только ключи
            Z(n) = strKey_Val(0)
        Else

            ' только значения ключей
            If UBound(strKey_Val) = 1 Then
                Z(n) = strKey_Val(1)
            End If
        End If

    Next

    Erase strKeyAndVal
    Erase strKey_Val

    If n = -1 Then

        ReDim Z(0)

    End If

    LoadIniSectionKeys = Z
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub NormIniFile
'! Description (Описание)  :   [Привидение ини файла в "читабельный" вид]
'! Parameters  (Переменные):   sFileName (String)
'!--------------------------------------------------------------------------------
Public Sub NormIniFile(ByVal sFileName As String)

    Dim nf          As Long
    Dim ub          As Long
    Dim sBuffer     As String
    Dim slArray()   As String
    Dim sOutArray() As String

    If GetFileSizeByPath(sFileName) Then
        nf = FreeFile
        
        Open sFileName For Binary Access Read Lock Write As nf
        sBuffer = String$(LOF(nf), 0&)
        Get nf, 1, sBuffer
        Close nf
        
        slArray = Split(sBuffer, vbNewLine)
        ub = &HFFFF

        For nf = 0 To UBound(slArray)

            If Len(slArray(nf)) Then
                ub = ub + IIf(Left$(slArray(nf), vbNull) = Chr$(&H5B), 2, vbNull)

                ReDim Preserve sOutArray(ub)

                sOutArray(ub) = slArray(nf)
            End If

        Next

        sBuffer = Join(sOutArray, vbNewLine)
        DeleteFiles sFileName
        
        nf = FreeFile
        Open sFileName For Binary Access Write Lock Read As nf
        Put nf, 1, sBuffer
        Close nf
    End If

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ReadFromINI
'! Description (Описание)  :   [use to read/write ini/inf file]
'! Parameters  (Переменные):   strSection (String)
'                              strKey (String)
'                              strfullpath (String)
'                              strDefault (String = vbNullString)
'!--------------------------------------------------------------------------------
Public Function ReadFromINI(ByVal strSection As String, ByVal strKey As String, ByVal strFullPath As String, Optional ByVal strDefault As String = vbNullString) As String

    Dim strBuffer As String

    strBuffer = FillNullChar(1024)
    ReadFromINI = Left$(strBuffer, GetPrivateProfileString(strSection, ByVal LCase$(strKey), strDefault, strBuffer, Len(strBuffer), strFullPath))
    
End Function
