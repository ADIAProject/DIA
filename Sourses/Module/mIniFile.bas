Attribute VB_Name = "mIniFile"
Option Explicit

'Читает целый параметр из любого файла .INI
'Читает строку из любого файла .INI
'Записывает строку в любой файл .INI
'Читает список параметров и значений в секции
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
Public Function CheckIniSectionExists(ByVal strSection As String, ByVal strfullpath As String) As Boolean

    Dim strBuffer As String
    Dim nTemp     As Long

    strBuffer = FillNullChar(5120)
    nTemp = GetPrivateProfileSection(strSection, strBuffer, Len(strBuffer), strfullpath)

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
    Dim i_i                As Long
    Dim strManufSection    As String
    Dim sTemp              As String * 2048

    strSectionList_x = Split(strSectionList, "|")

    For i_i = 0 To UBound(strSectionList_x)
        strManufSection = strSectionList_x(i_i)
    
        If GetPrivateProfileSection(strManufSection, sTemp, 2048, strIniPath) = 0 Then
        
            If LenB(strTmp) Then
                strTmp = strTmp & "," & strManufSection
            Else
                strTmp = strManufSection
            End If
        End If

    Next

    If LenB(strTmp) Then
        GetIniEmptySectionFromList = strTmp
    Else
        GetIniEmptySectionFromList = "-"
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

    Dim lngValue As Long

    lngValue = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If lngValue = 9999 Then
        lngValue = lngValueDefault
    End If

    GetIniValueBoolean = CBool(lngValue)
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

    Dim lngValue As Long

    lngValue = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If lngValue = 9999 Then
        lngValue = lngValueDefault
    End If

    GetIniValueLong = lngValue
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
'                              FirstValue (Boolean)   - если требуется прочитать только первую строку в секции
'! Return Value:Возвр. знач.:  Малый буфер или Нет секции если есть ошибки в работе функции. Иначе возвращает массив переменная=значение
'!--------------------------------------------------------------------------------
Public Function GetSectionMass(ByVal SekName As String, ByVal IniFileName As String, Optional ByVal FirstValue As Boolean)

    Dim strBuffer        As String * 32767
    Dim strTemp          As String
    Dim intTemp          As Long
    Dim intTempSmallBuff As Long
    Dim intSize          As Long
    Dim Index            As Long
    Dim arrSection()     As String
    Dim arrSectionTemp() As String
    Dim Key              As String
    Dim Value            As String
    Dim str              As String
    Dim lpKeyValue()     As String
    Dim miRavnoPosition  As Long

    On Error GoTo PROC_ERR

    Index = 1
    intSize = GetPrivateProfileSection(SekName, strBuffer, 32767, IniFileName)
    strTemp = Left$(strBuffer, intSize)

    If FirstValue Then

        ReDim arrSection(1, 2)

        arrSectionTemp = Split(strTemp, vbNullChar)
        intTempSmallBuff = InStrRev(strTemp, vbNullChar)

        If intTempSmallBuff Then
            str = arrSectionTemp(0)
            miRavnoPosition = InStr(str, "=")

            If miRavnoPosition Then
                Key = Left$(str, miRavnoPosition - 1)
                Value = Mid$(str, miRavnoPosition + 1)
            Else
                Key = str
                Value = str
            End If

            arrSection(Index, 1) = Key
            arrSection(Index, 2) = Value
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
                str = Left$(strTemp, intTemp)

                If InStr(str, "---") Then
                    Key = "Строка без ID"
                    Value = "Строка без ID"
                    GoTo Save_StrKey
                End If

                miRavnoPosition = InStr(str, "=")

                If miRavnoPosition Then
                    Key = Left$(str, miRavnoPosition - 1)
                    Value = Mid$(str, miRavnoPosition + 1)
                Else
                    Key = TrimNull(str)
                    Value = TrimNull(str)
                End If

Save_StrKey:
                arrSection(Index, 1) = Key
                arrSection(Index, 2) = Value
                Index = Index + 1
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

        ReDim arrSection(Index, 2)

        arrSection(Index, 1) = "no_section"
        arrSection(Index, 2) = "no_section"
    End If

    IndexDevIDMass = Index
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
Public Function LoadIniSectionKeys(ByVal strSection As String, ByVal strfullpath As String, Optional ByVal mbKeys As Boolean = True) As String()

    Dim KeyAndVal() As String
    Dim Key_Val()   As String
    Dim strBuffer   As String
    Dim intx        As Long
    Dim Z()         As String
    Dim n           As Long

    n = -1
    strBuffer = FillNullChar(5120)
    GetPrivateProfileSection strSection, strBuffer, Len(strBuffer), strfullpath
    KeyAndVal = Split(strBuffer, vbNullChar)

    For intx = LBound(KeyAndVal) To UBound(KeyAndVal)

        If LenB(KeyAndVal(intx)) = 0 Then

            Exit For

        End If

        Key_Val = Split(KeyAndVal(intx), "=")

        If UBound(Key_Val) = -1 Then

            Exit For

        End If

        n = n + 1

        ReDim Preserve Z(n)

        If mbKeys Then
            ' только ключи
            Z(n) = Key_Val(0)
        Else

            ' только значения ключей
            If UBound(Key_Val) = 1 Then
                Z(n) = Key_Val(1)
            End If
        End If

    Next

    Erase KeyAndVal
    Erase Key_Val

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

    nf = FreeFile

    If Not FileLen(sFileName) = 0& Then
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
Public Function ReadFromINI(ByVal strSection As String, ByVal strKey As String, ByVal strfullpath As String, Optional ByVal strDefault As String = vbNullString) As String

    Dim strBuffer As String

    strBuffer = FillNullChar(1024)
    ReadFromINI = Left$(strBuffer, GetPrivateProfileString(strSection, ByVal LCase$(strKey), strDefault, strBuffer, Len(strBuffer), strfullpath))
End Function
