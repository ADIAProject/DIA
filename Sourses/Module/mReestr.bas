Attribute VB_Name = "mReestr"
Option Explicit

' Работа с реестром
Public regParam                      As String

Private Const REG_SZ                 As Long = 1
Private Const REG_EXPAND_SZ          As Long = 2
Private Const REG_MULTI_SZ           As Long = 7
Private Const REG_DWORD              As Long = 4
Private Const ERROR_SUCCESS          As Long = 0
Private Const REG_BINARY             As Long = 3

Public Const HKEY_CLASSES_ROOT       As Long = &H80000000
Public Const HKEY_CURRENT_USER       As Long = &H80000001
Public Const HKEY_LOCAL_MACHINE      As Long = &H80000002

Private Const KEY_ALL_ACCESS         As Long = &H3F
Private Const ERROR_NONE             As Long = 0
Private Const STANDARD_RIGHTS_READ   As Long = &H20000
Private Const KEY_QUERY_VALUE        As Long = &H1
Private Const KEY_ENUMERATE_SUB_KEYS As Long = &H8
Private Const KEY_NOTIFY             As Long = &H10
Private Const SYNCHRONIZE            As Long = &H100000
Private Const KEY_READ               As Long = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))

' for wow64 access
Private Const KEY_WOW64_32KEY        As Long = &H100
Private Const KEY_WOW64_64KEY        As Long = &H200
Private Const KEY_CREATE_SUB_KEY     As Long = &H4
Private Const KEY_SET_VALUE          As Long = &H2

Private Declare Function RegEnumKey Lib "advapi32.dll" Alias "RegEnumKeyA" (ByVal hkey As Long, ByVal dwIndex As Long, ByVal lpName As String, ByVal cbName As Long) As Long
Private Declare Function RegOpenKey Lib "advapi32.dll" Alias "RegOpenKeyA" (ByVal hkey As Long, ByVal lpSubKey As String, phkResult As Long) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, ByRef phkResult As Long) As Long
Private Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, ByRef lpType As Long, ByVal lpData As String, ByRef lpcbData As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hkey As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hkey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, phkResult As Long, lpdwDisposition As Long) As Long

' Note that if you declare the lpData parameter as String, you must pass it By Value.
Private Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hkey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetKeyValue
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   KeyRoot (Long)
'                              KeyName (String)
'                              SubKeyRef (String)
'                              mbReadKeyRights (Boolean = False)
'!--------------------------------------------------------------------------------
Public Function GetKeyValue(ByVal KeyRoot As Long, ByVal KeyName As String, ByVal SubKeyRef As String, Optional ByVal mbReadKeyRights As Boolean = False) As String

    Dim I                    As Long
    Dim RC                   As Long
    Dim hkey                 As Long
    Dim sKeyVal              As String
    Dim sKeyVal_x()          As String
    Dim lKeyValType          As Long
    Dim tmpVal               As String
    Dim KeyValSize           As Long
    Dim GetKeyValueMultiSZ() As String
    Dim intTemp              As Long
    Dim intTempSmallBuff     As Long
    Dim Index                As Long
    Dim strstr               As String
    Dim lngKeyRights         As Long

    ' Loop Counter
    ' Return Code
    ' Handle To An Open Registry Key
    ' Data Type Of A Registry Key
    ' Tempory Storage For A Registry Key Value
    ' Size Of Registry Key Variable
    ' Open RegKey Under KeyRoot {HKEY_LOCAL_MACHINE...}
    '------------------------------------------------------------
    If mbReadKeyRights Then
        lngKeyRights = KEY_READ
    Else
        lngKeyRights = KEY_READ Or KEY_ALL_ACCESS
    End If

    RC = RegOpenKeyEx(KeyRoot, KeyName, 0, lngKeyRights, hkey)

    ' Open Registry Key
    If RC <> ERROR_SUCCESS Then
        GoTo GetKeyError
    End If

    ' Handle Error...
    tmpVal = String$(1024, 0)
    ' Allocate Variable Space
    KeyValSize = 1024
    ' Mark Variable Size
    '------------------------------------------------------------
    ' Retrieve Registry Key Value...
    '------------------------------------------------------------
    RC = RegQueryValueEx(hkey, SubKeyRef, 0, lKeyValType, tmpVal, KeyValSize)

    ' Get/Create Key Value
    If RC <> ERROR_SUCCESS Then
        GoTo GetKeyError
    End If

    ' Handle Errors
    '------------------------------------------------------------
    ' Determine Key Value Type For Conversion...
    '------------------------------------------------------------
    Select Case lKeyValType

            ' Search Data Types...
        Case REG_SZ, REG_EXPAND_SZ
            'tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)
            ' String Registry Key Data Type
            sKeyVal = TrimNull(tmpVal)

            ' Copy String Value
        Case REG_DWORD
            'tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)
            tmpVal = TrimNull(tmpVal)

            ' Double Word Registry Key Data Type
            If LenB(tmpVal) Then

                For I = Len(tmpVal) To 1 Step -1
                    ' Convert Each Bit
                    ' Build Value Char. By Char.
                    sKeyVal = sKeyVal + Hex$(Asc(Mid$(tmpVal, I, 1)))
                Next

                sKeyVal = Format$("&h" & sKeyVal)
                ' Convert Double Word To String
            Else
                sKeyVal = 0
            End If

        Case REG_MULTI_SZ
            intTemp = InStr(tmpVal, str2vbNullChar)
            tmpVal = Left$(tmpVal, intTemp)
            sKeyVal_x = Split(tmpVal, vbNullChar)

            ReDim GetKeyValueMultiSZ(UBound(sKeyVal_x))

            Do Until LenB(tmpVal) = 0
                intTempSmallBuff = InStrRev(tmpVal, vbNullChar)

                If intTempSmallBuff Then
                    intTemp = InStr(tmpVal, vbNullChar)
                    strstr = Left$(tmpVal, intTemp)
                    GetKeyValueMultiSZ(Index) = TrimNull(strstr)
                    Index = Index + 1
                    tmpVal = Mid$(tmpVal, intTemp + 1, Len(tmpVal))
                Else
                    sKeyVal = vbNullString
                End If

            Loop

            ReDim Preserve GetKeyValueMultiSZ(Index)

            For I = LBound(GetKeyValueMultiSZ) To UBound(GetKeyValueMultiSZ) - 1

                If LenB(sKeyVal) Then
                    If LenB(GetKeyValueMultiSZ(I)) Then
                        sKeyVal = sKeyVal & " | " & GetKeyValueMultiSZ(I)
                    End If

                Else
                    sKeyVal = GetKeyValueMultiSZ(I)
                End If

            Next

        Case Else
            'tmpVal = Left$(tmpVal, InStr(tmpVal, vbNullChar) - 1)
            sKeyVal = TrimNull(tmpVal)
    End Select

    GetKeyValue = Trim$(sKeyVal)
    ' Return Value
    RC = RegCloseKey(hkey)

    ' Close Registry Key
    Exit Function

    ' Exit
GetKeyError:
    ' Cleanup After An Error Has Occured...
    GetKeyValue = vbNullString
    If mbDebugDetail Then DebugMode "Error read RegistryParam: Key:" & KeyName & " Param:" & SubKeyRef & " Error: №" & RC & " - " & ApiErrorText(RC)
    ' Set Return Val To Empty String
    RC = RegCloseKey(hkey)
    ' Close Registry Key
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function GetRegString
'! Description (Описание)  :   [Получение значение ключа из реестра]
'! Parameters  (Переменные):   hkey (Long)
'                              strSubKey (String)
'                              strValueName (String)
'!--------------------------------------------------------------------------------
Public Function GetRegString(hkey As Long, strSubKey As String, strValueName As String) As String

    Dim strSetting As String
    Dim lngDataLen As Long
    Dim lngRes     As Long

    If RegOpenKey(hkey, strSubKey, lngRes) = ERROR_SUCCESS Then
        strSetting = FillNullChar(MAX_PATH)
        lngDataLen = MAX_PATH

        If RegQueryValueEx(lngRes, strValueName, ByVal 0, REG_EXPAND_SZ, ByVal strSetting, lngDataLen) = ERROR_SUCCESS Then
            If lngDataLen > 1 Then
                'GetRegString = Left$(strSetting, lngDataLen - 1)
                GetRegString = TrimNull(strSetting)
            End If
        End If

        If RegCloseKey(lngRes) <> ERROR_SUCCESS Then
            MsgBox "RegCloseKey Failed: " & strSubKey, vbCritical, strProductName
        End If
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function IsAppPresent
'! Description (Описание)  :   [Присутствует ли данное приложение в реестре]
'! Parameters  (Переменные):   strSubKey (String)
'                              strValueName (String)
'!--------------------------------------------------------------------------------
Public Function IsAppPresent(strSubKey As String, strValueName As String) As Boolean
    regParam = GetRegString(HKEY_CLASSES_ROOT, strSubKey, strValueName)
    IsAppPresent = CBool(Len(regParam))
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function ListKey
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   hkey (Variant)
'                              Key (Variant)
'                              mbReadKeyRights (Boolean = True) As String()
'!--------------------------------------------------------------------------------
Public Function ListKey(ByVal hkey, ByVal Key, Optional ByVal mbReadKeyRights As Boolean = True) As String()

    Dim strValue      As String
    Dim lDataLen      As Long
    Dim lResult       As Long
    Dim lValueLen     As Long
    Dim lCurIdx       As Long
    Dim lRetVal       As Long
    Dim hKeyResult    As Long
    Dim K()           As String
    Dim lngKeyRights  As Long
    Dim miMaxCountArr As Long

    'Variable to hold current enumerated key
    'Length of data
    'Result of RegEnumKey
    'Current Index which gets incremented with each pass through the loop
    'Result of RegOpenKeyEx
    If mbReadKeyRights Then
        lngKeyRights = KEY_READ
    Else
        lngKeyRights = KEY_READ Or KEY_ALL_ACCESS
    End If

    ' максимальное кол-во элементов в массиве
    miMaxCountArr = 500

    ReDim K(miMaxCountArr)

    lRetVal = RegOpenKeyEx(hkey, Key, 0, lngKeyRights, hKeyResult)

    'Open key with Full Access Rights
    If lRetVal = ERROR_SUCCESS Then
        'data Length
        lDataLen = 64
        lValueLen = 64

        Do
            strValue = String$(lValueLen, 0)
            'get current key's value
            lResult = RegEnumKey(hKeyResult, lCurIdx, strValue, lDataLen)

            'Enumerate keys
            'if successful, add current enumerated key to the txtEnumKeys textbox
            If lResult = ERROR_SUCCESS Then

                ' Если записей в массиве становится больше чем объявлено, то увеличиваем размерность массива
                If lCurIdx = miMaxCountArr Then
                    miMaxCountArr = miMaxCountArr + miMaxCountArr

                    ReDim Preserve K(miMaxCountArr)

                End If

                K(lCurIdx) = Replace$(strValue, vbNullChar, vbNullString)
            End If

            'Increment counter for next enumeration
            lCurIdx = lCurIdx + 1
            'continue while successful
        Loop While lResult = ERROR_SUCCESS

        'Close key
        RegCloseKey hKeyResult
        'If lRetVal is unsuccessful
    Else
        MsgBox "Cannot Open Key"
    End If

    ' Итоговое переобъявление массива на реальное кол-во записей
    If lCurIdx Then
        ReDim Preserve K(lCurIdx - 1)
    Else
        ReDim Preserve K(0)
    End If

    ListKey = K
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub SetRegBin
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Key (Long)
'                              SubName (String)
'                              ValueName (String)
'                              Value() (Byte)
'!--------------------------------------------------------------------------------
Public Sub SetRegBin(ByVal Key As Long, ByVal SubName As String, ByVal ValueName As String, Value() As Byte)

    Dim hkey As Long
    Dim sA   As SECURITY_ATTRIBUTES

    If RegCreateKeyEx(Key, SubName, 0, "", 0, KEY_CREATE_SUB_KEY Or KEY_SET_VALUE, sA, hkey, 0) <> 0 Then

        Exit Sub

    End If

    RegSetValueEx hkey, ValueName, 0, REG_BINARY, Value(0), UBound(Value) + 1
    RegCloseKey hkey
End Sub
