Attribute VB_Name = "mIniFile"
Option Explicit

'������ ����� �������� �� ������ ����� .INI
'������ ������ �� ������ ����� .INI
'���������� ������ � ����� ���� .INI
'������ ������ ���������� � �������� � ������
Private IndexDevIDMass                  As Long
Private Declare Function GetPrivateProfileSection _
                          Lib "kernel32.dll" _
                              Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, _
                                                                 ByVal lpReturnedString As String, _
                                                                 ByVal nSize As Long, _
                                                                 ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileInt _
                          Lib "kernel32.dll" _
                              Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, _
                                                             ByVal lpKeyName As String, _
                                                             ByVal nDefault As Long, _
                                                             ByVal lpFileName As String) As Long

Private Declare Function GetPrivateProfileString _
                          Lib "kernel32.dll" _
                              Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                ByVal lpKeyName As String, _
                                                                ByVal lpDefault As String, _
                                                                ByVal lpReturnedString As String, _
                                                                ByVal nSize As Long, _
                                                                ByVal lpFileName As String) As Long
Private Declare Function GetPrivateProfileStringW _
                            Lib "kernel32" (ByVal lpApplicationName As Long, _
                                            ByVal lpKeyName As Long, _
                                            ByVal lpDefault As Long, _
                                            ByVal lpReturnedString As Long, _
                                            ByVal nSize As Long, _
                                            ByVal lpFileName As Long) As Long

Private Declare Function WritePrivateProfileString _
                          Lib "kernel32.dll" _
                              Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, _
                                                                  ByVal lpKeyName As String, _
                                                                  ByVal lpString As Any, _
                                                                  ByVal lpFileName As String) As Long

'sub to load all keys from an ini section into a listbox.
Public Function CheckIniSectionExists(ByVal strSection As String, _
                                      ByVal strfullpath As String) As Boolean

Dim strBuffer                           As String
Dim nTemp                               As Long

    strBuffer = String$(5 * 1024, vbNullChar)
    nTemp = GetPrivateProfileSection(strSection, strBuffer, Len(strBuffer), strfullpath)

    If nTemp > 0 Then
        CheckIniSectionExists = True
    Else
        CheckIniSectionExists = False

    End If

End Function

' ��������� Boolean �������� ���������� ini-����� � ���������� ���������
Public Function GetIniValueBoolean(ByVal strIniPath As String, _
                                   ByVal strIniSection As String, _
                                   ByVal strIniValue As String, _
                                   ByVal lngValueDefault As Long) As Boolean

Dim LngValue                            As Long

    LngValue = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If LngValue = 9999 Then
        LngValue = lngValueDefault

    End If

    GetIniValueBoolean = CBool(LngValue)

End Function

' ��������� Long �������� ���������� ini-����� � ���������� ���������
Public Function GetIniValueLong(ByVal strIniPath As String, _
                                ByVal strIniSection As String, _
                                ByVal strIniValue As String, _
                                ByVal lngValueDefault As Long) As Long

Dim LngValue                            As Long

    LngValue = IniLongPrivate(strIniSection, strIniValue, strIniPath)

    If LngValue = 9999 Then
        LngValue = lngValueDefault

    End If

    GetIniValueLong = LngValue

End Function

' ��������� String �������� ���������� ini-����� � ���������� ���������
Public Function GetIniValueString(ByVal strIniPath As String, _
                                  ByVal strIniSection As String, _
                                  ByVal strIniValue As String, _
                                  ByVal strValueDefault As String) As String

Dim strValue                            As String

    strValue = IniStringPrivate(strIniSection, strIniValue, strIniPath)

    If strValue = "no_key" Then
        strValue = strValueDefault

    End If

    GetIniValueString = strValue

End Function

'! -----------------------------------------------------------
'!  �������     :  GetSectionMass
'!  ����������  :  SekName As String, IniFileName As String, Optional FirstValue As Boolean
'!                 SekName - ��� ������ (������� �� �����������)
'!                 FirstValue   - ���� ��������� ��������� ������ ������ ������ � ������
'!                 IniFileName - ��� ����� .ini (���� ���� � ����� �� ������,���� ������ � ����� Windows)
'!  �����. ����.:  ����� ����� ��� ��� ������ ���� ���� ������ � ������ �������. ����� ���������� ������ ����������=��������
'!  ��������    :  ������ ����� �������� � ���������� � ������ � ��������� ������ .INI
'! -----------------------------------------------------------
Public Function GetSectionMass(ByVal SekName As String, _
                               ByVal IniFileName As String, _
                               Optional ByVal FirstValue As Boolean)

Dim strBuffer                           As String * 32767
Dim strTemp                             As String
Dim intTemp                             As Long
Dim intTempSmallBuff                    As Long
Dim intSize                             As Long
Dim Index                               As Long
Dim arrSection()                        As String
Dim arrSectionTemp()                    As String
Dim Key                                 As String
Dim Value                               As String
Dim str                                 As String
Dim lpKeyValue()                        As String
Dim miRavnoPosition                     As Long

    On Error GoTo PROC_ERR

    Index = 1
    intSize = GetPrivateProfileSection(SekName, strBuffer, 32767, IniFileName)
    strTemp = Left$(strBuffer, intSize)

    If FirstValue Then
        ReDim arrSection(1, 2) As String
        arrSectionTemp = Split(strTemp, vbNullChar)
        intTempSmallBuff = InStrRev(strTemp, vbNullChar)

        If intTempSmallBuff > 0 Then
            str = arrSectionTemp(0)
            miRavnoPosition = InStr(str, "=")

            If miRavnoPosition > 0 Then
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
            ReDim arrSection(1, 2) As String
            arrSection(1, 1) = "Small Buffer"
            arrSection(1, 2) = "Small Buffer"
            IndexDevIDMass = 1
            GoTo IF_EXIT

        End If

    End If

    If LenB(strTemp) > 0 Then
        lpKeyValue = Split(strTemp, vbNullChar)
        ReDim arrSection(UBound(lpKeyValue), 2) As String

        Do Until LenB(strTemp) = 0
            intTempSmallBuff = InStrRev(strTemp, vbNullChar)

            If intTempSmallBuff > 0 Then
                intTemp = InStr(strTemp, vbNullChar)
                str = Left$(strTemp, intTemp)

                If InStr(str, "---") Then
                    Key = "������ ��� ID"
                    Value = "������ ��� ID"
                    GoTo Save_StrKey

                End If

                miRavnoPosition = InStr(str, "=")

                If miRavnoPosition > 0 Then
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
                ReDim arrSection(1, 2) As String
                arrSection(1, 1) = "Small Buffer"
                arrSection(1, 2) = "Small Buffer"
                IndexDevIDMass = 1
                GoTo IF_EXIT

            End If

        Loop
    Else
        ReDim arrSection(Index, 2) As String
        arrSection(Index, 1) = "No section"
        arrSection(Index, 2) = "No section"

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

'������� ��� ����� � �������� ������ � ��������� ����� .INI
'������ ������� � ���� ������!
'-------------------------------------------------
'SekName - ��� ������ (������� �� �����������)
'IniFileName - ��� ����� .ini (���� ���� � ����� �� ������,���� ������ � ����� Windows)
Public Function IniDelAllKeyPrivate(SekName As String, IniFileName As String)

Dim nTemp                               As Long

    nTemp = WritePrivateProfileString(SekName, vbNullString, vbNullString, IniFileName)

End Function

'! -----------------------------------------------------------
'!  �������     :  IniLongPrivate
'!  ����������  :  SekName As String, KeyName As String, IniFileName As String
'!                 SekName - ��� ������ (������� �� �����������)
'!                 KeyName - ��� ����� (������� �� �����������)
'!                 IniFileName - ��� ����� .ini (���� ���� � ����� �� ������,���� ������ � ����� Windows)
'!  �����. ����.:  As Long
'!                 9999    - ������������ �������� ��������, ���� ���� �� ������
'!  ��������    :  ������ ����� �������� �� ������ ����� .INI
'! -----------------------------------------------------------
'--------------------------------------------------
Public Function IniLongPrivate(ByVal SekName As String, _
                               ByVal KeyName As String, _
                               ByVal IniFileName As String) As Long
    IniLongPrivate = GetPrivateProfileInt(SekName, KeyName, 9999, IniFileName)

End Function

'! -----------------------------------------------------------
'!  �������     :  IniStringPrivate
'!  ����������  :  SekName As String, KeyName As String, IniFileName As String
'!                 SekName - ��� ������ (������� �� �����������)
'!                 KeyName - ��� ����� (������� �� �����������)
'!                 IniFileName - ��� ����� .ini (���� ���� � ����� �� ������,���� ������ � ����� Windows)
'!  �����. ����.:  As String
'!                 "no_key"    - ������������ �������� ��������, ���� ���� �� ������
'!  ��������    :  ������ ��������� �������� �� ������ ����� .INI
'! -----------------------------------------------------------
Public Function IniStringPrivate(ByVal SekName As String, _
                                 ByVal KeyName As String, _
                                 ByVal IniFileName As String) As String

'��������� �����(��� �������� �����)
Dim sTemp(4096) As Byte
Dim nTemp                               As Long

    '� �� ��������� ���������� �������� � ������ �����
    '����������� - �������� �� ����� ���� ������ 4096 ��������
    nTemp = GetPrivateProfileStringW(StrPtr(SekName), StrPtr(KeyName), StrPtr("no_key"), VarPtr(sTemp(0)), -1, StrPtr(IniFileName))

    IniStringPrivate = Left$(sTemp(), nTemp * 2)
    IniStringPrivate = TrimNull(IniStringPrivate)
    Erase sTemp

End Function

'! -----------------------------------------------------------
'!  �������     :  IniWriteStrPrivate
'!  ����������  :  SekName As String, KeyName As String, Param As String, IniFileName As String
'!                 SekName - ��� ������ (������� �� �����������)
'!                 KeyName - ��� ����� (������� �� �����������)
'!                 Param   - ��������,������������ � ���� (�� ������ ������)
'!                 IniFileName - ��� ����� .ini (���� ���� � ����� �� ������,���� ������ � ����� Windows)
'!  �����. ����.:  As Long
'!  ��������    :  ���������� ��������� �������� � ����� ���� .INI
'! -----------------------------------------------------------
Public Sub IniWriteStrPrivate(ByVal SekName As String, _
                              ByVal KeyName As String, _
                              ByVal Param As String, _
                              ByVal IniFileName As String)
    WritePrivateProfileString SekName, KeyName, Param, IniFileName

End Sub

'sub to load all keys from an ini section into a listbox.
Public Function LoadIniSectionKeys(ByVal strSection As String, _
                                   ByVal strfullpath As String, _
                                   Optional ByVal mbKeys As Boolean = True) As String()

Dim KeyAndVal()                         As String
Dim Key_Val()                           As String
Dim strBuffer                           As String
Dim intx                                As Long
Dim Z()                                 As String
Dim n                                   As Long

    n = -1
    strBuffer = String$(5 * 1024, vbNullChar)
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
            ' ������ �����
            Z(n) = Key_Val(0)
        Else

            ' ������ �������� ������
            If UBound(Key_Val) = 1 Then
                Z(n) = Key_Val(1)

            End If

        End If

    Next
    Erase KeyAndVal
    Erase Key_Val

    If n = -1 Then
        ReDim Z(0) As String

    End If

    LoadIniSectionKeys = Z

End Function

'! -----------------------------------------------------------
'!  �������     :  NormFile
'!  ����������  :  sFileName As String
'!  ��������    :  ���������� ��� ����� � "�����������" ���
'! -----------------------------------------------------------
Public Sub NormIniFile(ByVal sFileName As String)

Dim nf                                  As Long
Dim ub                                  As Long
Dim sBuffer                             As String
Dim slArray()                           As String
Dim sOutArray()                         As String

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

'# use to read/write ini/inf file #
Public Function ReadFromINI(ByVal strSection As String, _
                            ByVal strKey As String, _
                            ByVal strfullpath As String, _
                            Optional ByVal strDefault As String = vbNullString) As String

Dim strBuffer                           As String

    strBuffer = String$(750, vbNullChar)
    ReadFromINI = Left$(strBuffer, GetPrivateProfileString(strSection, ByVal LCase$(strKey), strDefault, strBuffer, Len(strBuffer), strfullpath))

End Function

' �������� �� �� ��� ������ ������
Public Function IniSectionIsEmpty(strSection As String, strIni As String) As Boolean

Dim sTemp                               As String * 2048
Dim nTemp                               As Long

    '� �� ��������� ���������� �������� � ������ �����
    nTemp = GetPrivateProfileSection(strSection, sTemp, 2048, strIni)
    IniSectionIsEmpty = nTemp = 0
    
End Function

' ��������� ������ ������ ������ �� ����������� ����� ������ ������
Public Function GetIniEmptySectionFromList(strSectionList As String, _
                                           strIniPath As String) As String

Dim strTmp                              As String
Dim strSectionList_x()                  As String
Dim i_i                                 As Long
Dim strManufSection                     As String

    strSectionList_x = Split(strSectionList, "|")

    For i_i = 0 To UBound(strSectionList_x)
        strManufSection = strSectionList_x(i_i)

        If IniSectionIsEmpty(strManufSection, strIniPath) Then
            If LenB(strTmp) > 0 Then
                strTmp = strTmp & "," & strManufSection
            Else
                strTmp = strManufSection

            End If

        End If

    Next

    If LenB(strTmp) > 0 Then
        GetIniEmptySectionFromList = strTmp
    Else
        GetIniEmptySectionFromList = "-"

    End If

End Function
