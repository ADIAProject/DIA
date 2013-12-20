Attribute VB_Name = "mStrings"
Option Explicit

Public Enum SplitCompareMethod
    [Split BinaryCompare] = VbCompareMethod.vbBinaryCompare         ' InStrB
    [Split CharacterCompare] = VbCompareMethod.vbDatabaseCompare    ' InStr(BinaryCompare)
End Enum

Private m_A()       As Long
Private m_AP        As Long
Private m_H(0 To 6) As Long
Private m_HP        As Long
Private Declare Sub GetMem4 Lib "msvbvm60" (ByVal Ptr As Long, Value As Long)
Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal Ptr As Long, ByVal Value As Long)
Private Declare Function InitStringArray Lib "oleaut32.dll" Alias "SafeArrayCreate" (Optional ByVal VarType As VbVarType = vbString, Optional ByVal Dims As Integer = 1, Optional saBound As Currency) As Long
Private Declare Function SysAllocStringByteLen Lib "oleaut32.dll" (ByVal Ptr As Long, ByVal Length As Long) As Long
Private Declare Function SysAllocStringLen Lib "oleaut32.dll" (ByVal Ptr As Long, ByVal Length As Long) As Long

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property API
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Module (String)
'                              Procedure (String)
'!--------------------------------------------------------------------------------
Private Property Get API(Module As String, Procedure As String) As Long

    Dim Handle    As Long
    Dim lngStrPtr As Long

    lngStrPtr = StrPtr(Module)
    Handle = GetModuleHandle(lngStrPtr)

    If Handle = 0 Then
        Handle = LoadLibrary(lngStrPtr)
    End If

    If Handle Then
        API = GetProcAddress(Handle, Procedure)
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Procedure
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   AddressOfDest (Long)
'!--------------------------------------------------------------------------------
Private Property Get Procedure(ByVal AddressOfDest As Long) As Long

    ' get correct pointer to procedure in IDE
    If Not InIDE() Then
        Procedure = AddressOfDest
    Else
        GetMem4 AddressOfDest + &H16&, Procedure
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Property Procedure
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   AddressOfDest (Long)
'                              AddressOfSrc (Long)
'!--------------------------------------------------------------------------------
Private Property Let Procedure(ByVal AddressOfDest As Long, ByVal AddressOfSrc As Long)

    Dim JMP As Currency
    Dim PID As Long

    ' get process handle
    PID = OpenProcess(&H1F0FFF, 0&, GetCurrentProcessId)

    If PID Then

        ' get correct pointer to procedure in IDE
        If InIDE() Then
            GetMem4 AddressOfDest + &H16&, AddressOfDest
        End If

        Debug.Assert App.hInstance
        ' ASM JMP (0xE9) followed by bytes to jump in memory
        JMP = (&HE9& * 0.0001) + (AddressOfSrc - AddressOfDest - 5@) * 0.0256
        ' write the JMP over the destination procedure
        WriteProcessMemory PID, ByVal AddressOfDest, JMP, 5
        ' close process handle
        CloseHandle PID
    End If

End Property

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function Split
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Expression (String)
'                              Delimiter (String = " ")
'                              Limit (Long = -1)
'                              Compare (SplitCompareMethod) As String()
'!--------------------------------------------------------------------------------
Public Function Split(Expression As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As SplitCompareMethod) As String()
    Procedure(AddressOf mStrings.Split) = Procedure(AddressOf mStrings.z_Split)
    Split = mStrings.Split(Expression, Delimiter, Limit, Compare)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function z_Split
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Expression (String)
'                              Delimiter (String = " ")
'                              Limit (Long = -1)
'                              Compare (SplitCompareMethod)
'!--------------------------------------------------------------------------------
Public Function z_Split(Expression As String, Optional Delimiter As String = " ", Optional ByVal Limit As Long = -1, Optional ByVal Compare As SplitCompareMethod) As Long

    Static R() As Long
    Static RP  As Long

    Dim P()    As Long
    Dim C      As Long
    Dim i      As Long
    Dim j      As Long
    Dim K      As Long
    Dim LD     As Long
    Dim LD2    As Long
    Dim LE     As Long
    Dim PL     As Long
    Dim PS     As Long

    ' get pointer
    PS = StrPtr(Expression)
    ' length information
    LE = LenB(Expression)
    LD = LenB(Delimiter)

    ' unlimited or limited?
    If Limit = -1 Then
        If LD Then
            Limit = LE \ LD + 1
        End If
    End If

    ' validate lengths and limit
    If LE > 0 Then
        If LD > 0 Then
            If Limit >= 0 Then

                ' pointer to R array
                If RP = 0 Then
                    RP = ArrPtr(R)
                End If

                ' generic safe array hack
                If m_AP = 0 Then
                    ' array variable pointer
                    m_AP = ArrPtr(m_A)
                    ' create a safe array header
                    m_H(0) = vbLong
                    m_H(1) = &H800001
                    m_H(2) = 4
                    m_H(5) = &H7FFFFFFF
                    ' header pointer
                    m_HP = VarPtr(m_H(1))
                End If

                ' set pointer to array
                PutLong m_AP, m_HP

                ' find the first item
                If Limit > 1 Then
                    If Compare = [Split BinaryCompare] Then

                        Do
                            i = InStrB(i + 1, Expression, Delimiter)
                        Loop Until (i And 1) = 1 Or (i = 0)

                    Else
                        'NOT COMPARE...
                        i = InStr(Expression, Delimiter)
                    End If
                End If

                ' did we find an item?
                If i Then

                    ReDim P(3) As Long

                    ' space for knowing the positions
                    PL = (Limit \ 96)

                    If PL > 8191 Then
                        PL = 8191
                    End If

                    If PL > UBound(P) Then

                        ReDim Preserve P(0 To PL)

                    End If

                    ' InStrB?
                    If Compare = [Split BinaryCompare] Then

                        For C = 0 To Limit

                            ' make sure will always have enough items
                            If C >= PL Then
                                PL = PL + C

                                ReDim Preserve P(PL)

                            End If

                            ' exit if nothing found
                            If i = 0 Then

                                Exit For

                            End If

                            ' remember position
                            P(C) = i - 1
                            ' find next
                            i = i + LD - 1

                            Do
                                i = InStrB(i + 1, Expression, Delimiter)
                            Loop Until (i And 1) = 1 Or (i = 0)

                        Next C

                    Else
                        ' InStr'NOT COMPARE...
                        LD2 = LD \ 2

                        For C = 0 To Limit

                            ' make sure will always have enough items
                            If C >= PL Then
                                PL = PL + C

                                ReDim Preserve P(PL)

                            End If

                            ' exit if nothing found
                            If i = 0 Then

                                Exit For

                            End If

                            ' remember position
                            P(C) = (i - 1) * 2
                            ' find next
                            i = InStr(i + LD2, Expression, Delimiter)
                        Next C

                    End If

                    P(C) = LE
                    ' make space for the new items
                    z_Split = InitStringArray(, , (C + 1) * 0.0001)
                    ' set pointer
                    m_H(4) = RP
                    m_A(0) = z_Split
                    ' keep it simple, stupid!
                    i = 0

                    For C = 0 To C
                        K = P(C)
                        j = K - i

                        If j Then
                            R(C) = SysAllocStringByteLen(PS + i, j)
                        End If

                        i = K + LD
                    Next C

                Else
                    'I = FALSE/0
                    ' one item
                    z_Split = InitStringArray(, , 0.0001)
                    ' set pointer
                    m_H(4) = RP
                    m_A(0) = z_Split
                    R(0) = SysAllocStringByteLen(PS, LE)
                End If

                ' clean up z_Split reference
                m_A(0) = 0
                ' clean up safe array reference
                m_H(4) = m_AP
                m_A(0) = 0
            Else
                z_Split = InitStringArray
            End If

        Else
            z_Split = InitStringArray
        End If

    Else
        z_Split = InitStringArray
    End If

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub PutLong
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   Ptr (Long)
'                              Value (Long)
'!--------------------------------------------------------------------------------
Public Sub PutLong(ByVal Ptr As Long, ByVal Value As Long)
    Procedure(AddressOf mStrings.PutLong) = API("msvbvm60", "PutMem4")
    PutLong Ptr, Value
End Sub
