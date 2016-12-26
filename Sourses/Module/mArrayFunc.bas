Attribute VB_Name = "mArrayFunc"
Option Explicit

' Not add to project (if not DBS) - option for compile
#Const mbIDE_DBSProject = False
'*********************************
'Api declare for sorting array function ShellSortAny
' return value for function CompareValues
Public Enum eCompareResult
    crLess = -1&
    crEqual = 0&
    crGreater = 1&
End Enum

'VB lacks any support for procedure calling using an address, but the good ol - CallWindowProc will do just fine!
Private Declare Function CompareValues Lib "user32.dll" Alias "CallWindowProcW" (ByVal CompareFunc As Long, ByVal First As Long, ByVal Second As Long, ByVal unused1 As Long, ByVal unused2 As Long) As eCompareResult
'General purpose CopyMemory, but optimized for our purposes using byval longs - since we are working with pointers
Private Declare Sub CopyMemoryByVal Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Dst As Long, ByVal Src As Long, ByVal ByteCount As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Any, ByVal Length As Long)

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BinarySearch
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   strArray() (String)
'                              strSearch (String)
'!--------------------------------------------------------------------------------
Public Function BinarySearch(strArray() As String, ByVal strSearch As String) As Long

    Dim lngFirst        As Long
    Dim lngLast         As Long
    Dim lngMiddle       As Long
    Dim bolInverseOrder As Boolean
                
    lngFirst = LBound(strArray)
    lngLast = UBound(strArray)
    bolInverseOrder = (strArray(lngFirst) > strArray(lngLast))
    BinarySearch = lngFirst - 1

    Do
        lngMiddle = (lngFirst + lngLast) \ 2

        If StrComp(strArray(lngMiddle), strSearch) = 0 Then
            BinarySearch = lngMiddle

            Exit Do

        ElseIf ((StrComp(strArray(lngMiddle), strSearch) < 0) Xor bolInverseOrder) Then
            lngFirst = lngMiddle + 1
        Else
            lngLast = lngMiddle - 1
        End If

    Loop Until lngFirst > lngLast

End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub QuickSortMDArray
'! Description (Описание)  :   [Written by Ellis Dee
'                               Sort a 2-dimensional array on either dimension
'                               Sample usage to sort on column 4
'                               Dim MyArray(1 to 1000, 1 to 5) As Long
'                               QuickSort MyArray, 2, 4
'                               Dim MyArray(1 to 5, 1 to 1000) As Long
'                               QuickSort MyArray, 1, 4]
'! Parameters  (Переменные):   pArray (Variant)
'                              pbytDimension (Byte)
'                              plngColumn (Long)
'                              plngLeft (Long)
'                              plngRight (Long)
'!--------------------------------------------------------------------------------
Public Sub QuickSortMDArray(pArray As Variant, pbytDimension As Byte, plngColumn As Long, Optional ByVal plngLeft As Long, Optional ByVal plngRight As Long)

    Dim i            As Long
    Dim lngFirst     As Long
    Dim lngLast      As Long
    Dim vFirst       As Variant
    Dim vMid         As Variant
    Dim vLast        As Variant
    Dim lDim()       As Long
    Dim bytCol       As Byte
    Dim bytRow       As Byte

    ReDim lDim(1 To 2)
    bytRow = -pbytDimension + 3
    bytCol = pbytDimension

    If plngRight = 0 Then
        plngLeft = LBound(pArray, bytRow)
        plngRight = UBound(pArray, bytRow)
    End If

    lngFirst = plngLeft
    lngLast = plngRight
    lDim(bytRow) = (plngLeft + plngRight) \ 2
    lDim(bytCol) = plngColumn
    vMid = pArray(lDim(1), lDim(2))

    Do
        lDim(bytRow) = lngFirst
        lDim(bytCol) = plngColumn

        Do While pArray(lDim(1), lDim(2)) < vMid And lngFirst < plngRight
            lngFirst = lngFirst + 1
            lDim(bytRow) = lngFirst
        Loop

        lDim(bytRow) = lngLast

        Do While vMid < pArray(lDim(1), lDim(2)) And lngLast > plngLeft
            lngLast = lngLast - 1
            lDim(bytRow) = lngLast
        Loop

        If lngFirst <= lngLast Then

            For i = LBound(pArray, bytCol) To UBound(pArray, bytCol)
                lDim(bytCol) = i
                lDim(bytRow) = lngFirst
                vFirst = pArray(lDim(1), lDim(2))
                lDim(bytRow) = lngLast
                vLast = pArray(lDim(1), lDim(2))
                pArray(lDim(1), lDim(2)) = vFirst
                lDim(bytRow) = lngFirst
                pArray(lDim(1), lDim(2)) = vLast
            Next

            lngFirst = lngFirst + 1
            lngLast = lngLast - 1
        End If

    Loop Until lngFirst > lngLast

    If plngLeft < lngLast Then QuickSortMDArray pArray, pbytDimension, plngColumn, plngLeft, lngLast
    If lngFirst < plngRight Then QuickSortMDArray pArray, pbytDimension, plngColumn, lngFirst, plngRight
End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SaveAnyStringArray2File
'! Description (Описание)  :   [My Function for Save Any String Array Any Dimension to File]
'! Parameters  (Переменные):   StrPathFile (String)
'                              MyArray() (String)
'                              strDelimiter (String = vbTab)
'!--------------------------------------------------------------------------------
Public Function SaveAnyStringArray2File(ByVal strPathFile As String, MyArray() As String, Optional ByVal strDelimiter As String = vbTab) As Boolean

    Dim hiIndex       As Long
    Dim loIndex       As Long
    Dim strResultAll  As String
    Dim strLine       As String
    Dim i             As Long
    Dim iii           As Long

    If mbDebugStandart Then DebugMode vbTab & "SaveAnyStringArray2File-Start"
    hiIndex = UBound(MyArray, 2)
    loIndex = UBound(MyArray, 1)

    For i = 0 To hiIndex
        strLine = vbNullString

        For iii = 0 To loIndex
            AppendStr strLine, MyArray(iii, i), strDelimiter
        Next

        AppendStr strResultAll, strLine, vbNewLine
    Next

    If LenB(strResultAll) Then
        '---------------Выводим итог в файл-----
        FileWriteData strPathFile, strResultAll
        
        If mbDebugStandart Then DebugMode vbTab & "ListLocalHwid:" & vbNewLine & "**************************************************************************" & vbNewLine & strResultAll & vbNewLine & _
                                    "**************************************************************************"
        SaveAnyStringArray2File = True
    End If

    If mbDebugStandart Then DebugMode vbTab & "SaveAnyStringArray2File-End"
End Function

#If Not mbIDE_DBSProject Then
'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function SaveHwidsArray2File
'! Description (Описание)  :   [My Function for Save Any String Array Any Dimension to File]
'! Parameters  (Переменные):   StrPathFile (String)
'                              MyArray() (arrHwidsStruct)
'!--------------------------------------------------------------------------------
Public Function SaveHwidsArray2File(ByVal strPathFile As String, MyArray() As arrHwidsStruct) As Boolean

    Dim strResultAll  As String
    Dim strLine       As String
    Dim i             As Long

    If mbDebugDetail Then DebugMode "SaveHwidsArray2File-Start: ToFile: " & strPathFile

    For i = 0 To UBound(MyArray)
        strLine = vbNullString

        With arrHwidsLocal(i)
            AppendStr strLine, .HWID, vbTab
            AppendStr strLine, .DevName, vbTab
            AppendStr strLine, .Status, vbTab
            AppendStr strLine, .VerLocal, vbTab
            AppendStr strLine, .HWIDOrig, vbTab
            AppendStr strLine, .Provider, vbTab
            AppendStr strLine, .HWIDCompat, vbTab
            AppendStr strLine, .Description, vbTab
            AppendStr strLine, .PriznakSravnenia, vbTab
            AppendStr strLine, .InfSection, vbTab
            AppendStr strLine, .HWIDCutting, vbTab
            AppendStr strLine, .HWIDMatches, vbTab
            AppendStr strLine, .InfName, vbTab
            AppendStr strLine, .DPsList, vbTab
            AppendStr strLine, .DRVScore, vbTab
        End With

        AppendStr strResultAll, strLine, vbNewLine
    Next

    If LenB(strResultAll) Then
        '---------------Выводим итог в файл-----
        FileWriteData strPathFile, strResultAll
        If mbDebugStandart Then DebugMode "SaveHwidsArray2File-ListLocalHwid:" & vbNewLine & "**************************************************************************" & vbNewLine & strResultAll & vbNewLine & _
                                    "**************************************************************************"
        SaveHwidsArray2File = True
    Else
        If mbDebugDetail Then DebugMode "SaveHwidsArray2File-False: NO DATA"
    End If

End Function
#End If

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Sub ShellSortAny
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   piArrPtr (Long)
'                              piElementCount (Long)
'                              piBytesPerElement (Integer)
'                              piCompareProcAddr (Long)
'!--------------------------------------------------------------------------------
Public Sub ShellSortAny(ByVal piArrPtr As Long, ByVal piElementCount As Long, ByVal piBytesPerElement As Integer, ByVal piCompareProcAddr As Long)

    Dim liDist         As Long
    Dim liDistBytes    As Long
    Dim liValuePtr     As Long
    Dim liBufferPtr    As Long
    Dim liPtr          As Long
    Dim liPtr2         As Long
    Dim liLastValuePtr As Long
    Dim lyBuffer()     As Byte

    'Dim our buffer for enough bytes to hold one element
    ReDim lyBuffer(0 To piBytesPerElement - 1)

    'Get the pointer to the first element
    liBufferPtr = VarPtr(lyBuffer(0))

    'Find the initial value for liDist
    Do
        liDist = liDist + liDist + liDist + 1&
    Loop Until liDist > piElementCount

    'get the last valid pointer
    liLastValuePtr = piArrPtr + piElementCount * piBytesPerElement - piBytesPerElement

    Do
        'Reduce liDist by two thirds
        liDist = liDist \ 3
        'Get the number of bytes
        liDistBytes = liDist * piBytesPerElement

        'Loop through each pointer in our current section
        For liValuePtr = piArrPtr + liDistBytes To liLastValuePtr Step piBytesPerElement

            'Compare the current value with the immediately previous value, to see if they're in the correct order
            If CompareValues(piCompareProcAddr, liValuePtr - liDistBytes, liValuePtr, 0&, 0&) = crGreater Then
                'If the wrong order, then copy the current value to the buffer
                CopyMemoryByVal liBufferPtr, liValuePtr, piBytesPerElement
                'Set our temp pointer to the current value
                liPtr = liValuePtr
                'Set the other temp pointer to the beginning of the section
                liPtr2 = liPtr - liDistBytes

                Do
                    'Copy the first value to the current value
                    CopyMemoryByVal liPtr, liPtr2, piBytesPerElement
                    'Adjust the pointers
                    liPtr = liPtr2
                    liPtr2 = liPtr2 - liDistBytes

                    'Make sure we're in-bounds
                    If liPtr2 < piArrPtr Then Exit Do
                    'Keep going as long as we're in order
                Loop While CompareValues(piCompareProcAddr, liPtr2, liBufferPtr, 0&, 0&) = crGreater

                'put the buffered value back in
                CopyMemoryByVal liPtr, liBufferPtr, piBytesPerElement
            End If

        Next

    Loop Until liDist = 1&

End Sub

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function CompareString
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   First (String)
'                              Second (String)
'                              unused1 (Long)
'                              unused2 (Long)
'!--------------------------------------------------------------------------------
Public Function CompareString(First As String, Second As String, unused1 As Long, unused2 As Long) As eCompareResult
    CompareString = StrComp(First, Second, vbBinaryCompare)
End Function

'!--------------------------------------------------------------------------------
'! Procedure   (Функция)   :   Function BinarySearchLong
'! Description (Описание)  :   [type_description_here]
'! Parameters  (Переменные):   lngArray() (Long)
'                              lngHash (Long)
'!--------------------------------------------------------------------------------
Public Function BinarySearchLong(lngArray() As Long, ByVal lngHash As Long) As Long

    Dim lngFirst        As Long
    Dim lngLast         As Long
    Dim lngMiddle       As Long
    Dim bolInverseOrder As Boolean
                
    lngFirst = LBound(lngArray)
    lngLast = UBound(lngArray)
    bolInverseOrder = (lngArray(lngFirst) > lngArray(lngLast))
    BinarySearchLong = lngFirst - 1

    Do
        lngMiddle = (lngFirst + lngLast) \ 2

        If lngArray(lngMiddle) = lngHash Then
            BinarySearchLong = lngMiddle

            Exit Do

        ElseIf ((lngArray(lngMiddle) < lngHash) Xor bolInverseOrder) Then
            lngFirst = lngMiddle + 1
        Else
            lngLast = lngMiddle - 1
        End If

    Loop Until lngFirst > lngLast

End Function

Public Sub CopyStringArray(ByRef Dest() As String, Src() As String, Optional StartIndex As Long = -1)
   Dim tmpArr() As String, VarSize&, NewUbound&
  
   ' Создаем полную копию копируемого массива
   tmpArr = Src
   
   ' Определяем число добавляемых элементов массива
   VarSize = (UBound(Src) - LBound(Src) + 1)
  
   If StartIndex = -1 Then StartIndex = LBound(Dest)
  
   ' Возможно, копируемый массив не влезет в массив назначения....
   If UBound(Dest) < StartIndex + VarSize Then
      '  ... и тогда увеличим его
      NewUbound = StartIndex + VarSize
      ReDim Preserve Dest(LBound(Dest) To NewUbound)
   End If
   
   ' Копируем указатели временного массива на место указателей основного
   CopyMemory ByVal VarPtr(Dest(StartIndex)), ByVal VarPtr(tmpArr(LBound(tmpArr))), VarSize * 4
   
   ' Обнуляем указатели временного массива (рекомендуется, чтобы VB не удалил данные)
   ZeroMemory ByVal VarPtr(tmpArr(LBound(tmpArr))), VarSize * 4
End Sub
