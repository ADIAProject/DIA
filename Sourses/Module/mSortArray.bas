Attribute VB_Name = "mSortArray"
Option Explicit

Public Enum eCompareResult
    crLess = -1&
    crEqual = 0&
    crGreater = 1&
End Enum

'VB lacks any support for procedure calling using an address, but the good ol'
'CallWindowProc will do just fine!
Private Declare Function CompareValues Lib "user32.dll" Alias "CallWindowProcW" (ByVal CompareFunc As Long, ByVal First As Long, ByVal Second As Long, ByVal unused1 As Long, ByVal unused2 As Long) As eCompareResult

'General purpose CopyMemory, but optimized for our purposes using byval longs
'since we are working with pointers
Private Declare Sub CopyMemoryByVal Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Dst As Long, ByVal Src As Long, ByVal ByteCount As Long)

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
    ReDim lyBuffer(0 To piBytesPerElement - 1) As Byte

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
    'CompareString = StrComp(First, Second, vbTextCompare)
    CompareString = StrComp(First, Second, vbBinaryCompare)
End Function

'Public Function CompareStringApi(First As String, Second As String, unused1 As Long, unused2 As Long) As eCompareResult
'    'CompareStringApi = StrCmp(StrPtr(First), StrPtr(Second))
'    CompareStringApi = lstrcmp(StrPtr(First), StrPtr(Second))
'End Function
